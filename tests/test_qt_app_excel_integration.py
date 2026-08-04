from __future__ import annotations

import os
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from openpyxl import Workbook  # noqa: E402
from PySide6.QtCore import Qt  # noqa: E402
from PySide6.QtWidgets import QApplication, QMessageBox  # noqa: E402

from excel_privacy_cleaner.excel_processor import EXCEL_NLP_DETECTION_KIND, ExcelPrivacyProcessor  # noqa: E402
from excel_privacy_cleaner.qt_app import ExcelPrivacyCleanerWindow  # noqa: E402


def assert_true(condition: bool, message: str) -> None:
    if not condition:
        raise AssertionError(message)


def create_ai_candidate_fixture(path: Path) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Sheet1"
    sheet["A1"] = "備考"
    sheet["A2"] = "場所　アルプススチール様1階応接室"
    # A second, unambiguous finding so the sheet still has something left to
    # convert after the AI candidate above gets excluded.
    sheet["B1"] = "連絡先"
    sheet["B2"] = "連絡先電話は090-1111-2222です。"
    workbook.save(path)


def _silence_message_boxes(monkeypatch) -> list[tuple[str, str]]:
    calls: list[tuple[str, str]] = []
    for method in ("information", "warning", "critical", "question"):
        def make_stub(name: str):
            def stub(*_args, **_kwargs):
                calls.append((name, ""))
                return QMessageBox.Yes if name == "question" else None
            return stub
        monkeypatch.setattr(QMessageBox, method, staticmethod(make_stub(method)))
    return calls


def test_excel_ai_candidate_exclude_checkbox_resolves_review_required_block(monkeypatch) -> None:
    app = QApplication.instance() or QApplication([])
    calls = _silence_message_boxes(monkeypatch)

    with tempfile.TemporaryDirectory(prefix="qt_app_excel_") as tmpdir:
        tmp = Path(tmpdir)
        source = tmp / "fixture.xlsx"
        create_ai_candidate_fixture(source)

        window = ExcelPrivacyCleanerWindow()
        try:
            window.set_source(source)
            assert_true(isinstance(window.processor, ExcelPrivacyProcessor), "Selecting a .xlsx should create an ExcelPrivacyProcessor")

            window.scan_file()
            ai_row = next(row for row, finding in enumerate(window.findings) if finding.detection_kind == EXCEL_NLP_DETECTION_KIND)
            assert_true(not window.findings[ai_row].enabled, "AI-sourced candidate should default to disabled")

            # Leaving it unresolved must block conversion.
            window.convert_file()
            assert_true(any(name == "critical" for name, _ in calls), "Leaving an AI candidate unreviewed must block conversion")
            calls.clear()

            # Check "変換しない" for that row -- this is the fix for the block.
            window.table.item(ai_row, 1).setCheckState(Qt.Checked)
            assert_true(window.findings[ai_row].excluded, "Finding should now be marked excluded")
            assert_true(not window.findings[ai_row].enabled, "Excluded finding must stay disabled")
            assert_true(
                window.table.item(ai_row, 0).checkState() == Qt.Unchecked,
                "Checking 変換しない must clear 変換する for the same row",
            )
            assert_true("除外" in window.findings[ai_row].reason, "Reason text should mention the exclusion")

            window.convert_file()
            assert_true(any(name == "information" for name, _ in calls), "Conversion should now succeed")
            assert_true(window.history.count() == 1, "Conversion should append one history entry")
        finally:
            window.processor.cleanup()
            window.close()
            app.processEvents()
