from __future__ import annotations

import os
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from docx import Document  # noqa: E402
from PySide6.QtCore import Qt  # noqa: E402
from PySide6.QtWidgets import QApplication, QMessageBox  # noqa: E402

from excel_privacy_cleaner.qt_app import ExcelPrivacyCleanerWindow  # noqa: E402
from excel_privacy_cleaner.word_processor import WordPrivacyProcessor  # noqa: E402


def assert_true(condition: bool, message: str) -> None:
    if not condition:
        raise AssertionError(message)


def create_gui_fixture(path: Path) -> None:
    document = Document()
    document.add_paragraph("連絡先電話は090-1111-2222です。")
    document.add_paragraph("佐藤 花子")
    document.save(path)


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


def test_word_scan_and_convert_via_gui(monkeypatch) -> None:
    app = QApplication.instance() or QApplication([])
    calls = _silence_message_boxes(monkeypatch)

    with tempfile.TemporaryDirectory(prefix="qt_app_word_") as tmpdir:
        tmp = Path(tmpdir)
        source = tmp / "fixture.docx"
        create_gui_fixture(source)

        window = ExcelPrivacyCleanerWindow()
        try:
            window.set_source(source)
            assert_true(isinstance(window.processor, WordPrivacyProcessor), "Selecting a .docx should create a WordPrivacyProcessor")
            assert_true(window.is_word_source(), "is_word_source() should be True after selecting a .docx")

            window.scan_file()
            assert_true(len(window.word_decisions) == 2, "Fixture should yield exactly two candidates")
            assert_true(window.table.rowCount() == 2, "Table should show one row per candidate")
            assert_true(len(window.findings) == 2, "Adapter should produce one Finding per WordReplacementDecision")

            phone_row = next(row for row, finding in enumerate(window.findings) if finding.entity_type == "電話番号")
            name_row = next(row for row, finding in enumerate(window.findings) if finding.entity_type == "氏名")
            assert_true(window.word_decisions[phone_row].enabled, "High-confidence phone candidate should default to enabled")
            assert_true(not window.word_decisions[name_row].enabled, "Low-confidence bare name candidate should default to disabled (review-required)")

            # Column 3 (entity_type) must not be user-editable for Word rows,
            # since WordCandidate.category is immutable.
            entity_item = window.table.item(name_row, 3)
            assert_true(entity_item is not None and not (entity_item.flags() & Qt.ItemIsEditable), "Word rows must not allow editing the category column")

            # Simulate a reviewer approving the review-required candidate via the checkbox.
            checkbox_item = window.table.item(name_row, 0)
            assert_true(checkbox_item is not None, "Checkbox item should exist")
            checkbox_item.setCheckState(Qt.Checked)

            window.convert_file()

            assert_true(any(name == "information" for name, _ in calls), "A success message box should have been shown")
            assert_true(window.history.count() == 1, "Conversion should append one history entry")
        finally:
            window.processor.cleanup()
            window.close()
            app.processEvents()
