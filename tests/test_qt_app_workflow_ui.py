from __future__ import annotations

import os
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

import fitz  # noqa: E402
from openpyxl import Workbook  # noqa: E402
from PySide6.QtWidgets import QApplication  # noqa: E402

from excel_privacy_cleaner.qt_app import ExcelPrivacyCleanerWindow  # noqa: E402


def assert_true(condition: bool, message: str) -> None:
    if not condition:
        raise AssertionError(message)


def create_simple_xlsx(path: Path) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet["A1"] = "連絡先"
    sheet["A2"] = "連絡先電話は090-1111-2222です。"
    workbook.save(path)


def create_simple_pdf(path: Path) -> None:
    doc = fitz.open()
    page = doc.new_page()
    page.insert_text((72, 72), "Filler text so the page counts as a text-layer page.", fontsize=12)
    doc.save(path)
    doc.close()


def test_scan_and_convert_buttons_are_staged() -> None:
    # Reproduces "ファーストアクションが分かりにくい": the scan/convert
    # buttons used to always be clickable regardless of state. They should
    # now only enable once the previous step in the workflow is done.
    app = QApplication.instance() or QApplication([])
    window = ExcelPrivacyCleanerWindow()
    try:
        window.show()
        app.processEvents()

        assert_true(not window.scan_button.isEnabled(), "検査開始 should start disabled with no file selected")
        assert_true(not window.convert_button.isEnabled(), "出力 should start disabled with no file selected")

        with tempfile.TemporaryDirectory(prefix="workflow_ui_") as tmpdir:
            source = Path(tmpdir) / "fixture.xlsx"
            create_simple_xlsx(source)

            window.set_source(source)
            assert_true(window.scan_button.isEnabled(), "検査開始 should enable once a file is selected")
            assert_true(not window.convert_button.isEnabled(), "出力 should stay disabled until scanned")

            window.scan_file()
            assert_true(window.convert_button.isEnabled(), "出力 should enable once the scan produced findings")
    finally:
        window.processor.cleanup()
        window.close()
        app.processEvents()


def test_summary_counts_reflect_findings() -> None:
    app = QApplication.instance() or QApplication([])
    window = ExcelPrivacyCleanerWindow()
    try:
        assert_true(window.summary_total_label.text() == "検出 0件", "Summary should start at zero")

        with tempfile.TemporaryDirectory(prefix="workflow_ui_summary_") as tmpdir:
            source = Path(tmpdir) / "fixture.xlsx"
            create_simple_xlsx(source)
            window.set_source(source)
            window.scan_file()

            assert_true(len(window.findings) > 0, "Fixture should yield at least one finding")
            assert_true(
                window.summary_total_label.text() == f"検出 {len(window.findings)}件",
                "Summary total should match the number of findings",
            )
            enabled_count = sum(1 for finding in window.findings if finding.enabled)
            assert_true(
                window.summary_target_label.text() == f"変換対象 {enabled_count}件",
                "Summary target count should match enabled findings",
            )

            # Toggling a checkbox should update the live count too.
            checkbox = window._row_checkbox(0, 0)
            was_checked = checkbox.isChecked()
            checkbox.setChecked(not was_checked)
            expected = enabled_count + (1 if not was_checked else -1)
            assert_true(
                window.summary_target_label.text() == f"変換対象 {expected}件",
                "Summary target count should update live when a checkbox is toggled",
            )
    finally:
        window.processor.cleanup()
        window.close()
        app.processEvents()


def test_settings_panel_starts_collapsed_and_toggles() -> None:
    app = QApplication.instance() or QApplication([])
    window = ExcelPrivacyCleanerWindow()
    try:
        window.show()
        app.processEvents()

        assert_true(not window.settings_panel.isVisible(), "設定 panel should start collapsed")
        assert_true("分析継続用" in window.settings_toggle_button.text(), "Toggle button should summarize the current settings")

        window._toggle_settings_panel()
        app.processEvents()
        assert_true(window.settings_panel.isVisible(), "設定 panel should expand after toggling")

        window._toggle_settings_panel()
        app.processEvents()
        assert_true(not window.settings_panel.isVisible(), "設定 panel should collapse again after toggling twice")
    finally:
        window.processor.cleanup()
        window.close()
        app.processEvents()


def test_pdf_review_button_is_hidden_for_non_pdf_and_shown_for_pdf() -> None:
    # Reproduces "PDF候補確認がわかりにくい": it's now part of the primary
    # workflow row, only visible for PDF sources, and shows how many pages
    # still need review.
    app = QApplication.instance() or QApplication([])
    window = ExcelPrivacyCleanerWindow()
    try:
        window.show()
        app.processEvents()

        with tempfile.TemporaryDirectory(prefix="workflow_ui_pdf_") as tmpdir:
            xlsx_source = Path(tmpdir) / "fixture.xlsx"
            create_simple_xlsx(xlsx_source)
            window.set_source(xlsx_source)
            app.processEvents()
            assert_true(not window.pdf_review_button.isVisible(), "PDF review button should stay hidden for non-PDF sources")
            assert_true(not window.mode_note.isVisible(), "PDF assistance note should stay hidden for non-PDF sources")

            pdf_source = Path(tmpdir) / "fixture.pdf"
            create_simple_pdf(pdf_source)
            window.set_source(pdf_source)
            app.processEvents()
            assert_true(window.pdf_review_button.isVisible(), "PDF review button should appear for PDF sources")
            assert_true(window.mode_note.isVisible(), "PDF assistance note should appear for PDF sources")
            assert_true(not window.pdf_review_button.isEnabled(), "PDF review button should stay disabled before scanning")

            window.scan_file()
            app.processEvents()
            assert_true(window.pdf_review_button.isEnabled(), "PDF review button should enable after scanning")
            assert_true("残り" in window.pdf_review_button.text(), "PDF review button should show the remaining page count")
    finally:
        window.processor.cleanup()
        window.close()
        app.processEvents()
