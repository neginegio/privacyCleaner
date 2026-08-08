from __future__ import annotations

import os
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from openpyxl import Workbook  # noqa: E402
from PySide6.QtCore import Qt  # noqa: E402
from PySide6.QtTest import QTest  # noqa: E402
from PySide6.QtWidgets import QApplication  # noqa: E402

from excel_privacy_cleaner.qt_app import ExcelPrivacyCleanerWindow  # noqa: E402


def assert_true(condition: bool, message: str) -> None:
    if not condition:
        raise AssertionError(message)


def create_simple_fixture(path: Path) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Sheet1"
    sheet["A1"] = "連絡先"
    sheet["A2"] = "連絡先電話は090-1111-2222です。"
    workbook.save(path)


def test_clicking_anywhere_in_checkbox_cell_toggles_it() -> None:
    # Reproduces the report that only the small native checkbox glyph
    # (near the cell's left edge) responded to clicks, not the rest of the
    # cell -- see qt_app.py's _CheckboxCell, which uses a real centered
    # QCheckBox and handles the click on the whole cell widget instead.
    app = QApplication.instance() or QApplication([])

    with tempfile.TemporaryDirectory(prefix="qt_app_checkbox_click_") as tmpdir:
        tmp = Path(tmpdir)
        source = tmp / "fixture.xlsx"
        create_simple_fixture(source)

        window = ExcelPrivacyCleanerWindow()
        try:
            window.set_source(source)
            window.scan_file()
            assert_true(window.table.rowCount() > 0, "Fixture should yield at least one candidate")

            checkbox = window._row_checkbox(0, 0)
            assert_true(checkbox is not None, "Checkbox should exist")
            before = checkbox.isChecked()

            cell_widget = window.table.cellWidget(0, 0)
            # Click near the right edge of the cell, away from where the
            # small native checkbox glyph renders (near the left edge) --
            # this is exactly the spot that used to not respond at all.
            click_point = cell_widget.rect().center()
            click_point.setX(cell_widget.rect().right() - 2)
            QTest.mouseClick(cell_widget, Qt.LeftButton, Qt.NoModifier, click_point)

            after = checkbox.isChecked()
            assert_true(after != before, "Clicking anywhere in the checkbox cell should toggle it")

            QTest.mouseClick(cell_widget, Qt.LeftButton, Qt.NoModifier, click_point)
            assert_true(checkbox.isChecked() == before, "A second click should toggle it back")
        finally:
            window.processor.cleanup()
            window.close()
            app.processEvents()


def test_checkbox_cell_is_horizontally_centered() -> None:
    app = QApplication.instance() or QApplication([])

    with tempfile.TemporaryDirectory(prefix="qt_app_checkbox_center_") as tmpdir:
        tmp = Path(tmpdir)
        source = tmp / "fixture.xlsx"
        create_simple_fixture(source)

        window = ExcelPrivacyCleanerWindow()
        try:
            window.show()
            window.set_source(source)
            window.scan_file()
            assert_true(window.table.rowCount() > 0, "Fixture should yield at least one candidate")
            app.processEvents()

            checkbox = window._row_checkbox(0, 0)
            cell_widget = window.table.cellWidget(0, 0)
            assert_true(checkbox is not None and cell_widget is not None, "Checkbox cell should exist")

            cell_center_x = cell_widget.rect().center().x()
            checkbox_center_x = checkbox.geometry().center().x()
            assert_true(
                abs(cell_center_x - checkbox_center_x) <= 2,
                f"Checkbox should be horizontally centered in its cell (cell center={cell_center_x}, checkbox center={checkbox_center_x})",
            )
        finally:
            window.processor.cleanup()
            window.close()
            app.processEvents()


def test_removed_ui_elements_are_gone() -> None:
    app = QApplication.instance() or QApplication([])
    window = ExcelPrivacyCleanerWindow()
    try:
        assert_true(not hasattr(window, "history"), "変換履歴 list widget should be removed")
        assert_true(not hasattr(window, "toggle_selected"), "選択行を切替 handler should be removed")
        assert_true(not hasattr(window, "clear_history"), "履歴消去 handler should be removed")
    finally:
        window.processor.cleanup()
        window.close()
        app.processEvents()
