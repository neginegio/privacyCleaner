from __future__ import annotations

import csv
import sys
from datetime import datetime
from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtGui import QCloseEvent, QIcon
from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QComboBox,
    QFileDialog,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QListWidget,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from .excel_processor import ExcelPrivacyProcessor, ProcessingOptions, write_findings_csv
from .models import Finding


SUPPORTED_EXTENSIONS = {".xlsx", ".xlsm"}


def asset_path(relative_path: str) -> Path:
    base_path = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parents[2]))
    return base_path / relative_path


class ExcelPrivacyCleanerWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("Excel 機密情報削除 - Presidio")
        icon_path = asset_path("assets/app_icon.ico")
        if icon_path.exists():
            self.setWindowIcon(QIcon(str(icon_path)))
        self.resize(1180, 760)
        self.setMinimumSize(980, 640)
        self.setAcceptDrops(True)

        self.processor = ExcelPrivacyProcessor()
        self.source_path: Path | None = None
        self.findings: list[Finding] = []

        self.path_label = QLabel("未選択")
        self.path_label.setTextInteractionFlags(Qt.TextSelectableByMouse)
        self.status_label = QLabel("待機中: 外部クラウドへ送信しません。")
        self.mode_combo = QComboBox()
        self.business_secret_checkbox = QCheckBox("企業機密も変換する")
        self.scope_combo = QComboBox()
        self.mode_note = QLabel("")
        self.history = QListWidget()
        self.table = QTableWidget(0, 8)
        self._build_ui()

    def _build_ui(self) -> None:
        root = QWidget()
        layout = QVBoxLayout(root)
        layout.setContentsMargins(14, 14, 14, 14)
        layout.setSpacing(10)

        title = QLabel("Excel ファイルを選択してください。Presidio による検出はこの PC 内だけで行い、原本は上書きしません。")
        title.setStyleSheet("font-size: 15px; font-weight: 600;")
        layout.addWidget(title)

        mode_row = QHBoxLayout()
        self.mode_combo.addItem("分析継続用", "analysis")
        self.mode_combo.addItem("外部共有用", "external")
        self.scope_combo.addItem("このファイル内だけ", "file")
        self.scope_combo.addItem("今回アップロードした一連のファイル内", "batch")
        self.scope_combo.addItem("プロジェクト内", "project")
        self.business_secret_checkbox.setChecked(False)
        self.mode_combo.currentIndexChanged.connect(self.update_mode_note)
        self.business_secret_checkbox.stateChanged.connect(self.update_mode_note)
        mode_row.addWidget(QLabel("処理モード:"))
        mode_row.addWidget(self.mode_combo)
        mode_row.addWidget(QLabel("仮名化範囲:"))
        mode_row.addWidget(self.scope_combo)
        mode_row.addWidget(self.business_secret_checkbox)
        mode_row.addStretch(1)
        layout.addLayout(mode_row)

        self.mode_note.setWordWrap(True)
        self.mode_note.setStyleSheet("border: 1px solid #fde68a; padding: 6px; background: #fffbeb; color: #713f12;")
        layout.addWidget(self.mode_note)
        self.update_mode_note()

        file_row = QHBoxLayout()
        choose_button = QPushButton("Excelを選択")
        scan_button = QPushButton("検査開始")
        convert_button = QPushButton("確認済みを変換保存")
        choose_button.clicked.connect(self.choose_file)
        scan_button.clicked.connect(self.scan_file)
        convert_button.clicked.connect(self.convert_file)
        file_row.addWidget(QLabel("ファイル:"))
        file_row.addWidget(self.path_label, 1)
        file_row.addWidget(choose_button)
        file_row.addWidget(scan_button)
        file_row.addWidget(convert_button)
        layout.addLayout(file_row)

        action_row = QHBoxLayout()
        toggle_button = QPushButton("選択行を切替")
        all_button = QPushButton("すべて変換")
        none_button = QPushButton("すべて除外")
        export_csv_button = QPushButton("検出結果CSV出力")
        clear_history_button = QPushButton("履歴消去")
        toggle_button.clicked.connect(self.toggle_selected)
        all_button.clicked.connect(lambda: self.set_all_enabled(True))
        none_button.clicked.connect(lambda: self.set_all_enabled(False))
        export_csv_button.clicked.connect(self.export_findings_csv)
        clear_history_button.clicked.connect(self.clear_history)
        action_row.addWidget(toggle_button)
        action_row.addWidget(all_button)
        action_row.addWidget(none_button)
        action_row.addWidget(export_csv_button)
        action_row.addStretch(1)
        action_row.addWidget(clear_history_button)
        layout.addLayout(action_row)

        headers = ["変換", "シート", "セル", "種類", "検査", "検出値", "変換後", "理由"]
        self.table.setHorizontalHeaderLabels(headers)
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setSelectionMode(QTableWidget.ExtendedSelection)
        self.table.verticalHeader().setVisible(False)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self.table.horizontalHeader().setSectionResizeMode(5, QHeaderView.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(7, QHeaderView.Stretch)
        self.table.setColumnWidth(0, 54)
        self.table.setColumnWidth(1, 110)
        self.table.setColumnWidth(2, 64)
        self.table.setColumnWidth(3, 74)
        self.table.setColumnWidth(4, 90)
        self.table.setColumnWidth(6, 130)
        self.table.itemDoubleClicked.connect(lambda _item: self.toggle_selected())
        layout.addWidget(self.table, 1)

        history_label = QLabel("変換履歴")
        history_label.setStyleSheet("font-weight: 600;")
        layout.addWidget(history_label)
        self.history.setMaximumHeight(110)
        layout.addWidget(self.history)

        self.status_label.setStyleSheet("border: 1px solid #cbd5e1; padding: 5px; background: #f8fafc;")
        layout.addWidget(self.status_label)

        self.setCentralWidget(root)

    def choose_file(self) -> None:
        filename, _ = QFileDialog.getOpenFileName(
            self,
            "検査する Excel を選択",
            "",
            "Excel files (*.xlsx *.xlsm)",
        )
        if filename:
            self.set_source(Path(filename))

    def set_source(self, path: Path) -> None:
        if path.suffix.lower() not in SUPPORTED_EXTENSIONS:
            QMessageBox.warning(self, "形式エラー", "対応形式は .xlsx / .xlsm です。")
            return
        self.processor.cleanup()
        self.source_path = path
        self.findings = []
        self.path_label.setText(str(path))
        self.refresh_table()
        self.status_label.setText("選択済み: 検査開始を押してください。")

    def scan_file(self) -> None:
        if self.source_path is None:
            QMessageBox.warning(self, "ファイル未選択", "Excel ファイルを選択してください。")
            return
        try:
            self.status_label.setText("検査中: Presidio カスタム Recognizer でローカル解析しています...")
            QApplication.processEvents()
            options = self.current_options()
            self.findings = self.processor.scan(self.source_path, options=options)
            self.refresh_table()
            formula_count = self.processor.enabled_formula_replacement_count(self.findings, options=options)
            formula_note = f" 数式文字列化予定: {formula_count} 件。" if formula_count else " 数式は維持します。"
            self.status_label.setText(f"検査完了: {len(self.findings)} 件を検出しました。変換対象を確認してください。{formula_note}")
        except Exception as exc:
            QMessageBox.critical(self, "検査エラー", str(exc))
            self.status_label.setText("検査エラー")

    def convert_file(self) -> None:
        if self.source_path is None:
            QMessageBox.warning(self, "ファイル未選択", "Excel ファイルを選択してください。")
            return
        try:
            self.update_findings_from_table()
            self.status_label.setText("変換中: 一時コピーへ置換を適用しています...")
            QApplication.processEvents()
            options = self.current_options()
            result = self.processor.convert_with_artifacts(self.source_path, self.findings, options=options)
            output_path = result.excel_path
            converted_count = result.converted_count
            formula_note = f"  数式維持 {result.formula_maintained_count} 件" if result.formula_maintained_count else ""
            self.history.insertItem(
                0,
                f"{datetime.now():%Y/%m/%d %H:%M:%S}  {options.mode_label}  {converted_count} 件変換{formula_note}  {output_path.name}  一時ファイル削除済み",
            )
            self.status_label.setText(f"保存完了: {output_path}")
            QMessageBox.information(
                self,
                "保存完了",
                "匿名化済み Excel、検出・変換結果CSV、処理報告書を保存しました。\n\n"
                f"Excel: {result.excel_path}\nCSV: {result.csv_path}\n報告書: {result.report_path}\n\n"
                "原本は上書きしていません。一時コピーは削除済みです。",
            )
        except Exception as exc:
            QMessageBox.critical(self, "変換エラー", str(exc))
            self.status_label.setText("変換エラー")

    def export_findings_csv(self) -> None:
        if not self.findings:
            QMessageBox.information(self, "CSV出力", "検出結果がありません。先に検査を実行してください。")
            return

        self.update_findings_from_table()
        default_name = self._default_csv_name()
        filename, _ = QFileDialog.getSaveFileName(
            self,
            "検出結果CSVを保存",
            default_name,
            "CSV files (*.csv)",
        )
        if not filename:
            return

        path = Path(filename)
        if path.suffix.lower() != ".csv":
            path = path.with_suffix(".csv")

        try:
            write_findings_csv(path, self.findings)
            self.status_label.setText(f"CSV出力完了: {path}")
            QMessageBox.information(self, "CSV出力完了", f"検出結果CSVを保存しました。\n\n{path}")
        except Exception as exc:
            QMessageBox.critical(self, "CSV出力エラー", str(exc))
            self.status_label.setText("CSV出力エラー")

    def refresh_table(self) -> None:
        self.table.setRowCount(0)
        for row, finding in enumerate(self.findings):
            self.table.insertRow(row)
            enabled = QTableWidgetItem()
            enabled.setFlags(Qt.ItemIsUserCheckable | Qt.ItemIsEnabled | Qt.ItemIsSelectable)
            enabled.setCheckState(Qt.Checked if finding.enabled else Qt.Unchecked)
            self.table.setItem(row, 0, enabled)

            values = [
                finding.sheet,
                finding.cell,
                finding.entity_type,
                finding.detection_kind,
                finding.original,
                finding.replacement,
                finding.reason,
            ]
            for offset, value in enumerate(values, start=1):
                item = QTableWidgetItem(value)
                if offset != 6:
                    item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                self.table.setItem(row, offset, item)

    def update_findings_from_table(self) -> None:
        for row, finding in enumerate(self.findings):
            enabled_item = self.table.item(row, 0)
            replacement_item = self.table.item(row, 6)
            finding.enabled = enabled_item is not None and enabled_item.checkState() == Qt.Checked
            if replacement_item is not None and replacement_item.text().strip():
                finding.replacement = replacement_item.text().strip()

    def toggle_selected(self) -> None:
        rows = sorted({index.row() for index in self.table.selectedIndexes()})
        for row in rows:
            item = self.table.item(row, 0)
            if item is not None:
                item.setCheckState(Qt.Unchecked if item.checkState() == Qt.Checked else Qt.Checked)

    def set_all_enabled(self, enabled: bool) -> None:
        for row in range(self.table.rowCount()):
            item = self.table.item(row, 0)
            if item is not None:
                item.setCheckState(Qt.Checked if enabled else Qt.Unchecked)

    def clear_history(self) -> None:
        self.history.clear()
        self.status_label.setText("履歴を消去しました。")

    def current_options(self) -> ProcessingOptions:
        mode = str(self.mode_combo.currentData())
        return ProcessingOptions(
            mode=mode,
            transform_business_secrets=self.business_secret_checkbox.isChecked() or mode == "external",
            pseudonym_scope=str(self.scope_combo.currentData()),
        )

    def update_mode_note(self) -> None:
        options = self.current_options()
        self.business_secret_checkbox.setEnabled(options.is_analysis)
        if options.is_analysis:
            self.mode_note.setText(
                "分析継続用では、金額、数量、原価、評価などの分析項目を維持します。"
                "外部へ提供する場合は、企業機密情報の追加変換を確認してください。"
            )
        else:
            self.mode_note.setText("外部共有用では、既存の匿名化方針に近い形で企業機密項目も変換対象にします。")

    def _default_csv_name(self) -> str:
        folder = self.source_path.parent if self.source_path else Path.cwd()
        stem = self.source_path.stem if self.source_path else "検出結果"
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        return str(folder / f"{stem}_検出結果_{timestamp}.csv")

    def _write_findings_csv(self, path: Path) -> None:
        headers = [
            "変換",
            "処理状態",
            "シート",
            "セル",
            "種類",
            "検査",
            "検出値",
            "変換後",
            "理由",
            "問題種別",
            "期待する判定",
            "期待する変換",
            "メモ",
        ]
        with path.open("w", encoding="utf-8-sig", newline="") as output:
            writer = csv.writer(output)
            writer.writerow(headers)
            for finding in self.findings:
                status = "確認候補" if finding.detection_kind == "確認候補" else ("自動変換" if finding.enabled else "除外")
                writer.writerow(
                    [
                        "対象" if finding.enabled else "除外",
                        status,
                        finding.sheet,
                        finding.cell,
                        finding.entity_type,
                        finding.detection_kind,
                        finding.original,
                        finding.replacement,
                        finding.reason,
                        "",
                        "",
                        "",
                        "",
                    ]
                )

    def dragEnterEvent(self, event) -> None:  # type: ignore[override]
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event) -> None:  # type: ignore[override]
        urls = event.mimeData().urls()
        if not urls:
            return
        path = Path(urls[0].toLocalFile())
        self.set_source(path)

    def closeEvent(self, event: QCloseEvent) -> None:
        self.processor.cleanup()
        event.accept()


def main() -> int:
    app = QApplication(sys.argv)
    icon_path = asset_path("assets/app_icon.ico")
    if icon_path.exists():
        app.setWindowIcon(QIcon(str(icon_path)))
    window = ExcelPrivacyCleanerWindow()
    window.show()
    return app.exec()
