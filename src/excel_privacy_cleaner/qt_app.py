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
from .pdf_processor import (
    PDF_ASSISTANCE_NOTICE,
    PDF_REDACTION_MODES,
    PdfPrivacyProcessor,
    final_output_status,
    validate_ocr_environment,
    write_pdf_findings_csv,
)
from .pdf_review_dialog import PdfCandidateReviewDialog
from .resources import resource_path
from .word_processor import (
    WORD_SUPPORTED_EXTENSION,
    WordPrivacyProcessor,
    WordReplacementDecision,
    word_candidate_location_label,
    word_finding_reason,
    word_finding_status,
    write_word_findings_csv,
)


EXCEL_EXTENSIONS = {".xlsx", ".xlsm"}
PDF_EXTENSIONS = {".pdf"}
WORD_EXTENSIONS = {WORD_SUPPORTED_EXTENSION}
SUPPORTED_EXTENSIONS = EXCEL_EXTENSIONS | PDF_EXTENSIONS | WORD_EXTENSIONS


def asset_path(relative_path: str) -> Path:
    return resource_path(relative_path)


def _finding_from_word_decision(decision: WordReplacementDecision) -> Finding:
    candidate = decision.candidate
    return Finding(
        enabled=decision.enabled,
        sheet=word_candidate_location_label(candidate),
        cell="",
        entity_type=candidate.category,
        detection_kind=word_finding_status(decision),
        original=candidate.text,
        replacement=decision.replacement,
        reason=word_finding_reason(decision),
        start=candidate.char_start,
        end=candidate.char_end,
    )


class ExcelPrivacyCleanerWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("hoso Privacy Cleaner")
        icon_path = asset_path("assets/app_icon.ico")
        if icon_path.exists():
            self.setWindowIcon(QIcon(str(icon_path)))
        self.resize(1180, 760)
        self.setMinimumSize(980, 640)
        self.setAcceptDrops(True)

        self.processor: ExcelPrivacyProcessor | PdfPrivacyProcessor | WordPrivacyProcessor = ExcelPrivacyProcessor()
        self.source_path: Path | None = None
        self.findings: list[Finding] = []
        self.word_decisions: list[WordReplacementDecision] = []

        self.path_label = QLabel("未選択")
        self.path_label.setTextInteractionFlags(Qt.TextSelectableByMouse)
        self.status_label = QLabel("待機中: 外部クラウドへ送信しません。")
        self.mode_combo = QComboBox()
        self.business_secret_checkbox = QCheckBox("企業機密も変換する")
        self.scope_combo = QComboBox()
        self.pdf_redaction_combo = QComboBox()
        self.pdf_review_button = QPushButton("PDF候補確認")
        self.mode_note = QLabel("")
        self.history = QListWidget()
        self.table = QTableWidget(0, 9)
        self._build_ui()

    def _build_ui(self) -> None:
        root = QWidget()
        layout = QVBoxLayout(root)
        layout.setContentsMargins(14, 14, 14, 14)
        layout.setSpacing(10)

        title = QLabel("Excel / PDF / Word ファイルを選択してください。検出はこの PC 内だけで行い、原本は上書きしません。")
        title.setStyleSheet("font-size: 15px; font-weight: 600;")
        layout.addWidget(title)

        mode_row = QHBoxLayout()
        self.mode_combo.addItem("分析継続用", "analysis")
        self.mode_combo.addItem("外部共有用", "external")
        self.scope_combo.addItem("このファイル内だけ", "file")
        self.scope_combo.addItem("今回アップロードした一連のファイル内", "batch")
        self.scope_combo.addItem("プロジェクト内", "project")
        for mode_key, mode_label in PDF_REDACTION_MODES.items():
            self.pdf_redaction_combo.addItem(mode_label, mode_key)
        self.business_secret_checkbox.setChecked(False)
        self.mode_combo.currentIndexChanged.connect(self.update_mode_note)
        self.business_secret_checkbox.stateChanged.connect(self.update_mode_note)
        mode_row.addWidget(QLabel("処理モード:"))
        mode_row.addWidget(self.mode_combo)
        mode_row.addWidget(QLabel("仮名化範囲:"))
        mode_row.addWidget(self.scope_combo)
        mode_row.addWidget(self.business_secret_checkbox)
        mode_row.addWidget(QLabel("PDF匿名化方法:"))
        mode_row.addWidget(self.pdf_redaction_combo)
        mode_row.addStretch(1)
        layout.addLayout(mode_row)

        self.mode_note.setWordWrap(True)
        self.mode_note.setStyleSheet("border: 1px solid #fde68a; padding: 6px; background: #fffbeb; color: #713f12;")
        layout.addWidget(self.mode_note)
        self.update_mode_note()

        file_row = QHBoxLayout()
        choose_button = QPushButton("Excel/PDF/Wordを選択")
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
        self.pdf_review_button.clicked.connect(self.open_pdf_review)
        action_row.addWidget(toggle_button)
        action_row.addWidget(all_button)
        action_row.addWidget(none_button)
        action_row.addWidget(self.pdf_review_button)
        action_row.addWidget(export_csv_button)
        action_row.addStretch(1)
        action_row.addWidget(clear_history_button)
        layout.addLayout(action_row)

        headers = ["変換する", "変換しない", "シート", "セル", "種類", "検査", "検出値", "変換後", "理由"]
        self.table.setHorizontalHeaderLabels(headers)
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setSelectionMode(QTableWidget.ExtendedSelection)
        self.table.verticalHeader().setVisible(False)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self.table.horizontalHeader().setSectionResizeMode(6, QHeaderView.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(8, QHeaderView.Stretch)
        self.table.setColumnWidth(0, 64)
        self.table.setColumnWidth(1, 64)
        self.table.setColumnWidth(2, 110)
        self.table.setColumnWidth(3, 64)
        self.table.setColumnWidth(4, 74)
        self.table.setColumnWidth(5, 90)
        self.table.setColumnWidth(7, 130)
        self.table.itemDoubleClicked.connect(lambda _item: self.toggle_selected())
        self.table.itemChanged.connect(self._on_table_item_changed)
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
            "検査するファイルを選択",
            "",
            "Supported files (*.xlsx *.xlsm *.pdf *.docx);;Excel files (*.xlsx *.xlsm);;PDF files (*.pdf);;Word files (*.docx)",
        )
        if filename:
            self.set_source(Path(filename))

    def set_source(self, path: Path) -> None:
        if path.suffix.lower() not in SUPPORTED_EXTENSIONS:
            QMessageBox.warning(self, "形式エラー", "対応形式は .xlsx / .xlsm / .pdf / .docx です。")
            return
        self.processor.cleanup()
        if path.suffix.lower() in PDF_EXTENSIONS:
            self.processor = PdfPrivacyProcessor()
        elif path.suffix.lower() in WORD_EXTENSIONS:
            self.processor = WordPrivacyProcessor()
        else:
            self.processor = ExcelPrivacyProcessor()
        self.source_path = path
        self.findings = []
        self.word_decisions = []
        self.path_label.setText(str(path))
        self.update_mode_note()
        self.refresh_table()
        if path.suffix.lower() in PDF_EXTENSIONS:
            ocr_errors = validate_ocr_environment()
            if ocr_errors:
                QMessageBox.warning(self, "PDF OCR設定エラー", "\n".join(ocr_errors))
                self.status_label.setText("PDF選択済み: OCR設定に不足があります。")
                return
            self.status_label.setText("PDF選択済み: 全ページ確認前提の匿名化支援です。検査開始を押してください。")
            return
        self.status_label.setText("選択済み: 検査開始を押してください。")

    def scan_file(self) -> None:
        if self.source_path is None:
            QMessageBox.warning(self, "ファイル未選択", "ファイルを選択してください。")
            return
        busy_cursor = False
        try:
            if self.is_pdf_source():
                self.status_label.setText("PDF検査中: OCRとローカル検出を実行しています。最終出力には全ページ確認が必要です...")
            else:
                self.status_label.setText("検査中: Presidio カスタム Recognizer でローカル解析しています...")
            QApplication.setOverrideCursor(Qt.WaitCursor)
            busy_cursor = True
            QApplication.processEvents()
            options = self.current_options()
            if self.is_word_source() and isinstance(self.processor, WordPrivacyProcessor):
                self.word_decisions = self.processor.scan(self.source_path, options=options)
                self.findings = [_finding_from_word_decision(decision) for decision in self.word_decisions]
            else:
                self.findings = self.processor.scan(self.source_path, options=options)
            self.refresh_table()
            formula_count = (
                self.processor.enabled_formula_replacement_count(self.findings, options=options)
                if isinstance(self.processor, ExcelPrivacyProcessor)
                else 0
            )
            formula_note = (
                f" 数式文字列化予定: {formula_count} 件。"
                if formula_count
                else (
                    " PDFは全ページ確認が必要です。PDF候補確認を開いてください。"
                    if self.is_pdf_source()
                    else (" Word固有の項目はありません。" if self.is_word_source() else " 数式は維持します。")
                )
            )
            if self.is_pdf_source():
                self.status_label.setText(
                    f"PDF検査完了: {len(self.findings)} 件を検出しました。"
                    "PDF候補確認で全ページを確認するまで、匿名化済みPDFとして出力できません。"
                )
            elif self.is_word_source():
                review_required_count = sum(1 for finding in self.findings if finding.detection_kind == "要確認(未処理)")
                self.status_label.setText(
                    f"Word検査完了: {len(self.findings)} 件を検出しました。"
                    f"要確認候補 {review_required_count} 件は変換前に承認(チェック)してください。"
                )
            else:
                self.status_label.setText(f"検査完了: {len(self.findings)} 件を検出しました。変換対象を確認してください。{formula_note}")
            self.update_pdf_review_button()
        except Exception as exc:
            if busy_cursor:
                QApplication.restoreOverrideCursor()
                busy_cursor = False
            QMessageBox.critical(self, "検査エラー", str(exc))
            self.status_label.setText("検査エラー")
            self.update_pdf_review_button()
        finally:
            if busy_cursor:
                QApplication.restoreOverrideCursor()

    def convert_file(self) -> None:
        if self.source_path is None:
            QMessageBox.warning(self, "ファイル未選択", "ファイルを選択してください。")
            return
        busy_cursor = False
        try:
            self.update_findings_from_table()
            self.status_label.setText("変換中: 一時コピーへ置換を適用しています...")
            options = self.current_options()
            if self.is_word_source() and isinstance(self.processor, WordPrivacyProcessor):
                self._sync_word_decisions_from_findings()
                QApplication.setOverrideCursor(Qt.WaitCursor)
                busy_cursor = True
                QApplication.processEvents()
                result = self.processor.convert(self.source_path, self.word_decisions)
                QApplication.restoreOverrideCursor()
                busy_cursor = False
                output_path = result.output_path
                converted_count = result.converted_run_count + result.converted_property_count
                self.history.insertItem(
                    0,
                    f"{datetime.now():%Y/%m/%d %H:%M:%S}  {options.mode_label}  {converted_count} 件変換  {output_path.name}  一時ファイル削除済み",
                )
                self.status_label.setText(f"保存完了: {output_path}")
                warning_note = ("\n\n警告:\n" + "\n".join(result.warnings)) if result.warnings else ""
                QMessageBox.information(
                    self,
                    "保存完了",
                    "匿名化済み Word、検出・変換結果CSV、処理報告書を保存しました。\n\n"
                    f"Word: {result.output_path}\nCSV: {result.csv_path}\n報告書: {result.report_path}"
                    f"{warning_note}\n\n"
                    "原本は上書きしていません。一時コピーは削除済みです。",
                )
            elif self.is_pdf_source() and isinstance(self.processor, PdfPrivacyProcessor):
                can_output, reasons = final_output_status(
                    self.findings,
                    self.processor.page_quality,
                    self.processor.confirmed_pages,
                    self.processor.page_review_state,
                )
                summary = self._pdf_output_summary(can_output, reasons)
                if not can_output:
                    QMessageBox.warning(self, "PDF出力不可", summary)
                    self.status_label.setText("PDF出力不可")
                    return
                if QMessageBox.question(self, "PDF最終出力確認", summary) != QMessageBox.Yes:
                    self.status_label.setText("PDF出力をキャンセルしました。")
                    return
                QApplication.setOverrideCursor(Qt.WaitCursor)
                busy_cursor = True
                QApplication.processEvents()
                result = self.processor.convert_with_artifacts(
                    self.source_path,
                    self.findings,
                    options=options,
                    redaction_mode=str(self.pdf_redaction_combo.currentData()),
                )
                QApplication.restoreOverrideCursor()
                busy_cursor = False
                output_path = result.pdf_path
                self.history.insertItem(
                    0,
                    f"{datetime.now():%Y/%m/%d %H:%M:%S}  PDF  {result.converted_count} 件変換  {output_path.name}  一時ファイル削除済み",
                )
                self.status_label.setText(f"保存完了: {output_path}")
                QMessageBox.information(
                    self,
                    "保存完了",
                    "全ページ確認済みのPDFとして、匿名化済みPDF、検出・変換結果CSV、処理報告書を保存しました。\n\n"
                    f"PDF: {result.pdf_path}\nCSV: {result.csv_path}\n報告書: {result.report_path}\n\n"
                    "PDF OCR匿名化は支援機能です。報告書で確認状態と検証状態を確認してください。\n"
                    "原本は上書きしていません。一時コピーは削除済みです。",
                )
            else:
                QApplication.setOverrideCursor(Qt.WaitCursor)
                busy_cursor = True
                QApplication.processEvents()
                result = self.processor.convert_with_artifacts(self.source_path, self.findings, options=options)
                QApplication.restoreOverrideCursor()
                busy_cursor = False
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
            if busy_cursor:
                QApplication.restoreOverrideCursor()
                busy_cursor = False
            QMessageBox.critical(self, "変換エラー", str(exc))
            self.status_label.setText("変換エラー")
        finally:
            if busy_cursor:
                QApplication.restoreOverrideCursor()

    def export_findings_csv(self) -> None:
        if not self.findings:
            QMessageBox.information(self, "CSV出力", "検出結果がありません。先に検査を実行してください。")
            return

        self.update_findings_from_table()
        if self.is_word_source():
            self._sync_word_decisions_from_findings()
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
            if self.is_pdf_source():
                write_pdf_findings_csv(path, self.findings)
            elif self.is_word_source():
                write_word_findings_csv(path, self.word_decisions)
            else:
                write_findings_csv(path, self.findings)
            self.status_label.setText(f"CSV出力完了: {path}")
            QMessageBox.information(self, "CSV出力完了", f"検出結果CSVを保存しました。\n\n{path}")
        except Exception as exc:
            QMessageBox.critical(self, "CSV出力エラー", str(exc))
            self.status_label.setText("CSV出力エラー")

    def open_pdf_review(self) -> None:
        if not self.is_pdf_source() or not isinstance(self.processor, PdfPrivacyProcessor):
            QMessageBox.information(self, "PDF候補確認", "PDFファイルの検査後に利用できます。")
            return
        if not self.source_path or not self.processor.temp_pdf:
            QMessageBox.information(self, "PDF候補確認", "先にPDF検査を実行してください。")
            return
        dialog = PdfCandidateReviewDialog(self.processor, self.findings, self)
        dialog.exec()
        self.refresh_table()
        self.status_label.setText("PDF候補確認を反映しました。全ページが確認済みになるまで最終出力できません。")

    def refresh_table(self) -> None:
        self.table.blockSignals(True)
        self.table.setRowCount(0)
        is_word = self.is_word_source()
        # Word candidate categories are fixed by detection and can't be edited
        # after the fact (WordCandidate is immutable), unlike Excel/PDF's
        # entity_type, so column 4 stays read-only for Word rows.
        editable_offsets = {7} if is_word else {4, 7}
        for row, finding in enumerate(self.findings):
            self.table.insertRow(row)
            enabled = QTableWidgetItem()
            enabled.setFlags(Qt.ItemIsUserCheckable | Qt.ItemIsEnabled | Qt.ItemIsSelectable)
            enabled.setCheckState(Qt.Checked if finding.enabled else Qt.Unchecked)
            self.table.setItem(row, 0, enabled)

            # "変換しない" (reviewed-and-excluded) is a Word-only concept --
            # Excel/PDF have no third state, so the checkbox stays absent
            # (not just unchecked) for their rows.
            excluded = QTableWidgetItem()
            if is_word:
                excluded.setFlags(Qt.ItemIsUserCheckable | Qt.ItemIsEnabled | Qt.ItemIsSelectable)
                is_excluded = row < len(self.word_decisions) and self.word_decisions[row].excluded
                excluded.setCheckState(Qt.Checked if is_excluded else Qt.Unchecked)
            else:
                excluded.setFlags(Qt.ItemIsSelectable)
            self.table.setItem(row, 1, excluded)

            values = [
                finding.sheet,
                finding.cell,
                finding.entity_type,
                finding.detection_kind,
                finding.original,
                finding.replacement,
                finding.reason,
            ]
            for offset, value in enumerate(values, start=2):
                item = QTableWidgetItem(value)
                if offset not in editable_offsets:
                    item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                self.table.setItem(row, offset, item)
        self.table.blockSignals(False)

    def update_findings_from_table(self) -> None:
        for row, finding in enumerate(self.findings):
            enabled_item = self.table.item(row, 0)
            entity_item = self.table.item(row, 4)
            replacement_item = self.table.item(row, 7)
            finding.enabled = enabled_item is not None and enabled_item.checkState() == Qt.Checked
            if entity_item is not None and entity_item.text().strip():
                finding.entity_type = entity_item.text().strip()
            if replacement_item is not None and replacement_item.text().strip():
                finding.replacement = replacement_item.text().strip()

    def _on_table_item_changed(self, item: QTableWidgetItem) -> None:
        if item.column() not in (0, 1):
            return
        row = item.row()
        other_column = 1 - item.column()
        if item.checkState() == Qt.Checked:
            other_item = self.table.item(row, other_column)
            if other_item is not None and other_item.checkState() == Qt.Checked:
                self.table.blockSignals(True)
                other_item.setCheckState(Qt.Unchecked)
                self.table.blockSignals(False)
        if self.is_word_source():
            self._refresh_word_row_status(row)

    def _refresh_word_row_status(self, row: int) -> None:
        if row >= len(self.word_decisions):
            return
        enabled_item = self.table.item(row, 0)
        excluded_item = self.table.item(row, 1)
        decision = self.word_decisions[row]
        decision.enabled = enabled_item is not None and enabled_item.checkState() == Qt.Checked
        decision.excluded = bool(
            excluded_item is not None and excluded_item.checkState() == Qt.Checked and not decision.enabled
        )
        if row < len(self.findings):
            self.findings[row].enabled = decision.enabled
        self.table.blockSignals(True)
        status_item = self.table.item(row, 5)
        if status_item is not None:
            status_item.setText(word_finding_status(decision))
        reason_item = self.table.item(row, 8)
        if reason_item is not None:
            reason_item.setText(word_finding_reason(decision))
        self.table.blockSignals(False)

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
        if self.is_pdf_source():
            self.mode_note.setText(self.mode_note.text() + "\n" + PDF_ASSISTANCE_NOTICE)
        self.pdf_redaction_combo.setEnabled(self.is_pdf_source())
        self.update_pdf_review_button()

    def is_pdf_source(self) -> bool:
        return self.source_path is not None and self.source_path.suffix.lower() in PDF_EXTENSIONS

    def is_word_source(self) -> bool:
        return self.source_path is not None and self.source_path.suffix.lower() in WORD_EXTENSIONS

    def _sync_word_decisions_from_findings(self) -> None:
        for row, (decision, finding) in enumerate(zip(self.word_decisions, self.findings)):
            decision.enabled = finding.enabled
            excluded_item = self.table.item(row, 1)
            decision.excluded = bool(
                excluded_item is not None and excluded_item.checkState() == Qt.Checked and not decision.enabled
            )
            if finding.replacement.strip():
                decision.replacement = finding.replacement.strip()

    def update_pdf_review_button(self) -> None:
        self.pdf_review_button.setEnabled(self.is_pdf_source() and isinstance(self.processor, PdfPrivacyProcessor) and bool(self.processor.temp_pdf))

    def _default_csv_name(self) -> str:
        folder = self.source_path.parent if self.source_path else Path.cwd()
        stem = self.source_path.stem if self.source_path else "検出結果"
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        return str(folder / f"{stem}_検出結果_{timestamp}.csv")

    def _pdf_output_summary(self, can_output: bool, reasons: list[str]) -> str:
        if not isinstance(self.processor, PdfPrivacyProcessor):
            return ""
        total_pages = self.processor.page_count
        state_counts: dict[str, int] = {}
        for state in self.processor.page_review_state.values():
            state_counts[state] = state_counts.get(state, 0) + 1
        completed_like = sum(
            state_counts.get(state, 0)
            for state in ("REVIEWED_NO_SENSITIVE_DATA", "REVIEWED_WITH_REDACTIONS", "COMPLETED")
        )
        unresolved_count = sum(1 for finding in self.findings if finding.detection_kind in {"確認候補", "REVIEW_REQUIRED"})
        failed_count = state_counts.get("FAILED_UNRESOLVED", 0)
        manual_count = sum(1 for finding in self.findings if finding.detection_kind == "MANUAL")
        selected_count = sum(1 for finding in self.findings if finding.enabled)
        text = [
            "PDF OCR匿名化は全ページ確認前提の支援機能です。",
            "未確認ページ、未確認候補、OCR不能またはFAILED未対応ページ、検証失敗ページが残っている場合は出力できません。",
            f"全ページ数: {total_pages}",
            f"確認済みページ数: {completed_like}",
            f"UNREVIEWEDページ数: {state_counts.get('UNREVIEWED', 0)}",
            f"未確認候補数: {unresolved_count}",
            f"FAILED_UNRESOLVEDページ数: {failed_count}",
            f"VERIFICATION_FAILEDページ数: {state_counts.get('VERIFICATION_FAILED', 0)}",
            f"手動追加範囲数: {manual_count}",
            f"匿名化処理予定件数: {selected_count}",
            "再OCR/残存検証: 出力後に区分して記録",
            f"出力可否: {'出力可能（全ページ確認済み）' if can_output else '出力不可'}",
        ]
        if reasons:
            text.append("理由: " + " / ".join(reasons))
        return "\n".join(text)

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
