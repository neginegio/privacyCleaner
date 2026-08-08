from __future__ import annotations

import csv
import sys
from datetime import datetime
from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtGui import QBrush, QCloseEvent, QColor, QIcon
from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QComboBox,
    QFileDialog,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from .excel_processor import EXCEL_NLP_DETECTION_KIND, ExcelPrivacyProcessor, ProcessingOptions, write_findings_csv
from .models import Finding
from .pdf_ocr_support import (
    CANDIDATE_AUTO,
    CANDIDATE_MANUAL,
    CANDIDATE_REVIEW,
    USER_APPROVED,
    USER_REJECTED,
)
from .pdf_processor import (
    PDF_ASSISTANCE_NOTICE,
    PDF_REDACTION_MODES,
    PdfPrivacyProcessor,
    final_output_status,
    pdf_review_state_path,
    validate_ocr_environment,
    write_pdf_findings_csv,
)
from .pdf_review_dialog import PdfCandidateReviewDialog
from .pptx_processor import (
    PPTX_SUPPORTED_EXTENSION,
    PptxPrivacyProcessor,
    PptxReplacementDecision,
    pptx_candidate_location_label,
    pptx_finding_reason,
    pptx_finding_status,
    write_pptx_findings_csv,
)
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
PPTX_EXTENSIONS = {PPTX_SUPPORTED_EXTENSION}
SUPPORTED_EXTENSIONS = EXCEL_EXTENSIONS | PDF_EXTENSIONS | WORD_EXTENSIONS | PPTX_EXTENSIONS

# PDF findings store detection_kind using internal English constants
# (USER_APPROVED, etc.) so evaluation tooling and the CSV/audit exports can
# keep comparing against stable values. This maps them to Japanese labels
# for on-screen display only.
#
# 検査欄は常に「今のレビュー状態」だけを表す方針: CANDIDATE_AUTO(高信頼度の
# 自動確定)は人が未確認のまま変換される点で USER_APPROVED と同じ状態なので、
# 同じ「承認済み」ラベルにまとめる。検出方法(GiNZA/PDFテキスト層パターン一
# 致等)は理由欄でのみ示す。"確認候補" は旧バージョンのPDF側リテラル値の後方
# 互換用。
PDF_DETECTION_KIND_LABELS = {
    CANDIDATE_REVIEW: "要確認",
    "確認候補": "要確認",
    USER_APPROVED: "承認済み",
    USER_REJECTED: "却下済み",
    CANDIDATE_MANUAL: "手動追加",
    CANDIDATE_AUTO: "承認済み",
    "MERGED": "結合済み",
}


def display_detection_kind(value: str) -> str:
    return PDF_DETECTION_KIND_LABELS.get(value, value)


# Same highlight color used by the PDF candidate review dialog's "未確認" rows.
UNRESOLVED_ROW_COLOR = QColor("#fef3c7")

# 出力ボタンを画面内で唯一のアクセント色付きボタンにして、ワークフローの
# 最終ゴールであることを視覚的に示す。
PRIMARY_BUTTON_STYLE = (
    "QPushButton { background: #2563eb; color: white; font-weight: 600; padding: 6px 14px; border-radius: 4px; border: none; }"
    " QPushButton:disabled { background: #cbd5e1; color: #64748b; }"
    " QPushButton:hover:!disabled { background: #1d4ed8; }"
)
WARNING_BUTTON_STYLE = (
    "QPushButton { background: #fef3c7; color: #92400e; font-weight: 600; padding: 6px 14px; border: 1px solid #f59e0b; border-radius: 4px; }"
    " QPushButton:hover:!disabled { background: #fde68a; }"
)


class _CheckboxCell(QWidget):
    """A checkbox centered in a table cell that toggles on a click anywhere
    in the cell, not just the small native checkbox glyph.

    QTableWidgetItem's built-in checkbox rendering doesn't honor
    setTextAlignment(Qt.AlignCenter) for the check indicator's position in
    this app's style, and only the indicator's own tiny native hit-rect
    (near the cell's left edge) responds to clicks. Using a real QCheckBox
    inside a centered layout fixes the positioning; making the checkbox
    itself mouse-transparent and handling the click on this container
    instead avoids the checkbox's own native click-to-toggle firing a
    second time on top of ours when a click happens to land on the glyph.
    """

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.checkbox = QCheckBox(self)
        self.checkbox.setAttribute(Qt.WA_TransparentForMouseEvents, True)
        self.checkbox.setFocusPolicy(Qt.NoFocus)
        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addStretch(1)
        layout.addWidget(self.checkbox)
        layout.addStretch(1)

    def mousePressEvent(self, event) -> None:  # type: ignore[override]
        if event.button() == Qt.LeftButton and self.checkbox.isEnabled():
            self.checkbox.toggle()
            event.accept()
            return
        super().mousePressEvent(event)


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
        excluded=decision.excluded,
    )


def _finding_from_pptx_decision(decision: PptxReplacementDecision) -> Finding:
    candidate = decision.candidate
    return Finding(
        enabled=decision.enabled,
        sheet=pptx_candidate_location_label(candidate),
        cell="",
        entity_type=candidate.category,
        detection_kind=pptx_finding_status(decision),
        original=candidate.text,
        replacement=decision.replacement,
        reason=pptx_finding_reason(decision),
        start=candidate.char_start,
        end=candidate.char_end,
        excluded=decision.excluded,
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

        self.processor: ExcelPrivacyProcessor | PdfPrivacyProcessor | WordPrivacyProcessor | PptxPrivacyProcessor = ExcelPrivacyProcessor()
        self.source_path: Path | None = None
        self.findings: list[Finding] = []
        self.word_decisions: list[WordReplacementDecision] = []
        self.pptx_decisions: list[PptxReplacementDecision] = []

        self.path_label = QLabel("未選択")
        self.path_label.setTextInteractionFlags(Qt.TextSelectableByMouse)
        self.status_label = QLabel("待機中: 外部クラウドへ送信しません。")
        self.mode_combo = QComboBox()
        self.business_secret_checkbox = QCheckBox("企業機密も変換する")
        self.scope_combo = QComboBox()
        self.pdf_redaction_combo = QComboBox()
        self.choose_button = QPushButton("匿名化したいファイルを選択")
        self.scan_button = QPushButton("検査開始")
        self.convert_button = QPushButton("匿名化したファイルを出力")
        self.pdf_review_button = QPushButton("PDFページを確認")
        self.pdf_review_button.setToolTip("PDFの各ページを1ページずつ確認しながら、候補を承認・却下します(PDF検査後に有効化)。")
        self.settings_toggle_button = QPushButton()
        self.settings_panel = QWidget()
        self.mode_note = QLabel("")
        self.summary_total_label = QLabel()
        self.summary_unresolved_label = QLabel()
        self.summary_target_label = QLabel()
        self.table = QTableWidget(0, 9)
        self._settings_expanded = False
        self._build_ui()

    def _build_ui(self) -> None:
        root = QWidget()
        layout = QVBoxLayout(root)
        layout.setContentsMargins(14, 14, 14, 14)
        layout.setSpacing(10)

        title = QLabel("Excel / Word / PowerPoint / PDF ファイルを選択してください。検出はこの PC 内だけで行い、原本は上書きしません。")
        title.setStyleSheet("font-size: 15px; font-weight: 600;")
        layout.addWidget(title)

        file_row = QHBoxLayout()
        file_row.addWidget(QLabel("ファイル:"))
        file_row.addWidget(self.path_label, 1)
        layout.addLayout(file_row)

        # 主な操作は「選択 → 検査 → (PDFのみ)ページ確認 → 出力」の順に左から
        # 並べ、まだ押せない段階のボタンは無効化する。どのボタンが今押せる
        # かだけで次にすることが伝わるようにするため。
        workflow_row = QHBoxLayout()
        self.choose_button.setToolTip("検査するファイルを選択します(Excel/Word/PowerPoint/PDF)。")
        self.scan_button.setToolTip("選択したファイルを検査し、個人情報・機密情報の候補を検出します。")
        self.convert_button.setToolTip("承認済みの候補を匿名化して保存します。原本のファイルは上書きしません。")
        self.choose_button.clicked.connect(self.choose_file)
        self.scan_button.clicked.connect(self.scan_file)
        self.convert_button.clicked.connect(self.convert_file)
        self.scan_button.setEnabled(False)
        self.convert_button.setEnabled(False)
        self.convert_button.setStyleSheet(PRIMARY_BUTTON_STYLE)
        self.pdf_review_button.clicked.connect(self.open_pdf_review)
        self.pdf_review_button.setVisible(False)
        workflow_row.addWidget(self.choose_button)
        workflow_row.addWidget(self.scan_button)
        workflow_row.addWidget(self.pdf_review_button)
        workflow_row.addWidget(self.convert_button)
        workflow_row.addStretch(1)
        layout.addLayout(workflow_row)

        action_row = QHBoxLayout()
        all_button = QPushButton("すべて変換")
        none_button = QPushButton("すべて除外")
        export_csv_button = QPushButton("検出結果CSV出力")
        all_button.setToolTip("検出された全ての候補の「変換する」をチェックします。")
        none_button.setToolTip("検出された全ての候補の「変換する」を解除します。")
        export_csv_button.setToolTip("現在の検出結果の一覧をCSVファイルとして保存します。")
        all_button.clicked.connect(lambda: self.set_all_enabled(True))
        none_button.clicked.connect(lambda: self.set_all_enabled(False))
        export_csv_button.clicked.connect(self.export_findings_csv)
        action_row.addWidget(all_button)
        action_row.addWidget(none_button)
        action_row.addWidget(export_csv_button)
        action_row.addStretch(1)
        layout.addLayout(action_row)

        summary_row = QHBoxLayout()
        for label, style in (
            (self.summary_total_label, "background: #f8fafc; border: 1px solid #cbd5e1;"),
            (self.summary_unresolved_label, "background: #fef3c7; border: 1px solid #f59e0b; color: #92400e;"),
            (self.summary_target_label, "background: #f8fafc; border: 1px solid #cbd5e1;"),
        ):
            label.setStyleSheet(f"{style} padding: 6px 10px; border-radius: 4px;")
            summary_row.addWidget(label)
        summary_row.addStretch(1)
        layout.addLayout(summary_row)
        self._refresh_summary_counts()

        # 処理設定(処理モード・仮名化範囲・企業機密・PDF匿名化方法)は普段は
        # 折りたたみ、現在の設定を1行要約したボタンだけを表示する。各項目の
        # 説明はツールチップに譲る。
        self.settings_toggle_button.setFlat(True)
        self.settings_toggle_button.setStyleSheet("QPushButton { text-align: left; color: #475569; } QPushButton:hover { color: #1e293b; }")
        self.settings_toggle_button.clicked.connect(self._toggle_settings_panel)
        layout.addWidget(self.settings_toggle_button)

        settings_layout = QHBoxLayout(self.settings_panel)
        settings_layout.setContentsMargins(0, 0, 0, 0)
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
        self.mode_combo.setToolTip(
            "分析継続用: 金額・数量・原価・評価などの分析項目は維持したまま、氏名や会社名などの識別情報だけを変換します。"
            "社内での分析継続を想定した設定です。\n\n"
            "外部共有用: 既存の匿名化方針に近い形で、企業機密にあたる項目(金額・数量など)も含めて変換対象にします。"
            "社外へ提出・共有する場合に選びます。"
        )
        self.scope_combo.setToolTip(
            "同じ人物・会社名などに、常に同じ仮名(個人001、法人001など)を割り当てる範囲を選びます。\n\n"
            "このファイル内だけ: 今回のファイル1件の中でだけ仮名を統一します。\n"
            "今回アップロードした一連のファイル内: 同じ操作で選んだ複数ファイルの間でも仮名を統一します。\n"
            "プロジェクト内: さらに広い範囲(プロジェクト単位)で仮名を統一します。"
        )
        self.business_secret_checkbox.setToolTip(
            "分析継続用モードでも、企業機密情報(金額・数量・原価・評価など)を追加で変換対象にします。"
            "外部共有用モードでは常にオンとして扱われます。"
        )
        self.pdf_redaction_combo.setToolTip(
            "PDFの匿名化箇所をどのように置き換えるかを選びます。仮名化(個人001などに置き換え)・部分マスキング・黒塗り・"
            "白塗り・完全削除から選べます。PDFファイルを選んでいるときだけ有効です。"
        )
        settings_layout.addWidget(QLabel("処理モード:"))
        settings_layout.addWidget(self.mode_combo)
        settings_layout.addWidget(QLabel("仮名化範囲:"))
        settings_layout.addWidget(self.scope_combo)
        settings_layout.addWidget(self.business_secret_checkbox)
        settings_layout.addWidget(QLabel("PDF匿名化方法:"))
        settings_layout.addWidget(self.pdf_redaction_combo)
        settings_layout.addStretch(1)
        self.settings_panel.setVisible(False)
        layout.addWidget(self.settings_panel)
        self.update_mode_note()

        # PDFの支援機能である旨の注意書きは、PDFを選んでいるときだけ表示する
        # (常設の説明文は上の設定要約とツールチップに譲った)。
        self.mode_note.setWordWrap(True)
        self.mode_note.setStyleSheet("border: 1px solid #f59e0b; padding: 6px; background: #fffbeb; color: #92400e;")
        self.mode_note.setVisible(False)
        layout.addWidget(self.mode_note)

        headers = ["変換する", "変換しない", "シート", "セル", "種類", "検査", "検出値", "変換後", "理由"]
        self.table.setHorizontalHeaderLabels(headers)
        self.table.horizontalHeaderItem(0).setToolTip("チェックした候補を匿名化の対象にします。")
        self.table.horizontalHeaderItem(1).setToolTip("チェックした候補を、確認済みのうえで原文のまま維持します(Word/Excel/PowerPoint)。")
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
        layout.addWidget(self.table, 1)

        self.status_label.setStyleSheet("border: 1px solid #cbd5e1; padding: 5px; background: #f8fafc;")
        layout.addWidget(self.status_label)

        self.setCentralWidget(root)

    def choose_file(self) -> None:
        filename, _ = QFileDialog.getOpenFileName(
            self,
            "検査するファイルを選択",
            "",
            "Supported files (*.xlsx *.xlsm *.docx *.pptx *.pdf);;Excel files (*.xlsx *.xlsm);;Word files (*.docx);;PowerPoint files (*.pptx);;PDF files (*.pdf)",
        )
        if filename:
            self.set_source(Path(filename))

    def set_source(self, path: Path) -> None:
        if path.suffix.lower() not in SUPPORTED_EXTENSIONS:
            QMessageBox.warning(self, "形式エラー", "対応形式は .xlsx / .xlsm / .docx / .pptx / .pdf です。")
            return
        self.processor.cleanup()
        if path.suffix.lower() in PDF_EXTENSIONS:
            self.processor = PdfPrivacyProcessor()
        elif path.suffix.lower() in WORD_EXTENSIONS:
            self.processor = WordPrivacyProcessor()
        elif path.suffix.lower() in PPTX_EXTENSIONS:
            self.processor = PptxPrivacyProcessor()
        else:
            self.processor = ExcelPrivacyProcessor()
        self.source_path = path
        self.findings = []
        self.word_decisions = []
        self.pptx_decisions = []
        self.path_label.setText(str(path))
        self.scan_button.setEnabled(True)
        self.convert_button.setEnabled(False)
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
            elif self.is_pptx_source() and isinstance(self.processor, PptxPrivacyProcessor):
                self.pptx_decisions = self.processor.scan(self.source_path, options=options)
                self.findings = [_finding_from_pptx_decision(decision) for decision in self.pptx_decisions]
            else:
                self.findings = self.processor.scan(self.source_path, options=options)
            restored_note = ""
            if self.is_pdf_source() and isinstance(self.processor, PdfPrivacyProcessor):
                restored_note = self._restore_pdf_review_state()
            self.refresh_table()
            self.convert_button.setEnabled(bool(self.findings))
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
                    f"{restored_note}"
                )
            elif self.is_word_source():
                review_required_count = sum(1 for finding in self.findings if finding.detection_kind == "要確認(未処理)")
                self.status_label.setText(
                    f"Word検査完了: {len(self.findings)} 件を検出しました。"
                    f"要確認候補 {review_required_count} 件は変換前に承認(チェック)してください。"
                )
            elif self.is_pptx_source():
                review_required_count = sum(1 for finding in self.findings if finding.detection_kind == "要確認(未処理)")
                self.status_label.setText(
                    f"PPTX検査完了: {len(self.findings)} 件を検出しました。"
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
            self.convert_button.setEnabled(bool(self.findings))
            self.update_pdf_review_button()
        finally:
            if busy_cursor:
                QApplication.restoreOverrideCursor()

    def _restore_pdf_review_state(self) -> str:
        """同じ元PDFを再検査したとき、前回このアプリで保存した承認/却下/枠の
        調整結果を自動で復元する。ファイルが無い・入力PDFが別物・壊れている
        場合は何もせず、通常の未確認状態のまま検査結果を使う。
        """
        if self.source_path is None or not isinstance(self.processor, PdfPrivacyProcessor):
            return ""
        state_path = pdf_review_state_path(self.source_path)
        if not state_path.exists():
            return ""
        try:
            self.processor.import_review_state(state_path, self.source_path, self.findings)
        except Exception:
            return ""
        return " 前回このPDFを確認したときの承認・却下・枠の状態を復元しました。"

    def _save_pdf_review_state(self) -> None:
        if self.source_path is None or not isinstance(self.processor, PdfPrivacyProcessor):
            return
        state_path = pdf_review_state_path(self.source_path)
        try:
            self.processor.export_review_state(state_path, self.source_path, self.findings)
        except Exception:
            pass

    def convert_file(self) -> None:
        if self.source_path is None:
            QMessageBox.warning(self, "ファイル未選択", "ファイルを選択してください。")
            return
        busy_cursor = False
        try:
            self.update_findings_from_table()
            options = self.current_options()

            pdf_can_output = True
            pdf_reasons: list[str] = []
            if self.is_pdf_source() and isinstance(self.processor, PdfPrivacyProcessor):
                pdf_can_output, pdf_reasons = final_output_status(
                    self.findings,
                    self.processor.page_quality,
                    self.processor.confirmed_pages,
                    self.processor.page_review_state,
                )
                if not pdf_can_output:
                    QMessageBox.warning(self, "PDF出力不可", self._pdf_output_summary(pdf_can_output, pdf_reasons))
                    self.status_label.setText("PDF出力不可")
                    return

            if QMessageBox.question(self, "出力確認", self._output_confirmation_message(options)) != QMessageBox.Yes:
                self.status_label.setText("出力をキャンセルしました。")
                return

            self.status_label.setText("変換中: 一時コピーへ置換を適用しています...")
            if self.is_word_source() and isinstance(self.processor, WordPrivacyProcessor):
                self._sync_word_decisions_from_findings()
                QApplication.setOverrideCursor(Qt.WaitCursor)
                busy_cursor = True
                QApplication.processEvents()
                result = self.processor.convert(self.source_path, self.word_decisions)
                QApplication.restoreOverrideCursor()
                busy_cursor = False
                output_path = result.output_path
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
            elif self.is_pptx_source() and isinstance(self.processor, PptxPrivacyProcessor):
                self._sync_pptx_decisions_from_findings()
                QApplication.setOverrideCursor(Qt.WaitCursor)
                busy_cursor = True
                QApplication.processEvents()
                result = self.processor.convert(self.source_path, self.pptx_decisions)
                QApplication.restoreOverrideCursor()
                busy_cursor = False
                output_path = result.output_path
                self.status_label.setText(f"保存完了: {output_path}")
                warning_note = ("\n\n警告:\n" + "\n".join(result.warnings)) if result.warnings else ""
                QMessageBox.information(
                    self,
                    "保存完了",
                    "匿名化済み PPTX、検出・変換結果CSV、処理報告書を保存しました。\n\n"
                    f"PPTX: {result.output_path}\nCSV: {result.csv_path}\n報告書: {result.report_path}"
                    f"{warning_note}\n\n"
                    "原本は上書きしていません。一時コピーは削除済みです。",
                )
            elif self.is_pdf_source() and isinstance(self.processor, PdfPrivacyProcessor):
                # pdf_can_output was already confirmed further up, before the
                # shared 出力確認 dialog -- asking the user to confirm output
                # settings for a conversion that can't even run yet would be
                # backwards.
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
        elif self.is_pptx_source():
            self._sync_pptx_decisions_from_findings()
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
            elif self.is_pptx_source():
                write_pptx_findings_csv(path, self.pptx_decisions)
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
        self._save_pdf_review_state()
        self.refresh_table()
        self.update_pdf_review_button()
        self.convert_button.setEnabled(bool(self.findings))
        self.status_label.setText("PDF候補確認を反映しました。全ページが確認済みになるまで最終出力できません。")

    def _is_unresolved_row(self, finding: Finding, is_word: bool, is_excel: bool, is_pptx: bool = False) -> bool:
        if is_word or is_pptx:
            return finding.detection_kind == "要確認(未処理)"
        if is_excel:
            return (
                finding.detection_kind in {"確認候補", EXCEL_NLP_DETECTION_KIND}
                and not finding.enabled
                and not finding.excluded
            )
        if isinstance(self.processor, PdfPrivacyProcessor):
            return finding.detection_kind in {CANDIDATE_REVIEW, "確認候補"}
        return False

    def refresh_table(self) -> None:
        self.table.blockSignals(True)
        self.table.setRowCount(0)
        is_word = self.is_word_source()
        is_pptx = self.is_pptx_source()
        is_excel = isinstance(self.processor, ExcelPrivacyProcessor)
        # Word/PPTX candidate categories are fixed by detection and can't be
        # edited after the fact (WordCandidate/PptxCandidate are immutable),
        # unlike Excel/PDF's entity_type, so column 4 stays read-only for
        # their rows.
        editable_offsets = {7} if (is_word or is_pptx) else {4, 7}
        for row, finding in enumerate(self.findings):
            self.table.insertRow(row)
            is_unresolved = self._is_unresolved_row(finding, is_word, is_excel, is_pptx)

            enabled_cell = _CheckboxCell()
            enabled_cell.checkbox.setChecked(finding.enabled)
            enabled_cell.checkbox.toggled.connect(lambda _checked, r=row: self._on_checkbox_toggled(r, 0))
            self._style_checkbox_cell(enabled_cell, is_unresolved)
            self.table.setCellWidget(row, 0, enabled_cell)

            # "変換しない" (reviewed-and-excluded) is a Word/PPTX/Excel concept
            # -- PDF has its own separate page-by-page review dialog and no
            # third state here, so the checkbox stays absent (not just
            # unchecked) for its rows.
            excluded_cell = _CheckboxCell()
            if is_word or is_pptx or is_excel:
                if is_word:
                    is_excluded = row < len(self.word_decisions) and self.word_decisions[row].excluded
                elif is_pptx:
                    is_excluded = row < len(self.pptx_decisions) and self.pptx_decisions[row].excluded
                else:
                    is_excluded = finding.excluded
                excluded_cell.checkbox.setChecked(is_excluded)
                excluded_cell.checkbox.toggled.connect(lambda _checked, r=row: self._on_checkbox_toggled(r, 1))
            else:
                excluded_cell.checkbox.setEnabled(False)
                excluded_cell.checkbox.setVisible(False)
            self._style_checkbox_cell(excluded_cell, is_unresolved)
            self.table.setCellWidget(row, 1, excluded_cell)

            values = [
                finding.sheet,
                finding.cell,
                finding.entity_type,
                display_detection_kind(finding.detection_kind),
                finding.original,
                finding.replacement,
                finding.reason,
            ]
            for offset, value in enumerate(values, start=2):
                item = QTableWidgetItem(value)
                if offset not in editable_offsets:
                    item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                if is_unresolved:
                    item.setBackground(UNRESOLVED_ROW_COLOR)
                self.table.setItem(row, offset, item)
        self.table.blockSignals(False)
        self._refresh_summary_counts()

    def _row_checkbox(self, row: int, column: int) -> QCheckBox | None:
        widget = self.table.cellWidget(row, column)
        return widget.checkbox if isinstance(widget, _CheckboxCell) else None

    def _style_checkbox_cell(self, cell: "_CheckboxCell", is_unresolved: bool) -> None:
        cell.setStyleSheet(f"background-color: {UNRESOLVED_ROW_COLOR.name()};" if is_unresolved else "")

    def update_findings_from_table(self) -> None:
        for row, finding in enumerate(self.findings):
            checkbox = self._row_checkbox(row, 0)
            entity_item = self.table.item(row, 4)
            replacement_item = self.table.item(row, 7)
            finding.enabled = checkbox is not None and checkbox.isChecked()
            if entity_item is not None and entity_item.text().strip():
                finding.entity_type = entity_item.text().strip()
            if replacement_item is not None and replacement_item.text().strip():
                finding.replacement = replacement_item.text().strip()

    def _on_checkbox_toggled(self, row: int, column: int) -> None:
        checkbox = self._row_checkbox(row, column)
        if checkbox is None:
            return
        other_column = 1 - column
        if checkbox.isChecked():
            other_checkbox = self._row_checkbox(row, other_column)
            if other_checkbox is not None and other_checkbox.isChecked():
                other_checkbox.blockSignals(True)
                other_checkbox.setChecked(False)
                other_checkbox.blockSignals(False)
        if self.is_word_source():
            self._refresh_word_row_status(row)
        elif self.is_pptx_source():
            self._refresh_pptx_row_status(row)
        elif isinstance(self.processor, ExcelPrivacyProcessor):
            self._refresh_excel_row_status(row)
        elif isinstance(self.processor, PdfPrivacyProcessor):
            self._refresh_pdf_row_status(row)
        self._refresh_summary_counts()

    def _refresh_pdf_row_status(self, row: int) -> None:
        if row >= len(self.findings):
            return
        checkbox = self._row_checkbox(row, 0)
        finding = self.findings[row]
        finding.enabled = checkbox is not None and checkbox.isChecked()
        # 検査欄は常に「今のレビュー状態」を表す方針: 手動追加枠と結合済みの
        # 項目以外は、メイン一覧のチェックボックス操作でも
        # USER_APPROVED/USER_REJECTED に切り替える(PDF候補確認ダイアログの
        # set_selected_enabled と同じ規則)。
        if finding.detection_kind not in {CANDIDATE_MANUAL, "MERGED"}:
            finding.detection_kind = USER_APPROVED if finding.enabled else USER_REJECTED
        self.table.blockSignals(True)
        status_item = self.table.item(row, 5)
        if status_item is not None:
            status_item.setText(display_detection_kind(finding.detection_kind))
        self._refresh_row_highlight(row, self._is_unresolved_row(finding, is_word=False, is_excel=False))
        self.table.blockSignals(False)

    def _refresh_row_highlight(self, row: int, is_unresolved: bool) -> None:
        for column in range(self.table.columnCount()):
            if column in (0, 1):
                cell_widget = self.table.cellWidget(row, column)
                if isinstance(cell_widget, _CheckboxCell):
                    self._style_checkbox_cell(cell_widget, is_unresolved)
                continue
            cell_item = self.table.item(row, column)
            if cell_item is not None:
                cell_item.setBackground(UNRESOLVED_ROW_COLOR if is_unresolved else QBrush())

    def _refresh_excel_row_status(self, row: int) -> None:
        if row >= len(self.findings):
            return
        enabled_checkbox = self._row_checkbox(row, 0)
        excluded_checkbox = self._row_checkbox(row, 1)
        finding = self.findings[row]
        finding.enabled = enabled_checkbox is not None and enabled_checkbox.isChecked()
        finding.excluded = bool(
            excluded_checkbox is not None and excluded_checkbox.isChecked() and not finding.enabled
        )
        if finding.detection_kind not in {"確認候補", EXCEL_NLP_DETECTION_KIND}:
            return
        base_reason = finding.reason.split(" / 要確認候補を利用者が確認のうえ")[0]
        if finding.excluded:
            finding.reason = base_reason + " / 要確認候補を利用者が確認のうえ除外(原文を維持)"
        elif finding.enabled:
            finding.reason = base_reason + " / 要確認候補を利用者が確認のうえ承認"
        else:
            finding.reason = base_reason
        self.table.blockSignals(True)
        reason_item = self.table.item(row, 8)
        if reason_item is not None:
            reason_item.setText(finding.reason)
        self._refresh_row_highlight(row, self._is_unresolved_row(finding, is_word=False, is_excel=True))
        self.table.blockSignals(False)

    def _refresh_word_row_status(self, row: int) -> None:
        if row >= len(self.word_decisions):
            return
        enabled_checkbox = self._row_checkbox(row, 0)
        excluded_checkbox = self._row_checkbox(row, 1)
        decision = self.word_decisions[row]
        decision.enabled = enabled_checkbox is not None and enabled_checkbox.isChecked()
        decision.excluded = bool(
            excluded_checkbox is not None and excluded_checkbox.isChecked() and not decision.enabled
        )
        status = word_finding_status(decision)
        if row < len(self.findings):
            self.findings[row].enabled = decision.enabled
            # word_finding_status() is re-derived from `decision` on every
            # call rather than stored anywhere, so keep the Finding's own
            # detection_kind (read by _is_unresolved_row for the row
            # highlight) in sync too instead of only updating the cell text.
            self.findings[row].detection_kind = status
        self.table.blockSignals(True)
        status_item = self.table.item(row, 5)
        if status_item is not None:
            status_item.setText(status)
        reason_item = self.table.item(row, 8)
        if reason_item is not None:
            reason_item.setText(word_finding_reason(decision))
        if row < len(self.findings):
            self._refresh_row_highlight(row, self._is_unresolved_row(self.findings[row], is_word=True, is_excel=False))
        self.table.blockSignals(False)

    def _refresh_pptx_row_status(self, row: int) -> None:
        if row >= len(self.pptx_decisions):
            return
        enabled_checkbox = self._row_checkbox(row, 0)
        excluded_checkbox = self._row_checkbox(row, 1)
        decision = self.pptx_decisions[row]
        decision.enabled = enabled_checkbox is not None and enabled_checkbox.isChecked()
        decision.excluded = bool(
            excluded_checkbox is not None and excluded_checkbox.isChecked() and not decision.enabled
        )
        status = pptx_finding_status(decision)
        if row < len(self.findings):
            self.findings[row].enabled = decision.enabled
            # pptx_finding_status() is re-derived from `decision` on every
            # call rather than stored anywhere, so keep the Finding's own
            # detection_kind (read by _is_unresolved_row for the row
            # highlight) in sync too instead of only updating the cell text.
            self.findings[row].detection_kind = status
        self.table.blockSignals(True)
        status_item = self.table.item(row, 5)
        if status_item is not None:
            status_item.setText(status)
        reason_item = self.table.item(row, 8)
        if reason_item is not None:
            reason_item.setText(pptx_finding_reason(decision))
        if row < len(self.findings):
            self._refresh_row_highlight(row, self._is_unresolved_row(self.findings[row], is_word=False, is_excel=False, is_pptx=True))
        self.table.blockSignals(False)

    def set_all_enabled(self, enabled: bool) -> None:
        for row in range(self.table.rowCount()):
            checkbox = self._row_checkbox(row, 0)
            if checkbox is not None:
                checkbox.setChecked(enabled)

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
        self.pdf_redaction_combo.setEnabled(self.is_pdf_source())
        # 分析継続用/外部共有用の一般的な説明はツールチップに譲り、ここでは
        # PDF選択時の支援機能に関する注意書きだけを表示する。
        self.mode_note.setText(PDF_ASSISTANCE_NOTICE)
        self.mode_note.setVisible(self.is_pdf_source())
        self.settings_toggle_button.setText(f"設定: {self._settings_summary_text()}  {'▴' if self._settings_expanded else '▾'}")
        self.update_pdf_review_button()

    def _settings_summary_text(self) -> str:
        options = self.current_options()
        parts = [
            options.mode_label,
            "企業機密も変換する" if options.transform_business_secrets else "企業機密は変換しない",
        ]
        if self.is_pdf_source():
            parts.append(PDF_REDACTION_MODES.get(str(self.pdf_redaction_combo.currentData()), ""))
        return "・".join(part for part in parts if part)

    def _toggle_settings_panel(self) -> None:
        self._settings_expanded = not self._settings_expanded
        self.settings_panel.setVisible(self._settings_expanded)
        self.update_mode_note()

    def _refresh_summary_counts(self) -> None:
        total = len(self.findings)
        target = sum(1 for finding in self.findings if finding.enabled)
        is_word = self.is_word_source()
        is_pptx = self.is_pptx_source()
        is_excel = isinstance(self.processor, ExcelPrivacyProcessor)
        unresolved = sum(
            1 for finding in self.findings if self._is_unresolved_row(finding, is_word, is_excel, is_pptx)
        )
        self.summary_total_label.setText(f"検出 {total}件")
        self.summary_unresolved_label.setText(f"要確認 {unresolved}件")
        self.summary_target_label.setText(f"変換対象 {target}件")

    def is_pdf_source(self) -> bool:
        return self.source_path is not None and self.source_path.suffix.lower() in PDF_EXTENSIONS

    def is_word_source(self) -> bool:
        return self.source_path is not None and self.source_path.suffix.lower() in WORD_EXTENSIONS

    def is_pptx_source(self) -> bool:
        return self.source_path is not None and self.source_path.suffix.lower() in PPTX_EXTENSIONS

    def _sync_word_decisions_from_findings(self) -> None:
        for row, (decision, finding) in enumerate(zip(self.word_decisions, self.findings)):
            decision.enabled = finding.enabled
            excluded_checkbox = self._row_checkbox(row, 1)
            decision.excluded = bool(
                excluded_checkbox is not None and excluded_checkbox.isChecked() and not decision.enabled
            )
            if finding.replacement.strip():
                decision.replacement = finding.replacement.strip()

    def _sync_pptx_decisions_from_findings(self) -> None:
        for row, (decision, finding) in enumerate(zip(self.pptx_decisions, self.findings)):
            decision.enabled = finding.enabled
            excluded_checkbox = self._row_checkbox(row, 1)
            decision.excluded = bool(
                excluded_checkbox is not None and excluded_checkbox.isChecked() and not decision.enabled
            )
            if finding.replacement.strip():
                decision.replacement = finding.replacement.strip()

    def update_pdf_review_button(self) -> None:
        is_pdf_scanned = self.is_pdf_source() and isinstance(self.processor, PdfPrivacyProcessor) and bool(self.processor.temp_pdf)
        self.pdf_review_button.setVisible(self.is_pdf_source())
        self.pdf_review_button.setEnabled(is_pdf_scanned)
        if not is_pdf_scanned:
            self.pdf_review_button.setText("PDFページを確認")
            self.pdf_review_button.setStyleSheet("")
            return
        remaining = self._pdf_unreviewed_page_count()
        if remaining > 0:
            self.pdf_review_button.setText(f"PDFページを確認(残り{remaining}ページ)")
            self.pdf_review_button.setStyleSheet(WARNING_BUTTON_STYLE)
        else:
            self.pdf_review_button.setText("PDFページを確認(確認済み)")
            self.pdf_review_button.setStyleSheet("")

    def _pdf_unreviewed_page_count(self) -> int:
        if not isinstance(self.processor, PdfPrivacyProcessor):
            return 0
        completed_like = {"REVIEWED_NO_SENSITIVE_DATA", "REVIEWED_WITH_REDACTIONS", "COMPLETED"}
        return sum(
            1
            for page_index in range(self.processor.page_count)
            if self.processor.page_review_state.get(page_index) not in completed_like
        )

    def _default_csv_name(self) -> str:
        folder = self.source_path.parent if self.source_path else Path.cwd()
        stem = self.source_path.stem if self.source_path else "検出結果"
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        return str(folder / f"{stem}_検出結果_{timestamp}.csv")

    def _output_confirmation_message(self, options: ProcessingOptions) -> str:
        lines = [
            "次の設定で出力します。よろしいですか？",
            "",
            f"処理モード: {options.mode_label}",
            f"企業機密も変換する: {'はい' if options.transform_business_secrets else 'いいえ'}",
        ]
        if self.is_pdf_source():
            redaction_label = PDF_REDACTION_MODES.get(str(self.pdf_redaction_combo.currentData()), "")
            lines.append(f"PDF匿名化方法: {redaction_label}")
        return "\n".join(lines)

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
