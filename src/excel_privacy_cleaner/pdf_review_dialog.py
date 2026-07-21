from __future__ import annotations

from pathlib import Path

try:
    import fitz
except ImportError:  # pragma: no cover
    fitz = None  # type: ignore[assignment]

from PySide6.QtCore import Qt
from PySide6.QtGui import QColor, QImage, QPen, QPixmap
from PySide6.QtWidgets import (
    QCheckBox,
    QComboBox,
    QDialog,
    QDoubleSpinBox,
    QFormLayout,
    QGraphicsPixmapItem,
    QGraphicsRectItem,
    QGraphicsScene,
    QGraphicsView,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QSplitter,
    QSpinBox,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from .models import Finding
from .pdf_ocr_support import (
    CANDIDATE_AUTO,
    CANDIDATE_MANUAL,
    CANDIDATE_REVIEW,
    PAGE_COMPLETED,
    PAGE_REVIEWED_NO_SENSITIVE_DATA,
    PAGE_REVIEWED_WITH_REDACTIONS,
    QUALITY_FAILED,
    QUALITY_REVIEW,
    USER_APPROVED,
    USER_REJECTED,
)
from .pdf_processor import PdfLocation, PdfPrivacyProcessor


class PdfCandidateReviewDialog(QDialog):
    def __init__(self, processor: PdfPrivacyProcessor, findings: list[Finding], parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("PDF候補確認")
        self.resize(1180, 780)
        self.processor = processor
        self.findings = findings
        self.items: dict[QGraphicsRectItem, Finding] = {}
        self.page_index = 0
        self.scale = 1.0

        self.scene = QGraphicsScene(self)
        self.view = QGraphicsView(self.scene)
        self.view.setDragMode(QGraphicsView.RubberBandDrag)
        self.view.setRenderHints(self.view.renderHints())

        self.page_combo = QComboBox()
        for index in range(max(self.processor.page_count, 1)):
            quality = self.processor.page_quality.get(index)
            suffix = f" - {quality.verdict}" if quality else ""
            self.page_combo.addItem(f"ページ{index + 1}{suffix}", index)
        self.page_combo.currentIndexChanged.connect(self.reload_page)

        self.page_spin = QSpinBox()
        self.page_spin.setRange(1, max(self.processor.page_count, 1))
        self.page_spin.valueChanged.connect(lambda value: self.page_combo.setCurrentIndex(max(value - 1, 0)))
        self.hide_reviewed_checkbox = QCheckBox("確認済みページを非表示")
        self.hide_reviewed_checkbox.stateChanged.connect(self.refresh_page_table)

        self.page_table = QTableWidget(0, 10)
        self.page_table.setHorizontalHeaderLabels(["頁", "種別", "品質", "状態", "候補", "自動", "未確認", "手動", "検証", "警告"])
        self.page_table.setMaximumHeight(170)
        self.page_table.itemSelectionChanged.connect(self.goto_selected_page_row)

        self.quality_label = QLabel("")
        self.quality_label.setWordWrap(True)

        self.entity_combo = QComboBox()
        self.entity_combo.addItems(["氏名", "氏名カナ", "会社名", "住所", "電話番号", "メールアドレス", "銀行口座", "APIキー", "パスワード", "手動追加"])
        self.replacement_edit = QLineEdit()

        self.x_spin = self._spinbox()
        self.y_spin = self._spinbox()
        self.w_spin = self._spinbox()
        self.h_spin = self._spinbox()

        approve_button = QPushButton("承認")
        reject_button = QPushButton("解除")
        apply_button = QPushButton("種類/変換/枠を反映")
        add_button = QPushButton("新しい匿名化範囲")
        merge_button = QPushButton("複数枠を結合")
        no_sensitive_button = QPushButton("このページに匿名化対象なし")
        checked_button = QPushButton("確認完了")
        next_unresolved_button = QPushButton("次の未確認候補")
        next_unreviewed_button = QPushButton("次の未確認ページ")
        previous_unreviewed_button = QPushButton("前の未確認ページ")
        next_review_button = QPushButton("次のREVIEWページ")
        next_failed_button = QPushButton("次のFAILEDページ")
        close_button = QPushButton("閉じる")

        approve_button.clicked.connect(lambda: self.set_selected_enabled(True))
        reject_button.clicked.connect(lambda: self.set_selected_enabled(False))
        apply_button.clicked.connect(self.apply_selected_edits)
        add_button.clicked.connect(self.add_manual_item)
        merge_button.clicked.connect(self.merge_selected_items)
        no_sensitive_button.clicked.connect(self.mark_page_no_sensitive)
        checked_button.clicked.connect(self.mark_page_checked)
        next_unresolved_button.clicked.connect(self.goto_next_unresolved)
        next_unreviewed_button.clicked.connect(lambda: self.goto_unreviewed_page(forward=True))
        previous_unreviewed_button.clicked.connect(lambda: self.goto_unreviewed_page(forward=False))
        next_review_button.clicked.connect(lambda: self.goto_next_quality(QUALITY_REVIEW))
        next_failed_button.clicked.connect(lambda: self.goto_next_quality(QUALITY_FAILED))
        close_button.clicked.connect(self.accept)
        self.scene.selectionChanged.connect(self.load_selected_item)

        right = QVBoxLayout()
        right.addWidget(QLabel("ページ"))
        right.addWidget(self.page_combo)
        right.addWidget(self.page_spin)
        right.addWidget(self.hide_reviewed_checkbox)
        right.addWidget(self.page_table)
        right.addWidget(self.quality_label)

        form = QFormLayout()
        form.addRow("種類", self.entity_combo)
        form.addRow("変換後", self.replacement_edit)
        form.addRow("X", self.x_spin)
        form.addRow("Y", self.y_spin)
        form.addRow("幅", self.w_spin)
        form.addRow("高さ", self.h_spin)
        right.addLayout(form)
        right.addWidget(approve_button)
        right.addWidget(reject_button)
        right.addWidget(apply_button)
        right.addWidget(add_button)
        right.addWidget(merge_button)
        nav = QHBoxLayout()
        nav.addWidget(next_unresolved_button)
        nav.addWidget(next_unreviewed_button)
        nav.addWidget(previous_unreviewed_button)
        nav.addWidget(next_review_button)
        nav.addWidget(next_failed_button)
        right.addLayout(nav)
        right.addWidget(no_sensitive_button)
        right.addWidget(checked_button)
        right.addStretch(1)
        right.addWidget(close_button)

        right_widget = QWidget()
        right_widget.setLayout(right)

        splitter = QSplitter()
        splitter.addWidget(self.view)
        splitter.addWidget(right_widget)
        splitter.setStretchFactor(0, 1)

        layout = QVBoxLayout(self)
        legend = QLabel("緑: AUTO_CONFIRMED / 黄: REVIEW_REQUIRED / 赤: FAILEDページ / 青: 手動追加")
        layout.addWidget(legend)
        layout.addWidget(splitter, 1)

        self.reload_page()
        self.refresh_page_table()

    def accept(self) -> None:
        self.sync_items_to_locations()
        super().accept()

    def reload_page(self) -> None:
        self.sync_items_to_locations()
        self.scene.clear()
        self.items = {}
        self.page_index = int(self.page_combo.currentData() or 0)
        if self.page_spin.value() != self.page_index + 1:
            self.page_spin.blockSignals(True)
            self.page_spin.setValue(self.page_index + 1)
            self.page_spin.blockSignals(False)
        self.processor.last_review_page = self.page_index
        pixmap = self._render_page(self.page_index)
        self.scale = pixmap.width() / max(self._page_rect_width(self.page_index), 1)
        self.scene.addItem(QGraphicsPixmapItem(pixmap))
        self._load_quality_label()
        for finding in self.findings:
            location = self.processor.locations.get((finding.sheet, finding.cell, finding.original))
            if location is None or location.page_index != self.page_index:
                continue
            self._add_rect_item(finding, location)
        self.scene.setSceneRect(self.scene.itemsBoundingRect())
        self.refresh_page_table()

    def sync_items_to_locations(self) -> None:
        for item, finding in list(self.items.items()):
            rect = item.mapRectToScene(item.rect())
            pdf_rect = (
                rect.x() / self.scale,
                rect.y() / self.scale,
                (rect.x() + rect.width()) / self.scale,
                (rect.y() + rect.height()) / self.scale,
            )
            self.processor.locations[(finding.sheet, finding.cell, finding.original)] = PdfLocation(self.page_index, pdf_rect)
            item.setPos(0, 0)
            item.setRect(rect)

    def set_selected_enabled(self, enabled: bool) -> None:
        for item in self._selected_rect_items():
            finding = self.items[item]
            finding.enabled = enabled
            if finding.detection_kind in {CANDIDATE_REVIEW, USER_APPROVED, USER_REJECTED}:
                finding.detection_kind = USER_APPROVED if enabled else USER_REJECTED
            self._style_item(item, finding)
        self.refresh_page_table()

    def apply_selected_edits(self) -> None:
        items = self._selected_rect_items()
        if not items:
            return
        for item in items:
            finding = self.items[item]
            finding.entity_type = self.entity_combo.currentText()
            if self.replacement_edit.text().strip():
                finding.replacement = self.replacement_edit.text().strip()
            rect = self._spin_rect()
            item.setPos(0, 0)
            item.setRect(*rect)
            self._style_item(item, finding)
        self.sync_items_to_locations()
        self.refresh_page_table()

    def add_manual_item(self) -> None:
        page_width = self._page_rect_width(self.page_index)
        rect = (page_width * 0.18, 120.0, page_width * 0.55, 160.0)
        finding = self.processor.add_manual_redaction(self.findings, self.page_index, rect)
        item = self._add_rect_item(finding, self.processor.locations[(finding.sheet, finding.cell, finding.original)])
        item.setSelected(True)
        self.load_selected_item()
        self.refresh_page_table()

    def merge_selected_items(self) -> None:
        items = self._selected_rect_items()
        if len(items) < 2:
            QMessageBox.information(self, "枠の結合", "結合する枠を2つ以上選択してください。")
            return
        union = items[0].mapRectToScene(items[0].rect())
        for item in items[1:]:
            union = union.united(item.mapRectToScene(item.rect()))
        primary = self.items[items[0]]
        primary.enabled = True
        primary.reason = f"{primary.reason} / 複数枠を結合"
        items[0].setPos(0, 0)
        items[0].setRect(union)
        for item in items[1:]:
            finding = self.items[item]
            finding.enabled = False
            finding.detection_kind = "MERGED"
            self.scene.removeItem(item)
            self.items.pop(item, None)
        self.sync_items_to_locations()
        items[0].setSelected(True)
        self.refresh_page_table()

    def mark_page_checked(self) -> None:
        unresolved = [
            finding
            for finding in self.findings
            if _page_index(finding.sheet) == self.page_index and finding.detection_kind in {"確認候補", CANDIDATE_REVIEW}
        ]
        if unresolved:
            QMessageBox.warning(self, "確認未完了", "未確認候補が残っています。承認または解除してから確認完了にしてください。")
            return
        if self._page_has_redactions(self.page_index):
            self.processor.mark_page_reviewed_with_redactions(self.page_index)
        else:
            self.processor.mark_page_no_sensitive_data(self.page_index)
        self.refresh_page_table()
        self._load_quality_label()

    def mark_page_no_sensitive(self) -> None:
        if self._page_has_redactions(self.page_index):
            if QMessageBox.question(self, "匿名化対象なし", "このページには有効な匿名化範囲があります。すべて解除して対象なしにしますか？") != QMessageBox.Yes:
                return
            for finding in self.findings:
                if _page_index(finding.sheet) == self.page_index:
                    finding.enabled = False
                    if finding.detection_kind in {CANDIDATE_REVIEW, USER_APPROVED, USER_REJECTED}:
                        finding.detection_kind = USER_REJECTED
        self.processor.mark_page_no_sensitive_data(self.page_index)
        self.refresh_page_table()
        self._load_quality_label()

    def goto_selected_page_row(self) -> None:
        rows = self.page_table.selectionModel().selectedRows() if self.page_table.selectionModel() else []
        if not rows:
            return
        row = rows[0].row()
        item = self.page_table.item(row, 0)
        if item is None:
            return
        page_index = int(item.text()) - 1
        if 0 <= page_index < self.page_combo.count():
            self.page_combo.setCurrentIndex(page_index)

    def goto_next_unresolved(self) -> None:
        current = self.page_index
        pages = sorted({_page_index(finding.sheet) for finding in self.findings if finding.detection_kind in {"確認候補", CANDIDATE_REVIEW}})
        self._goto_next_page_in(pages, current)

    def goto_next_quality(self, verdict: str) -> None:
        current = self.page_index
        pages = sorted(index for index, quality in self.processor.page_quality.items() if quality.verdict == verdict)
        self._goto_next_page_in(pages, current)

    def goto_unreviewed_page(self, forward: bool = True) -> None:
        current = self.page_index
        pages = sorted(
            index
            for index in range(max(self.processor.page_count, 1))
            if self.processor.page_review_state.get(index) not in {
                PAGE_REVIEWED_NO_SENSITIVE_DATA,
                PAGE_REVIEWED_WITH_REDACTIONS,
                PAGE_COMPLETED,
            }
        )
        if forward:
            self._goto_next_page_in(pages, current)
            return
        if not pages:
            QMessageBox.information(self, "ページ移動", "該当ページはありません。")
            return
        for page in reversed(pages):
            if page < current:
                self.page_combo.setCurrentIndex(page)
                return
        self.page_combo.setCurrentIndex(pages[-1])

    def _goto_next_page_in(self, pages: list[int], current: int) -> None:
        if not pages:
            QMessageBox.information(self, "ページ移動", "該当ページはありません。")
            return
        for page in pages:
            if page > current:
                self.page_combo.setCurrentIndex(page)
                return
        self.page_combo.setCurrentIndex(pages[0])

    def refresh_page_table(self) -> None:
        self.page_table.blockSignals(True)
        try:
            self.page_table.setRowCount(0)
            for index in range(max(self.processor.page_count, 1)):
                state = self.processor.page_review_state.get(index, "UNREVIEWED")
                if self.hide_reviewed_checkbox.isChecked() and state in {
                    PAGE_REVIEWED_NO_SENSITIVE_DATA,
                    PAGE_REVIEWED_WITH_REDACTIONS,
                    PAGE_COMPLETED,
                }:
                    continue
                row = self.page_table.rowCount()
                self.page_table.insertRow(row)
                quality = self.processor.page_quality.get(index)
                page_findings = [finding for finding in self.findings if _page_index(finding.sheet) == index]
                auto_count = sum(1 for finding in page_findings if finding.detection_kind == CANDIDATE_AUTO)
                unresolved_count = sum(1 for finding in page_findings if finding.detection_kind in {"確認候補", CANDIDATE_REVIEW})
                manual_count = sum(1 for finding in page_findings if finding.detection_kind == CANDIDATE_MANUAL)
                values = [
                    str(index + 1),
                    self.processor.page_modes.get(index, "不明"),
                    quality.verdict if quality else "未評価",
                    state,
                    str(len(page_findings)),
                    str(auto_count),
                    str(unresolved_count),
                    str(manual_count),
                    self.processor.page_verification_state.get(index, "NOT_EVALUATED"),
                    quality.warning_reason if quality else "",
                ]
                for col, value in enumerate(values):
                    item = QTableWidgetItem(value)
                    item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                    self.page_table.setItem(row, col, item)
                if index == self.page_index:
                    self.page_table.selectRow(row)
        finally:
            self.page_table.blockSignals(False)

    def load_selected_item(self) -> None:
        items = self._selected_rect_items()
        if not items:
            return
        item = items[0]
        finding = self.items[item]
        index = self.entity_combo.findText(finding.entity_type)
        if index >= 0:
            self.entity_combo.setCurrentIndex(index)
        self.replacement_edit.setText(finding.replacement)
        rect = item.mapRectToScene(item.rect())
        self.x_spin.setValue(rect.x())
        self.y_spin.setValue(rect.y())
        self.w_spin.setValue(rect.width())
        self.h_spin.setValue(rect.height())

    def _add_rect_item(self, finding: Finding, location: PdfLocation) -> QGraphicsRectItem:
        x0, y0, x1, y1 = location.rect
        item = QGraphicsRectItem(x0 * self.scale, y0 * self.scale, (x1 - x0) * self.scale, (y1 - y0) * self.scale)
        item.setFlags(QGraphicsRectItem.ItemIsMovable | QGraphicsRectItem.ItemIsSelectable | QGraphicsRectItem.ItemSendsGeometryChanges)
        self._style_item(item, finding)
        self.scene.addItem(item)
        self.items[item] = finding
        return item

    def _style_item(self, item: QGraphicsRectItem, finding: Finding) -> None:
        if finding.detection_kind == CANDIDATE_AUTO:
            color = QColor(20, 160, 80)
        elif finding.detection_kind == CANDIDATE_MANUAL:
            color = QColor(30, 100, 220)
        elif finding.detection_kind in {CANDIDATE_REVIEW, USER_APPROVED, USER_REJECTED}:
            color = QColor(230, 180, 20)
        else:
            color = QColor(180, 70, 70)
        pen = QPen(color, 3 if finding.enabled else 1.5)
        item.setPen(pen)
        item.setBrush(Qt.NoBrush)

    def _selected_rect_items(self) -> list[QGraphicsRectItem]:
        return [item for item in self.scene.selectedItems() if isinstance(item, QGraphicsRectItem) and item in self.items]

    def _page_has_redactions(self, page_index: int) -> bool:
        return any(
            finding.enabled and _page_index(finding.sheet) == page_index and finding.detection_kind != USER_REJECTED
            for finding in self.findings
        )

    def _spinbox(self) -> QDoubleSpinBox:
        spin = QDoubleSpinBox()
        spin.setRange(0, 10000)
        spin.setDecimals(1)
        spin.setSingleStep(2.0)
        return spin

    def _spin_rect(self):
        return self.x_spin.value(), self.y_spin.value(), max(self.w_spin.value(), 1), max(self.h_spin.value(), 1)

    def _render_page(self, page_index: int) -> QPixmap:
        if fitz is None or self.processor.temp_pdf is None:
            raise RuntimeError("PDFページを表示できません。")
        doc = fitz.open(self.processor.temp_pdf)
        try:
            page = doc[page_index]
            pix = page.get_pixmap(dpi=120, alpha=False)
            image = QImage(pix.samples, pix.width, pix.height, pix.stride, QImage.Format_RGB888).copy()
            return QPixmap.fromImage(image)
        finally:
            doc.close()

    def _page_rect_width(self, page_index: int) -> float:
        if fitz is None or self.processor.temp_pdf is None:
            return 1.0
        doc = fitz.open(self.processor.temp_pdf)
        try:
            return float(doc[page_index].rect.width)
        finally:
            doc.close()

    def _load_quality_label(self) -> None:
        quality = self.processor.page_quality.get(self.page_index)
        if not quality:
            self.quality_label.setText("ページ品質: 未評価")
            return
        self.quality_label.setText(
            f"ページ品質: {quality.verdict}\n"
            f"確認状態: {self.processor.page_review_state.get(self.page_index, 'UNREVIEWED')} / "
            f"検証状態: {self.processor.page_verification_state.get(self.page_index, 'NOT_EVALUATED')}\n"
            f"OCR文字数: {quality.ocr_text_chars} / 座標付き単語: {quality.coordinate_word_count} / "
            f"候補: {quality.candidate_count}\n"
            f"警告: {quality.warning_reason}"
        )
        if quality.verdict == QUALITY_FAILED:
            self.quality_label.setStyleSheet("color: #991b1b; font-weight: 600;")
        else:
            self.quality_label.setStyleSheet("")


def _page_index(label: str) -> int:
    import re

    match = re.search(r"(\d+)", label)
    return int(match.group(1)) - 1 if match else 0
