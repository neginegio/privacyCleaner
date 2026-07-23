from __future__ import annotations

from pathlib import Path

try:
    import fitz
except ImportError:  # pragma: no cover
    fitz = None  # type: ignore[assignment]

from PySide6.QtCore import QRectF, Qt
from PySide6.QtGui import QColor, QImage, QKeySequence, QPen, QPixmap
from PySide6.QtWidgets import (
    QAbstractItemView,
    QComboBox,
    QDialog,
    QDoubleSpinBox,
    QFormLayout,
    QGraphicsPixmapItem,
    QGraphicsRectItem,
    QGraphicsScene,
    QGraphicsView,
    QGridLayout,
    QGroupBox,
    QHeaderView,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QScrollArea,
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


class ZoomableGraphicsView(QGraphicsView):
    def __init__(self, scene: QGraphicsScene, parent: QWidget | None = None) -> None:
        super().__init__(scene, parent)
        self.zoom_factor = 1.0
        self.zoom_changed_callback = None
        self.setTransformationAnchor(QGraphicsView.AnchorUnderMouse)
        self.setResizeAnchor(QGraphicsView.AnchorUnderMouse)

    def wheelEvent(self, event) -> None:  # type: ignore[override]
        if event.modifiers() & Qt.ControlModifier:
            factor = 1.15 if event.angleDelta().y() > 0 else 1 / 1.15
            self.set_zoom(self.zoom_factor * factor)
            event.accept()
            return
        super().wheelEvent(event)

    def set_zoom(self, value: float) -> None:
        new_value = min(max(value, 0.25), 6.0)
        self.resetTransform()
        self.scale(new_value, new_value)
        self.zoom_factor = new_value
        if self.zoom_changed_callback:
            self.zoom_changed_callback(self.zoom_factor)

    def reset_zoom(self) -> None:
        self.set_zoom(1.0)


class ResizableRectItem(QGraphicsRectItem):
    HANDLE_MARGIN = 8.0
    MIN_SIZE = 6.0

    def __init__(self, x: float, y: float, width: float, height: float) -> None:
        super().__init__(x, y, width, height)
        self.resize_mode = ""
        self.resize_start_rect = QRectF()
        self.resize_start_scene_pos = None
        self.geometry_changed_callback = None

    def hoverMoveEvent(self, event) -> None:  # type: ignore[override]
        self.setCursor(self._cursor_for_mode(self._resize_mode_at(event.pos())))
        super().hoverMoveEvent(event)

    def hoverLeaveEvent(self, event) -> None:  # type: ignore[override]
        self.unsetCursor()
        super().hoverLeaveEvent(event)

    def mousePressEvent(self, event) -> None:  # type: ignore[override]
        self.resize_mode = self._resize_mode_at(event.pos())
        if self.resize_mode:
            self.resize_start_rect = QRectF(self.rect())
            self.resize_start_scene_pos = event.scenePos()
            event.accept()
            return
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event) -> None:  # type: ignore[override]
        if self.resize_mode and self.resize_start_scene_pos is not None:
            delta = event.scenePos() - self.resize_start_scene_pos
            rect = QRectF(self.resize_start_rect)
            if "l" in self.resize_mode:
                rect.setLeft(min(rect.left() + delta.x(), rect.right() - self.MIN_SIZE))
            if "r" in self.resize_mode:
                rect.setRight(max(rect.right() + delta.x(), rect.left() + self.MIN_SIZE))
            if "t" in self.resize_mode:
                rect.setTop(min(rect.top() + delta.y(), rect.bottom() - self.MIN_SIZE))
            if "b" in self.resize_mode:
                rect.setBottom(max(rect.bottom() + delta.y(), rect.top() + self.MIN_SIZE))
            self.setRect(rect)
            if self.geometry_changed_callback:
                self.geometry_changed_callback()
            event.accept()
            return
        super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event) -> None:  # type: ignore[override]
        self.resize_mode = ""
        self.resize_start_scene_pos = None
        if self.geometry_changed_callback:
            self.geometry_changed_callback()
        super().mouseReleaseEvent(event)

    def itemChange(self, change, value):  # type: ignore[override]
        result = super().itemChange(change, value)
        if change == QGraphicsRectItem.ItemPositionHasChanged and self.geometry_changed_callback:
            self.geometry_changed_callback()
        return result

    def _resize_mode_at(self, pos) -> str:
        rect = self.rect()
        margin = self.HANDLE_MARGIN
        left = abs(pos.x() - rect.left()) <= margin
        right = abs(pos.x() - rect.right()) <= margin
        top = abs(pos.y() - rect.top()) <= margin
        bottom = abs(pos.y() - rect.bottom()) <= margin
        inside_x = rect.left() - margin <= pos.x() <= rect.right() + margin
        inside_y = rect.top() - margin <= pos.y() <= rect.bottom() + margin
        if top and left:
            return "tl"
        if top and right:
            return "tr"
        if bottom and left:
            return "bl"
        if bottom and right:
            return "br"
        if left and inside_y:
            return "l"
        if right and inside_y:
            return "r"
        if top and inside_x:
            return "t"
        if bottom and inside_x:
            return "b"
        return ""

    def _cursor_for_mode(self, mode: str):
        if mode in {"tl", "br"}:
            return Qt.SizeFDiagCursor
        if mode in {"tr", "bl"}:
            return Qt.SizeBDiagCursor
        if mode in {"l", "r"}:
            return Qt.SizeHorCursor
        if mode in {"t", "b"}:
            return Qt.SizeVerCursor
        return Qt.ArrowCursor


class PdfCandidateReviewDialog(QDialog):
    def __init__(self, processor: PdfPrivacyProcessor, findings: list[Finding], parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("PDF候補確認")
        self.setWindowFlags(
            self.windowFlags()
            | Qt.WindowMaximizeButtonHint
            | Qt.WindowMinimizeButtonHint
            | Qt.WindowCloseButtonHint
        )
        self.setSizeGripEnabled(True)
        self.resize(1420, 900)
        self.processor = processor
        self.findings = findings
        self.items: dict[QGraphicsRectItem, Finding] = {}
        self.candidate_rows: list[Finding] = []
        self._syncing_selection = False
        self.page_index = 0
        self.scale = 1.0

        self.scene = QGraphicsScene(self)
        self.view = ZoomableGraphicsView(self.scene)
        self.view.setDragMode(QGraphicsView.RubberBandDrag)
        self.view.setRenderHints(self.view.renderHints())
        self.zoom_label = QLabel("100%")
        self.zoom_label.setMinimumWidth(48)
        reset_zoom_button = QPushButton("100%")
        reset_zoom_button.setToolTip("PDF表示倍率を100%に戻します。Ctrl+マウスホイールでも拡大縮小できます。")
        reset_zoom_button.clicked.connect(self.view.reset_zoom)
        self.view.zoom_changed_callback = self._update_zoom_label

        self.page_combo = QComboBox()
        for index in range(max(self.processor.page_count, 1)):
            quality = self.processor.page_quality.get(index)
            suffix = f" - {quality.verdict}" if quality else ""
            self.page_combo.addItem(f"ページ{index + 1}{suffix}", index)
        self.page_combo.currentIndexChanged.connect(self.reload_page)

        self.page_spin = QSpinBox()
        self.page_spin.setRange(1, max(self.processor.page_count, 1))
        self.page_spin.valueChanged.connect(lambda value: self.page_combo.setCurrentIndex(max(value - 1, 0)))

        self.selected_label = QLabel("選択中: なし")
        self.selected_label.setWordWrap(False)
        self.selected_label.setStyleSheet("font-weight: 600; color: #1f2937;")

        self.candidate_table = QTableWidget(0, 5)
        self.candidate_table.setHorizontalHeaderLabels(["状態", "種類", "変換後", "理由", "候補文字"])
        self.candidate_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.candidate_table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.candidate_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.candidate_table.horizontalHeader().setStretchLastSection(True)
        self.candidate_table.itemSelectionChanged.connect(self.select_candidate_row)

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
        delete_manual_button = QPushButton("手動枠を削除")
        merge_button = QPushButton("複数枠を結合")
        no_sensitive_button = QPushButton("このページに匿名化対象無し(Ctrl+D)")
        checked_button = QPushButton("確認完了")
        next_unresolved_button = QPushButton("次の未確認候補")
        next_unreviewed_button = QPushButton("次の未確認ページ")
        previous_unreviewed_button = QPushButton("前の未確認ページ")
        next_review_button = QPushButton("次のREVIEWページ")
        next_failed_button = QPushButton("次のFAILEDページ")
        close_button = QPushButton("閉じる")

        approve_button.clicked.connect(lambda: self.set_selected_enabled(True, advance_after=True))
        reject_button.clicked.connect(lambda: self.set_selected_enabled(False))
        apply_button.clicked.connect(self.apply_selected_edits)
        add_button.clicked.connect(self.add_manual_item)
        delete_manual_button.clicked.connect(self.delete_selected_manual_items)
        merge_button.clicked.connect(self.merge_selected_items)
        no_sensitive_button.clicked.connect(self.mark_page_no_sensitive)
        no_sensitive_button.setShortcut(QKeySequence("Ctrl+D"))
        no_sensitive_button.setToolTip("このページに匿名化対象がないものとして確認します。ショートカット: Ctrl+D")
        checked_button.clicked.connect(self.mark_page_checked)
        next_unresolved_button.clicked.connect(self.goto_next_unresolved)
        next_unreviewed_button.clicked.connect(lambda: self.goto_unreviewed_page(forward=True))
        previous_unreviewed_button.clicked.connect(lambda: self.goto_unreviewed_page(forward=False))
        next_review_button.clicked.connect(lambda: self.goto_next_quality(QUALITY_REVIEW))
        next_failed_button.clicked.connect(lambda: self.goto_next_quality(QUALITY_FAILED))
        close_button.clicked.connect(self.accept)
        self.scene.selectionChanged.connect(self.on_scene_selection_changed)

        page_nav = QHBoxLayout()
        page_nav.addWidget(QLabel("ページ"))
        page_nav.addWidget(self.page_combo, 1)
        page_nav.addWidget(self.page_spin)
        page_nav.addWidget(previous_unreviewed_button)
        page_nav.addWidget(next_unreviewed_button)
        page_nav.addWidget(next_unresolved_button)
        page_nav.addWidget(next_review_button)
        page_nav.addWidget(next_failed_button)
        page_nav.addWidget(QLabel("倍率"))
        page_nav.addWidget(self.zoom_label)
        page_nav.addWidget(reset_zoom_button)

        left = QVBoxLayout()
        left.addLayout(page_nav)
        left.addWidget(self.view, 1)
        legend = QLabel("緑: 自動候補 / 黄: 要確認 / 赤: FAILED系 / 青: 手動追加")
        legend.setStyleSheet("color: #4b5563;")
        left.addWidget(legend)
        left_widget = QWidget()
        left_widget.setLayout(left)

        right = QVBoxLayout()
        right.setSpacing(10)

        candidate_group = QGroupBox("現在ページの候補")
        candidate_group.setMinimumHeight(170)
        candidate_group.setMaximumHeight(240)
        candidate_layout = QVBoxLayout()
        candidate_layout.addWidget(self.candidate_table)
        candidate_group.setLayout(candidate_layout)
        right.addWidget(candidate_group)

        form = QFormLayout()
        form.addRow(self.selected_label)
        form.addRow("種類", self.entity_combo)
        form.addRow("変換後", self.replacement_edit)
        position_widget = QWidget()
        position_layout = QHBoxLayout(position_widget)
        position_layout.setContentsMargins(0, 0, 0, 0)
        position_layout.addWidget(QLabel("X"))
        position_layout.addWidget(self.x_spin)
        position_layout.addWidget(QLabel("Y"))
        position_layout.addWidget(self.y_spin)
        form.addRow("位置", position_widget)
        size_widget = QWidget()
        size_layout = QHBoxLayout(size_widget)
        size_layout.setContentsMargins(0, 0, 0, 0)
        size_layout.addWidget(QLabel("幅"))
        size_layout.addWidget(self.w_spin)
        size_layout.addWidget(QLabel("高さ"))
        size_layout.addWidget(self.h_spin)
        form.addRow("サイズ", size_widget)
        edit_group = QGroupBox("選択中の候補・枠")
        edit_group.setLayout(form)
        right.addWidget(edit_group)

        action_group = QGroupBox("候補操作")
        action_grid = QGridLayout()
        action_grid.addWidget(approve_button, 0, 0)
        action_grid.addWidget(reject_button, 0, 1)
        action_grid.addWidget(apply_button, 1, 0, 1, 2)
        action_grid.addWidget(add_button, 2, 0)
        action_grid.addWidget(delete_manual_button, 2, 1)
        action_grid.addWidget(merge_button, 3, 0, 1, 2)
        action_group.setLayout(action_grid)
        right.addWidget(action_group)

        page_done_group = QGroupBox("ページ確認")
        page_done_layout = QVBoxLayout()
        page_done_layout.addWidget(no_sensitive_button)
        page_done_layout.addWidget(checked_button)
        page_done_group.setLayout(page_done_layout)
        right.addWidget(page_done_group)
        right.addStretch(1)
        right.addWidget(close_button)

        right_widget = QWidget()
        right_widget.setMinimumWidth(420)
        right_widget.setLayout(right)
        right_scroll = QScrollArea()
        right_scroll.setWidgetResizable(True)
        right_scroll.setFrameShape(QScrollArea.NoFrame)
        right_scroll.setWidget(right_widget)

        splitter = QSplitter(Qt.Horizontal)
        splitter.addWidget(left_widget)
        splitter.addWidget(right_scroll)
        splitter.setStretchFactor(0, 1)
        splitter.setStretchFactor(1, 0)
        splitter.setSizes([1040, 440])

        layout = QVBoxLayout(self)
        notice = QLabel(
            "PDF OCR匿名化は全ページ確認前提の支援機能です。"
            "未確認ページ、未確認候補、FAILED未対応ページが残っている場合は最終出力できません。"
        )
        notice.setWordWrap(True)
        notice.setStyleSheet("border: 1px solid #fde68a; padding: 6px; background: #fffbeb; color: #713f12;")
        layout.addWidget(notice)
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
        self.refresh_candidate_table()

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

    def set_selected_enabled(self, enabled: bool, advance_after: bool = False) -> None:
        items = self._selected_rect_items()
        current_finding = self.items[items[0]] if items else None
        for item in items:
            finding = self.items[item]
            finding.enabled = enabled
            if finding.detection_kind in {CANDIDATE_REVIEW, USER_APPROVED, USER_REJECTED}:
                finding.detection_kind = USER_APPROVED if enabled else USER_REJECTED
            self._style_item(item, finding)
        self.sync_items_to_locations()
        self.refresh_page_table()
        self.refresh_candidate_table()
        if advance_after and current_finding is not None:
            self.select_next_candidate_after(current_finding)
        else:
            self.load_selected_item()

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
        self.refresh_candidate_table()
        self.load_selected_item()

    def add_manual_item(self) -> None:
        page_width = self._page_rect_width(self.page_index)
        rect = (page_width * 0.18, 120.0, page_width * 0.55, 160.0)
        finding = self.processor.add_manual_redaction(self.findings, self.page_index, rect)
        item = self._add_rect_item(finding, self.processor.locations[(finding.sheet, finding.cell, finding.original)])
        self.scene.clearSelection()
        item.setSelected(True)
        self.load_selected_item()
        self.refresh_page_table()
        self.refresh_candidate_table()

    def delete_selected_manual_items(self) -> None:
        items = self._selected_rect_items()
        if not items:
            QMessageBox.information(self, "手動枠の削除", "削除する手動追加枠を選択してください。")
            return
        non_manual = [self.items[item] for item in items if self.items[item].detection_kind != CANDIDATE_MANUAL]
        if non_manual:
            QMessageBox.information(self, "手動枠の削除", "自動検出候補は削除できません。不要な候補は「解除」を押してください。")
            return
        if QMessageBox.question(self, "手動枠の削除", "選択中の手動追加枠を削除しますか？") != QMessageBox.Yes:
            return
        for item in items:
            finding = self.items.pop(item)
            key = (finding.sheet, finding.cell, finding.original)
            self.processor.locations.pop(key, None)
            if finding in self.findings:
                self.findings.remove(finding)
            self.scene.removeItem(item)
        self.refresh_page_table()
        self.refresh_candidate_table()
        self.load_selected_item()

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
        self.refresh_candidate_table()

    def mark_page_checked(self) -> None:
        self.sync_items_to_locations()
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
        self.goto_unreviewed_page(forward=True)

    def mark_page_no_sensitive(self) -> None:
        self.sync_items_to_locations()
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
        self.refresh_candidate_table()
        self._load_quality_label()
        self.goto_unreviewed_page(forward=True)

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
        return

    def refresh_candidate_table(self) -> None:
        selected_finding = self._selected_finding()
        self._syncing_selection = True
        self.candidate_table.blockSignals(True)
        try:
            self.candidate_rows = [finding for finding in self.findings if _page_index(finding.sheet) == self.page_index]
            self.candidate_table.setRowCount(0)
            for row, finding in enumerate(self.candidate_rows):
                self.candidate_table.insertRow(row)
                state = "有効" if finding.enabled else "解除"
                values = [
                    state,
                    finding.entity_type,
                    finding.replacement,
                    finding.reason,
                    _compact_text(finding.original),
                ]
                for col, value in enumerate(values):
                    item = QTableWidgetItem(value)
                    item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                    self.candidate_table.setItem(row, col, item)
                if finding is selected_finding:
                    self.candidate_table.selectRow(row)
        finally:
            self.candidate_table.blockSignals(False)
            self._syncing_selection = False

    def select_candidate_row(self) -> None:
        if self._syncing_selection:
            return
        rows = self.candidate_table.selectionModel().selectedRows() if self.candidate_table.selectionModel() else []
        if not rows:
            return
        row = rows[0].row()
        if row >= len(self.candidate_rows):
            return
        finding = self.candidate_rows[row]
        self.select_finding_item(finding)

    def select_next_candidate_after(self, current_finding: Finding) -> None:
        if current_finding in self.candidate_rows:
            start = self.candidate_rows.index(current_finding) + 1
            for finding in self.candidate_rows[start:]:
                self.select_finding_item(finding)
                return
        self.select_finding_item(current_finding)

    def select_finding_item(self, finding: Finding) -> None:
        for item, item_finding in self.items.items():
            if item_finding is finding:
                self._syncing_selection = True
                try:
                    self.scene.clearSelection()
                    item.setSelected(True)
                    self.view.centerOn(item)
                finally:
                    self._syncing_selection = False
                self.load_selected_item()
                return

    def load_selected_item(self) -> None:
        items = self._selected_rect_items()
        if not items:
            self.selected_label.setText("選択中: なし")
            self.selected_label.setToolTip("")
            return
        item = items[0]
        finding = self.items[item]
        self.selected_label.setText(f"選択中: {_compact_text(finding.original, 30)}")
        self.selected_label.setToolTip(str(finding.original))
        index = self.entity_combo.findText(finding.entity_type)
        if index >= 0:
            self.entity_combo.setCurrentIndex(index)
        self.replacement_edit.setText(finding.replacement)
        rect = item.mapRectToScene(item.rect())
        self.x_spin.setValue(rect.x())
        self.y_spin.setValue(rect.y())
        self.w_spin.setValue(rect.width())
        self.h_spin.setValue(rect.height())
        self.refresh_candidate_table()

    def on_scene_selection_changed(self) -> None:
        for item, finding in list(self.items.items()):
            self._style_item(item, finding)
        self.load_selected_item()

    def _add_rect_item(self, finding: Finding, location: PdfLocation) -> QGraphicsRectItem:
        x0, y0, x1, y1 = location.rect
        item = ResizableRectItem(x0 * self.scale, y0 * self.scale, (x1 - x0) * self.scale, (y1 - y0) * self.scale)
        item.setFlags(QGraphicsRectItem.ItemIsMovable | QGraphicsRectItem.ItemIsSelectable | QGraphicsRectItem.ItemSendsGeometryChanges)
        item.setAcceptHoverEvents(True)
        item.geometry_changed_callback = self.load_selected_item
        item.setToolTip(f"{finding.entity_type}: {_compact_text(finding.original, 60)}")
        self._style_item(item, finding)
        self.scene.addItem(item)
        self.items[item] = finding
        return item

    def _style_item(self, item: QGraphicsRectItem, finding: Finding) -> None:
        if item.isSelected():
            color = QColor(37, 99, 235)
            pen = QPen(color, 5)
            item.setZValue(20)
            item.setPen(pen)
            item.setBrush(Qt.NoBrush)
            return
        if finding.detection_kind == CANDIDATE_AUTO:
            color = QColor(20, 160, 80)
        elif finding.detection_kind == CANDIDATE_MANUAL:
            color = QColor(30, 100, 220)
        elif finding.detection_kind in {CANDIDATE_REVIEW, USER_APPROVED, USER_REJECTED}:
            color = QColor(230, 180, 20)
        else:
            color = QColor(180, 70, 70)
        pen = QPen(color, 3 if finding.enabled else 1.5)
        item.setZValue(5 if finding.enabled else 1)
        item.setPen(pen)
        item.setBrush(Qt.NoBrush)

    def _selected_rect_items(self) -> list[QGraphicsRectItem]:
        return [item for item in self.scene.selectedItems() if isinstance(item, QGraphicsRectItem) and item in self.items]

    def _selected_finding(self) -> Finding | None:
        items = self._selected_rect_items()
        return self.items[items[0]] if items else None

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
        spin.setMaximumWidth(92)
        return spin

    def _spin_rect(self):
        return self.x_spin.value(), self.y_spin.value(), max(self.w_spin.value(), 1), max(self.h_spin.value(), 1)

    def _update_zoom_label(self, zoom_factor: float) -> None:
        self.zoom_label.setText(f"{int(round(zoom_factor * 100))}%")

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
        # Detailed page state is shown in the page list to avoid duplicating the same
        # information in the crowded right-hand review panel.
        return


def _page_index(label: str) -> int:
    import re

    match = re.search(r"(\d+)", label)
    return int(match.group(1)) - 1 if match else 0


def _compact_text(text: str, limit: int = 36) -> str:
    compact = " ".join(str(text).split())
    if len(compact) <= limit:
        return compact
    return compact[: max(limit - 1, 1)] + "..."
