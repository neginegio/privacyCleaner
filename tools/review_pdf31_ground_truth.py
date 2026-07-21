from __future__ import annotations

import argparse
import csv
import difflib
import hashlib
import json
import os
import sys
from dataclasses import asdict, dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

import fitz  # noqa: E402
from PySide6.QtCore import Qt  # noqa: E402
from PySide6.QtGui import QColor, QImage, QPen, QPixmap  # noqa: E402
from PySide6.QtWidgets import (  # noqa: E402
    QApplication,
    QCheckBox,
    QComboBox,
    QDialog,
    QDoubleSpinBox,
    QFileDialog,
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
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from excel_privacy_cleaner.pdf_processor import ocr_page_text_and_words  # noqa: E402
from pdf31_review_fixtures import REPRESENTATIVE_PAGE_SPECS  # noqa: E402


SOURCE = Path("PDFテストデータ２.pdf")
OUTPUT_DIR = Path("ocr_quality_outputs")
STATE_JSON = OUTPUT_DIR / "PDF31代表6ページ_正解レビュー状態.json"
FINAL_GT_CSV = OUTPUT_DIR / "PDF31代表6ページ_ユーザー確認済み正解データ_v1.csv"
FINAL_STATUS_CSV = OUTPUT_DIR / "ユーザー確認状態一覧.csv"
EXCLUDED_CSV = OUTPUT_DIR / "対象外一覧.csv"
FINAL_REPORT_TXT = OUTPUT_DIR / "正解データ確定報告書.txt"

STATUS_CONFIRMED = "ユーザー確認済み"
STATUS_EXCLUDED = "対象外"
STATUS_NEEDS_COORD = "要座標修正"
STATUS_PENDING = "ユーザー確認待ち"
STATUS_RELEASE = "解除候補"
STATUS_HOLD = "保留"
FINAL_STATUSES = {STATUS_CONFIRMED, STATUS_EXCLUDED}
REVIEW_ORDER = {
    STATUS_PENDING: 0,
    STATUS_NEEDS_COORD: 1,
    STATUS_RELEASE: 2,
    STATUS_HOLD: 3,
    STATUS_EXCLUDED: 4,
    STATUS_CONFIRMED: 5,
}


@dataclass
class GroundTruthRecord:
    truth_id: str
    page: int
    occurrence: int
    same_info_id: str
    text: str
    entity_type: str
    mode: str
    reason: str
    required: str
    rect: list[float]
    replacement: str
    preserve: str
    protected_area: str
    user_status: str
    note: str
    source: str = "initial"
    updated_at: str = ""


@dataclass
class ReviewState:
    schema: str
    version: str
    source_pdf: str
    source_sha256: str
    reviewer: str
    last_truth_id: str
    records: list[GroundTruthRecord] = field(default_factory=list)


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def normalize_text(value: str) -> str:
    return "".join(value.split()).replace("（", "(").replace("）", ")")


def default_records() -> list[GroundTruthRecord]:
    return [
        GroundTruthRecord(
            truth_id=spec.truth_id,
            page=spec.page,
            occurrence=spec.occurrence,
            same_info_id=spec.same_info_id,
            text=spec.text,
            entity_type=spec.entity_type,
            mode=spec.mode,
            reason=spec.reason,
            required=spec.required,
            rect=[float(value) for value in spec.rect],
            replacement=spec.replacement,
            preserve=spec.preserve,
            protected_area=spec.protected_area,
            user_status=spec.user_status,
            note=spec.note,
        )
        for spec in REPRESENTATIVE_PAGE_SPECS
    ]


def load_state(path: Path | None = None, source_pdf: Path = SOURCE) -> ReviewState:
    path = path or STATE_JSON
    source_hash = sha256_file(source_pdf)
    if path.exists():
        payload = json.loads(path.read_text(encoding="utf-8"))
        records = [GroundTruthRecord(**row) for row in payload["records"]]
        state = ReviewState(
            schema=payload["schema"],
            version=payload.get("version", "draft"),
            source_pdf=payload["source_pdf"],
            source_sha256=payload["source_sha256"],
            reviewer=payload.get("reviewer", os.environ.get("USERNAME", "")),
            last_truth_id=payload.get("last_truth_id", ""),
            records=records,
        )
        if state.source_sha256 != source_hash:
            raise RuntimeError("入力PDFのSHA-256がレビュー状態JSONと一致しません。")
        return state
    return ReviewState(
        schema="pdf31_ground_truth_review_state_v1",
        version="draft",
        source_pdf=str(source_pdf),
        source_sha256=source_hash,
        reviewer=os.environ.get("USERNAME", ""),
        last_truth_id="",
        records=default_records(),
    )


def save_state(state: ReviewState, path: Path | None = None) -> None:
    path = path or STATE_JSON
    OUTPUT_DIR.mkdir(exist_ok=True)
    payload = asdict(state)
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def status_counts(records: list[GroundTruthRecord]) -> dict[str, int]:
    counts: dict[str, int] = {}
    for record in records:
        counts[record.user_status] = counts.get(record.user_status, 0) + 1
    return counts


def ordered_records(records: list[GroundTruthRecord], include_confirmed: bool = False) -> list[GroundTruthRecord]:
    selected = records if include_confirmed else [row for row in records if row.user_status != STATUS_CONFIRMED]
    return sorted(selected, key=lambda row: (REVIEW_ORDER.get(row.user_status, 99), row.page, row.truth_id))


def words_in_rect(words: list[tuple[Any, ...]], rect: list[float]) -> list[tuple[Any, ...]]:
    x0, y0, x1, y1 = rect
    selected: list[tuple[Any, ...]] = []
    for word in words:
        wx0, wy0, wx1, wy1 = (float(word[0]), float(word[1]), float(word[2]), float(word[3]))
        cx = (wx0 + wx1) / 2
        cy = (wy0 + wy1) / 2
        if x0 <= cx <= x1 and y0 <= cy <= y1:
            selected.append(word)
    return selected


def ocr_text_for_rect(words: list[tuple[Any, ...]], record: GroundTruthRecord) -> tuple[str, float]:
    inside = "".join(str(word[4]) for word in words_in_rect(words, record.rect))
    ratio = difflib.SequenceMatcher(None, normalize_text(record.text), normalize_text(inside)).ratio() if inside else 0.0
    return inside, ratio


def write_records_csv(path: Path, records: list[GroundTruthRecord]) -> None:
    headers = [
        "正解ID",
        "ページ番号",
        "出現番号",
        "同一情報ID",
        "匿名化対象文字列",
        "情報種別",
        "適用モード",
        "対象理由",
        "必須度",
        "正解座標",
        "期待する変換",
        "残すべき文字列",
        "匿名化してはいけない領域",
        "ユーザー確認状態",
        "備考",
    ]
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.writer(handle)
        writer.writerow(headers)
        for row in records:
            writer.writerow(
                [
                    row.truth_id,
                    row.page,
                    row.occurrence,
                    row.same_info_id,
                    row.text,
                    row.entity_type,
                    row.mode,
                    row.reason,
                    row.required,
                    ",".join(f"{value:.1f}" for value in row.rect),
                    row.replacement,
                    row.preserve,
                    row.protected_area,
                    row.user_status,
                    row.note,
                ]
            )


def write_final_review_images(source_pdf: Path, records: list[GroundTruthRecord]) -> list[Path]:
    paths: list[Path] = []
    doc = fitz.open(source_pdf)
    try:
        for page in sorted({record.page for record in records}):
            page_obj = doc[page - 1]
            shape = page_obj.new_shape()
            for record in [row for row in records if row.page == page]:
                color = (0.0, 0.65, 0.22) if record.user_status == STATUS_CONFIRMED else (0.45, 0.45, 0.45)
                rect = fitz.Rect(record.rect)
                shape.draw_rect(rect)
                shape.finish(color=color, width=1.4)
                page_obj.insert_text(fitz.Point(rect.x0, max(rect.y0 - 2, 8)), record.truth_id, fontsize=6.5, color=color)
            shape.commit()
            output = OUTPUT_DIR / f"PDF31代表ページ_{page}_正解枠確認_v1.png"
            page_obj.get_pixmap(dpi=144, alpha=False).save(output)
            paths.append(output)
    finally:
        doc.close()
    return paths


def completion_errors(records: list[GroundTruthRecord]) -> list[str]:
    counts = status_counts(records)
    errors: list[str] = []
    if len(records) != 49:
        errors.append(f"状態別件数の合計が49件ではありません: {len(records)}")
    if counts.get(STATUS_NEEDS_COORD, 0):
        errors.append(f"要座標修正が残っています: {counts[STATUS_NEEDS_COORD]}件")
    if counts.get(STATUS_PENDING, 0):
        errors.append(f"ユーザー確認待ちが残っています: {counts[STATUS_PENDING]}件")
    if counts.get(STATUS_HOLD, 0):
        errors.append(f"保留が残っています: {counts[STATUS_HOLD]}件")
    unfinished = [status for status in counts if status not in FINAL_STATUSES]
    if unfinished:
        errors.append(f"正式状態以外が残っています: {', '.join(f'{status}={counts[status]}' for status in unfinished)}")
    return errors


def finalize_outputs(state: ReviewState, reviewer: str, source_pdf: Path = SOURCE) -> None:
    records = state.records
    errors = completion_errors(records)
    if errors:
        raise RuntimeError("正解データはまだ確定できません:\n- " + "\n- ".join(errors))
    confirmed = [record for record in records if record.user_status == STATUS_CONFIRMED]
    excluded = [record for record in records if record.user_status == STATUS_EXCLUDED]
    OUTPUT_DIR.mkdir(exist_ok=True)
    write_records_csv(FINAL_GT_CSV, confirmed)
    write_records_csv(FINAL_STATUS_CSV, records)
    write_records_csv(EXCLUDED_CSV, excluded)
    write_final_review_images(source_pdf, records)
    counts = status_counts(records)
    modes = ", ".join(sorted({record.mode for record in confirmed}))
    lines = [
        "正解データ確定報告書",
        f"確認者: {reviewer}",
        f"確認日時: {datetime.now():%Y-%m-%d %H:%M:%S}",
        f"入力PDFのSHA-256: {sha256_file(source_pdf)}",
        "正解データのバージョン: v1",
        f"ユーザー確認済み矩形数: {counts.get(STATUS_CONFIRMED, 0)}",
        f"対象外矩形数: {counts.get(STATUS_EXCLUDED, 0)}",
        f"要修正矩形数: {counts.get(STATUS_NEEDS_COORD, 0)}",
        f"確認待ち矩形数: {counts.get(STATUS_PENDING, 0)}",
        f"対象モード: {modes}",
    ]
    FINAL_REPORT_TXT.write_text("\n".join(lines) + "\n", encoding="utf-8")


class GroundTruthReviewDialog(QDialog):
    def __init__(self, state: ReviewState, source_pdf: Path = SOURCE, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("PDF31代表6ページ 正解データ確認")
        self.resize(1320, 840)
        self.state = state
        self.source_pdf = source_pdf
        self.records = state.records
        self.current: GroundTruthRecord | None = None
        self.form_rect_snapshot: list[float] | None = None
        self.page_index = 0
        self.scale = 1.0
        self.page_pixmap: QPixmap | None = None
        self.items: dict[QGraphicsRectItem, GroundTruthRecord] = {}
        self.ocr_words: dict[int, list[tuple[Any, ...]]] = {}

        self.scene = QGraphicsScene(self)
        self.view = QGraphicsView(self.scene)
        self.view.setDragMode(QGraphicsView.RubberBandDrag)
        self.zoom_label = QLabel("拡大画像")
        self.zoom_label.setMinimumSize(320, 220)
        self.zoom_label.setAlignment(Qt.AlignCenter)
        self.zoom_label.setStyleSheet("border: 1px solid #cbd5e1; background: #f8fafc;")

        self.include_confirmed = QCheckBox("確認済み12件も再確認対象に含める")
        self.include_confirmed.stateChanged.connect(self.refresh_table)
        self.record_table = QTableWidget(0, 7)
        self.record_table.setHorizontalHeaderLabels(["順", "状態", "頁", "正解ID", "期待文字列", "種類", "一致率"])
        self.record_table.itemSelectionChanged.connect(self.goto_selected_record)

        self.reviewer_edit = QLineEdit(self.state.reviewer)
        self.truth_id_label = QLabel("-")
        self.page_label = QLabel("-")
        self.context_label = QLabel("-")
        self.ocr_label = QLabel("-")
        self.ocr_label.setWordWrap(True)
        self.status_combo = QComboBox()
        self.status_combo.addItems([STATUS_CONFIRMED, STATUS_EXCLUDED, STATUS_NEEDS_COORD, STATUS_HOLD, STATUS_PENDING, STATUS_RELEASE])
        self.text_edit = QLineEdit()
        self.entity_edit = QLineEdit()
        self.mode_edit = QLineEdit()
        self.note_edit = QLineEdit()
        self.x_spin = self._spinbox()
        self.y_spin = self._spinbox()
        self.w_spin = self._spinbox()
        self.h_spin = self._spinbox()

        self._build_ui()
        self.scene.selectionChanged.connect(self.on_scene_selection_changed)
        self.refresh_table()
        first = ordered_records(self.records, include_confirmed=False)
        if first:
            self.select_record(first[0])

    def _build_ui(self) -> None:
        left = QVBoxLayout()
        left.addWidget(self.include_confirmed)
        left.addWidget(self.record_table, 1)

        nav_row = QHBoxLayout()
        previous_button = QPushButton("前へ")
        next_button = QPushButton("次へ")
        previous_button.clicked.connect(lambda: self.goto_offset(-1))
        next_button.clicked.connect(lambda: self.goto_offset(1))
        nav_row.addWidget(previous_button)
        nav_row.addWidget(next_button)
        left.addLayout(nav_row)

        left_widget = QWidget()
        left_widget.setLayout(left)

        center = QVBoxLayout()
        center.addWidget(QLabel("ページ全体画像"))
        center.addWidget(self.view, 1)
        center.addWidget(QLabel("拡大画像"))
        center.addWidget(self.zoom_label)
        center_widget = QWidget()
        center_widget.setLayout(center)

        right = QVBoxLayout()
        form = QFormLayout()
        form.addRow("確認者", self.reviewer_edit)
        form.addRow("ページ番号", self.page_label)
        form.addRow("正解ID", self.truth_id_label)
        form.addRow("期待文字列", self.text_edit)
        form.addRow("情報種別", self.entity_edit)
        form.addRow("適用モード", self.mode_edit)
        form.addRow("現在の状態", self.status_combo)
        form.addRow("枠内OCR文字", self.ocr_label)
        form.addRow("前後の項目", self.context_label)
        form.addRow("X", self.x_spin)
        form.addRow("Y", self.y_spin)
        form.addRow("幅", self.w_spin)
        form.addRow("高さ", self.h_spin)
        form.addRow("備考", self.note_edit)
        right.addLayout(form)

        action_rows = [
            [("承認", self.approve_current), ("対象外", self.exclude_current)],
            [("要座標修正", self.needs_coord_current), ("保留", self.hold_current)],
            [("枠を反映", self.apply_edits), ("枠を削除", self.delete_current)],
            [("新しい枠の追加", self.add_record), ("複数枠に分割", self.split_current)],
            [("選択枠を結合", self.merge_selected), ("保存", self.save)],
            [("正式版出力", self.finalize), ("閉じる", self.accept)],
        ]
        for row in action_rows:
            layout = QHBoxLayout()
            for label, slot in row:
                button = QPushButton(label)
                button.clicked.connect(slot)
                layout.addWidget(button)
            right.addLayout(layout)
        right.addStretch(1)
        right_widget = QWidget()
        right_widget.setLayout(right)

        splitter = QSplitter()
        splitter.addWidget(left_widget)
        splitter.addWidget(center_widget)
        splitter.addWidget(right_widget)
        splitter.setStretchFactor(0, 0)
        splitter.setStretchFactor(1, 1)
        splitter.setStretchFactor(2, 0)

        layout = QVBoxLayout(self)
        legend = QLabel("確認順序: ユーザー確認待ち8件 → 要座標修正28件 → 解除候補/保留。緑=確認済み、灰=対象外/解除候補、赤=要座標修正、黄=確認待ち/保留")
        layout.addWidget(legend)
        layout.addWidget(splitter, 1)

    def refresh_table(self) -> None:
        selected_id = self.current.truth_id if self.current else ""
        rows = ordered_records(self.records, include_confirmed=self.include_confirmed.isChecked())
        self.record_table.blockSignals(True)
        try:
            self.record_table.setRowCount(0)
            for index, record in enumerate(rows, start=1):
                row = self.record_table.rowCount()
                self.record_table.insertRow(row)
                words = self.words_for_page(record.page)
                _inside, ratio = ocr_text_for_rect(words, record)
                values = [str(index), record.user_status, str(record.page), record.truth_id, record.text, record.entity_type, f"{ratio:.3f}"]
                for column, value in enumerate(values):
                    item = QTableWidgetItem(value)
                    item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                    self.record_table.setItem(row, column, item)
                if record.truth_id == selected_id:
                    self.record_table.selectRow(row)
        finally:
            self.record_table.blockSignals(False)

    def goto_selected_record(self) -> None:
        rows = self.record_table.selectionModel().selectedRows() if self.record_table.selectionModel() else []
        if not rows:
            return
        truth_id = self.record_table.item(rows[0].row(), 3).text()
        record = next((row for row in self.records if row.truth_id == truth_id), None)
        if record:
            self.select_record(record)

    def select_record(self, record: GroundTruthRecord) -> None:
        self.sync_scene_rects()
        self.current = record
        self.state.last_truth_id = record.truth_id
        if self.page_index != record.page - 1:
            self.load_page(record.page - 1)
        self.populate_form(record)
        for item, item_record in self.items.items():
            item.setSelected(item_record.truth_id == record.truth_id)
            self.style_item(item, item_record)
        self.update_zoom()

    def load_page(self, page_index: int, sync_existing: bool = True) -> None:
        if sync_existing:
            self.sync_scene_rects()
        self.page_index = page_index
        self.scene.clear()
        self.items = {}
        self.page_pixmap = self.render_page(page_index)
        self.scale = self.page_pixmap.width() / max(self.page_width(page_index), 1)
        self.scene.addItem(QGraphicsPixmapItem(self.page_pixmap))
        for record in self.records:
            if record.page == page_index + 1:
                item = self.add_rect_item(record)
                item.setSelected(self.current is not None and record.truth_id == self.current.truth_id)
        self.scene.setSceneRect(self.scene.itemsBoundingRect())

    def populate_form(self, record: GroundTruthRecord) -> None:
        self.page_label.setText(str(record.page))
        self.truth_id_label.setText(record.truth_id)
        self.text_edit.setText(record.text)
        self.entity_edit.setText(record.entity_type)
        self.mode_edit.setText(record.mode)
        index = self.status_combo.findText(record.user_status)
        if index >= 0:
            self.status_combo.setCurrentIndex(index)
        self.note_edit.setText(record.note)
        x0, y0, x1, y1 = record.rect
        self.x_spin.setValue(x0)
        self.y_spin.setValue(y0)
        self.w_spin.setValue(max(x1 - x0, 1))
        self.h_spin.setValue(max(y1 - y0, 1))
        self.form_rect_snapshot = [x0, y0, x1, y1]
        inside, ratio = ocr_text_for_rect(self.words_for_page(record.page), record)
        self.ocr_label.setText(f"一致率: {ratio:.3f}\n{inside or '(空)'}")
        self.context_label.setText(self.context_text(record))

    def context_text(self, record: GroundTruthRecord) -> str:
        page_records = sorted([row for row in self.records if row.page == record.page], key=lambda row: row.truth_id)
        index = page_records.index(record)
        previous_text = page_records[index - 1].truth_id + " " + page_records[index - 1].text if index > 0 else "(なし)"
        next_text = page_records[index + 1].truth_id + " " + page_records[index + 1].text if index + 1 < len(page_records) else "(なし)"
        return f"前: {previous_text}\n次: {next_text}"

    def apply_edits(self) -> None:
        if not self.current:
            return
        resolved_rect = self.resolve_current_edit_rect()
        self.sync_scene_rects()
        self.current.text = self.text_edit.text().strip()
        self.current.entity_type = self.entity_edit.text().strip()
        self.current.mode = self.mode_edit.text().strip()
        self.current.user_status = self.status_combo.currentText()
        self.current.note = self.note_edit.text().strip()
        self.current.updated_at = datetime.now().isoformat(timespec="seconds")
        self.current.rect = resolved_rect
        self.load_page(self.current.page - 1, sync_existing=False)
        self.populate_form(self.current)
        self.refresh_table()
        self.state.reviewer = self.reviewer_edit.text().strip()
        save_state(self.state)

    def resolve_current_edit_rect(self) -> list[float]:
        form_rect = self.form_rect_values()
        scene_rect = self.current_scene_rect_values()
        snapshot = self.form_rect_snapshot or list(self.current.rect if self.current else form_rect)
        if self.rect_changed(form_rect, snapshot):
            return form_rect
        if scene_rect is not None and self.rect_changed(scene_rect, snapshot):
            return scene_rect
        return form_rect

    def form_rect_values(self) -> list[float]:
        x0 = self.x_spin.value()
        y0 = self.y_spin.value()
        return [x0, y0, x0 + self.w_spin.value(), y0 + self.h_spin.value()]

    def current_scene_rect_values(self) -> list[float] | None:
        if not self.current:
            return None
        for item, record in self.items.items():
            if record is self.current or record.truth_id == self.current.truth_id:
                rect = item.mapRectToScene(item.rect())
                return [
                    rect.x() / self.scale,
                    rect.y() / self.scale,
                    (rect.x() + rect.width()) / self.scale,
                    (rect.y() + rect.height()) / self.scale,
                ]
        return None

    @staticmethod
    def rect_changed(a: list[float], b: list[float], tolerance: float = 0.05) -> bool:
        return any(abs(left - right) > tolerance for left, right in zip(a, b))

    def approve_current(self) -> None:
        self.set_status(STATUS_CONFIRMED)

    def exclude_current(self) -> None:
        self.set_status(STATUS_EXCLUDED)

    def needs_coord_current(self) -> None:
        self.set_status(STATUS_NEEDS_COORD)

    def hold_current(self) -> None:
        self.set_status(STATUS_HOLD)

    def set_status(self, status: str) -> None:
        if not self.current:
            return
        self.apply_edits()
        self.current.user_status = status
        self.current.updated_at = datetime.now().isoformat(timespec="seconds")
        self.save(show_message=False)
        self.refresh_table()
        self.goto_offset(1)

    def add_record(self) -> None:
        page = self.current.page if self.current else 1
        existing = [row.truth_id for row in self.records if row.truth_id.startswith(f"P{page:02d}-ADD")]
        truth_id = f"P{page:02d}-ADD{len(existing) + 1:02d}"
        rect = [80.0, 120.0, 220.0, 150.0]
        if self.current:
            x0, y0, x1, y1 = self.current.rect
            rect = [x0 + 12, y0 + 12, x1 + 12, y1 + 12]
        record = GroundTruthRecord(
            truth_id=truth_id,
            page=page,
            occurrence=1,
            same_info_id=truth_id,
            text="",
            entity_type="会社名",
            mode="企業機密変換オン時",
            reason="手動追加",
            required="必須",
            rect=rect,
            replacement="",
            preserve="",
            protected_area="",
            user_status=STATUS_HOLD,
            note="新しい枠の追加",
            source="manual",
            updated_at=datetime.now().isoformat(timespec="seconds"),
        )
        self.records.append(record)
        self.refresh_table()
        self.select_record(record)

    def split_current(self) -> None:
        if not self.current:
            return
        x0, y0, x1, y1 = self.current.rect
        mid = (x0 + x1) / 2
        self.current.rect = [x0, y0, mid, y1]
        new_id = f"{self.current.truth_id}-B"
        record = GroundTruthRecord(**{**asdict(self.current), "truth_id": new_id, "rect": [mid, y0, x1, y1], "user_status": STATUS_HOLD, "note": "複数枠への分割", "source": "split"})
        self.records.append(record)
        self.refresh_table()
        self.select_record(record)

    def delete_current(self) -> None:
        if not self.current:
            return
        if QMessageBox.question(self, "枠の削除", f"{self.current.truth_id} を削除しますか？") != QMessageBox.Yes:
            return
        self.records.remove(self.current)
        self.current = None
        self.refresh_table()
        rows = ordered_records(self.records, include_confirmed=self.include_confirmed.isChecked())
        if rows:
            self.select_record(rows[0])

    def merge_selected(self) -> None:
        items = [item for item in self.scene.selectedItems() if isinstance(item, QGraphicsRectItem) and item in self.items]
        if len(items) < 2:
            QMessageBox.information(self, "枠の結合", "結合する枠を2つ以上選択してください。")
            return
        records = [self.items[item] for item in items]
        x0 = min(record.rect[0] for record in records)
        y0 = min(record.rect[1] for record in records)
        x1 = max(record.rect[2] for record in records)
        y1 = max(record.rect[3] for record in records)
        primary = records[0]
        primary.rect = [x0, y0, x1, y1]
        primary.note = (primary.note + " / 選択枠を結合").strip(" /")
        primary.updated_at = datetime.now().isoformat(timespec="seconds")
        for record in records[1:]:
            if record in self.records:
                self.records.remove(record)
        self.refresh_table()
        self.select_record(primary)

    def goto_offset(self, offset: int) -> None:
        rows = ordered_records(self.records, include_confirmed=self.include_confirmed.isChecked())
        if not rows:
            return
        if not self.current or self.current not in rows:
            self.select_record(rows[0])
            return
        index = rows.index(self.current)
        self.select_record(rows[(index + offset) % len(rows)])

    def save(self, show_message: bool = True) -> None:
        self.sync_scene_rects()
        self.state.reviewer = self.reviewer_edit.text().strip()
        save_state(self.state)
        if show_message:
            QMessageBox.information(self, "保存", f"レビュー状態を保存しました。\n{STATE_JSON}")

    def finalize(self) -> None:
        self.save(show_message=False)
        try:
            finalize_outputs(self.state, self.reviewer_edit.text().strip() or self.state.reviewer, self.source_pdf)
        except Exception as exc:
            QMessageBox.warning(self, "正式版出力不可", str(exc))
            return
        QMessageBox.information(self, "正式版出力", f"正式版を出力しました。\n{FINAL_GT_CSV}")

    def accept(self) -> None:
        self.save(show_message=False)
        super().accept()

    def add_rect_item(self, record: GroundTruthRecord) -> QGraphicsRectItem:
        x0, y0, x1, y1 = record.rect
        item = QGraphicsRectItem(x0 * self.scale, y0 * self.scale, (x1 - x0) * self.scale, (y1 - y0) * self.scale)
        item.setFlags(QGraphicsRectItem.ItemIsMovable | QGraphicsRectItem.ItemIsSelectable | QGraphicsRectItem.ItemSendsGeometryChanges)
        self.style_item(item, record)
        self.scene.addItem(item)
        self.items[item] = record
        return item

    def style_item(self, item: QGraphicsRectItem, record: GroundTruthRecord) -> None:
        color = {
            STATUS_CONFIRMED: QColor(0, 166, 80),
            STATUS_EXCLUDED: QColor(120, 120, 120),
            STATUS_RELEASE: QColor(120, 120, 120),
            STATUS_NEEDS_COORD: QColor(220, 40, 40),
            STATUS_PENDING: QColor(235, 170, 20),
            STATUS_HOLD: QColor(235, 170, 20),
        }.get(record.user_status, QColor(40, 110, 220))
        width = 4 if record is self.current else 2
        item.setPen(QPen(color, width))
        item.setBrush(Qt.NoBrush)

    def on_scene_selection_changed(self) -> None:
        selected = [item for item in self.scene.selectedItems() if isinstance(item, QGraphicsRectItem) and item in self.items]
        if not selected:
            return
        record = self.items[selected[0]]
        self.current = record
        self.populate_form(record)
        self.update_zoom()

    def sync_scene_rects(self) -> None:
        if not self.items:
            return
        for item, record in list(self.items.items()):
            rect = item.mapRectToScene(item.rect())
            record.rect = [
                rect.x() / self.scale,
                rect.y() / self.scale,
                (rect.x() + rect.width()) / self.scale,
                (rect.y() + rect.height()) / self.scale,
            ]
            item.setPos(0, 0)
            item.setRect(rect)

    def update_zoom(self) -> None:
        if not self.current or not self.page_pixmap:
            return
        x0, y0, x1, y1 = [value * self.scale for value in self.current.rect]
        margin = 80
        rect = self.page_pixmap.rect().intersected(
            self.page_pixmap.rect().adjusted(
                int(x0 - margin) - self.page_pixmap.rect().left(),
                int(y0 - margin) - self.page_pixmap.rect().top(),
                int(x1 + margin) - self.page_pixmap.rect().right(),
                int(y1 + margin) - self.page_pixmap.rect().bottom(),
            )
        )
        if rect.isEmpty():
            return
        crop = self.page_pixmap.copy(rect).scaled(self.zoom_label.size(), Qt.KeepAspectRatio, Qt.SmoothTransformation)
        self.zoom_label.setPixmap(crop)

    def words_for_page(self, page: int) -> list[tuple[Any, ...]]:
        if page not in self.ocr_words:
            doc = fitz.open(self.source_pdf)
            try:
                _text, words = ocr_page_text_and_words(doc[page - 1])
            finally:
                doc.close()
            self.ocr_words[page] = words
        return self.ocr_words[page]

    def render_page(self, page_index: int) -> QPixmap:
        doc = fitz.open(self.source_pdf)
        try:
            pix = doc[page_index].get_pixmap(dpi=120, alpha=False)
            image = QImage(pix.samples, pix.width, pix.height, pix.stride, QImage.Format_RGB888).copy()
            return QPixmap.fromImage(image)
        finally:
            doc.close()

    def page_width(self, page_index: int) -> float:
        doc = fitz.open(self.source_pdf)
        try:
            return float(doc[page_index].rect.width)
        finally:
            doc.close()

    def _spinbox(self) -> QDoubleSpinBox:
        spin = QDoubleSpinBox()
        spin.setRange(0, 10000)
        spin.setDecimals(1)
        spin.setSingleStep(2.0)
        return spin


def command_check() -> int:
    state = load_state()
    counts = status_counts(state.records)
    save_state(state)
    print(f"state_json={STATE_JSON}")
    print(f"total={len(state.records)}")
    for status in [STATUS_CONFIRMED, STATUS_NEEDS_COORD, STATUS_RELEASE, STATUS_PENDING, STATUS_HOLD, STATUS_EXCLUDED]:
        print(f"{status}={counts.get(status, 0)}")
    errors = completion_errors(state.records)
    print("finalizable=" + ("yes" if not errors else "no"))
    if errors:
        for error in errors:
            print(f"block_reason={error}")
    return 0


def command_finalize(reviewer: str) -> int:
    state = load_state()
    try:
        finalize_outputs(state, reviewer or state.reviewer or os.environ.get("USERNAME", ""), SOURCE)
    except Exception as exc:
        print(str(exc))
        return 1
    print(f"final_ground_truth_csv={FINAL_GT_CSV}")
    print(f"status_csv={FINAL_STATUS_CSV}")
    print(f"excluded_csv={EXCLUDED_CSV}")
    print(f"report={FINAL_REPORT_TXT}")
    return 0


def command_gui() -> int:
    state = load_state()
    app = QApplication(sys.argv)
    dialog = GroundTruthReviewDialog(state)
    dialog.show()
    return app.exec()


def main() -> int:
    parser = argparse.ArgumentParser(description="PDF31代表6ページの正解データ確認ツール")
    parser.add_argument("--check", action="store_true", help="現在のレビュー状態を確認し、JSONがなければ初期化する")
    parser.add_argument("--finalize", action="store_true", help="完了条件を満たす場合だけ正式版を出力する")
    parser.add_argument("--reviewer", default=os.environ.get("USERNAME", ""), help="確認者名")
    args = parser.parse_args()
    if args.check:
        return command_check()
    if args.finalize:
        return command_finalize(args.reviewer)
    return command_gui()


if __name__ == "__main__":
    raise SystemExit(main())
