from __future__ import annotations

import argparse
import csv
import hashlib
import json
import sys
import tkinter as tk
import zipfile
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from tkinter import messagebox, ttk
from xml.etree import ElementTree as ET


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from excel_privacy_cleaner.word_processor import WordParagraphText, extract_word_structure  # noqa: E402


CATEGORIES = ("会社名", "氏名", "住所", "電話番号", "メールアドレス", "銀行名")
GT_FIELDS = (
    "truth_id",
    "category",
    "location_id",
    "char_start",
    "char_end",
    "text",
    "story_type",
    "section_index",
    "container_type",
    "paragraph_index",
    "table_index",
    "row_index",
    "cell_index",
    "source",
    "property_name",
)


@dataclass(frozen=True)
class ReviewItem:
    source: str
    label: str
    text: str
    location_id: str
    story_type: str
    section_index: int | None
    container_type: str
    paragraph_index: int
    table_index: int | None = None
    row_index: int | None = None
    cell_index: int | None = None
    property_name: str = ""


def main() -> int:
    parser = argparse.ArgumentParser(description="Create independent Word holdout ground truth without reading candidates.")
    parser.add_argument("--docx", required=True, type=Path)
    parser.add_argument("--out-dir", required=True, type=Path)
    parser.add_argument("--holdout-id", default="Word Holdout001")
    parser.add_argument("--ground-truth-name", default="Holdout001_ground_truth.csv")
    args = parser.parse_args()

    args.out_dir.mkdir(parents=True, exist_ok=True)
    app = GroundTruthApp(args.docx, args.out_dir, args.holdout_id, args.ground_truth_name)
    app.mainloop()
    return 0


class GroundTruthApp(tk.Tk):
    def __init__(self, docx_path: Path, out_dir: Path, holdout_id: str, ground_truth_name: str) -> None:
        super().__init__()
        self.docx_path = docx_path
        self.out_dir = out_dir
        self.holdout_id = holdout_id
        self.ground_truth_path = out_dir / ground_truth_name
        self.meta_path = out_dir / f"{self.ground_truth_path.stem}_meta.json"
        self.unsupported_path = out_dir / f"{safe_file_stem(holdout_id)}_unsupported_review.json"
        self.truth_prefix = f"WORDH{holdout_number(holdout_id)}"
        self.review_status = "IN_PROGRESS"
        self.title(f"{holdout_id} Ground Truth")
        self.geometry("1180x760")

        self.inventory = extract_word_structure(docx_path)
        self.items = self._build_review_items()
        self.truths: list[dict[str, str]] = []
        self._load_existing()
        self._write_unsupported_review()
        self._build_ui()
        self._refresh_item_list()
        self._refresh_truth_list()
        if self.items:
            self.item_list.selection_set(0)
            self._show_selected_item()

    def _build_review_items(self) -> list[ReviewItem]:
        items = [self._paragraph_item(index, paragraph) for index, paragraph in enumerate(self.inventory.paragraphs)]
        props = self.inventory.document_properties
        for property_name in (
            "author",
            "last_modified_by",
            "title",
            "subject",
            "keywords",
            "category",
            "comments",
            "content_status",
            "identifier",
            "language",
            "version",
        ):
            value = str(getattr(props, property_name, "") or "")
            if value:
                items.append(
                    ReviewItem(
                        source="document_property",
                        label=f"property:{property_name}",
                        text=value,
                        location_id=f"document_property:None::property:None:None:None:0:None:/docProps/core.xml:{property_name}",
                        story_type="document_property",
                        section_index=None,
                        container_type="property",
                        paragraph_index=0,
                        property_name=property_name,
                    )
                )
        return items

    def _paragraph_item(self, display_index: int, paragraph: WordParagraphText) -> ReviewItem:
        return ReviewItem(
            source="paragraph",
            label=f"{display_index + 1:03d} {paragraph.story_type}/{paragraph.container_type}",
            text=paragraph.text,
            location_id=paragraph_location_id(paragraph),
            story_type=paragraph.story_type,
            section_index=paragraph.section_index,
            container_type=paragraph.container_type,
            paragraph_index=paragraph.paragraph_index,
            table_index=paragraph.table_index,
            row_index=paragraph.row_index,
            cell_index=paragraph.cell_index,
        )

    def _build_ui(self) -> None:
        self.columnconfigure(1, weight=1)
        self.rowconfigure(0, weight=1)

        left = ttk.Frame(self, padding=8)
        left.grid(row=0, column=0, sticky="ns")
        ttk.Label(left, text="Review Items").pack(anchor="w")
        self.item_list = tk.Listbox(left, width=38, height=34, exportselection=False)
        self.item_list.pack(fill="y", expand=True)
        self.item_list.bind("<<ListboxSelect>>", lambda _event: self._show_selected_item())

        center = ttk.Frame(self, padding=8)
        center.grid(row=0, column=1, sticky="nsew")
        center.columnconfigure(0, weight=1)
        center.rowconfigure(1, weight=1)
        self.location_label = ttk.Label(center, text="")
        self.location_label.grid(row=0, column=0, sticky="ew")
        self.text_widget = tk.Text(center, wrap="word", height=18, undo=False)
        self.text_widget.grid(row=1, column=0, sticky="nsew")
        self.text_widget.configure(font=("Yu Gothic UI", 11))

        controls = ttk.Frame(center)
        controls.grid(row=2, column=0, sticky="ew", pady=(8, 0))
        self.category_var = tk.StringVar(value=CATEGORIES[0])
        self.category_display_var = tk.StringVar(value=f"現在のカテゴリ: {self.category_var.get()}")
        self.category_var.trace_add("write", lambda *_args: self._refresh_category_display())
        self.category_display = ttk.Label(center, textvariable=self.category_display_var, font=("Yu Gothic UI", 14, "bold"))
        self.category_display.grid(row=3, column=0, sticky="ew", pady=(8, 0))
        ttk.Label(controls, text="Category").pack(side="left")
        ttk.Combobox(controls, textvariable=self.category_var, values=CATEGORIES, state="readonly", width=18).pack(side="left", padx=6)
        ttk.Button(controls, text="追加", command=self._add_truth).pack(side="left", padx=4)
        ttk.Button(controls, text="保存", command=self._save).pack(side="left", padx=4)
        ttk.Button(controls, text="レビュー完了", command=self._complete_review).pack(side="left", padx=4)

        right = ttk.Frame(self, padding=8)
        right.grid(row=0, column=2, sticky="ns")
        ttk.Label(right, text="Ground Truth").pack(anchor="w")
        self.truth_list = tk.Listbox(right, width=54, height=31, exportselection=False)
        self.truth_list.pack(fill="y", expand=True)
        ttk.Button(right, text="登録箇所へ移動", command=self._jump_to_truth).pack(anchor="e", pady=(6, 0))
        ttk.Button(right, text="現在カテゴリへ変更", command=self._update_truth_category).pack(anchor="e", pady=(6, 0))
        ttk.Button(right, text="選択を削除", command=self._delete_truth).pack(anchor="e", pady=(6, 0))
        self.status_label = ttk.Label(right, text="")
        self.status_label.pack(anchor="w", pady=(10, 0))

    def _refresh_item_list(self) -> None:
        self.item_list.delete(0, tk.END)
        for item in self.items:
            prefix = "PROP" if item.source == "document_property" else item.story_type.upper()
            self.item_list.insert(tk.END, f"{prefix} {item.label}")

    def _show_selected_item(self) -> None:
        item = self._current_item()
        if item is None:
            return
        self.location_label.configure(text=f"{item.label} | {item.location_id}")
        self.text_widget.configure(state="normal")
        self.text_widget.delete("1.0", tk.END)
        self.text_widget.insert("1.0", item.text)

    def _current_item(self) -> ReviewItem | None:
        selected = self.item_list.curselection()
        if not selected:
            return None
        return self.items[int(selected[0])]

    def _add_truth(self) -> None:
        item = self._current_item()
        if item is None:
            return
        try:
            start_index = self.text_widget.index("sel.first")
            end_index = self.text_widget.index("sel.last")
        except tk.TclError:
            messagebox.showwarning("Selection required", "匿名化対象の文字列を選択してください。")
            return
        start = index_to_offset(self.text_widget, start_index)
        end = index_to_offset(self.text_widget, end_index)
        if end <= start:
            messagebox.showwarning("Invalid selection", "選択範囲が空です。")
            return
        selected_text = item.text[start:end]
        truth = {
            "truth_id": f"{self.truth_prefix}-{len(self.truths) + 1:04d}",
            "category": self.category_var.get(),
            "location_id": item.location_id,
            "char_start": str(start),
            "char_end": str(end),
            "text": selected_text,
            "story_type": item.story_type,
            "section_index": none_to_empty(item.section_index),
            "container_type": item.container_type,
            "paragraph_index": str(item.paragraph_index),
            "table_index": none_to_empty(item.table_index),
            "row_index": none_to_empty(item.row_index),
            "cell_index": none_to_empty(item.cell_index),
            "source": item.source,
            "property_name": item.property_name,
        }
        self.truths.append(truth)
        self._renumber_truths()
        self._refresh_truth_list()

    def _delete_truth(self) -> None:
        selected = self.truth_list.curselection()
        if not selected:
            return
        del self.truths[int(selected[0])]
        self._renumber_truths()
        self._refresh_truth_list()

    def _update_truth_category(self) -> None:
        selected = self.truth_list.curselection()
        if not selected:
            messagebox.showwarning("Selection required", "変更するground truthを右側一覧で選択してください。")
            return
        index = int(selected[0])
        self.truths[index]["category"] = self.category_var.get()
        self.review_status = "IN_PROGRESS"
        self._refresh_truth_list()
        self.truth_list.selection_set(index)

    def _jump_to_truth(self) -> None:
        selected = self.truth_list.curselection()
        if not selected:
            messagebox.showwarning("Selection required", "表示するground truthを右側一覧で選択してください。")
            return
        truth = self.truths[int(selected[0])]
        for index, item in enumerate(self.items):
            if item.location_id == truth.get("location_id"):
                self.item_list.selection_clear(0, tk.END)
                self.item_list.selection_set(index)
                self.item_list.see(index)
                self._show_selected_item()
                start = int(truth["char_start"])
                end = int(truth["char_end"])
                self.text_widget.tag_remove(tk.SEL, "1.0", tk.END)
                self.text_widget.tag_add(tk.SEL, f"1.0 + {start} chars", f"1.0 + {end} chars")
                self.text_widget.see(f"1.0 + {start} chars")
                self.category_var.set(truth.get("category", CATEGORIES[0]))
                return
        messagebox.showwarning("Not found", "この登録のlocation_idに対応する表示項目が見つかりません。")

    def _renumber_truths(self) -> None:
        for index, truth in enumerate(self.truths, start=1):
            truth["truth_id"] = f"{self.truth_prefix}-{index:04d}"

    def _refresh_truth_list(self) -> None:
        self.truth_list.delete(0, tk.END)
        for truth in self.truths:
            text_length = int(truth["char_end"]) - int(truth["char_start"])
            self.truth_list.insert(
                tk.END,
                f"{truth['truth_id']} {truth['category']} {truth['story_type']} {truth['char_start']}-{truth['char_end']} len={text_length}",
            )
        self.status_label.configure(text=f"review_status={self.review_status} / truths={len(self.truths)}")

    def _refresh_category_display(self) -> None:
        self.category_display_var.set(f"現在のカテゴリ: {self.category_var.get()}")

    def _load_existing(self) -> None:
        if self.meta_path.exists():
            meta = json.loads(self.meta_path.read_text(encoding="utf-8"))
            self.review_status = str(meta.get("review_status", "IN_PROGRESS"))
        if not self.ground_truth_path.exists():
            return
        with self.ground_truth_path.open(encoding="utf-8-sig", newline="") as input_file:
            self.truths = list(csv.DictReader(input_file))

    def _save(self) -> None:
        self._write_truths()
        self._write_meta()
        self._refresh_truth_list()
        messagebox.showinfo("Saved", f"保存しました。\nreview_status={self.review_status}")

    def _complete_review(self) -> None:
        if not messagebox.askyesno("Complete review", "全ページ・header/footer・document propertiesの確認が完了しましたか？"):
            return
        self.review_status = "COMPLETED"
        self._write_truths()
        self._write_meta()
        self._refresh_truth_list()
        messagebox.showinfo("Completed", "レビュー完了として保存しました。")

    def _write_truths(self) -> None:
        with self.ground_truth_path.open("w", encoding="utf-8-sig", newline="") as output_file:
            writer = csv.DictWriter(output_file, fieldnames=GT_FIELDS)
            writer.writeheader()
            for truth in self.truths:
                writer.writerow(truth)

    def _write_meta(self) -> None:
        payload = {
            "holdout_id": self.holdout_id,
            "review_status": self.review_status,
            "updated_at_utc": datetime.now(timezone.utc).isoformat(),
            "docx_sha256": sha256(self.docx_path),
            "ground_truth_path": str(self.ground_truth_path),
            "ground_truth_sha256": sha256(self.ground_truth_path) if self.ground_truth_path.exists() else "",
            "truth_count": len(self.truths),
            "category_counts": category_counts(self.truths),
            "candidate_reference_policy": "This tool does not read candidate CSV/JSON/count/category/span data.",
        }
        self.meta_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")

    def _write_unsupported_review(self) -> None:
        payload = {
            "holdout_id": self.holdout_id,
            "docx_sha256": sha256(self.docx_path),
            "unsupported_review": inspect_unsupported_text_presence(self.docx_path),
            "note": "Text content is not copied into this file; only existence and non-empty text flags are recorded.",
        }
        self.unsupported_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def paragraph_location_id(paragraph: WordParagraphText) -> str:
    values = [
        paragraph.story_type,
        str(paragraph.section_index),
        paragraph.header_footer_type or "",
        paragraph.container_type,
        str(paragraph.table_index),
        str(paragraph.row_index),
        str(paragraph.cell_index),
        str(paragraph.paragraph_index),
        str(paragraph.cell_paragraph_index),
        paragraph.part_name,
        paragraph.element_path,
    ]
    return ":".join(values)


def index_to_offset(widget: tk.Text, index: str) -> int:
    return len(widget.get("1.0", index))


def none_to_empty(value: int | None) -> str:
    return "" if value is None else str(value)


def holdout_number(holdout_id: str) -> str:
    digits = "".join(char for char in holdout_id if char.isdigit())
    return digits[-3:].zfill(3) if digits else "000"


def safe_file_stem(value: str) -> str:
    return "".join(char if char.isalnum() else "_" for char in value).strip("_") or "word_holdout"


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as input_file:
        for chunk in iter(lambda: input_file.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def category_counts(truths: list[dict[str, str]]) -> dict[str, int]:
    counts = {category: 0 for category in CATEGORIES}
    for truth in truths:
        category = truth.get("category", "")
        if category in counts:
            counts[category] += 1
    return counts


def inspect_unsupported_text_presence(docx_path: Path) -> list[dict[str, object]]:
    interesting_parts = (
        "word/footnotes.xml",
        "word/endnotes.xml",
        "word/comments.xml",
        "word/settings.xml",
        "word/styles.xml",
        "word/numbering.xml",
    )
    results: list[dict[str, object]] = []
    with zipfile.ZipFile(docx_path) as archive:
        names = set(archive.namelist())
        for part_name in interesting_parts:
            if part_name not in names:
                continue
            xml = archive.read(part_name)
            text_count, nonempty_text_count = count_text_nodes(xml)
            results.append(
                {
                    "part_name": part_name,
                    "exists": True,
                    "text_node_count": text_count,
                    "nonempty_text_node_count": nonempty_text_count,
                    "classification": "実際の非空テキストあり" if nonempty_text_count else "part/element存在のみ",
                }
            )
        for part_name in sorted(name for name in names if name.startswith("word/") and name.endswith(".xml")):
            if part_name in interesting_parts:
                continue
            xml = archive.read(part_name)
            if b"txbxContent" not in xml and b"<w:ins" not in xml and b"<w:del" not in xml and b"fldChar" not in xml:
                continue
            text_count, nonempty_text_count = count_text_nodes(xml)
            results.append(
                {
                    "part_name": part_name,
                    "exists": True,
                    "text_node_count": text_count,
                    "nonempty_text_node_count": nonempty_text_count,
                    "classification": "実際の非空テキストあり" if nonempty_text_count else "part/element存在のみ",
                }
            )
    return results


def count_text_nodes(xml: bytes) -> tuple[int, int]:
    try:
        root = ET.fromstring(xml)
    except ET.ParseError:
        return 0, 0
    text_count = 0
    nonempty_text_count = 0
    for element in root.iter():
        if element.tag.endswith("}t") or element.tag.endswith("}instrText") or element.tag.endswith("}delText"):
            text_count += 1
            if (element.text or "").strip():
                nonempty_text_count += 1
    return text_count, nonempty_text_count


if __name__ == "__main__":
    raise SystemExit(main())
