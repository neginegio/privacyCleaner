from __future__ import annotations

import sys
import traceback
from pathlib import Path
from tkinter import BOTH, END, LEFT, RIGHT, TOP, X, filedialog, messagebox, ttk
import tkinter as tk

from .excel_processor import ExcelPrivacyProcessor
from .models import Finding


SUPPORTED_EXTENSIONS = {".xlsx", ".xlsm"}


class ExcelPrivacyCleanerApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Excel 機密情報削除 - Presidio")
        self.geometry("1180x760")
        self.minsize(980, 640)

        self.processor = ExcelPrivacyProcessor()
        self.source_path: Path | None = None
        self.findings: list[Finding] = []
        self.item_to_index: dict[str, int] = {}

        self._build_ui()
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    def _build_ui(self) -> None:
        container = ttk.Frame(self, padding=12)
        container.pack(fill=BOTH, expand=True)

        title = ttk.Label(
            container,
            text="Excel ファイルを選択してください。Presidio による検出はこの PC 内だけで行い、原本は上書きしません。",
            font=("", 11, "bold"),
        )
        title.pack(fill=X, pady=(0, 10))

        controls = ttk.Frame(container)
        controls.pack(fill=X, pady=(0, 8))

        self.path_var = tk.StringVar()
        ttk.Entry(controls, textvariable=self.path_var, state="readonly").pack(side=LEFT, fill=X, expand=True)
        ttk.Button(controls, text="Excelを選択", command=self.choose_file).pack(side=LEFT, padx=(8, 0))
        ttk.Button(controls, text="検査開始", command=self.scan_file).pack(side=LEFT, padx=(8, 0))
        ttk.Button(controls, text="確認済みを変換保存", command=self.convert_file).pack(side=LEFT, padx=(8, 0))

        selection_controls = ttk.Frame(container)
        selection_controls.pack(fill=X, pady=(0, 8))
        ttk.Button(selection_controls, text="選択行を切替", command=self.toggle_selected).pack(side=LEFT)
        ttk.Button(selection_controls, text="すべて変換", command=lambda: self.set_all_enabled(True)).pack(side=LEFT, padx=(8, 0))
        ttk.Button(selection_controls, text="すべて除外", command=lambda: self.set_all_enabled(False)).pack(side=LEFT, padx=(8, 0))
        ttk.Button(selection_controls, text="履歴消去", command=self.clear_history).pack(side=RIGHT)

        columns = ("enabled", "sheet", "cell", "type", "kind", "original", "replacement", "reason")
        self.tree = ttk.Treeview(container, columns=columns, show="headings", height=16)
        headings = {
            "enabled": "変換",
            "sheet": "シート",
            "cell": "セル",
            "type": "種類",
            "kind": "検査",
            "original": "検出値",
            "replacement": "変換後",
            "reason": "理由",
        }
        widths = {
            "enabled": 55,
            "sheet": 100,
            "cell": 65,
            "type": 70,
            "kind": 80,
            "original": 240,
            "replacement": 120,
            "reason": 210,
        }
        for column in columns:
            self.tree.heading(column, text=headings[column])
            self.tree.column(column, width=widths[column], minwidth=40, stretch=column in {"original", "reason"})
        self.tree.pack(fill=BOTH, expand=True)
        self.tree.bind("<Double-1>", lambda _event: self.toggle_selected())
        self.tree.bind("<Return>", lambda _event: self.toggle_selected())

        edit_panel = ttk.Frame(container)
        edit_panel.pack(fill=X, pady=(8, 8))
        ttk.Label(edit_panel, text="変換後").pack(side=LEFT)
        self.replacement_var = tk.StringVar()
        ttk.Entry(edit_panel, textvariable=self.replacement_var).pack(side=LEFT, fill=X, expand=True, padx=(8, 8))
        ttk.Button(edit_panel, text="選択行へ反映", command=self.apply_replacement_to_selected).pack(side=LEFT)

        history_frame = ttk.LabelFrame(container, text="変換履歴")
        history_frame.pack(fill=BOTH, expand=False)
        self.history = tk.Listbox(history_frame, height=5)
        self.history.pack(fill=BOTH, expand=True)

        self.status_var = tk.StringVar(value="待機中: 外部クラウドへ送信しません。")
        ttk.Label(container, textvariable=self.status_var, relief="sunken", anchor="w").pack(side=TOP, fill=X, pady=(8, 0))

        self.tree.bind("<<TreeviewSelect>>", self._load_selected_replacement)

    def choose_file(self) -> None:
        filename = filedialog.askopenfilename(
            title="検査する Excel を選択",
            filetypes=[("Excel files", "*.xlsx *.xlsm"), ("All files", "*.*")],
        )
        if filename:
            self.set_source(Path(filename))

    def set_source(self, path: Path) -> None:
        if path.suffix.lower() not in SUPPORTED_EXTENSIONS:
            messagebox.showerror("形式エラー", "Python/Presidio 版は .xlsx / .xlsm に対応しています。")
            return
        self.processor.cleanup()
        self.source_path = path
        self.findings = []
        self.item_to_index.clear()
        self.path_var.set(str(path))
        self._refresh_tree()
        self.status_var.set("選択済み: 検査開始を押してください。")

    def scan_file(self) -> None:
        if self.source_path is None:
            messagebox.showwarning("ファイル未選択", "Excel ファイルを選択してください。")
            return
        try:
            self.status_var.set("検査中: Presidio カスタム Recognizer でローカル解析しています...")
            self.update_idletasks()
            self.findings = self.processor.scan(self.source_path)
            self._refresh_tree()
            self.status_var.set(f"検査完了: {len(self.findings)} 件を検出しました。変換対象を確認してください。")
        except Exception as exc:
            self._show_exception("検査エラー", exc)

    def convert_file(self) -> None:
        if self.source_path is None:
            messagebox.showwarning("ファイル未選択", "Excel ファイルを選択してください。")
            return
        try:
            self.status_var.set("変換中: 一時コピーへ置換を適用しています...")
            self.update_idletasks()
            output_path = self.processor.convert(self.source_path, self.findings)
            enabled_count = sum(1 for finding in self.findings if finding.enabled)
            self.history.insert(
                0,
                f"{output_path.name} / {enabled_count} 件変換 / 原本上書きなし / 一時ファイル削除済み",
            )
            self.status_var.set(f"保存完了: {output_path}")
            messagebox.showinfo(
                "保存完了",
                f"匿名化済み Excel を保存しました。\n\n{output_path}\n\n原本は上書きしていません。一時コピーは削除済みです。",
            )
        except Exception as exc:
            self._show_exception("変換エラー", exc)

    def toggle_selected(self) -> None:
        for item in self.tree.selection():
            index = self.item_to_index.get(item)
            if index is None:
                continue
            self.findings[index].enabled = not self.findings[index].enabled
        self._refresh_tree()

    def set_all_enabled(self, enabled: bool) -> None:
        for finding in self.findings:
            finding.enabled = enabled
        self._refresh_tree()

    def apply_replacement_to_selected(self) -> None:
        value = self.replacement_var.get()
        if not value:
            return
        for item in self.tree.selection():
            index = self.item_to_index.get(item)
            if index is not None:
                self.findings[index].replacement = value
        self._refresh_tree()

    def clear_history(self) -> None:
        self.history.delete(0, END)
        self.status_var.set("履歴を消去しました。")

    def _refresh_tree(self) -> None:
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.item_to_index.clear()
        for index, finding in enumerate(self.findings):
            item = self.tree.insert(
                "",
                END,
                values=(
                    "対象" if finding.enabled else "除外",
                    finding.sheet,
                    finding.cell,
                    finding.entity_type,
                    finding.detection_kind,
                    finding.original,
                    finding.replacement,
                    finding.reason,
                ),
            )
            self.item_to_index[item] = index

    def _load_selected_replacement(self, _event: object | None = None) -> None:
        selection = self.tree.selection()
        if not selection:
            return
        index = self.item_to_index.get(selection[0])
        if index is not None:
            self.replacement_var.set(self.findings[index].replacement)

    def _show_exception(self, title: str, exc: Exception) -> None:
        self.status_var.set(title)
        traceback.print_exc()
        messagebox.showerror(title, str(exc))

    def _on_close(self) -> None:
        self.processor.cleanup()
        self.destroy()


def main() -> int:
    app = ExcelPrivacyCleanerApp()
    app.mainloop()
    return 0


if __name__ == "__main__":
    sys.exit(main())
