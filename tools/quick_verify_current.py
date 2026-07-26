from __future__ import annotations

import sys
from pathlib import Path

from openpyxl import load_workbook

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from excel_privacy_cleaner.excel_processor import ExcelPrivacyProcessor


SOURCE = Path("Excel匿名化アプリ_テストデータ.xlsx")
CHECKS = [
    ("01_顧客マスタ", "O2"),
    ("01_顧客マスタ", "O3"),
    ("02_自由記述", "D2"),
    ("02_自由記述", "D5"),
    ("02_自由記述", "E5"),
    ("02_自由記述", "D11"),
    ("04_取引機密", "D3"),
    ("05_表記ゆれ", "C2"),
    ("05_表記ゆれ", "C3"),
    ("05_表記ゆれ", "C4"),
    ("05_表記ゆれ", "C5"),
    ("05_表記ゆれ", "C17"),
    ("05_表記ゆれ", "C19"),
    ("05_表記ゆれ", "C21"),
]


def main() -> int:
    processor = ExcelPrivacyProcessor()
    findings = processor.scan(SOURCE)
    candidates = [finding for finding in findings if finding.detection_kind == "確認候補"]
    print("findings", len(findings), "candidates", len(candidates))
    print("candidate_preview", [(f.sheet, f.cell, f.original, f.replacement) for f in candidates[:12]])

    blocked = False
    if not candidates:
        raise AssertionError("確認候補が検出されませんでした。")
    if candidates:
        candidates[0].enabled = False
        try:
            processor.convert(SOURCE, findings, Path("tmp_verify"))
        except Exception as exc:
            blocked = "未確認候補" in str(exc)
            print("blocked", str(exc)[:300])
    if not blocked:
        raise AssertionError("未確認候補が無効な状態で変換がブロックされませんでした。")

    for finding in findings:
        if finding.detection_kind == "確認候補":
            finding.enabled = True

    output_path = processor.convert(SOURCE, findings, Path("tmp_verify"))
    print("out", output_path)
    workbook = load_workbook(output_path, data_only=False)
    for sheet_name, coordinate in CHECKS:
        print(sheet_name, coordinate, workbook[sheet_name][coordinate].value)
    workbook.close()
    processor.cleanup()
    print("blocked_ok", blocked)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
