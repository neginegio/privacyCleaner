from __future__ import annotations

import csv
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from excel_privacy_cleaner.excel_processor import ExcelPrivacyProcessor  # noqa: E402


def safe(value: str) -> str:
    return str(value).encode("unicode_escape").decode("ascii")


def read_csv(path: Path) -> list[dict[str, str]]:
    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        return list(csv.DictReader(handle))


def main() -> int:
    root = Path.cwd()
    workbook_path = next(path for path in root.glob("*.xlsx") if "202607" not in path.name)
    missed_path = root / "Excel匿名化アプリ_未検出項目一覧.csv"
    wrong_path = root / "Excel匿名化アプリ_テストデータ_検出結果_期待値追記済み.csv"

    findings = ExcelPrivacyProcessor().scan(workbook_path)
    by_cell = {(finding.sheet, finding.cell): [] for finding in findings}
    for finding in findings:
        by_cell.setdefault((finding.sheet, finding.cell), []).append(finding)

    missed_rows = read_csv(missed_path)
    fixed_missed = [row for row in missed_rows if by_cell.get((row["シート"], row["セル"]))]
    still_missed = [row for row in missed_rows if not by_cell.get((row["シート"], row["セル"]))]

    wrong_rows = read_csv(wrong_path)
    excluded_rows = [
        row
        for row in wrong_rows
        if row["期待する判定"] in {"変換対象外", "要確認・自動変換しない"}
    ]
    still_detected = []
    for row in excluded_rows:
        cell_findings = by_cell.get((row["シート"], row["セル"]), [])
        if any(finding.original == row["検出値"] for finding in cell_findings):
            still_detected.append(row)

    print(f"findings={len(findings)}")
    print(f"missed_fixed={len(fixed_missed)}/{len(missed_rows)}")
    print(f"missed_remaining={len(still_missed)}")
    for row in still_missed[:30]:
        print("MISS", safe(row["シート"]), row["セル"], safe(row["項目分類"]), safe(row["未検出値"]))
    print(f"excluded_still_detected={len(still_detected)}/{len(excluded_rows)}")
    for row in still_detected[:30]:
        print("FALSE", safe(row["シート"]), row["セル"], safe(row["問題種別"]), safe(row["検出値"]))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
