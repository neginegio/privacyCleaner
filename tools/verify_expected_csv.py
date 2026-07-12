from __future__ import annotations

import argparse
import csv
import sys
from collections import defaultdict
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from excel_privacy_cleaner.excel_processor import ExcelPrivacyProcessor, ProcessingOptions  # noqa: E402


EXCLUDED_JUDGEMENTS = {"変換対象外", "要確認・自動変換しない"}


def read_csv(path: Path) -> list[dict[str, str]]:
    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        return list(csv.DictReader(handle))


def finding_text(finding) -> str:
    return f"{finding.sheet}!{finding.cell} {finding.entity_type}/{finding.detection_kind}: {finding.original!r} -> {finding.replacement!r}"


def main() -> int:
    parser = argparse.ArgumentParser(description="Verify current scan results against an expected feedback CSV.")
    parser.add_argument("source", type=Path)
    parser.add_argument("expected_csv", type=Path)
    parser.add_argument("--missed-csv", type=Path)
    parser.add_argument("--mode", choices=("analysis", "external"), default="external")
    parser.add_argument("--max-details", type=int, default=40)
    args = parser.parse_args()

    options = ProcessingOptions(mode=args.mode, transform_business_secrets=args.mode == "external")
    findings = ExcelPrivacyProcessor().scan(args.source, options=options)
    by_cell: dict[tuple[str, str], list] = defaultdict(list)
    for finding in findings:
        by_cell[(finding.sheet, finding.cell)].append(finding)

    rows = read_csv(args.expected_csv)
    failures: list[str] = []
    excluded_rows = [row for row in rows if row.get("期待する判定") in EXCLUDED_JUDGEMENTS]
    target_rows = [
        row
        for row in rows
        if row.get("期待する判定", "").startswith("変換対象") and row.get("期待する判定") not in EXCLUDED_JUDGEMENTS
    ]

    for row in target_rows:
        cell_findings = by_cell.get((row["シート"], row["セル"]), [])
        if not any(finding.enabled for finding in cell_findings):
            failures.append(f"TARGET_MISSING {row['シート']}!{row['セル']} expected={row.get('期待する判定', '')}")

    for row in excluded_rows:
        cell_findings = by_cell.get((row["シート"], row["セル"]), [])
        original = row.get("検出値", "")
        bad_matches = [finding for finding in cell_findings if finding.enabled and finding.original == original]
        if bad_matches:
            failures.append(f"EXCLUDED_STILL_ENABLED {finding_text(bad_matches[0])}")

    if args.missed_csv and args.missed_csv.exists():
        missed_rows = read_csv(args.missed_csv)
        for row in missed_rows:
            cell_findings = by_cell.get((row["シート"], row["セル"]), [])
            if not cell_findings:
                failures.append(f"STILL_MISSED {row['シート']}!{row['セル']} value={row.get('未検出値', '')!r}")

    print(f"findings={len(findings)}")
    print(f"expected_rows={len(rows)}")
    print(f"target_rows={len(target_rows)}")
    print(f"excluded_rows={len(excluded_rows)}")
    print(f"failures={len(failures)}")
    for failure in failures[: args.max_details]:
        print(failure)
    return 1 if failures else 0


if __name__ == "__main__":
    raise SystemExit(main())
