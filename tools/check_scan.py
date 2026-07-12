from __future__ import annotations

import sys
from pathlib import Path

from openpyxl import load_workbook

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from excel_privacy_cleaner.excel_processor import ExcelPrivacyProcessor  # noqa: E402


def safe(value: object) -> str:
    return str(value).encode("unicode_escape").decode("ascii")


def main() -> int:
    workbook_path = next(path for path in Path.cwd().glob("*.xlsx") if "202607" not in path.name)
    workbook = load_workbook(workbook_path, data_only=True)
    for sheet in workbook.worksheets:
        print("##", safe(sheet.title))
        print([(cell.coordinate, safe(cell.value)) for cell in sheet[1]])
    workbook.close()

    findings = ExcelPrivacyProcessor().scan(workbook_path)
    print("finding_count", len(findings))
    for finding in findings:
        print(
            safe(finding.sheet),
            finding.cell,
            safe(finding.entity_type),
            safe(finding.detection_kind),
            safe(finding.original),
            "->",
            safe(finding.replacement),
        )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
