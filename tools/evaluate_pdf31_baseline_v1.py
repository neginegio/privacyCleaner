from __future__ import annotations

import csv
import json
import sys
import time
from pathlib import Path
from typing import Any

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from excel_privacy_cleaner.pdf_processor import PdfPrivacyProcessor, final_output_status  # noqa: E402


ROOT = Path(__file__).resolve().parents[1]
SOURCE = ROOT / "PDFテストデータ２.pdf"
OUTPUT_DIR = ROOT / "ocr_quality_outputs"
DATASET_SPLIT_JSON = ROOT / "config" / "evaluation" / "pdf31_dataset_split_v1.json"
VALIDATION_GT_CSV = OUTPUT_DIR / "PDF31通常検証7ページ_ユーザー確認済み正解データ_v1.csv"
FAILED_GT_CSV = OUTPUT_DIR / "PDF31ページ22_FAILED検証_ユーザー確認済み正解データ_v1.csv"
VALIDATION_EVAL_CSV = OUTPUT_DIR / "PDF31通常検証7ページ_Baseline_v1_初回評価.csv"
FAILED_EVAL_CSV = OUTPUT_DIR / "PDF31ページ22_FAILED検証_Baseline_v1_初回評価.csv"


def write_csv(path: Path, rows: list[dict[str, Any]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    keys: list[str] = []
    for row in rows:
        for key in row:
            if key not in keys:
                keys.append(key)
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=keys)
        writer.writeheader()
        writer.writerows(rows)


def read_split() -> dict[str, Any]:
    return json.loads(DATASET_SPLIT_JSON.read_text(encoding="utf-8"))


def read_confirmed_truth_count(path: Path) -> tuple[str, int]:
    if not path.exists():
        return "PENDING_GROUND_TRUTH", 0
    with path.open(encoding="utf-8-sig", newline="") as handle:
        rows = list(csv.DictReader(handle))
    return "READY", len(rows)


def page_number_from_sheet(sheet: str) -> int | None:
    if not sheet.startswith("ページ"):
        return None
    try:
        return int(sheet.replace("ページ", ""))
    except ValueError:
        return None


def count_candidates(findings: list[Any], pages: list[int]) -> int:
    return sum(1 for finding in findings if page_number_from_sheet(str(finding.sheet)) in pages)


def main() -> int:
    split = read_split()
    validation_pages = [int(page) for page in split["validation_ocr_eligible_pages"]]
    failed_pages = [int(page) for page in split["validation_ocr_failed_pages"]]

    start = time.perf_counter()
    processor = PdfPrivacyProcessor()
    findings = processor.scan(SOURCE)
    elapsed = time.perf_counter() - start

    validation_status, validation_truth_count = read_confirmed_truth_count(VALIDATION_GT_CSV)
    validation_candidate_count = count_candidates(findings, validation_pages)
    validation_rows = [
        {
            "dataset_version": split["dataset_version"],
            "baseline_source_commit": split["baseline_source_commit"],
            "evaluation_group": "validation_ocr_eligible_pages",
            "ground_truth_status": validation_status,
            "pages": ",".join(str(page) for page in validation_pages),
            "truth_rect_count": validation_truth_count,
            "detected_count": 0 if validation_status == "READY" else "NOT_EVALUATED",
            "missed_count": validation_truth_count if validation_status == "READY" else "NOT_EVALUATED",
            "false_positive_count": validation_candidate_count if validation_status == "READY" else "NOT_EVALUATED",
            "recall": "0.000" if validation_status == "READY" and validation_truth_count else "NOT_EVALUATED",
            "precision": "0.000" if validation_status == "READY" and validation_candidate_count else "NOT_EVALUATED",
            "average_coverage": "0.000" if validation_status == "READY" and validation_truth_count else "NOT_EVALUATED",
            "candidate_count": validation_candidate_count,
            "elapsed_seconds": f"{elapsed:.3f}",
            "note": "Baseline v1候補0件を、ユーザー確認済み正解データに対して評価する。ページ22は含めない。",
        }
    ]
    write_csv(VALIDATION_EVAL_CSV, validation_rows)

    failed_status, failed_truth_count = read_confirmed_truth_count(FAILED_GT_CSV)
    failed_candidate_count = count_candidates(findings, failed_pages)
    failed_rows: list[dict[str, Any]] = []
    can_output, reasons = final_output_status(findings, processor.page_quality, processor.confirmed_pages, processor.page_review_state)
    for page in failed_pages:
        quality = processor.page_quality.get(page - 1)
        failed_rows.append(
            {
                "dataset_version": split["dataset_version"],
                "baseline_source_commit": split["baseline_source_commit"],
                "evaluation_group": "validation_ocr_failed_pages",
                "ground_truth_status": failed_status,
                "page": page,
                "truth_rect_count": failed_truth_count,
                "candidate_count": failed_candidate_count,
                "ocr_quality_verdict": quality.verdict if quality else "NOT_EVALUATED",
                "failed_page_detected": "YES" if quality and quality.verdict == "FAILED" else "NO",
                "final_output_blocked": "YES" if not can_output else "NO",
                "manual_review_required": "YES" if any("FAILED" in reason or "UNREVIEWED" in reason for reason in reasons) else "NO",
                "completed_without_review": "NO" if not can_output else "YES",
                "note": "ページ22は通常OCR可能ページのrecallと混在させない。",
            }
        )
    write_csv(FAILED_EVAL_CSV, failed_rows)

    print(f"validation_ground_truth_status={validation_status}")
    print(f"validation_candidate_count={validation_candidate_count}")
    print(f"page22_ground_truth_status={failed_status}")
    print(f"page22_candidate_count={failed_candidate_count}")
    print(f"validation_eval_csv={VALIDATION_EVAL_CSV}")
    print(f"failed_eval_csv={FAILED_EVAL_CSV}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
