from __future__ import annotations

import csv
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "tools"))

from excel_privacy_cleaner.pdf_context_rules import context_candidates_for_page  # noqa: E402
from check_pdf_generalization import audit_detection_independence  # noqa: E402


ROOT = Path(__file__).resolve().parents[1]
SPLIT_JSON = ROOT / "config" / "evaluation" / "pdf31_dataset_split_v1.json"
BASELINE_JSON = ROOT / "docs" / "evaluation_baselines" / "PDF匿名化検出_Baseline_v1.json"
BASELINE_CSV = ROOT / "docs" / "evaluation_baselines" / "PDF匿名化検出_Baseline_v1.csv"


def main() -> int:
    split = json.loads(SPLIT_JSON.read_text(encoding="utf-8"))
    assert split["dataset_version"] == "pdf31_dataset_split_v1"
    assert split["baseline_source_commit"] == "adf1fa3"
    assert split["development_pages"] == [5, 6, 10, 11, 13, 15]
    assert split["validation_ocr_eligible_pages"] == [1, 2, 3, 7, 8, 16, 31]
    assert split["validation_ocr_failed_pages"] == [22]
    assert split["final_test_pages"] == [4, 9, 12, 14, 17, 18, 19, 20, 21, 23, 24, 25, 26, 27, 28, 29, 30]
    all_pages = (
        split["development_pages"]
        + split["validation_ocr_eligible_pages"]
        + split["validation_ocr_failed_pages"]
        + split["final_test_pages"]
    )
    assert sorted(all_pages) == list(range(1, 32))
    assert len(set(all_pages)) == 31

    baseline = json.loads(BASELINE_JSON.read_text(encoding="utf-8"))
    assert baseline["baseline_source_commit"] == "adf1fa3"
    assert baseline["baseline_results"]["development_formal_evaluation"]["formal_evaluation_count"] == 49
    assert baseline["baseline_results"]["development_formal_evaluation"]["detected_count"] == 0
    assert baseline["baseline_results"]["development_formal_evaluation"]["missed_count"] == 49
    assert baseline["baseline_results"]["validation_ocr_eligible_initial_candidates"]["generated_candidate_count"] == 0
    assert baseline["baseline_results"]["validation_ocr_failed_page22_initial_candidates"]["generated_candidate_count"] == 0
    with BASELINE_CSV.open(encoding="utf-8-sig", newline="") as handle:
        rows = list(csv.DictReader(handle))
    assert any(row["項目"] == "independence_audit_pass" and row["値"] == "6" for row in rows)

    assert context_candidates_for_page(1, object()) == []
    audit = audit_detection_independence()
    assert len(audit) >= 9
    assert all(value == "PASS" for value in audit.values())
    print("pdf_baseline_v1_tests=passed")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
