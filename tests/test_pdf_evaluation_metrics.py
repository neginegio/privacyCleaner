from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "tools"))

import evaluate_pdf31_development_generic as dev_eval  # noqa: E402


def truth(truth_id: str, rect: tuple[float, float, float, float]) -> dev_eval.TruthRect:
    return dev_eval.TruthRect(
        truth_id=truth_id,
        page=1,
        original_text="synthetic",
        entity_type="会社名",
        mode="常時",
        required="必須",
        structure="synthetic",
        rect=rect,
    )


def candidate(name: str, rect: tuple[float, float, float, float]) -> dev_eval.CandidateRect:
    return dev_eval.CandidateRect(
        page=1,
        original=name,
        normalized=name,
        entity_type="会社名",
        rule_name="synthetic_rule",
        confidence="MEDIUM",
        reason="synthetic",
        rect=rect,
    )


def main() -> int:
    truths = [
        truth("T1", (0, 0, 10, 10)),
        truth("T2", (0, 10, 10, 20)),
        truth("T3", (50, 50, 60, 60)),
    ]
    candidates = [
        candidate("C1", (0, 0, 10, 20)),  # One candidate matches T1 and T2.
        candidate("C2", (0, 0, 10, 10)),  # A second candidate also matches T1.
        candidate("C3", (100, 100, 110, 110)),
    ]
    metrics = dev_eval.evaluate(truths, candidates)
    assert metrics["ground_truth_count"] == 3
    assert metrics["candidate_count"] == 3
    assert metrics["matched_ground_truth_count"] == 2
    assert metrics["missed_ground_truth_count"] == 1
    assert metrics["true_positive_candidate_count"] == 2
    assert metrics["false_positive_candidate_count"] == 1
    assert metrics["matched_ground_truth_count"] + metrics["missed_ground_truth_count"] == metrics["ground_truth_count"]
    assert (
        metrics["true_positive_candidate_count"] + metrics["false_positive_candidate_count"] == metrics["candidate_count"]
    )
    assert 0.0 <= metrics["recall"] <= 1.0
    assert 0.0 <= metrics["candidate_precision"] <= 1.0
    assert metrics["legacy_ground_truth_per_candidate_ratio"] == metrics["matched_ground_truth_count"] / metrics["candidate_count"]
    print("pdf_evaluation_metric_tests=passed")
    return 0


def test_pdf_evaluation_metrics() -> None:
    assert main() == 0


if __name__ == "__main__":
    raise SystemExit(main())
