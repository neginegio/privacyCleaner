from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "tools"))

import evaluate_pdf31_development_generic as dev_eval  # noqa: E402
from check_pdf_generalization import audit_detection_independence  # noqa: E402


def signature() -> list[tuple[object, ...]]:
    pages = dev_eval.read_split_pages()
    candidates = dev_eval.generate_candidates(pages)
    return [
        (
            candidate.page,
            candidate.entity_type,
            candidate.rule_name,
            tuple(round(value, 1) for value in candidate.rect),
            candidate.normalized,
        )
        for candidate in candidates
    ]


def main() -> int:
    first = signature()
    second = signature()
    assert first == second, "Candidate generation must be deterministic for the same input"
    assert first, "Generic detection should improve on Baseline v1 candidate_count=0"
    audit = audit_detection_independence()
    assert len(audit) == 6
    assert all(value == "PASS" for value in audit.values())
    print("pdf_generic_detection_tests=passed")
    print(f"candidate_count={len(first)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
