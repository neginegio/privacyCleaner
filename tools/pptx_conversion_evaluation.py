"""Core evaluation helpers for PPTX candidate detection/conversion, mirroring
tools/word_conversion_evaluation.py's design.

This is groundwork only (per docs/pptx_anonymization_v1_design.md's
"テストデータ分割" section): no dev/validation/holdout PPTX files or ground
truth exist yet, so there is no CLI entry point here. Once real data is
available, build tools/evaluate_pptx_holdout.py on top of
run_conversion_and_check_residual() below, following the same shape as
tools/evaluate_word_holdout.py (candidate_key-based ground-truth matching,
guard classification instead of raw exception text, recall/precision from
the metric set the design doc defines).
"""

from __future__ import annotations

import sys
from dataclasses import dataclass
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from excel_privacy_cleaner.pptx_processor import (  # noqa: E402
    PptxCandidate,
    PptxConversionResult,
    PptxPrivacyProcessor,
    PptxReplacementDecision,
    extract_pptx_structure,
    _find_residual_text,
)


TruthKey = tuple[str, str, int, int]


@dataclass(frozen=True)
class ConversionResidualReport:
    conversion_result: PptxConversionResult | None
    converted_matched_candidate_count: int
    convert_internal_residual_warning: bool
    matched_truth_readable_text_residual: tuple[str, ...]
    matched_truth_internal_xml_residual: dict[str, list[str]]
    verdict: str  # "pass" | "fail" | "blocked"
    blocked_guard: str = ""  # short classification only, never the raw guard message (which may embed candidate text)


_GUARD_CLASSIFIERS = (
    ("先に検査を実行してください。", "not_scanned"),
    ("未確認候補が", "review_required_guard"),
    ("未対応領域が検出されたため外部共有用の出力を停止しました。", "unsupported_features_guard_external"),
    ("構造が変化しました。再スキャンしてください。", "integrity_guard"),
    ("検出候補の範囲が重複しています", "overlap_guard"),
    ("の検出候補の範囲が重複しています。", "overlap_guard_property"),
    ("内部XML残存検査で匿名化対象の原文が検出されたため外部共有用の出力を停止しました。", "residual_guard_external"),
)


def _classify_guard(message: str) -> str:
    for prefix, tag in _GUARD_CLASSIFIERS:
        if prefix in message:
            return tag
    return "unknown_guard"


def candidate_key(candidate: PptxCandidate) -> TruthKey:
    return (candidate.location_id, candidate.category, candidate.char_start, candidate.char_end)


def enable_all_decisions(decisions: list[PptxReplacementDecision]) -> list[PptxReplacementDecision]:
    # Approves every candidate, as if a reviewer confirmed each one,
    # mirroring tests/test_pptx_replacement.py's _enable_review_required.
    for decision in decisions:
        decision.enabled = True
    return decisions


def run_conversion_and_check_residual(
    pptx_path: Path,
    truth_keys: set[TruthKey],
    output_dir: Path,
) -> ConversionResidualReport:
    processor = PptxPrivacyProcessor()
    decisions = processor.scan(pptx_path)
    enable_all_decisions(decisions)
    matched = [decision for decision in decisions if decision.enabled and candidate_key(decision.candidate) in truth_keys]

    try:
        result = processor.convert(pptx_path, decisions, output_dir=output_dir, write_artifacts=False)
    except RuntimeError as exc:
        # A guard correctly refused to convert. This is itself a valid,
        # honest evaluation outcome -- record it, never the raw message
        # (which may embed candidate original text), and never route around
        # it by picking a candidate to disable on the evaluator's behalf.
        return ConversionResidualReport(
            conversion_result=None,
            converted_matched_candidate_count=len(matched),
            convert_internal_residual_warning=False,
            matched_truth_readable_text_residual=(),
            matched_truth_internal_xml_residual={},
            verdict="blocked",
            blocked_guard=_classify_guard(str(exc)),
        )

    output_inventory = extract_pptx_structure(result.output_path)
    combined_text = "\n".join(paragraph.text for paragraph in output_inventory.paragraphs)
    readable_residual = tuple(sorted({decision.candidate.text for decision in matched if decision.candidate.text in combined_text}))

    xml_checked = [decision for decision in matched if decision.candidate.source in ("paragraph", "document_property")]
    xml_residual = {part: sorted(categories) for part, categories in _find_residual_text(result.output_path, xml_checked).items()}

    warning_residual = any("残存" in warning for warning in result.warnings)
    verdict = "pass" if not readable_residual and not xml_residual and not warning_residual else "fail"

    return ConversionResidualReport(
        conversion_result=result,
        converted_matched_candidate_count=len(matched),
        convert_internal_residual_warning=warning_residual,
        matched_truth_readable_text_residual=readable_residual,
        matched_truth_internal_xml_residual=xml_residual,
        verdict=verdict,
    )


def compute_metrics(
    ground_truth_keys: set[TruthKey],
    candidate_keys: set[TruthKey],
) -> dict[str, float | int]:
    """The metric set docs/pptx_anonymization_v1_design.md's "評価方法"
    section defines. Structural-position + char-range + category match
    (via TruthKey), not PDF-style bounding-box coverage.
    """
    matched = ground_truth_keys & candidate_keys
    ground_truth_count = len(ground_truth_keys)
    candidate_count = len(candidate_keys)
    matched_ground_truth_count = len(matched)
    missed_ground_truth_count = ground_truth_count - matched_ground_truth_count
    true_positive_candidate_count = matched_ground_truth_count
    false_positive_candidate_count = candidate_count - true_positive_candidate_count
    recall = matched_ground_truth_count / ground_truth_count if ground_truth_count else 0.0
    candidate_precision = true_positive_candidate_count / candidate_count if candidate_count else 0.0
    return {
        "ground_truth_count": ground_truth_count,
        "candidate_count": candidate_count,
        "matched_ground_truth_count": matched_ground_truth_count,
        "missed_ground_truth_count": missed_ground_truth_count,
        "true_positive_candidate_count": true_positive_candidate_count,
        "false_positive_candidate_count": false_positive_candidate_count,
        "recall": recall,
        "candidate_precision": candidate_precision,
    }
