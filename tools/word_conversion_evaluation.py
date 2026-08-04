from __future__ import annotations

import sys
from dataclasses import dataclass
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from excel_privacy_cleaner.word_processor import (  # noqa: E402
    WordCandidate,
    WordConversionResult,
    WordPrivacyProcessor,
    WordReplacementDecision,
    extract_word_structure,
    _find_residual_text,
)


TruthKey = tuple[str, str, int, int]


@dataclass(frozen=True)
class ConversionResidualReport:
    conversion_result: WordConversionResult
    converted_matched_candidate_count: int
    convert_internal_residual_warning: bool
    matched_truth_readable_text_residual: tuple[str, ...]
    matched_truth_internal_xml_residual: dict[str, list[str]]
    verdict: str  # "pass" | "fail"


def candidate_key(candidate: WordCandidate) -> TruthKey:
    return (candidate.location_id, candidate.category, candidate.char_start, candidate.char_end)


def enable_all_non_hyperlink_decisions(decisions: list[WordReplacementDecision]) -> list[WordReplacementDecision]:
    # Approves every non-hyperlink candidate, as if a reviewer confirmed each one,
    # mirroring tests/test_word_replacement.py's _enable_review_required.
    for decision in decisions:
        if decision.candidate.source != "hyperlink_target":
            decision.enabled = True
    return decisions


def run_conversion_and_check_residual(
    docx_path: Path,
    truth_keys: set[TruthKey],
    output_dir: Path,
) -> ConversionResidualReport:
    processor = WordPrivacyProcessor()
    decisions = processor.scan(docx_path)
    enable_all_non_hyperlink_decisions(decisions)
    matched = [decision for decision in decisions if decision.enabled and candidate_key(decision.candidate) in truth_keys]

    result = processor.convert(docx_path, decisions, output_dir=output_dir, write_artifacts=False)

    output_inventory = extract_word_structure(result.output_path)
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
