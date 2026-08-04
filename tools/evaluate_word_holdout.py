from __future__ import annotations

import argparse
import csv
import hashlib
import json
import sys
from collections import Counter, defaultdict
from datetime import datetime, timezone
from pathlib import Path


CATEGORIES = ("会社名", "氏名", "住所", "電話番号", "メールアドレス", "銀行名")
FORMAL_STORIES = ("body", "header", "footer")

def main() -> int:
    parser = argparse.ArgumentParser(description="Evaluate Word holdout data without writing sensitive details to git-managed files.")
    parser.add_argument("--holdout-id", required=True)
    parser.add_argument("--phase-commit", default="ca7b01e")
    parser.add_argument("--expected-word-sha256", required=True)
    parser.add_argument("--expected-candidate-sha256", required=True)
    parser.add_argument("--expected-ground-truth-sha256", required=True)
    parser.add_argument("--docx", required=True, type=Path)
    parser.add_argument("--candidates", required=True, type=Path)
    parser.add_argument("--ground-truth", required=True, type=Path)
    parser.add_argument("--private-out-dir", required=True, type=Path)
    parser.add_argument("--docs-out-dir", default=Path("docs/evaluation_baselines"), type=Path)
    parser.add_argument("--public-stem", required=True)
    parser.add_argument(
        "--check-conversion-residual",
        action="store_true",
        help="Also run scan()/convert() on --docx and check for post-conversion residual text (design doc step 10 dimension).",
    )
    args = parser.parse_args()

    word_sha = sha256(args.docx)
    candidate_sha = sha256(args.candidates)
    ground_truth_sha = sha256(args.ground_truth)
    assert_equal(word_sha, args.expected_word_sha256, "Word SHA-256 mismatch")
    assert_equal(candidate_sha, args.expected_candidate_sha256, "candidate SHA-256 mismatch")
    assert_equal(ground_truth_sha, args.expected_ground_truth_sha256, "ground truth SHA-256 mismatch")

    candidates = read_candidates(args.candidates)
    truths = read_truths(args.ground_truth)
    formal_candidates = [candidate for candidate in candidates if is_formal_location(candidate["location_id"])]
    formal_truths = [truth for truth in truths if is_formal_truth(truth)]
    metadata_truths = [truth for truth in truths if truth.get("source") == "document_property" or truth.get("story_type") == "document_property"]

    metrics, category_results, matched_keys, missed_keys, false_positive_keys = evaluate(formal_candidates, formal_truths)
    missed_details, fp_details = analyze_errors(formal_candidates, formal_truths, missed_keys, false_positive_keys)
    missed_cause_counts = dict(sorted(Counter(detail["cause_category"] for detail in missed_details).items()))
    fp_cause_counts = dict(sorted(Counter(detail["cause_category"] for detail in fp_details).items()))

    conversion_residual_private: dict[str, object] | None = None
    conversion_residual_public: dict[str, object] = {
        "conversion_residual_check_performed": False,
        "conversion_residual_verdict": "not_run",
        "converted_matched_candidate_count": 0,
    }
    if args.check_conversion_residual:
        sys.path.insert(0, str(Path(__file__).resolve().parent))
        from word_conversion_evaluation import run_conversion_and_check_residual  # noqa: E402 (deliberately lazy, see module docstring note in the plan)

        truth_keys = {
            (str(truth["location_id"]), normalize_category(str(truth["category"])), int(truth["char_start"]), int(truth["char_end"]))
            for truth in truths
        }
        conversion_check_dir = args.private_out_dir / "conversion_check"
        conversion_check_dir.mkdir(parents=True, exist_ok=True)
        report = run_conversion_and_check_residual(args.docx, truth_keys, conversion_check_dir)
        conversion_residual_private = {
            "verdict": report.verdict,
            "blocked_guard": report.blocked_guard,
            "converted_matched_candidate_count": report.converted_matched_candidate_count,
            "convert_internal_residual_warning": report.convert_internal_residual_warning,
            "matched_truth_readable_text_residual_count": len(report.matched_truth_readable_text_residual),
            "matched_truth_internal_xml_residual_parts": sorted(report.matched_truth_internal_xml_residual),
            "warnings": list(report.conversion_result.warnings) if report.conversion_result else [],
        }
        conversion_residual_public = {
            "conversion_residual_check_performed": True,
            "conversion_residual_verdict": report.verdict,
            "conversion_residual_blocked_guard": report.blocked_guard,
            "converted_matched_candidate_count": report.converted_matched_candidate_count,
        }

    args.private_out_dir.mkdir(parents=True, exist_ok=True)
    private_prefix = args.public_stem.replace("Word匿名化検出_", "")
    private_details_path = args.private_out_dir / f"{private_prefix}_evaluation_error_details_private.csv"
    write_private_details(private_details_path, missed_details, fp_details)
    private_summary_path = args.private_out_dir / f"{private_prefix}_evaluation_summary_private.json"
    private_summary_path.write_text(
        json.dumps(
            {
                "word_sha256": word_sha,
                "candidate_sha256": candidate_sha,
                "ground_truth_sha256": ground_truth_sha,
                "private_error_details_sha256": sha256(private_details_path),
                "metrics": metrics,
                "category_results": category_results,
                "missed_cause_counts": missed_cause_counts,
                "fp_cause_counts": fp_cause_counts,
                "conversion_residual_check": conversion_residual_private,
            },
            ensure_ascii=False,
            indent=2,
        )
        + "\n",
        encoding="utf-8",
    )

    public_payload = {
        "baseline_name": f"Word candidate detection {args.holdout_id}",
        "holdout_id": args.holdout_id,
        "evaluated_at_utc": datetime.now(timezone.utc).isoformat(),
        "phase2_1_frozen_commit": args.phase_commit,
        "word_sha256": word_sha,
        "candidate_sha256": candidate_sha,
        "ground_truth_sha256": ground_truth_sha,
        "candidate_count_fixed_before_evaluation": len(candidates),
        "ground_truth_count_completed_before_evaluation": len(truths),
        "formal_scope": {
            "included_story_types": list(FORMAL_STORIES),
            "included_container_types": ["paragraph", "table_cell"],
            "excluded": ["document_property", "unsupported Word regions"],
        },
        "matching_method": {
            "true_positive": "same location_id, normalized category, exact char_start, exact char_end",
            "partial_match_is_true_positive": False,
            "one_candidate_can_match_multiple_truths": False,
        },
        "metrics": metrics,
        "category_results": category_results,
        "missed_cause_counts": missed_cause_counts,
        "false_positive_cause_counts": fp_cause_counts,
        "metadata_ground_truth_count": len(metadata_truths),
        "unsupported_handling": "Unsupported regions are recorded as a separate safety issue and are not included in the Phase 2.1 formal recall denominator.",
        "evaluation_interpretation": evaluation_interpretation(args.holdout_id, metrics, category_results, missed_cause_counts),
        "privacy_note": "This public record intentionally excludes original text, candidate text, ground truth text, and detailed spans containing real information.",
        "private_error_details_sha256": sha256(private_details_path),
        **conversion_residual_public,
    }
    args.docs_out_dir.mkdir(parents=True, exist_ok=True)
    write_public_records(args.docs_out_dir, args.public_stem, public_payload)
    print(json.dumps(public_payload, ensure_ascii=False, indent=2))
    return 0


def assert_equal(actual: str, expected: str, message: str) -> None:
    if actual.lower() != expected.lower():
        raise AssertionError(f"{message}: actual={actual} expected={expected}")


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as input_file:
        for chunk in iter(lambda: input_file.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def read_candidates(path: Path) -> list[dict[str, object]]:
    with path.open(encoding="utf-8-sig", newline="") as input_file:
        rows = []
        for row in csv.DictReader(input_file):
            row["char_start"] = int(str(row["char_start"]))
            row["char_end"] = int(str(row["char_end"]))
            rows.append(row)
        return rows


def read_truths(path: Path) -> list[dict[str, object]]:
    with path.open(encoding="utf-8-sig", newline="") as input_file:
        rows = []
        for row in csv.DictReader(input_file):
            row["char_start"] = int(str(row["char_start"]))
            row["char_end"] = int(str(row["char_end"]))
            row["category"] = normalize_category(str(row["category"]))
            rows.append(row)
        return rows


def is_formal_location(location_id: str) -> bool:
    return any(location_id.startswith(f"{story}:") for story in FORMAL_STORIES)


def is_formal_truth(truth: dict[str, object]) -> bool:
    return str(truth.get("story_type", "")) in FORMAL_STORIES and str(truth.get("container_type", "")) in {"paragraph", "table_cell"}


def normalize_category(category: str) -> str:
    return category.strip()


def evaluate(
    candidates: list[dict[str, object]],
    truths: list[dict[str, object]],
) -> tuple[dict[str, object], dict[str, dict[str, int]], set[tuple[str, str, int, int]], set[tuple[str, str, int, int]], set[tuple[str, str, int, int]]]:
    truth_keys = {(str(t["location_id"]), normalize_category(str(t["category"])), int(t["char_start"]), int(t["char_end"])) for t in truths}
    candidate_keys = {(str(c["location_id"]), normalize_category(str(c["category"])), int(c["char_start"]), int(c["char_end"])) for c in candidates}
    matched = truth_keys & candidate_keys
    missed = truth_keys - candidate_keys
    true_positive = candidate_keys & truth_keys
    false_positive = candidate_keys - truth_keys
    metrics = {
        "ground_truth_count": len(truth_keys),
        "candidate_count": len(candidate_keys),
        "matched_ground_truth_count": len(matched),
        "missed_ground_truth_count": len(missed),
        "true_positive_candidate_count": len(true_positive),
        "false_positive_candidate_count": len(false_positive),
        "recall": round(len(matched) / len(truth_keys), 3) if truth_keys else 0.0,
        "candidate_precision": round(len(true_positive) / len(candidate_keys), 3) if candidate_keys else 0.0,
        "exact_span_match_count": len(matched),
    }
    category_results: dict[str, dict[str, int]] = {category: {"ground_truth_count": 0, "candidate_count": 0, "matched": 0, "missed": 0, "true_positive_candidate": 0, "false_positive_candidate": 0} for category in CATEGORIES}
    for _location, category, _start, _end in truth_keys:
        category_results.setdefault(category, default_category_result())["ground_truth_count"] += 1
    for _location, category, _start, _end in candidate_keys:
        category_results.setdefault(category, default_category_result())["candidate_count"] += 1
    for _location, category, _start, _end in matched:
        category_results.setdefault(category, default_category_result())["matched"] += 1
    for _location, category, _start, _end in missed:
        category_results.setdefault(category, default_category_result())["missed"] += 1
    for _location, category, _start, _end in true_positive:
        category_results.setdefault(category, default_category_result())["true_positive_candidate"] += 1
    for _location, category, _start, _end in false_positive:
        category_results.setdefault(category, default_category_result())["false_positive_candidate"] += 1
    return metrics, dict(sorted(category_results.items())), matched, missed, false_positive


def default_category_result() -> dict[str, int]:
    return {"ground_truth_count": 0, "candidate_count": 0, "matched": 0, "missed": 0, "true_positive_candidate": 0, "false_positive_candidate": 0}


def analyze_errors(
    candidates: list[dict[str, object]],
    truths: list[dict[str, object]],
    missed_keys: set[tuple[str, str, int, int]],
    false_positive_keys: set[tuple[str, str, int, int]],
) -> tuple[list[dict[str, object]], list[dict[str, object]]]:
    candidates_by_location = defaultdict(list)
    truths_by_location = defaultdict(list)
    for candidate in candidates:
        candidates_by_location[str(candidate["location_id"])].append(candidate)
    for truth in truths:
        truths_by_location[str(truth["location_id"])].append(truth)

    missed_details = []
    for location_id, category, start, end in sorted(missed_keys):
        related = candidates_by_location.get(location_id, [])
        cause = classify_missed(category, start, end, related)
        missed_details.append(
            {
                "kind": "missed",
                "category": category,
                "location_id": location_id,
                "char_start": start,
                "char_end": end,
                "cause_category": cause,
                "related_candidate_count_same_location": len(related),
            }
        )
    fp_details = []
    for location_id, category, start, end in sorted(false_positive_keys):
        related_truths = truths_by_location.get(location_id, [])
        cause = classify_false_positive(category, start, end, related_truths)
        fp_details.append(
            {
                "kind": "false_positive",
                "category": category,
                "location_id": location_id,
                "char_start": start,
                "char_end": end,
                "cause_category": cause,
                "related_truth_count_same_location": len(related_truths),
            }
        )
    return missed_details, fp_details


def classify_missed(category: str, start: int, end: int, candidates: list[dict[str, object]]) -> str:
    if not candidates:
        return "候補自体が生成されなかった"
    overlapping = [candidate for candidate in candidates if ranges_overlap(start, end, int(candidate["char_start"]), int(candidate["char_end"]))]
    if not overlapping:
        return "候補自体が生成されなかった"
    if any(normalize_category(str(candidate["category"])) != category for candidate in overlapping):
        return "カテゴリ誤分類"
    if any(int(candidate["char_start"]) > start or int(candidate["char_end"]) < end for candidate in overlapping):
        return "span不足"
    if any(int(candidate["char_start"]) < start or int(candidate["char_end"]) > end for candidate in overlapping):
        return "span過大"
    return "その他"


def classify_false_positive(category: str, start: int, end: int, truths: list[dict[str, object]]) -> str:
    if not truths:
        return "一般語誤検出"
    overlapping = [truth for truth in truths if ranges_overlap(start, end, int(truth["char_start"]), int(truth["char_end"]))]
    if not overlapping:
        return "一般語誤検出"
    if any(normalize_category(str(truth["category"])) != category for truth in overlapping):
        return "カテゴリ誤分類"
    if any(start > int(truth["char_start"]) or end < int(truth["char_end"]) for truth in overlapping):
        return "span不足"
    if any(start < int(truth["char_start"]) or end > int(truth["char_end"]) for truth in overlapping):
        return "span過大"
    return "重複候補"


def ranges_overlap(left_start: int, left_end: int, right_start: int, right_end: int) -> bool:
    return left_start < right_end and right_start < left_end


def write_private_details(path: Path, missed_details: list[dict[str, object]], fp_details: list[dict[str, object]]) -> None:
    fieldnames = [
        "kind",
        "category",
        "location_id",
        "char_start",
        "char_end",
        "cause_category",
        "related_candidate_count_same_location",
        "related_truth_count_same_location",
    ]
    with path.open("w", encoding="utf-8-sig", newline="") as output_file:
        writer = csv.DictWriter(output_file, fieldnames=fieldnames)
        writer.writeheader()
        for detail in missed_details + fp_details:
            writer.writerow(detail)


def evaluation_interpretation(
    holdout_id: str,
    metrics: dict[str, object],
    category_results: dict[str, dict[str, int]],
    missed_cause_counts: dict[str, int],
) -> list[str]:
    missed_count = int(metrics["missed_ground_truth_count"])
    no_candidate_count = int(missed_cause_counts.get("候補自体が生成されなかった", 0))
    company = category_results.get("会社名", default_category_result())
    person = category_results.get("氏名", default_category_result())
    return [
        f"{holdout_id} is a completely unused real Word document evaluated after the Phase 2.1 frozen commit.",
        "Initial candidates were generated and fixed before ground truth creation.",
        "Human ground truth was created independently without viewing candidate contents.",
        "Candidate and ground truth were compared for the first time only after ground truth completion.",
        "This is an unknown-real-document evaluation for Phase 2.1.",
        "Development and post-tuning Validation 1.0/1.0 results must be kept separate from this holdout result.",
        "A single holdout document is not enough to make a general claim about Word-wide performance.",
        f"Company names and personal names dropped substantially: company matched {company['matched']}/{company['ground_truth_count']}, person matched {person['matched']}/{person['ground_truth_count']}.",
        f"{no_candidate_count} of {missed_count} missed ground truths had no generated candidate, so the drop is not explained only by strict exact-span matching.",
    ]


def write_public_records(out_dir: Path, public_stem: str, payload: dict[str, object]) -> None:
    base = out_dir / public_stem
    (base.with_suffix(".json")).write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    write_public_csv(base.with_suffix(".csv"), payload)
    write_public_md(base.with_suffix(".md"), payload)


def write_public_csv(path: Path, payload: dict[str, object]) -> None:
    with path.open("w", encoding="utf-8-sig", newline="") as output_file:
        writer = csv.writer(output_file)
        writer.writerow(["scope", "name", "metric", "value"])
        for key, value in payload["metrics"].items():
            writer.writerow(["overall", "all", key, value])
        for category, values in payload["category_results"].items():
            for key, value in values.items():
                writer.writerow(["category", category, key, value])
        for cause, count in payload["missed_cause_counts"].items():
            writer.writerow(["missed_cause", cause, "count", count])
        for cause, count in payload["false_positive_cause_counts"].items():
            writer.writerow(["false_positive_cause", cause, "count", count])
        writer.writerow(["metadata", "document_property", "ground_truth_count", payload["metadata_ground_truth_count"]])
        writer.writerow(["conversion_residual_check", "all", "performed", payload.get("conversion_residual_check_performed", False)])
        writer.writerow(["conversion_residual_check", "all", "verdict", payload.get("conversion_residual_verdict", "not_run")])
        writer.writerow(["conversion_residual_check", "all", "blocked_guard", payload.get("conversion_residual_blocked_guard", "")])
        writer.writerow(["conversion_residual_check", "all", "converted_matched_candidate_count", payload.get("converted_matched_candidate_count", 0)])


def write_public_md(path: Path, payload: dict[str, object]) -> None:
    metrics = payload["metrics"]
    lines = [
        f"# Word匿名化検出 {payload['holdout_id']}",
        "",
        f"Phase 2.1 frozen候補と、独立作成した{payload['holdout_id']} ground truthを初めて比較した正式ホールドアウト評価記録です。",
        "",
        "この公開記録には、実在する会社名、氏名、銀行名、本文、候補文字列、ground truth文字列を含めません。",
        "",
        "## 固定対象",
        "",
        f"- Phase 2.1 commit: `{payload['phase2_1_frozen_commit']}`",
        f"- Word SHA-256: `{payload['word_sha256']}`",
        f"- candidate SHA-256: `{payload['candidate_sha256']}`",
        f"- ground truth SHA-256: `{payload['ground_truth_sha256']}`",
        "",
        "## 評価上の解釈",
        "",
    ]
    for note in payload["evaluation_interpretation"]:
        lines.append(f"- {note}")
    lines.extend(["", "## Overall", "", "| 指標 | 値 |", "|---|---:|"])
    for key in (
        "ground_truth_count",
        "candidate_count",
        "matched_ground_truth_count",
        "missed_ground_truth_count",
        "true_positive_candidate_count",
        "false_positive_candidate_count",
        "recall",
        "candidate_precision",
        "exact_span_match_count",
    ):
        lines.append(f"| {key} | `{metrics[key]}` |")
    lines.extend(["", "## Category Results", "", "| カテゴリ | GT | Candidates | Matched | Missed | TP Candidate | FP Candidate |", "|---|---:|---:|---:|---:|---:|---:|"])
    for category, values in payload["category_results"].items():
        lines.append(
            f"| {category} | `{values['ground_truth_count']}` | `{values['candidate_count']}` | `{values['matched']}` | `{values['missed']}` | `{values['true_positive_candidate']}` | `{values['false_positive_candidate']}` |"
        )
    lines.extend(["", "## Cause Summary", "", "### Missed", "", "| 原因カテゴリ | 件数 |", "|---|---:|"])
    if payload["missed_cause_counts"]:
        for cause, count in payload["missed_cause_counts"].items():
            lines.append(f"| {cause} | `{count}` |")
    else:
        lines.append("| なし | `0` |")
    lines.extend(["", "### False Positive", "", "| 原因カテゴリ | 件数 |", "|---|---:|"])
    if payload["false_positive_cause_counts"]:
        for cause, count in payload["false_positive_cause_counts"].items():
            lines.append(f"| {cause} | `{count}` |")
    else:
        lines.append("| なし | `0` |")
    lines.extend(["", "## Conversion Residual Check(置換後残存なし)", ""])
    if payload.get("conversion_residual_check_performed"):
        lines.append(f"- verdict: `{payload['conversion_residual_verdict']}`")
        if payload.get("conversion_residual_blocked_guard"):
            lines.append(f"- blocked_guard: `{payload['conversion_residual_blocked_guard']}`")
        lines.append(f"- converted_matched_candidate_count: `{payload['converted_matched_candidate_count']}`")
    else:
        lines.append("- 未実施")
    lines.extend(
        [
            "",
            "## Notes",
            "",
            f"- metadata ground truth count: `{payload['metadata_ground_truth_count']}`",
            "- document propertiesとunsupported領域は本文系Phase 2.1評価のRecall/Precisionへ混ぜていません。",
            "- 評価結果を見た後の候補生成ルール変更は実施していません。",
        ]
    )
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")


if __name__ == "__main__":
    raise SystemExit(main())
