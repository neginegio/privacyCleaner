from __future__ import annotations

import csv
import json
import sys
from collections import defaultdict
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from excel_privacy_cleaner.word_processor import WordCandidate, candidates_for_inventory, extract_word_structure  # noqa: E402


ROOT = Path(__file__).resolve().parents[1]
VALIDATION_DOCX = ROOT / "docs" / "evaluation_baselines" / "word_phase2_validation" / "Word匿名化検出_Phase2_ValidationSynthetic.docx"
GROUND_TRUTH_CSV = ROOT / "docs" / "evaluation_baselines" / "word_phase2_validation" / "Word匿名化検出_Phase2_ValidationGroundTruth.csv"
INITIAL_JSON = ROOT / "docs" / "evaluation_baselines" / "Word匿名化検出_Phase2_ValidationInitial.json"
DEVELOPMENT_JSON = ROOT / "docs" / "evaluation_baselines" / "Word匿名化検出_Phase2_DevelopmentBaseline.json"
OUTPUT_JSON = ROOT / "docs" / "evaluation_baselines" / "Word匿名化検出_Phase2_1_Validation.json"
OUTPUT_CSV = ROOT / "docs" / "evaluation_baselines" / "Word匿名化検出_Phase2_1_Validation.csv"
OUTPUT_MD = ROOT / "docs" / "evaluation_baselines" / "Word匿名化検出_Phase2_1_Validation.md"


def main() -> int:
    truths = read_truths(GROUND_TRUTH_CSV)
    inventory = extract_word_structure(VALIDATION_DOCX)
    candidates = candidates_for_inventory(inventory)
    metrics, category_results, remaining_missed, remaining_false_positive = evaluate(candidates, truths)
    initial = json.loads(INITIAL_JSON.read_text(encoding="utf-8"))
    development = json.loads(DEVELOPMENT_JSON.read_text(encoding="utf-8"))
    payload = {
        "baseline_name": "Word candidate detection Phase 2.1 validation",
        "source": "working tree after 4fceb05 Record Word Phase 2 initial validation",
        "status": "Word candidate detection Phase 2.1 frozen",
        "evaluation_context": {
            "development_baseline_metrics": development["metrics"],
            "validation_initial_metrics": initial["metrics"],
            "phase2_1_post_tuning_validation_metrics": metrics,
            "interpretation": (
                "Phase 2.1 validation scores are post-tuning values after analyzing Validation Initial. "
                "They are not unknown Word generalization performance and are not holdout performance."
            ),
            "validation_dataset_status_after_phase2_1": "development/tuning evaluation data",
            "rule_freeze_scope": [
                "word_processor.py candidate generation rules",
                "Presidio-related candidate settings",
                "company/name/address regular expressions",
                "category conflict handling",
                "candidate span handling",
            ],
        },
        "validation_docx_sha256": initial["validation_docx_sha256"],
        "ground_truth_sha256": initial["ground_truth_sha256"],
        "initial_comparison": {
            "initial_metrics": initial["metrics"],
            "phase2_1_metrics": metrics,
            "initial_category_results": initial["category_results"],
            "phase2_1_category_results": category_results,
        },
        "metrics": metrics,
        "category_results": category_results,
        "remaining_missed": remaining_missed,
        "remaining_false_positive": remaining_false_positive,
        "method": {
            "formal_true_positive": "same location_id, normalized category, exact char_start, exact char_end",
            "validation_initial_record_rewritten": False,
        },
    }
    write_json(OUTPUT_JSON, payload)
    write_csv(OUTPUT_CSV, development["metrics"], initial["metrics"], metrics, initial["category_results"], category_results)
    write_markdown(OUTPUT_MD, payload)
    print("word_phase2_1_validation=completed")
    print("metrics", json.dumps(metrics, ensure_ascii=False, sort_keys=True))
    print("category_results", json.dumps(category_results, ensure_ascii=False, sort_keys=True))
    print("remaining_missed", json.dumps(remaining_missed, ensure_ascii=False, sort_keys=True))
    print("remaining_false_positive", json.dumps(remaining_false_positive, ensure_ascii=False, sort_keys=True))
    return 0


def read_truths(path: Path) -> list[dict[str, object]]:
    with path.open(encoding="utf-8-sig", newline="") as input_file:
        return list(csv.DictReader(input_file))


def evaluate(
    candidates: tuple[WordCandidate, ...],
    truths: list[dict[str, object]],
) -> tuple[dict[str, object], dict[str, dict[str, int]], list[dict[str, object]], list[dict[str, object]]]:
    truth_keys = {
        (str(truth["location_id"]), str(truth["category"]), int(truth["char_start"]), int(truth["char_end"]))
        for truth in truths
    }
    candidate_keys = {
        (candidate.location_id, candidate.category, candidate.char_start, candidate.char_end)
        for candidate in candidates
    }
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
    category_results: dict[str, dict[str, int]] = defaultdict(lambda: {"正解": 0, "検出": 0, "見逃し": 0, "TP候補": 0, "FP候補": 0})
    for _location, category, _start, _end in truth_keys:
        category_results[category]["正解"] += 1
    for _location, category, _start, _end in candidate_keys:
        category_results[category]["検出"] += 1
    for _location, category, _start, _end in missed:
        category_results[category]["見逃し"] += 1
    for _location, category, _start, _end in true_positive:
        category_results[category]["TP候補"] += 1
    for _location, category, _start, _end in false_positive:
        category_results[category]["FP候補"] += 1

    truth_by_key = {
        (str(truth["location_id"]), str(truth["category"]), int(truth["char_start"]), int(truth["char_end"])): truth
        for truth in truths
    }
    remaining_missed = [
        {
            "truth_id": str(truth_by_key[key]["truth_id"]),
            "category": key[1],
            "location_id": key[0],
            "char_start": key[2],
            "char_end": key[3],
        }
        for key in sorted(missed)
    ]
    candidate_by_key = {
        (candidate.location_id, candidate.category, candidate.char_start, candidate.char_end): candidate
        for candidate in candidates
    }
    remaining_false_positive = [
        {
            "category": key[1],
            "location_id": key[0],
            "char_start": key[2],
            "char_end": key[3],
            "detection_rule": candidate_by_key[key].detection_rule,
            "confidence": round(candidate_by_key[key].confidence, 3),
        }
        for key in sorted(false_positive)
    ]
    return metrics, dict(sorted(category_results.items())), remaining_missed, remaining_false_positive


def write_json(path: Path, payload: dict[str, object]) -> None:
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def write_csv(
    path: Path,
    development_metrics: dict[str, object],
    initial_metrics: dict[str, object],
    phase2_1_metrics: dict[str, object],
    initial_categories: dict[str, dict[str, int]],
    phase2_1_categories: dict[str, dict[str, int]],
) -> None:
    with path.open("w", encoding="utf-8-sig", newline="") as output:
        writer = csv.writer(output)
        writer.writerow(["scope", "name", "metric", "value"])
        writer.writerow(["context", "status", "phase2_1", "Word candidate detection Phase 2.1 frozen"])
        writer.writerow(["context", "interpretation", "phase2_1", "post-tuning validation; not unknown Word or holdout performance"])
        for metric in (
            "ground_truth_count",
            "matched_ground_truth_count",
            "false_positive_candidate_count",
            "recall",
            "candidate_precision",
        ):
            writer.writerow(["development_baseline", "all", metric, development_metrics[metric]])
        writer.writerow([])
        writer.writerow(["scope", "name", "metric", "initial", "phase2_1"])
        for metric in (
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
            writer.writerow(["overall", "all", metric, initial_metrics[metric], phase2_1_metrics[metric]])
        for category in sorted(set(initial_categories) | set(phase2_1_categories)):
            for metric in ("正解", "検出", "見逃し", "TP候補", "FP候補"):
                writer.writerow(
                    [
                        "category",
                        category,
                        metric,
                        initial_categories.get(category, {}).get(metric, 0),
                        phase2_1_categories.get(category, {}).get(metric, 0),
                    ]
                )


def write_markdown(path: Path, payload: dict[str, object]) -> None:
    context = payload["evaluation_context"]
    development_metrics = context["development_baseline_metrics"]
    comparison = payload["initial_comparison"]
    initial_metrics = comparison["initial_metrics"]
    phase2_1_metrics = comparison["phase2_1_metrics"]
    initial_categories = comparison["initial_category_results"]
    phase2_1_categories = comparison["phase2_1_category_results"]
    lines = [
        "# Word匿名化検出 Phase2.1 Validation",
        "",
        "Validation Initialを上書きせず、同一validation Wordと固定ground truthでPhase 2.1を評価した記録です。",
        "",
        "Phase 2.1の `recall=1.0` / `candidate_precision=1.0` は、Validation Initialの結果を原因分析して1回ルール改善した後のpost-tuning validation値です。未知Wordに対する汎化性能、完全ホールドアウト性能としては扱いません。",
        "",
        "このコミット以降、完全ホールドアウト評価が終了するまで候補生成ルールを凍結します。",
        "",
        "## Development Baseline",
        "",
        "| 指標 | 値 |",
        "|---|---:|",
        f"| ground_truth_count | `{development_metrics['ground_truth_count']}` |",
        f"| matched_ground_truth_count | `{development_metrics['matched_ground_truth_count']}` |",
        f"| false_positive_candidate_count | `{development_metrics['false_positive_candidate_count']}` |",
        f"| recall | `{development_metrics['recall']}` |",
        f"| candidate_precision | `{development_metrics['candidate_precision']}` |",
        "",
        "## Overall",
        "",
        "| 指標 | Validation Initial | Phase 2.1 |",
        "|---|---:|---:|",
    ]
    for metric in (
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
        lines.append(f"| {metric} | `{initial_metrics[metric]}` | `{phase2_1_metrics[metric]}` |")
    lines.extend(["", "## Category Results", "", "| カテゴリ | 指標 | Initial | Phase 2.1 |", "|---|---|---:|---:|"])
    for category in sorted(set(initial_categories) | set(phase2_1_categories)):
        for metric in ("正解", "検出", "見逃し", "TP候補", "FP候補"):
            initial_value = initial_categories.get(category, {}).get(metric, 0)
            phase_value = phase2_1_categories.get(category, {}).get(metric, 0)
            lines.append(f"| {category} | {metric} | `{initial_value}` | `{phase_value}` |")
    lines.extend(
        [
            "",
            "## Notes",
            "",
            "- Phase 2.1はValidationを100%にすることを目的にしていません。",
            "- Phase 2.1はValidation Initialの原因分析後に実施した1回の一般ルール改善結果です。",
            "- Validationデータは今後、Phase 2.1までの開発・調整済み評価データとして扱います。",
            "- Phase 2.1の結果は未知Word性能でも完全ホールドアウト性能でもありません。",
            "- 候補生成コードはvalidation ground truth、truth_id、固定位置、dataset splitを参照していません。",
            "- Phase 3 UI、Word書き換え、holdout評価は実施していません。",
        ]
    )
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")


if __name__ == "__main__":
    raise SystemExit(main())
