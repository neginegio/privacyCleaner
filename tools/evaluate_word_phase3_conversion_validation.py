from __future__ import annotations

import csv
import json
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))
sys.path.insert(0, str(Path(__file__).resolve().parent))

from excel_privacy_cleaner.word_processor import candidates_for_inventory, extract_word_structure  # noqa: E402
from word_conversion_evaluation import run_conversion_and_check_residual  # noqa: E402


ROOT = Path(__file__).resolve().parents[1]
VALIDATION_DOCX = ROOT / "docs" / "evaluation_baselines" / "word_phase2_validation" / "Word匿名化検出_Phase2_ValidationSynthetic.docx"
GROUND_TRUTH_CSV = ROOT / "docs" / "evaluation_baselines" / "word_phase2_validation" / "Word匿名化検出_Phase2_ValidationGroundTruth.csv"
FROZEN_PHASE2_1_JSON = ROOT / "docs" / "evaluation_baselines" / "Word匿名化検出_Phase2_1_Validation.json"
OUTPUT_JSON = ROOT / "docs" / "evaluation_baselines" / "Word匿名化検出_Phase3_ConversionValidation.json"
OUTPUT_CSV = ROOT / "docs" / "evaluation_baselines" / "Word匿名化検出_Phase3_ConversionValidation.csv"
OUTPUT_MD = ROOT / "docs" / "evaluation_baselines" / "Word匿名化検出_Phase3_ConversionValidation.md"


def main() -> int:
    truths = read_truths(GROUND_TRUTH_CSV)
    truth_keys = {(str(t["location_id"]), str(t["category"]), int(t["char_start"]), int(t["char_end"])) for t in truths}

    frozen = json.loads(FROZEN_PHASE2_1_JSON.read_text(encoding="utf-8"))
    frozen_metrics = frozen["metrics"]

    inventory = extract_word_structure(VALIDATION_DOCX)
    candidates = candidates_for_inventory(inventory)
    candidate_keys = {(c.location_id, c.category, c.char_start, c.char_end) for c in candidates}
    matched = truth_keys & candidate_keys
    reproduced_metrics = {
        "ground_truth_count": len(truth_keys),
        "candidate_count": len(candidate_keys),
        "matched_ground_truth_count": len(matched),
        "missed_ground_truth_count": len(truth_keys - candidate_keys),
        "true_positive_candidate_count": len(candidate_keys & truth_keys),
        "false_positive_candidate_count": len(candidate_keys - truth_keys),
        "recall": round(len(matched) / len(truth_keys), 3) if truth_keys else 0.0,
        "candidate_precision": round(len(matched) / len(candidate_keys), 3) if candidate_keys else 0.0,
        "exact_span_match_count": len(matched),
    }
    reproduction_matches = reproduced_metrics == frozen_metrics
    if not reproduction_matches:
        raise AssertionError(
            "Detection metrics no longer reproduce the frozen Phase 2.1 result. "
            f"frozen={frozen_metrics} reproduced={reproduced_metrics}"
        )

    with tempfile.TemporaryDirectory(prefix="word_phase3_conversion_validation_") as tmpdir:
        report = run_conversion_and_check_residual(VALIDATION_DOCX, truth_keys, Path(tmpdir))
        result = report.conversion_result

        payload = {
            "baseline_name": "Word candidate detection + conversion Phase 3 conversion validation",
            "status": "Word conversion + internal residual check validated (post Phase 2.1 frozen detection)",
            "validation_docx_sha256": frozen["validation_docx_sha256"],
            "ground_truth_sha256": frozen["ground_truth_sha256"],
            "reproduction_check": {
                "performed": True,
                "matches_phase2_1_frozen_metrics": reproduction_matches,
                "phase2_1_frozen_metrics": frozen_metrics,
            },
            "conversion_metrics": {
                "matched_ground_truth_count": len(matched),
                "converted_matched_candidate_count": report.converted_matched_candidate_count,
                "converted_run_count": result.converted_run_count,
                "converted_property_count": result.converted_property_count,
                "review_required_count": result.review_required_count,
                "skipped_hyperlink_target_count": result.skipped_hyperlink_target_count,
                "warnings": list(result.warnings),
            },
            "residual_check": {
                "convert_internal_residual_warning": report.convert_internal_residual_warning,
                "matched_truth_readable_text_residual_count": len(report.matched_truth_readable_text_residual),
                "matched_truth_internal_xml_residual_parts": sorted(report.matched_truth_internal_xml_residual),
                "verdict": report.verdict,
            },
            "method": {
                "conversion_approval": (
                    "All non-hyperlink candidates enabled (as if reviewer-approved), "
                    "mirroring tests/test_word_replacement.py::_enable_review_required"
                ),
                "mode": (
                    "analysis (ProcessingOptions default); external mode not usable here because this fixture "
                    "has known pre-existing unsupported_features (customXml + a tracked_changes false positive "
                    "on styles.xml), unrelated to this check"
                ),
            },
            "scope_note": (
                "This record covers development/validation-tier data only. Holdout evaluation "
                "(design doc step 10) has not been executed as part of this record; no holdout file, "
                "path, hash, or content was read to produce it."
            ),
        }
        write_json(OUTPUT_JSON, payload)
        write_csv(OUTPUT_CSV, payload)
        write_markdown(OUTPUT_MD, payload)

        print("word_phase3_conversion_validation=completed")
        print("verdict", report.verdict)
        print("conversion_metrics", json.dumps(payload["conversion_metrics"], ensure_ascii=False, sort_keys=True))
        print("residual_check", json.dumps(payload["residual_check"], ensure_ascii=False, sort_keys=True))
        return 0 if report.verdict == "pass" else 1


def read_truths(path: Path) -> list[dict[str, object]]:
    with path.open(encoding="utf-8-sig", newline="") as input_file:
        return list(csv.DictReader(input_file))


def write_json(path: Path, payload: dict[str, object]) -> None:
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def write_csv(path: Path, payload: dict[str, object]) -> None:
    with path.open("w", encoding="utf-8-sig", newline="") as output:
        writer = csv.writer(output)
        writer.writerow(["scope", "name", "metric", "value"])
        writer.writerow(["reproduction_check", "all", "matches_phase2_1_frozen_metrics", payload["reproduction_check"]["matches_phase2_1_frozen_metrics"]])
        for metric, value in payload["conversion_metrics"].items():
            if metric == "warnings":
                writer.writerow(["conversion_metrics", "all", "warning_count", len(value)])
                continue
            writer.writerow(["conversion_metrics", "all", metric, value])
        for metric, value in payload["residual_check"].items():
            if metric == "matched_truth_internal_xml_residual_parts":
                writer.writerow(["residual_check", "all", metric, "、".join(value)])
                continue
            writer.writerow(["residual_check", "all", metric, value])


def write_markdown(path: Path, payload: dict[str, object]) -> None:
    reproduction = payload["reproduction_check"]
    conversion = payload["conversion_metrics"]
    residual = payload["residual_check"]
    frozen_metrics = reproduction["phase2_1_frozen_metrics"]
    lines = [
        "# Word匿名化検出 Phase3 ConversionValidation",
        "",
        "Phase 2.1で凍結済みの検出結果(構造位置一致・文字範囲一致・カテゴリ一致)を再現したうえで、"
        "同一の検証用Word・ground truthに対して実際に scan() → convert() を実行し、"
        "設計書が定める4つ目の評価軸「置換後残存なし」を初めて検証した記録です。",
        "",
        "## 再現性確認",
        "",
        "| 指標 | Phase 2.1 凍結値 |",
        "|---|---:|",
    ]
    for metric in (
        "ground_truth_count",
        "candidate_count",
        "matched_ground_truth_count",
        "missed_ground_truth_count",
        "false_positive_candidate_count",
        "recall",
        "candidate_precision",
    ):
        lines.append(f"| {metric} | `{frozen_metrics[metric]}` |")
    lines.append("")
    lines.append(f"再現一致: `{reproduction['matches_phase2_1_frozen_metrics']}`(検出ロジックは変更されていません)")
    lines.extend(["", "## Conversion Metrics", "", "| 指標 | 値 |", "|---|---:|"])
    for metric in (
        "matched_ground_truth_count",
        "converted_matched_candidate_count",
        "converted_run_count",
        "converted_property_count",
        "review_required_count",
        "skipped_hyperlink_target_count",
    ):
        lines.append(f"| {metric} | `{conversion[metric]}` |")
    lines.append(f"| warning_count | `{len(conversion['warnings'])}` |")
    lines.extend(["", "## Residual Check(置換後残存なし)", "", "| 指標 | 値 |", "|---|---:|"])
    lines.append(f"| verdict | `{residual['verdict']}` |")
    lines.append(f"| convert_internal_residual_warning | `{residual['convert_internal_residual_warning']}` |")
    lines.append(f"| matched_truth_readable_text_residual_count | `{residual['matched_truth_readable_text_residual_count']}` |")
    parts = "、".join(residual["matched_truth_internal_xml_residual_parts"]) or "なし"
    lines.append(f"| matched_truth_internal_xml_residual_parts | `{parts}` |")
    lines.extend(
        [
            "",
            "## Notes",
            "",
            "- 本記録は開発/検証データのみを対象とします。完全ホールドアウト評価(設計書の手順10)はこの記録の一部として実施していません。",
            "- 候補生成ルールは変更していません(Phase 2.1 frozenのまま、再現性確認で照合済み)。",
            "- 変換承認は全非ハイパーリンク候補を有効化した上で実施しています(レビュー承認を模擬)。",
            "- 変換後の.docxおよびCSV/報告書/監査JSONの実ファイルはコミットしていません(一時ディレクトリで生成・破棄)。",
        ]
    )
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")


if __name__ == "__main__":
    raise SystemExit(main())
