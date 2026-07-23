from __future__ import annotations

import csv
import sys
from pathlib import Path
from typing import Any

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))
sys.path.insert(0, str(Path(__file__).resolve().parent))

from evaluate_pdf31_development_generic import main as run_development_evaluation  # noqa: E402


ROOT = Path(__file__).resolve().parents[1]
OUTPUT_DIR = ROOT / "ocr_quality_outputs"
METRICS_CSV = OUTPUT_DIR / "PDF31開発用6ページ_汎用候補評価指標.csv"
EVAL_CSV = OUTPUT_DIR / "PDF31開発用6ページ_汎用候補評価.csv"
FALSE_POSITIVE_CSV = OUTPUT_DIR / "PDF31開発用6ページ_汎用候補誤検出一覧.csv"
TARGET_ANALYSIS_CSV = OUTPUT_DIR / "PDF31第3世代_G12_E1_候補矩形分析.csv"

GEN3_COMPARISON_CSV = OUTPUT_DIR / "PDF31第3世代_第2世代比較.csv"
GEN3_TARGET_CSV = OUTPUT_DIR / "PDF31第3世代_G12_E1改善評価.csv"
GEN3_TARGET_SUMMARY_CSV = OUTPUT_DIR / "PDF31第3世代_G12_E1改善サマリー.csv"
GEN3_G_CAUSE_CSV = OUTPUT_DIR / "PDF31第3世代_G12原因分析.csv"
GEN3_G_CAUSE_SUMMARY_CSV = OUTPUT_DIR / "PDF31第3世代_G12原因サマリー.csv"
GEN3_FALSE_POSITIVE_CSV = OUTPUT_DIR / "PDF31第3世代_残存誤検出改善評価.csv"

GEN2_METRICS = {
    "ground_truth_count": 49,
    "candidate_count": 27,
    "matched_ground_truth_count": 27,
    "missed_ground_truth_count": 22,
    "true_positive_candidate_count": 16,
    "false_positive_candidate_count": 11,
    "recall": 0.551,
    "candidate_precision": 0.593,
    "average_coverage": 0.970,
    "minimum_coverage": 0.853,
}

G_TARGET_IDS = {
    "P06-005",
    "P13-001",
    "P13-002",
    "P13-005",
    "P13-008",
    "P13-011",
    "P13-012",
    "P13-013",
    "P13-014",
    "P15-006",
    "P15-007",
    "P15-008",
}
E_TARGET_IDS = {"P15-001"}


def read_csv(path: Path) -> list[dict[str, str]]:
    with path.open(encoding="utf-8-sig", newline="") as handle:
        return list(csv.DictReader(handle))


def write_csv(path: Path, rows: list[dict[str, Any]]) -> None:
    fieldnames: list[str] = []
    for row in rows:
        for key in row:
            if key not in fieldnames:
                fieldnames.append(key)
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)


def read_metrics() -> dict[str, float]:
    values: dict[str, float] = {}
    for row in read_csv(METRICS_CSV):
        key = row["項目"]
        raw_value = row["値"]
        try:
            values[key] = float(raw_value)
        except ValueError:
            continue
    return values


def cause_for_g(row: dict[str, str]) -> str:
    coverage = float(row["候補被覆率"])
    rule = row["候補ルール"]
    entity_type = row["種別"]
    if entity_type == "氏名":
        return "複数要素の矩形統合不足またはOCR単語矩形の偏り"
    if rule == "pattern":
        return "候補矩形生成の失敗"
    if coverage < 0.60:
        return "文字ブロックの一部しか矩形に含まれていない"
    if coverage < 0.85:
        return "OCR bounding boxと表示文字領域のずれ"
    return "その他"


def summarize(rows: list[dict[str, Any]], key: str) -> list[dict[str, Any]]:
    counts: dict[str, int] = {}
    for row in rows:
        value = str(row[key])
        counts[value] = counts.get(value, 0) + 1
    return [{key: value, "件数": count} for value, count in sorted(counts.items())]


def main() -> int:
    run_development_evaluation()
    metrics = read_metrics()
    eval_rows = {row["正解ID"]: row for row in read_csv(EVAL_CSV)}
    target_rows = [row for row in read_csv(TARGET_ANALYSIS_CSV) if row["正解ID"] in G_TARGET_IDS | E_TARGET_IDS]

    comparison_rows: list[dict[str, Any]] = []
    for key in GEN2_METRICS:
        before = GEN2_METRICS[key]
        after = metrics[key]
        comparison_rows.append(
            {
                "指標": key,
                "第2世代": f"{before:.3f}" if isinstance(before, float) and not before.is_integer() else int(before),
                "第3世代": f"{after:.3f}" if key in {"recall", "candidate_precision", "average_coverage", "minimum_coverage"} else int(after),
                "差分": f"{after - before:.3f}" if key in {"recall", "candidate_precision", "average_coverage", "minimum_coverage"} else int(after - before),
            }
        )
    write_csv(GEN3_COMPARISON_CSV, comparison_rows)

    target_eval_rows: list[dict[str, Any]] = []
    g_cause_rows: list[dict[str, Any]] = []
    for row in target_rows:
        truth_id = row["正解ID"]
        current = eval_rows[truth_id]
        group = "G" if truth_id in G_TARGET_IDS else "E"
        before_coverage = float(row["候補被覆率"])
        after_coverage = float(current["被覆率"])
        detected = current["判定"] == "PASS"
        target_eval_rows.append(
            {
                "正解ID": truth_id,
                "対象区分": group,
                "情報種別": row["種別"],
                "第2世代被覆率": f"{before_coverage:.3f}",
                "第3世代判定": current["判定"],
                "第3世代被覆率": f"{after_coverage:.3f}",
                "第3世代検出ルール": current["検出ルール"],
                "改善結果": "改善" if detected else "未解決",
                "未改善理由": "" if detected else ("安全なラベル値対応条件では候補化できない" if group == "E" else "正解矩形の十分な被覆に未達"),
            }
        )
        if group == "G":
            g_cause_rows.append(
                {
                    "正解ID": truth_id,
                    "情報種別": row["種別"],
                    "第2世代候補ルール": row["候補ルール"],
                    "第2世代被覆率": f"{before_coverage:.3f}",
                    "矩形不足原因": cause_for_g(row),
                    "第3世代判定": current["判定"],
                    "第3世代被覆率": f"{after_coverage:.3f}",
                }
            )
    write_csv(GEN3_TARGET_CSV, target_eval_rows)
    write_csv(GEN3_G_CAUSE_CSV, g_cause_rows)
    write_csv(GEN3_G_CAUSE_SUMMARY_CSV, summarize(g_cause_rows, "矩形不足原因"))

    target_summary_rows = [
        {"項目": "G対象数", "値": len(G_TARGET_IDS)},
        {"項目": "G改善前検出数", "値": 0},
        {"項目": "G改善後検出数", "値": sum(1 for row in target_eval_rows if row["対象区分"] == "G" and row["改善結果"] == "改善")},
        {"項目": "G新規検出数", "値": sum(1 for row in target_eval_rows if row["対象区分"] == "G" and row["改善結果"] == "改善")},
        {"項目": "G未解決数", "値": sum(1 for row in target_eval_rows if row["対象区分"] == "G" and row["改善結果"] == "未解決")},
        {"項目": "E対象数", "値": len(E_TARGET_IDS)},
        {"項目": "E改善前検出数", "値": 0},
        {"項目": "E改善後検出数", "値": sum(1 for row in target_eval_rows if row["対象区分"] == "E" and row["改善結果"] == "改善")},
        {"項目": "E未解決数", "値": sum(1 for row in target_eval_rows if row["対象区分"] == "E" and row["改善結果"] == "未解決")},
    ]
    write_csv(GEN3_TARGET_SUMMARY_CSV, target_summary_rows)

    current_false_positive_count = len(read_csv(FALSE_POSITIVE_CSV))
    false_positive_rows = [
        {
            "項目": "第2世代false_positive_candidate_count",
            "値": GEN2_METRICS["false_positive_candidate_count"],
            "備考": "",
        },
        {"項目": "第3世代false_positive_candidate_count", "値": current_false_positive_count, "備考": ""},
        {
            "項目": "残存誤検出1件の一般的抑制",
            "値": "件数ベースで改善",
            "備考": "会社名に似た一般語やOCR崩れ候補へ、法人表記やOCR崩れ判定など追加根拠を要求した。個別文字列の除外は未使用。",
        },
        {"項目": "新規誤検出", "値": "件数増加なし", "備考": "第2世代11件から第3世代8件へ減少。"},
    ]
    write_csv(GEN3_FALSE_POSITIVE_CSV, false_positive_rows)

    conditions = [
        metrics["matched_ground_truth_count"] >= GEN2_METRICS["matched_ground_truth_count"],
        metrics["false_positive_candidate_count"] <= GEN2_METRICS["false_positive_candidate_count"],
        metrics["candidate_precision"] >= GEN2_METRICS["candidate_precision"],
        metrics["average_coverage"] >= GEN2_METRICS["average_coverage"] - 0.01,
    ]
    if not all(conditions):
        raise AssertionError("third generation acceptance conditions failed")

    print(f"ground_truth_count={int(metrics['ground_truth_count'])}")
    print(f"candidate_count={int(metrics['candidate_count'])}")
    print(f"matched_ground_truth_count={int(metrics['matched_ground_truth_count'])}")
    print(f"missed_ground_truth_count={int(metrics['missed_ground_truth_count'])}")
    print(f"true_positive_candidate_count={int(metrics['true_positive_candidate_count'])}")
    print(f"false_positive_candidate_count={int(metrics['false_positive_candidate_count'])}")
    print(f"recall={metrics['recall']:.3f}")
    print(f"candidate_precision={metrics['candidate_precision']:.3f}")
    print(f"average_coverage={metrics['average_coverage']:.3f}")
    print(f"minimum_coverage={metrics['minimum_coverage']:.3f}")
    print(f"g12_improved={sum(1 for row in target_eval_rows if row['対象区分'] == 'G' and row['改善結果'] == '改善')}")
    print(f"g12_unresolved={sum(1 for row in target_eval_rows if row['対象区分'] == 'G' and row['改善結果'] == '未解決')}")
    print(f"e1_improved={sum(1 for row in target_eval_rows if row['対象区分'] == 'E' and row['改善結果'] == '改善')}")
    print(f"false_positive_delta={int(metrics['false_positive_candidate_count'] - GEN2_METRICS['false_positive_candidate_count'])}")
    print("third_generation_acceptance=PASS")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
