from __future__ import annotations

import csv
import json
import sys
from pathlib import Path
from typing import Any

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "tools"))

import evaluate_pdf31_development_generic as dev_eval  # noqa: E402


ROOT = Path(__file__).resolve().parents[1]
OUTPUT_DIR = ROOT / "ocr_quality_outputs"
AUDIT_JSON = OUTPUT_DIR / "PDFテストデータ２_匿名化済みPDF_20260723_112249_監査報告.json"
MANUAL_CAUSE_CSV = OUTPUT_DIR / "PDF31実手動追加14件_原因分析.csv"
REJECT_CAUSE_CSV = OUTPUT_DIR / "PDF31利用者解除6件_原因分析.csv"

MANUAL_IMPROVEMENT_CSV = OUTPUT_DIR / "PDF31第2世代_実手動追加11件改善評価.csv"
MANUAL_IMPROVEMENT_SUMMARY_CSV = OUTPUT_DIR / "PDF31第2世代_実手動追加11件改善サマリー.csv"
REJECT_IMPROVEMENT_CSV = OUTPUT_DIR / "PDF31第2世代_利用者解除6件改善評価.csv"
REJECT_IMPROVEMENT_SUMMARY_CSV = OUTPUT_DIR / "PDF31第2世代_利用者解除6件改善サマリー.csv"
OVERALL_COMPARISON_CSV = OUTPUT_DIR / "PDF31第2世代_改善前後比較.csv"

IN_SCOPE_MANUAL_CAUSES = {
    "文字が複数ブロックへ分割",
    "表列の対応関係を取得できなかった",
    "候補矩形生成の失敗",
}


def read_csv(path: Path) -> list[dict[str, str]]:
    with path.open(encoding="utf-8-sig", newline="") as handle:
        return list(csv.DictReader(handle))


def write_csv(path: Path, rows: list[dict[str, Any]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    fieldnames: list[str] = []
    for row in rows:
        for key in row:
            if key not in fieldnames:
                fieldnames.append(key)
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)


def audit_findings() -> list[dict[str, Any]]:
    return json.loads(AUDIT_JSON.read_text(encoding="utf-8"))["findings"]


def max_coverage_for(
    candidates: list[dev_eval.CandidateRect],
    page: int,
    entity_type: str,
    rect: tuple[float, float, float, float],
) -> tuple[float, str, str]:
    best_coverage = 0.0
    best_rule = ""
    best_confidence = ""
    for candidate in candidates:
        if candidate.page != page or candidate.entity_type != entity_type:
            continue
        item_coverage = dev_eval.coverage(rect, candidate.rect)
        if item_coverage > best_coverage:
            best_coverage = item_coverage
            best_rule = candidate.rule_name
            best_confidence = candidate.confidence
    return best_coverage, best_rule, best_confidence


def rect_from_finding(finding: dict[str, Any]) -> tuple[float, float, float, float]:
    values = finding["rect"]
    return float(values[0]), float(values[1]), float(values[2]), float(values[3])


def page_from_finding(finding: dict[str, Any]) -> int:
    value = str(finding["page"])
    if value.startswith("ページ"):
        value = value.replace("ページ", "", 1)
    return int(value)


def evaluate_manual_improvements(candidates: list[dev_eval.CandidateRect]) -> dict[str, Any]:
    manual_causes = read_csv(MANUAL_CAUSE_CSV)
    manual_findings = [
        finding
        for finding in audit_findings()
        if finding.get("confirmation_method") == "MANUAL" and finding.get("manual_added") is True
    ]
    rows: list[dict[str, Any]] = []
    for index, (cause_row, finding) in enumerate(zip(manual_causes, manual_findings, strict=True), start=1):
        cause = cause_row["原因分類"]
        page = page_from_finding(finding)
        entity_type = str(finding["entity_type"])
        before_coverage = float(cause_row.get("候補最大被覆率") or 0.0)
        after_coverage, after_rule, after_confidence = max_coverage_for(candidates, page, entity_type, rect_from_finding(finding))
        in_scope = cause in IN_SCOPE_MANUAL_CAUSES
        rows.append(
            {
                "分析ID": cause_row["分析ID"],
                "ページ番号": page,
                "情報種別": entity_type,
                "原因分類": cause,
                "今回対象": "YES" if in_scope else "NO",
                "改善前検出": "YES" if before_coverage >= 0.85 else "NO",
                "改善後検出": "YES" if after_coverage >= 0.85 else "NO",
                "改善前最大被覆率": f"{before_coverage:.3f}",
                "改善後最大被覆率": f"{after_coverage:.3f}",
                "改善後検出ルール": after_rule,
                "改善後信頼度": after_confidence,
                "残る理由": "" if after_coverage >= 0.85 or not in_scope else "OCR断片または表セル推定が現在の汎用条件に達していない",
            }
        )
    write_csv(MANUAL_IMPROVEMENT_CSV, rows)
    scoped = [row for row in rows if row["今回対象"] == "YES"]
    before_detected = sum(1 for row in scoped if row["改善前検出"] == "YES")
    after_detected = sum(1 for row in scoped if row["改善後検出"] == "YES")
    summary = {
        "評価対象": "実手動追加14件のうち構造改善対象11件",
        "対象件数": len(scoped),
        "改善前検出数": before_detected,
        "改善後検出数": after_detected,
        "新たに検出できた件数": sum(1 for row in scoped if row["改善前検出"] == "NO" and row["改善後検出"] == "YES"),
        "まだ検出できない件数": sum(1 for row in scoped if row["改善後検出"] == "NO"),
        "対象外件数": len(rows) - len(scoped),
    }
    write_csv(MANUAL_IMPROVEMENT_SUMMARY_CSV, [summary])
    return summary


def evaluate_rejected_improvements(candidates: list[dev_eval.CandidateRect]) -> dict[str, Any]:
    reject_causes = read_csv(REJECT_CAUSE_CSV)
    rejected_findings = [
        finding
        for finding in audit_findings()
        if finding.get("user_rejected") is True
    ]
    rows: list[dict[str, Any]] = []
    for cause_row, finding in zip(reject_causes, rejected_findings, strict=True):
        page = page_from_finding(finding)
        entity_type = str(finding["entity_type"])
        before_coverage = float(cause_row.get("候補最大被覆率") or 0.0)
        after_coverage, after_rule, after_confidence = max_coverage_for(candidates, page, entity_type, rect_from_finding(finding))
        rows.append(
            {
                "分析ID": cause_row["分析ID"],
                "ページ番号": page,
                "情報種別": entity_type,
                "評価区分": cause_row["評価区分"],
                "原因分類": cause_row["原因分類"],
                "改善前誤検出": "YES" if before_coverage >= 0.85 else "NO",
                "改善後誤検出": "YES" if after_coverage >= 0.85 else "NO",
                "改善前最大被覆率": f"{before_coverage:.3f}",
                "改善後最大被覆率": f"{after_coverage:.3f}",
                "改善後検出ルール": after_rule,
                "改善後信頼度": after_confidence,
            }
        )
    write_csv(REJECT_IMPROVEMENT_CSV, rows)
    before_fp = sum(1 for row in rows if row["改善前誤検出"] == "YES")
    after_fp = sum(1 for row in rows if row["改善後誤検出"] == "YES")
    summary = {
        "評価対象": "利用者解除6件",
        "対象件数": len(rows),
        "改善前誤検出数": before_fp,
        "改善後誤検出数": after_fp,
        "削減できた件数": before_fp - after_fp,
        "残った誤検出": after_fp,
        "新たに発生した誤検出": "NOT_EVALUATED",
    }
    write_csv(REJECT_IMPROVEMENT_SUMMARY_CSV, [summary])
    return summary


def write_overall_comparison(metrics: dict[str, Any]) -> None:
    rows = [
        {"指標": "ground_truth_count", "初期汎用ルール": dev_eval.INITIAL_GENERIC_METRICS["ground_truth_count"], "第2世代": metrics["ground_truth_count"], "差分": metrics["ground_truth_count"] - dev_eval.INITIAL_GENERIC_METRICS["ground_truth_count"]},
        {"指標": "candidate_count", "初期汎用ルール": dev_eval.INITIAL_GENERIC_METRICS["candidate_count"], "第2世代": metrics["candidate_count"], "差分": metrics["candidate_count"] - dev_eval.INITIAL_GENERIC_METRICS["candidate_count"]},
        {"指標": "matched_ground_truth_count", "初期汎用ルール": dev_eval.INITIAL_GENERIC_METRICS["matched_ground_truth_count"], "第2世代": metrics["matched_ground_truth_count"], "差分": metrics["matched_ground_truth_count"] - dev_eval.INITIAL_GENERIC_METRICS["matched_ground_truth_count"]},
        {"指標": "missed_ground_truth_count", "初期汎用ルール": dev_eval.INITIAL_GENERIC_METRICS["missed_ground_truth_count"], "第2世代": metrics["missed_ground_truth_count"], "差分": metrics["missed_ground_truth_count"] - dev_eval.INITIAL_GENERIC_METRICS["missed_ground_truth_count"]},
        {"指標": "true_positive_candidate_count", "初期汎用ルール": dev_eval.INITIAL_GENERIC_METRICS["true_positive_candidate_count"], "第2世代": metrics["true_positive_candidate_count"], "差分": metrics["true_positive_candidate_count"] - dev_eval.INITIAL_GENERIC_METRICS["true_positive_candidate_count"]},
        {"指標": "false_positive_candidate_count", "初期汎用ルール": dev_eval.INITIAL_GENERIC_METRICS["false_positive_candidate_count"], "第2世代": metrics["false_positive_candidate_count"], "差分": metrics["false_positive_candidate_count"] - dev_eval.INITIAL_GENERIC_METRICS["false_positive_candidate_count"]},
        {"指標": "recall", "初期汎用ルール": f"{dev_eval.INITIAL_GENERIC_METRICS['recall']:.3f}", "第2世代": f"{metrics['recall']:.3f}", "差分": f"{metrics['recall'] - dev_eval.INITIAL_GENERIC_METRICS['recall']:.3f}"},
        {"指標": "candidate_precision", "初期汎用ルール": f"{dev_eval.INITIAL_GENERIC_METRICS['candidate_precision']:.3f}", "第2世代": f"{metrics['candidate_precision']:.3f}", "差分": f"{metrics['candidate_precision'] - dev_eval.INITIAL_GENERIC_METRICS['candidate_precision']:.3f}"},
        {"指標": "average_coverage", "初期汎用ルール": f"{dev_eval.INITIAL_GENERIC_METRICS['average_coverage']:.3f}", "第2世代": f"{metrics['average_coverage']:.3f}", "差分": f"{metrics['average_coverage'] - dev_eval.INITIAL_GENERIC_METRICS['average_coverage']:.3f}"},
        {"指標": "minimum_coverage", "初期汎用ルール": f"{dev_eval.INITIAL_GENERIC_METRICS['minimum_coverage']:.3f}", "第2世代": f"{metrics['minimum_coverage']:.3f}", "差分": f"{metrics['minimum_coverage'] - dev_eval.INITIAL_GENERIC_METRICS['minimum_coverage']:.3f}"},
    ]
    write_csv(OVERALL_COMPARISON_CSV, rows)


def main() -> int:
    pages = dev_eval.read_split_pages()
    dev_truths = dev_eval.read_truths(set(pages))
    dev_candidates = dev_eval.generate_candidates(pages)
    metrics = dev_eval.evaluate(dev_truths, dev_candidates)
    extra_pages = sorted(
        {
            page_from_finding(finding)
            for finding in audit_findings()
            if finding.get("user_rejected") is True
        }
        - set(pages)
    )
    candidates = dev_eval.generate_candidates(pages + extra_pages)
    manual_summary = evaluate_manual_improvements(candidates)
    rejected_summary = evaluate_rejected_improvements(candidates)
    write_overall_comparison(metrics)
    print("manual_structural_in_scope_count", manual_summary["対象件数"])
    print("manual_before_detected", manual_summary["改善前検出数"])
    print("manual_after_detected", manual_summary["改善後検出数"])
    print("manual_newly_detected", manual_summary["新たに検出できた件数"])
    print("rejected_before_false_positive", rejected_summary["改善前誤検出数"])
    print("rejected_after_false_positive", rejected_summary["改善後誤検出数"])
    print("rejected_reduced", rejected_summary["削減できた件数"])
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
