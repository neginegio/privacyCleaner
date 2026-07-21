from __future__ import annotations

import csv
import json
import sys
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Any

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

import fitz  # noqa: E402

from excel_privacy_cleaner.pdf_context_rules import context_candidates_for_page  # noqa: E402
from excel_privacy_cleaner.pdf_ocr_support import CANDIDATE_REJECTED, detect_ocr_candidates  # noqa: E402
from excel_privacy_cleaner.pdf_processor import ocr_page_text_and_words  # noqa: E402


ROOT = Path(__file__).resolve().parents[1]
SOURCE = ROOT / "PDFテストデータ２.pdf"
OUTPUT_DIR = ROOT / "ocr_quality_outputs"
SPLIT_JSON = ROOT / "config" / "evaluation" / "pdf31_dataset_split_v1.json"
FORMAL_GT_CSV = OUTPUT_DIR / "PDF31代表6ページ_ユーザー確認済み正解データ_v1.csv"
DEV_CANDIDATES_CSV = OUTPUT_DIR / "PDF31開発用6ページ_汎用候補一覧.csv"
DEV_EVAL_CSV = OUTPUT_DIR / "PDF31開発用6ページ_汎用候補評価.csv"
DEV_METRICS_CSV = OUTPUT_DIR / "PDF31開発用6ページ_汎用候補評価指標.csv"
DEV_ENTITY_SUMMARY_CSV = OUTPUT_DIR / "PDF31開発用6ページ_情報種別別結果.csv"
DEV_RULE_SUMMARY_CSV = OUTPUT_DIR / "PDF31開発用6ページ_ルール別結果.csv"
DEV_TRUTH_CLASSIFICATION_CSV = OUTPUT_DIR / "PDF31開発用6ページ_正解分類集計.csv"
DEV_FALSE_POSITIVE_CLASSIFICATION_CSV = OUTPUT_DIR / "PDF31開発用6ページ_第2段階_誤検出分類.csv"
DEV_MISS_CLASSIFICATION_CSV = OUTPUT_DIR / "PDF31開発用6ページ_第2段階_見逃し分類.csv"
DEV_RULE_ANALYSIS_CSV = OUTPUT_DIR / "PDF31開発用6ページ_第2段階_ルール分析.csv"
DEV_CONFIDENCE_SUMMARY_CSV = OUTPUT_DIR / "PDF31開発用6ページ_第2段階_信頼度別結果.csv"
DEV_RULE_ABLATION_CSV = OUTPUT_DIR / "PDF31開発用6ページ_第2段階_ルール停止影響.csv"
DEV_COMPARISON_CSV = OUTPUT_DIR / "PDF31開発用6ページ_第2段階_改善前後比較.csv"
DEV_INITIAL_FALSE_POSITIVE_CAUSES_CSV = OUTPUT_DIR / "PDF31開発用6ページ_第2段階_初期誤検出47件分類.csv"

INITIAL_GENERIC_METRICS = {
    "candidate_count": 57,
    "detected_count": 10,
    "missed_count": 39,
    "false_positive_count": 47,
    "recall": 0.204,
    "precision": 0.175,
    "average_coverage": 0.945,
}

INITIAL_FALSE_POSITIVE_CAUSES = [
    {
        "分類": "見出し配下ブロックの横方向ゾーン未制限",
        "件数": 25,
        "主な該当ルール": "section_block",
        "改善方針": "見出しの近傍列だけを対象にし、横方向に離れた文字ブロックを除外",
    },
    {
        "分類": "一般語を会社名や氏名と誤認した",
        "件数": 10,
        "主な該当ルール": "section_block",
        "改善方針": "見出し語、業務項目語、表ヘッダー語を候補値から除外",
    },
    {
        "分類": "表の列ずれまたは隣接セルの誤取得",
        "件数": 5,
        "主な該当ルール": "section_block",
        "改善方針": "列境界と行値を分離し、過剰な行結合を抑制",
    },
    {
        "分類": "OCR誤認識",
        "件数": 4,
        "主な該当ルール": "section_block,pattern",
        "改善方針": "OCR揺れ法人表記の単独利用を抑え、構造情報がある場合だけ利用",
    },
    {
        "分類": "表全体をラベル値として巻き込んだ",
        "件数": 1,
        "主な該当ルール": "label_value",
        "改善方針": "ラベル右側値が長すぎる場合や複数の表ヘッダー語を含む場合は除外",
    },
    {
        "分類": "長文中の一部を会社名として誤認した",
        "件数": 1,
        "主な該当ルール": "pattern",
        "改善方針": "単独パターンは形式が明確な法人表記、住所、金融機関に限定",
    },
    {
        "分類": "住所の一部分だけを候補にした",
        "件数": 1,
        "主な該当ルール": "pattern",
        "改善方針": "住所候補は都道府県と市区町村など複数構成を要求",
    },
]


@dataclass(frozen=True)
class TruthRect:
    truth_id: str
    page: int
    original_text: str
    entity_type: str
    mode: str
    required: str
    structure: str
    rect: tuple[float, float, float, float]


@dataclass(frozen=True)
class CandidateRect:
    page: int
    original: str
    normalized: str
    entity_type: str
    rule_name: str
    confidence: str
    reason: str
    rect: tuple[float, float, float, float]


def read_split_pages() -> list[int]:
    split = json.loads(SPLIT_JSON.read_text(encoding="utf-8"))
    return [int(page) for page in split["development_pages"]]


def parse_rect(value: str) -> tuple[float, float, float, float]:
    parts = [float(part.strip()) for part in value.split(",")]
    if len(parts) != 4:
        raise ValueError(value)
    return parts[0], parts[1], parts[2], parts[3]


def read_truths(pages: set[int]) -> list[TruthRect]:
    with FORMAL_GT_CSV.open(encoding="utf-8-sig", newline="") as handle:
        rows = list(csv.DictReader(handle))
    truths: list[TruthRect] = []
    for row in rows:
        page = int(row["ページ番号"])
        if page not in pages:
            continue
        truths.append(
            TruthRect(
                truth_id=row["正解ID"],
                page=page,
                original_text=row["匿名化対象文字列"],
                entity_type=row["情報種別"],
                mode=row["適用モード"],
                required=row["必須度"],
                structure=row["対象理由"],
                rect=parse_rect(row["正解座標"]),
            )
        )
    return truths


def generate_candidates(pages: list[int]) -> list[CandidateRect]:
    candidates: list[CandidateRect] = []
    with fitz.open(SOURCE) as doc:
        for page_number in pages:
            page = doc[page_number - 1]
            _text, words = ocr_page_text_and_words(page)
            page_candidates = detect_ocr_candidates(words, page.rect)
            page_candidates.extend(context_candidates_for_page(page_number, page.rect, words))
            for candidate in page_candidates:
                if candidate.status == CANDIDATE_REJECTED:
                    continue
                rule_name = candidate.reason.split(":", 1)[0] if ":" in candidate.reason else candidate.reason.split("。", 1)[0]
                candidates.append(
                    CandidateRect(
                        page=page_number,
                        original=candidate.original,
                        normalized=candidate.normalized,
                        entity_type=candidate.entity_type,
                        rule_name=rule_name,
                        confidence=parse_confidence(candidate.reason),
                        reason=candidate.reason,
                        rect=tuple(float(value) for value in candidate.rect),
                    )
                )
    return candidates


def area(rect: tuple[float, float, float, float]) -> float:
    return max(rect[2] - rect[0], 0.0) * max(rect[3] - rect[1], 0.0)


def intersection(a: tuple[float, float, float, float], b: tuple[float, float, float, float]) -> float:
    x0 = max(a[0], b[0])
    y0 = max(a[1], b[1])
    x1 = min(a[2], b[2])
    y1 = min(a[3], b[3])
    return max(x1 - x0, 0.0) * max(y1 - y0, 0.0)


def coverage(truth: tuple[float, float, float, float], candidate: tuple[float, float, float, float]) -> float:
    return intersection(truth, candidate) / max(area(truth), 1.0)


def excess_ratio(truth: tuple[float, float, float, float], candidate: tuple[float, float, float, float]) -> float:
    candidate_area = max(area(candidate), 1.0)
    return max(candidate_area - intersection(truth, candidate), 0.0) / candidate_area


def normalize_text(value: str) -> str:
    import unicodedata

    return "".join(unicodedata.normalize("NFKC", value).split())


def parse_confidence(reason: str) -> str:
    marker = "信頼度="
    if marker not in reason:
        return "UNKNOWN"
    value = reason.split(marker, 1)[1].split("。", 1)[0].strip()
    return value if value in {"HIGH", "MEDIUM", "LOW"} else "UNKNOWN"


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


def write_truth_classification(truths: list[TruthRect]) -> None:
    counts: dict[tuple[str, str, str], int] = {}
    for truth in truths:
        key = (truth.entity_type, truth.mode, truth.structure)
        counts[key] = counts.get(key, 0) + 1
    rows = [
        {"情報種別": entity, "適用モード": mode, "文書構造": structure, "正解矩形数": count}
        for (entity, mode, structure), count in sorted(counts.items())
    ]
    write_csv(DEV_TRUTH_CLASSIFICATION_CSV, rows)


def evaluate(truths: list[TruthRect], candidates: list[CandidateRect]) -> dict[str, Any]:
    matched_candidate_indexes: set[int] = set()
    eval_rows: list[dict[str, Any]] = []
    coverage_values: list[float] = []
    excess_values: list[float] = []
    for truth in truths:
        best: tuple[int, CandidateRect, float, float] | None = None
        for index, candidate in enumerate(candidates):
            if candidate.page != truth.page or candidate.entity_type != truth.entity_type:
                continue
            item_coverage = coverage(truth.rect, candidate.rect)
            item_excess = excess_ratio(truth.rect, candidate.rect)
            if item_coverage >= 0.85 and (best is None or item_coverage > best[2]):
                best = (index, candidate, item_coverage, item_excess)
        if best:
            matched_candidate_indexes.add(best[0])
            coverage_values.append(best[2])
            excess_values.append(best[3])
        eval_rows.append(
            {
                "正解ID": truth.truth_id,
                "ページ番号": truth.page,
                "情報種別": truth.entity_type,
                "適用モード": truth.mode,
                "文書構造": truth.structure,
                "判定": "PASS" if best else "MISS",
                "被覆率": f"{best[2]:.3f}" if best else "0.000",
                "過剰率": f"{best[3]:.3f}" if best else "NOT_EVALUATED",
                "検出ルール": best[1].rule_name if best else "",
            }
        )
    false_positive_rows: list[dict[str, Any]] = []
    for index, candidate in enumerate(candidates):
        if index in matched_candidate_indexes:
            continue
        false_positive_rows.append(
            {
                "ページ番号": candidate.page,
                "情報種別": candidate.entity_type,
                "検出ルール": candidate.rule_name,
                "信頼度": candidate.confidence,
                "検出座標": ",".join(f"{value:.1f}" for value in candidate.rect),
                "分類": classify_false_positive(candidate),
                "理由": "開発用正解矩形に対応しない候補",
            }
        )
    write_csv(DEV_EVAL_CSV, eval_rows)
    write_csv(OUTPUT_DIR / "PDF31開発用6ページ_汎用候補誤検出一覧.csv", false_positive_rows)
    write_csv(DEV_FALSE_POSITIVE_CLASSIFICATION_CSV, false_positive_rows)
    detected = sum(1 for row in eval_rows if row["判定"] == "PASS")
    truth_count = len(truths)
    candidate_count = len(candidates)
    false_positive_count = len(false_positive_rows)
    return {
        "truth_count": truth_count,
        "candidate_count": candidate_count,
        "detected_count": detected,
        "missed_count": truth_count - detected,
        "false_positive_count": false_positive_count,
        "recall": detected / truth_count if truth_count else 0.0,
        "precision": detected / candidate_count if candidate_count else 0.0,
        "average_coverage": sum(coverage_values) / len(coverage_values) if coverage_values else 0.0,
        "minimum_coverage": min(coverage_values) if coverage_values else 0.0,
        "average_excess": sum(excess_values) / len(excess_values) if excess_values else 0.0,
        "maximum_excess": max(excess_values) if excess_values else 0.0,
        "eval_rows": eval_rows,
        "matched_candidate_indexes": matched_candidate_indexes,
        "false_positive_rows": false_positive_rows,
    }


def classify_false_positive(candidate: CandidateRect) -> str:
    value = candidate.normalized
    if is_header_or_label_only(value):
        return "見出しや項目名そのものを候補にした"
    if numeric_heavy(value):
        return "数字や割合を機密情報と誤認した"
    if len(value) <= 3:
        return "文字列が短すぎる"
    if len(value) > 60:
        return "文字列が長すぎる"
    if value in {"株式会社", "有限会社", "合同会社", "(株)", "(有)", "㈱", "㈲"}:
        return "法人表記だけを候補にした"
    if candidate.entity_type == "住所" and not looks_like_full_address(value):
        return "住所の一部分だけを候補にした"
    if candidate.rule_name.startswith("table_column"):
        return "表の列ずれ"
    if candidate.rule_name.startswith("label_value"):
        return "隣接セルの誤取得"
    if candidate.rule_name.startswith("section_block") and any(token in value for token in ("製品", "設計", "仕入", "仕上", "対象会社")):
        return "一般語を会社名や氏名と誤認した"
    if any(char in value for char in "0OIl|_") and any("\u3040" <= char <= "\u9fff" for char in value):
        return "OCR誤認識"
    return "その他"


def is_header_or_label_only(value: str) -> bool:
    labels = {
        "企業名称",
        "会社名",
        "所在地",
        "住所",
        "代表者",
        "取引銀行",
        "氏名",
        "株主",
        "販売先",
        "仕入先",
        "外注先",
        "対象会社",
    }
    return value.strip(" :：|/・-ー−－") in labels


def numeric_heavy(value: str) -> bool:
    import re

    return bool(re.fullmatch(r"[0-9０-９,.\-ー−－%％①-⑳\s]+", value))


def looks_like_full_address(value: str) -> bool:
    return any(token in value for token in ("都", "道", "府", "県")) and any(token in value for token in ("市", "区", "町", "村"))


def classify_misses(
    truths: list[TruthRect],
    candidates: list[CandidateRect],
    eval_rows: list[dict[str, Any]],
    ocr_text_by_page: dict[int, str],
) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    eval_by_id = {row["正解ID"]: row for row in eval_rows}
    for truth in truths:
        if eval_by_id[truth.truth_id]["判定"] == "PASS":
            continue
        same_page_entity = [
            candidate for candidate in candidates if candidate.page == truth.page and candidate.entity_type == truth.entity_type
        ]
        best_coverage = max((coverage(truth.rect, candidate.rect) for candidate in same_page_entity), default=0.0)
        ocr_text = ocr_text_by_page.get(truth.page, "")
        normalized_truth = normalize_text(truth.original_text)
        if normalized_truth and normalized_truth not in ocr_text:
            cause = "OCR文字列自体が取得できていない"
        elif best_coverage > 0:
            cause = "候補は生成したが、矩形が正解範囲と合っていない"
        elif any(token in truth.structure for token in ("列", "一覧", "構成")):
            cause = "表や列の対応関係を取得できていない"
        else:
            cause = "文字列は取得できているが、ルールが反応していない"
        rows.append(
            {
                "正解ID": truth.truth_id,
                "ページ番号": truth.page,
                "情報種別": truth.entity_type,
                "適用モード": truth.mode,
                "必須度": truth.required,
                "文書構造": truth.structure,
                "分類": cause,
                "最大被覆率": f"{best_coverage:.3f}",
            }
        )
    return rows


def extract_ocr_text_by_page(pages: list[int]) -> dict[int, str]:
    result: dict[int, str] = {}
    with fitz.open(SOURCE) as doc:
        for page_number in pages:
            page = doc[page_number - 1]
            text, _words = ocr_page_text_and_words(page)
            result[page_number] = normalize_text(text)
    return result


def write_candidate_csv(candidates: list[CandidateRect]) -> None:
    rows = [
        {
            "ページ番号": candidate.page,
            "候補文字列": candidate.original,
            "正規化文字列": candidate.normalized,
            "情報種別": candidate.entity_type,
            "検出ルール名": candidate.rule_name,
            "信頼度": candidate.confidence,
            "検出理由": candidate.reason,
            "座標": ",".join(f"{value:.1f}" for value in candidate.rect),
        }
        for candidate in candidates
    ]
    write_csv(DEV_CANDIDATES_CSV, rows)


def write_summaries(truths: list[TruthRect], candidates: list[CandidateRect], metrics: dict[str, Any]) -> None:
    metric_rows = [
        {"項目": "baseline_truth_count", "値": 49, "備考": "Baseline v1"},
        {"項目": "baseline_detected_count", "値": 0, "備考": "Baseline v1"},
        {"項目": "truth_count", "値": metrics["truth_count"], "備考": "開発用6ページ"},
        {"項目": "candidate_count", "値": metrics["candidate_count"], "備考": ""},
        {"項目": "detected_count", "値": metrics["detected_count"], "備考": ""},
        {"項目": "missed_count", "値": metrics["missed_count"], "備考": ""},
        {"項目": "false_positive_count", "値": metrics["false_positive_count"], "備考": ""},
        {"項目": "recall", "値": f"{metrics['recall']:.3f}", "備考": ""},
        {"項目": "precision", "値": f"{metrics['precision']:.3f}", "備考": ""},
        {"項目": "average_coverage", "値": f"{metrics['average_coverage']:.3f}", "備考": ""},
        {"項目": "minimum_coverage", "値": f"{metrics['minimum_coverage']:.3f}", "備考": ""},
        {"項目": "average_excess", "値": f"{metrics['average_excess']:.3f}", "備考": ""},
        {"項目": "maximum_excess", "値": f"{metrics['maximum_excess']:.3f}", "備考": ""},
    ]
    write_csv(DEV_METRICS_CSV, metric_rows)

    entity_rows: list[dict[str, Any]] = []
    for entity_type in sorted({truth.entity_type for truth in truths} | {candidate.entity_type for candidate in candidates}):
        truth_subset = [truth for truth in truths if truth.entity_type == entity_type]
        candidate_subset = [candidate for candidate in candidates if candidate.entity_type == entity_type]
        passed = [row for row in metrics["eval_rows"] if row["情報種別"] == entity_type and row["判定"] == "PASS"]
        entity_rows.append(
            {
                "情報種別": entity_type,
                "正解数": len(truth_subset),
                "候補数": len(candidate_subset),
                "検出数": len(passed),
                "見逃し数": len(truth_subset) - len(passed),
            }
        )
    write_csv(DEV_ENTITY_SUMMARY_CSV, entity_rows)

    rule_rows: list[dict[str, Any]] = []
    matched_rules = [row["検出ルール"] for row in metrics["eval_rows"] if row["判定"] == "PASS"]
    for rule_name in sorted({candidate.rule_name for candidate in candidates}):
        rule_candidates = [candidate for candidate in candidates if candidate.rule_name == rule_name]
        rule_rows.append(
            {
                "検出ルール": rule_name,
                "候補数": len(rule_candidates),
                "正解数": matched_rules.count(rule_name),
                "誤検出数": len(rule_candidates) - matched_rules.count(rule_name),
            }
        )
    write_csv(DEV_RULE_SUMMARY_CSV, rule_rows)

    matched_candidate_indexes: set[int] = metrics["matched_candidate_indexes"]
    rule_analysis_rows: list[dict[str, Any]] = []
    false_positive_rows = metrics["false_positive_rows"]
    for rule_name in sorted({candidate.rule_name for candidate in candidates}):
        rule_candidates = [candidate for candidate in candidates if candidate.rule_name == rule_name]
        rule_false_positive = [row for row in false_positive_rows if row["検出ルール"] == rule_name]
        candidate_indexes = [index for index, candidate in enumerate(candidates) if candidate.rule_name == rule_name]
        detected_by_rule = matched_rules.count(rule_name)
        main_cause = most_common([row["分類"] for row in rule_false_positive])
        action = "維持"
        if rule_false_positive and detected_by_rule == 0:
            action = "停止候補"
        elif len(rule_false_positive) > detected_by_rule * 3:
            action = "縮小"
        rule_analysis_rows.append(
            {
                "検出ルール": rule_name,
                "候補数": len(rule_candidates),
                "正解数": detected_by_rule,
                "誤検出数": len(rule_false_positive),
                "precision": f"{detected_by_rule / len(rule_candidates):.3f}" if rule_candidates else "0.000",
                "該当ページ数": len({candidate.page for candidate in rule_candidates}),
                "他ルールとの重複": duplicate_count_for_indexes(candidates, candidate_indexes),
                "誤検出の主な原因": main_cause,
                "改善方針": improvement_policy(rule_name, main_cause),
                "判断": action,
            }
        )
    write_csv(DEV_RULE_ANALYSIS_CSV, rule_analysis_rows)

    confidence_rows: list[dict[str, Any]] = []
    for confidence in ["HIGH", "MEDIUM", "LOW", "UNKNOWN"]:
        subset = [candidate for candidate in candidates if candidate.confidence == confidence]
        if not subset:
            continue
        subset_indexes = {index for index, candidate in enumerate(candidates) if candidate.confidence == confidence}
        matched = len(subset_indexes & matched_candidate_indexes)
        confidence_rows.append(
            {
                "信頼度": confidence,
                "候補数": len(subset),
                "検出候補数": matched,
                "precision": f"{matched / len(subset):.3f}",
            }
        )
    write_csv(DEV_CONFIDENCE_SUMMARY_CSV, confidence_rows)

    ablation_rows: list[dict[str, Any]] = []
    for rule_name in sorted({candidate.rule_name for candidate in candidates}):
        reduced = [candidate for candidate in candidates if candidate.rule_name != rule_name]
        reduced_metrics = evaluate_without_writing(truths, reduced)
        ablation_rows.append(
            {
                "停止ルール": rule_name,
                "候補数": reduced_metrics["candidate_count"],
                "検出数": reduced_metrics["detected_count"],
                "見逃し数": reduced_metrics["missed_count"],
                "誤検出数": reduced_metrics["false_positive_count"],
                "recall": f"{reduced_metrics['recall']:.3f}",
                "precision": f"{reduced_metrics['precision']:.3f}",
            }
        )
    write_csv(DEV_RULE_ABLATION_CSV, ablation_rows)

    comparison_rows = [
        {"指標": "候補数", "改善前": INITIAL_GENERIC_METRICS["candidate_count"], "改善後": metrics["candidate_count"], "差分": metrics["candidate_count"] - INITIAL_GENERIC_METRICS["candidate_count"]},
        {"指標": "検出数", "改善前": INITIAL_GENERIC_METRICS["detected_count"], "改善後": metrics["detected_count"], "差分": metrics["detected_count"] - INITIAL_GENERIC_METRICS["detected_count"]},
        {"指標": "見逃し数", "改善前": INITIAL_GENERIC_METRICS["missed_count"], "改善後": metrics["missed_count"], "差分": metrics["missed_count"] - INITIAL_GENERIC_METRICS["missed_count"]},
        {"指標": "誤検出数", "改善前": INITIAL_GENERIC_METRICS["false_positive_count"], "改善後": metrics["false_positive_count"], "差分": metrics["false_positive_count"] - INITIAL_GENERIC_METRICS["false_positive_count"]},
        {"指標": "recall", "改善前": f"{INITIAL_GENERIC_METRICS['recall']:.3f}", "改善後": f"{metrics['recall']:.3f}", "差分": f"{metrics['recall'] - INITIAL_GENERIC_METRICS['recall']:.3f}"},
        {"指標": "precision", "改善前": f"{INITIAL_GENERIC_METRICS['precision']:.3f}", "改善後": f"{metrics['precision']:.3f}", "差分": f"{metrics['precision'] - INITIAL_GENERIC_METRICS['precision']:.3f}"},
        {"指標": "平均被覆率", "改善前": f"{INITIAL_GENERIC_METRICS['average_coverage']:.3f}", "改善後": f"{metrics['average_coverage']:.3f}", "差分": f"{metrics['average_coverage'] - INITIAL_GENERIC_METRICS['average_coverage']:.3f}"},
    ]
    write_csv(DEV_COMPARISON_CSV, comparison_rows)
    write_csv(DEV_INITIAL_FALSE_POSITIVE_CAUSES_CSV, INITIAL_FALSE_POSITIVE_CAUSES)


def evaluate_without_writing(truths: list[TruthRect], candidates: list[CandidateRect]) -> dict[str, Any]:
    matched_candidate_indexes: set[int] = set()
    detected = 0
    for truth in truths:
        best: tuple[int, CandidateRect, float] | None = None
        for index, candidate in enumerate(candidates):
            if candidate.page != truth.page or candidate.entity_type != truth.entity_type:
                continue
            item_coverage = coverage(truth.rect, candidate.rect)
            if item_coverage >= 0.85 and (best is None or item_coverage > best[2]):
                best = (index, candidate, item_coverage)
        if best:
            detected += 1
            matched_candidate_indexes.add(best[0])
    candidate_count = len(candidates)
    false_positive_count = candidate_count - len(matched_candidate_indexes)
    truth_count = len(truths)
    return {
        "candidate_count": candidate_count,
        "detected_count": detected,
        "missed_count": truth_count - detected,
        "false_positive_count": false_positive_count,
        "recall": detected / truth_count if truth_count else 0.0,
        "precision": detected / candidate_count if candidate_count else 0.0,
    }


def most_common(values: list[str]) -> str:
    if not values:
        return ""
    counts: dict[str, int] = {}
    for value in values:
        counts[value] = counts.get(value, 0) + 1
    return sorted(counts.items(), key=lambda item: (-item[1], item[0]))[0][0]


def duplicate_count_for_indexes(candidates: list[CandidateRect], indexes: list[int]) -> int:
    count = 0
    for index in indexes:
        candidate = candidates[index]
        for other_index, other in enumerate(candidates):
            if other_index == index:
                continue
            if other.page != candidate.page or other.entity_type != candidate.entity_type:
                continue
            if coverage(candidate.rect, other.rect) >= 0.8:
                count += 1
                break
    return count


def improvement_policy(rule_name: str, cause: str) -> str:
    if rule_name.startswith("section_block"):
        return "見出し配下の横方向ゾーン、見出し語除外、業務一般語除外で縮小"
    if rule_name.startswith("table_column"):
        return "列境界と行値の分離を強化し、数値列や合計行を除外"
    if rule_name.startswith("label_value"):
        return "ラベル右側値が表全体を巻き込む場合は除外"
    if rule_name.startswith("pattern"):
        return "単独パターンは形式が明確なものへ限定"
    return cause


def main() -> int:
    start = time.perf_counter()
    pages = read_split_pages()
    truths = read_truths(set(pages))
    write_truth_classification(truths)
    candidates = generate_candidates(pages)
    write_candidate_csv(candidates)
    metrics = evaluate(truths, candidates)
    write_summaries(truths, candidates, metrics)
    ocr_text_by_page = extract_ocr_text_by_page(pages)
    write_csv(
        DEV_MISS_CLASSIFICATION_CSV,
        classify_misses(truths, candidates, metrics["eval_rows"], ocr_text_by_page),
    )
    print(f"development_pages={','.join(str(page) for page in pages)}")
    print(f"truth_count={metrics['truth_count']}")
    print(f"candidate_count={metrics['candidate_count']}")
    print(f"detected_count={metrics['detected_count']}")
    print(f"missed_count={metrics['missed_count']}")
    print(f"false_positive_count={metrics['false_positive_count']}")
    print(f"recall={metrics['recall']:.3f}")
    print(f"precision={metrics['precision']:.3f}")
    print(f"average_coverage={metrics['average_coverage']:.3f}")
    print(f"minimum_coverage={metrics['minimum_coverage']:.3f}")
    print(f"average_excess={metrics['average_excess']:.3f}")
    print(f"maximum_excess={metrics['maximum_excess']:.3f}")
    print(f"elapsed_seconds={time.perf_counter() - start:.3f}")
    print(f"metrics_csv={DEV_METRICS_CSV}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
