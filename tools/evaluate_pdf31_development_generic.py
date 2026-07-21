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


@dataclass(frozen=True)
class TruthRect:
    truth_id: str
    page: int
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
                "検出座標": ",".join(f"{value:.1f}" for value in candidate.rect),
                "理由": "開発用正解矩形に対応しない候補",
            }
        )
    write_csv(DEV_EVAL_CSV, eval_rows)
    write_csv(OUTPUT_DIR / "PDF31開発用6ページ_汎用候補誤検出一覧.csv", false_positive_rows)
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
        "eval_rows": eval_rows,
    }


def write_candidate_csv(candidates: list[CandidateRect]) -> None:
    rows = [
        {
            "ページ番号": candidate.page,
            "候補文字列": candidate.original,
            "正規化文字列": candidate.normalized,
            "情報種別": candidate.entity_type,
            "検出ルール名": candidate.rule_name,
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


def main() -> int:
    start = time.perf_counter()
    pages = read_split_pages()
    truths = read_truths(set(pages))
    write_truth_classification(truths)
    candidates = generate_candidates(pages)
    write_candidate_csv(candidates)
    metrics = evaluate(truths, candidates)
    write_summaries(truths, candidates, metrics)
    print(f"development_pages={','.join(str(page) for page in pages)}")
    print(f"truth_count={metrics['truth_count']}")
    print(f"candidate_count={metrics['candidate_count']}")
    print(f"detected_count={metrics['detected_count']}")
    print(f"missed_count={metrics['missed_count']}")
    print(f"false_positive_count={metrics['false_positive_count']}")
    print(f"recall={metrics['recall']:.3f}")
    print(f"precision={metrics['precision']:.3f}")
    print(f"average_coverage={metrics['average_coverage']:.3f}")
    print(f"elapsed_seconds={time.perf_counter() - start:.3f}")
    print(f"metrics_csv={DEV_METRICS_CSV}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
