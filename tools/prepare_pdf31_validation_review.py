from __future__ import annotations

import csv
import hashlib
import json
import re
import sys
import time
from collections import Counter
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Any

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

import fitz  # noqa: E402

from excel_privacy_cleaner.pdf_context_rules import context_candidates_for_page  # noqa: E402
from excel_privacy_cleaner.pdf_ocr_support import CANDIDATE_REJECTED, evaluate_page_quality  # noqa: E402
from excel_privacy_cleaner.pdf_ocr_support import detect_ocr_candidates  # noqa: E402
from excel_privacy_cleaner.pdf_processor import ocr_page_text_and_words  # noqa: E402


ROOT = Path(__file__).resolve().parents[1]
SOURCE = ROOT / "PDFテストデータ２.pdf"
SPLIT_JSON = ROOT / "config" / "evaluation" / "pdf31_dataset_split_v1.json"
OUTPUT_DIR = ROOT / "ocr_quality_outputs" / "PDF31検証用7ページ_初回候補レビュー"

CANDIDATES_CSV = OUTPUT_DIR / "PDF31検証用7ページ_初回候補一覧.csv"
PAGE_SUMMARY_CSV = OUTPUT_DIR / "PDF31検証用7ページ_ページ別候補数.csv"
ENTITY_SUMMARY_CSV = OUTPUT_DIR / "PDF31検証用7ページ_情報種別別候補数.csv"
RULE_SUMMARY_CSV = OUTPUT_DIR / "PDF31検証用7ページ_ルール別候補数.csv"
REVIEW_STATE_JSON = OUTPUT_DIR / "PDF31検証用7ページ_正解レビュー準備_state.json"
REPORT_TXT = OUTPUT_DIR / "PDF31検証用7ページ_初回候補レビュー準備_報告.txt"


@dataclass(frozen=True)
class CandidateRecord:
    candidate_id: str
    page: int
    entity_type: str
    rule_name: str
    confidence: str
    status: str
    original: str
    normalized: str
    replacement: str
    reason: str
    rect: tuple[float, float, float, float]


def read_validation_pages() -> list[int]:
    split = json.loads(SPLIT_JSON.read_text(encoding="utf-8"))
    return [int(page) for page in split["validation_ocr_eligible_pages"]]


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def parse_rule_name(reason: str) -> str:
    return reason.split(":", 1)[0] if ":" in reason else reason.split("。", 1)[0]


def parse_confidence(reason: str) -> str:
    marker = "信頼度="
    if marker not in reason:
        return "UNKNOWN"
    confidence = reason.split(marker, 1)[1].split("。", 1)[0].strip()
    return confidence if confidence in {"HIGH", "MEDIUM", "LOW"} else "UNKNOWN"


def generate_candidates(pages: list[int]) -> tuple[list[CandidateRecord], dict[int, Any]]:
    records: list[CandidateRecord] = []
    qualities: dict[int, Any] = {}
    with fitz.open(SOURCE) as doc:
        for page_number in pages:
            page = doc[page_number - 1]
            text, words = ocr_page_text_and_words(page)
            page_candidates = detect_ocr_candidates(words, page.rect)
            page_candidates.extend(context_candidates_for_page(page_number, page.rect, words))
            qualities[page_number] = evaluate_page_quality(page, text, words, page_candidates)
            active = [candidate for candidate in page_candidates if candidate.status != CANDIDATE_REJECTED]
            for index, candidate in enumerate(active, start=1):
                records.append(
                    CandidateRecord(
                        candidate_id=f"V{page_number:02d}-{index:03d}",
                        page=page_number,
                        entity_type=candidate.entity_type,
                        rule_name=parse_rule_name(candidate.reason),
                        confidence=parse_confidence(candidate.reason),
                        status=candidate.status,
                        original=candidate.original,
                        normalized=candidate.normalized,
                        replacement=candidate.replacement,
                        reason=candidate.reason,
                        rect=tuple(float(value) for value in candidate.rect),
                    )
                )
    return records, qualities


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


def rect_text(rect: tuple[float, float, float, float]) -> str:
    return ",".join(f"{value:.1f}" for value in rect)


def write_candidate_csv(records: list[CandidateRecord]) -> None:
    rows = [
        {
            "候補ID": record.candidate_id,
            "ページ番号": record.page,
            "情報種別": record.entity_type,
            "検出ルール": record.rule_name,
            "信頼度": record.confidence,
            "候補状態": record.status,
            "候補文字列": record.original,
            "正規化文字列": record.normalized,
            "変換候補": record.replacement,
            "座標": rect_text(record.rect),
            "検出理由": record.reason,
            "ユーザー確認状態": "未確認",
        }
        for record in records
    ]
    write_csv(CANDIDATES_CSV, rows)


def write_summaries(records: list[CandidateRecord], qualities: dict[int, Any], pages: list[int]) -> None:
    by_page = Counter(record.page for record in records)
    page_rows = []
    for page in pages:
        quality = qualities.get(page)
        page_rows.append(
            {
                "ページ番号": page,
                "候補数": by_page.get(page, 0),
                "品質判定": getattr(quality, "verdict", "UNKNOWN"),
                "OCR取得文字数": getattr(quality, "ocr_text_chars", ""),
                "座標付き単語数": getattr(quality, "coordinate_word_count", ""),
                "警告理由": getattr(quality, "warning_reason", ""),
            }
        )
    write_csv(PAGE_SUMMARY_CSV, page_rows)

    entity_counts = Counter(record.entity_type for record in records)
    write_csv(
        ENTITY_SUMMARY_CSV,
        [{"情報種別": entity, "候補数": count} for entity, count in sorted(entity_counts.items())],
    )

    rule_counts = Counter(record.rule_name for record in records)
    write_csv(
        RULE_SUMMARY_CSV,
        [{"検出ルール": rule, "候補数": count} for rule, count in sorted(rule_counts.items())],
    )


def color_for(record: CandidateRecord) -> tuple[float, float, float]:
    if record.confidence == "HIGH":
        return (0.0, 0.65, 0.0)
    if record.confidence == "MEDIUM":
        return (0.95, 0.65, 0.0)
    if record.confidence == "LOW":
        return (0.1, 0.35, 0.95)
    return (0.85, 0.1, 0.1)


def render_review_images(records: list[CandidateRecord], pages: list[int]) -> list[Path]:
    image_paths: list[Path] = []
    by_page: dict[int, list[CandidateRecord]] = {}
    for record in records:
        by_page.setdefault(record.page, []).append(record)
    with fitz.open(SOURCE) as doc:
        for page_number in pages:
            page = doc[page_number - 1]
            clean_pix = page.get_pixmap(matrix=fitz.Matrix(1.6, 1.6), alpha=False)
            clean_path = OUTPUT_DIR / f"page_{page_number:02d}_image.png"
            clean_pix.save(clean_path)

            annotated = fitz.open()
            annotated.insert_pdf(doc, from_page=page_number - 1, to_page=page_number - 1)
            annotated_page = annotated[0]
            shape = annotated_page.new_shape()
            for record in by_page.get(page_number, []):
                rect = fitz.Rect(record.rect)
                shape.draw_rect(rect)
                shape.finish(color=color_for(record), width=1.4)
                label_point = fitz.Point(rect.x0, max(rect.y0 - 2, 8))
                annotated_page.insert_text(
                    label_point,
                    record.candidate_id,
                    fontsize=6.5,
                    color=color_for(record),
                )
            shape.commit()
            pix = annotated_page.get_pixmap(matrix=fitz.Matrix(1.6, 1.6), alpha=False)
            annotated_path = OUTPUT_DIR / f"page_{page_number:02d}_candidates.png"
            pix.save(annotated_path)
            annotated.close()
            image_paths.append(annotated_path)
    return image_paths


def write_review_state(records: list[CandidateRecord], pages: list[int], image_paths: list[Path]) -> None:
    state = {
        "schema": "pdf31_validation_candidate_review_state_v1",
        "purpose": "validation_pages_initial_candidate_review",
        "source_pdf": str(SOURCE),
        "source_sha256": sha256_file(SOURCE),
        "dataset_split": "pdf31_dataset_split_v1.validation_ocr_eligible_pages",
        "pages": pages,
        "rule_change_allowed": False,
        "review_status": "not_started",
        "last_page": pages[0] if pages else None,
        "images": [str(path) for path in image_paths],
        "candidates": [
            {
                **asdict(record),
                "rect": [round(value, 1) for value in record.rect],
                "user_review_status": "未確認",
                "user_note": "",
            }
            for record in records
        ],
    }
    REVIEW_STATE_JSON.write_text(json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8")


def write_report(records: list[CandidateRecord], pages: list[int], image_paths: list[Path], started: float) -> None:
    by_page = Counter(record.page for record in records)
    by_entity = Counter(record.entity_type for record in records)
    by_rule = Counter(record.rule_name for record in records)
    lines = [
        "PDF31検証用7ページ 初回候補レビュー準備",
        f"対象ページ: {', '.join(str(page) for page in pages)}",
        "ページ22と最終テスト用17ページは処理していません。",
        "候補生成ルールは変更していません。",
        f"候補総数: {len(records)}",
        f"処理秒数: {time.perf_counter() - started:.3f}",
        "",
        "ページ別候補数:",
        *[f"- {page}: {by_page.get(page, 0)}" for page in pages],
        "",
        "情報種別別候補数:",
        *[f"- {entity}: {count}" for entity, count in sorted(by_entity.items())],
        "",
        "ルール別候補数:",
        *[f"- {rule}: {count}" for rule, count in sorted(by_rule.items())],
        "",
        "出力:",
        f"- {CANDIDATES_CSV}",
        f"- {PAGE_SUMMARY_CSV}",
        f"- {ENTITY_SUMMARY_CSV}",
        f"- {RULE_SUMMARY_CSV}",
        f"- {REVIEW_STATE_JSON}",
        *[f"- {path}" for path in image_paths],
    ]
    REPORT_TXT.write_text("\n".join(lines), encoding="utf-8")


def validate_pages(pages: list[int]) -> None:
    allowed = {1, 2, 3, 7, 8, 16, 31}
    if set(pages) != allowed:
        raise RuntimeError(f"Unexpected validation pages: {pages}")
    forbidden = {22, 4, 9, 12, 14, 17, 18, 19, 20, 21, 23, 24, 25, 26, 27, 28, 29, 30}
    if set(pages) & forbidden:
        raise RuntimeError("Forbidden pages are included")


def main() -> int:
    started = time.perf_counter()
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    pages = read_validation_pages()
    validate_pages(pages)
    records, qualities = generate_candidates(pages)
    write_candidate_csv(records)
    write_summaries(records, qualities, pages)
    image_paths = render_review_images(records, pages)
    write_review_state(records, pages, image_paths)
    write_report(records, pages, image_paths, started)

    print("pdf31_validation_review_preparation=completed")
    print("pages=" + ",".join(str(page) for page in pages))
    print(f"candidate_count={len(records)}")
    print(f"candidate_csv={CANDIDATES_CSV}")
    print(f"page_summary_csv={PAGE_SUMMARY_CSV}")
    print(f"entity_summary_csv={ENTITY_SUMMARY_CSV}")
    print(f"rule_summary_csv={RULE_SUMMARY_CSV}")
    print(f"review_state_json={REVIEW_STATE_JSON}")
    print("review_images=" + ",".join(str(path) for path in image_paths))
    print(f"report={REPORT_TXT}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
