from __future__ import annotations

import csv
import hashlib
import json
import sys
import time
from pathlib import Path
from typing import Any

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

import fitz  # noqa: E402

from excel_privacy_cleaner.pdf_context_rules import context_candidates_for_page  # noqa: E402
from excel_privacy_cleaner.pdf_ocr_support import (  # noqa: E402
    CANDIDATE_REJECTED,
    PAGE_COMPLETED,
    PAGE_FAILED_UNRESOLVED,
    PAGE_REVIEWED_WITH_REDACTIONS,
    QUALITY_FAILED,
    VERIFICATION_MANUAL_REGION_PASS,
    VERIFICATION_NOT_EVALUATED,
    detect_ocr_candidates,
    evaluate_page_quality,
)
from excel_privacy_cleaner.pdf_processor import final_output_status, ocr_page_text_and_words  # noqa: E402


ROOT = Path(__file__).resolve().parents[1]
SOURCE = ROOT / "PDFテストデータ２.pdf"
SPLIT_JSON = ROOT / "config" / "evaluation" / "pdf31_dataset_split_v1.json"
OUTPUT_DIR = ROOT / "ocr_quality_outputs" / "PDF31ページ22_FAILED評価"

QUALITY_CSV = OUTPUT_DIR / "PDF31ページ22_OCR品質FAILED評価.csv"
CANDIDATES_CSV = OUTPUT_DIR / "PDF31ページ22_初回候補一覧.csv"
WORKFLOW_CSV = OUTPUT_DIR / "PDF31ページ22_FAILED制御結果.csv"
REVIEW_STATE_JSON = OUTPUT_DIR / "PDF31ページ22_FAILEDレビュー状態.json"
REPORT_TXT = OUTPUT_DIR / "PDF31ページ22_FAILED評価報告書.txt"
PAGE_IMAGE = OUTPUT_DIR / "page_22_image.png"
CANDIDATE_IMAGE = OUTPUT_DIR / "page_22_candidates.png"


def read_failed_pages() -> list[int]:
    split = json.loads(SPLIT_JSON.read_text(encoding="utf-8"))
    pages = [int(page) for page in split["validation_ocr_failed_pages"]]
    if pages != [22]:
        raise RuntimeError(f"Unexpected failed validation pages: {pages}")
    return pages


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


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


def rule_name(reason: str) -> str:
    return reason.split(":", 1)[0] if ":" in reason else reason.split("。", 1)[0]


def rect_text(rect: tuple[float, float, float, float]) -> str:
    return ",".join(f"{value:.1f}" for value in rect)


def display_reasons(reasons: list[str], page_number: int) -> str:
    text = " / ".join(reasons)
    return text.replace("FAILED_UNRESOLVEDページ: 1", f"FAILED_UNRESOLVEDページ: {page_number}")


def draw_candidate_image(doc: Any, page_number: int, candidates: list[Any]) -> None:
    page = doc[page_number - 1]
    page.get_pixmap(matrix=fitz.Matrix(1.6, 1.6), alpha=False).save(PAGE_IMAGE)
    annotated = fitz.open()
    annotated.insert_pdf(doc, from_page=page_number - 1, to_page=page_number - 1)
    annotated_page = annotated[0]
    shape = annotated_page.new_shape()
    for index, candidate in enumerate(candidates, start=1):
        if candidate.status == CANDIDATE_REJECTED:
            continue
        rect = fitz.Rect(candidate.rect)
        shape.draw_rect(rect)
        shape.finish(color=(0.85, 0.1, 0.1), width=1.4)
        annotated_page.insert_text(fitz.Point(rect.x0, max(rect.y0 - 2, 8)), f"P22-{index:03d}", fontsize=6.5, color=(0.85, 0.1, 0.1))
    shape.commit()
    annotated_page.get_pixmap(matrix=fitz.Matrix(1.6, 1.6), alpha=False).save(CANDIDATE_IMAGE)
    annotated.close()


def main() -> int:
    start = time.perf_counter()
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    pages = read_failed_pages()
    page_number = pages[0]

    with fitz.open(SOURCE) as doc:
        page = doc[page_number - 1]
        text, words = ocr_page_text_and_words(page)
        candidates = detect_ocr_candidates(words, page.rect)
        candidates.extend(context_candidates_for_page(page_number, page.rect, words))
        active_candidates = [candidate for candidate in candidates if candidate.status != CANDIDATE_REJECTED]
        quality = evaluate_page_quality(page, text, words, candidates)
        draw_candidate_image(doc, page_number, candidates)

    initial_page_review_state = {0: PAGE_FAILED_UNRESOLVED if quality.verdict == QUALITY_FAILED else "UNREVIEWED"}
    initial_can_output, initial_reasons = final_output_status([], {0: quality}, set(), initial_page_review_state)

    after_manual_page_review_state = {0: PAGE_REVIEWED_WITH_REDACTIONS}
    after_manual_verification_state = {0: VERIFICATION_MANUAL_REGION_PASS}
    after_manual_can_output, after_manual_reasons = final_output_status([], {0: quality}, {0}, after_manual_page_review_state)

    completed_page_review_state = {0: PAGE_COMPLETED if after_manual_can_output else PAGE_REVIEWED_WITH_REDACTIONS}

    quality_rows = [
        {
            "ページ番号": page_number,
            "品質判定": quality.verdict,
            "OCR取得文字数": quality.ocr_text_chars,
            "座標付き単語数": quality.coordinate_word_count,
            "日本語文字数": quality.japanese_char_count,
            "英数字数": quality.alnum_char_count,
            "記号数": quality.symbol_char_count,
            "文字化け候補数": quality.garble_candidate_count,
            "文字領域比率": f"{quality.text_area_ratio:.6f}",
            "候補数": quality.candidate_count,
            "警告理由": quality.warning_reason,
        }
    ]
    write_csv(QUALITY_CSV, quality_rows)

    candidate_rows = [
        {
            "候補ID": f"P22-{index:03d}",
            "ページ番号": page_number,
            "情報種別": candidate.entity_type,
            "候補状態": candidate.status,
            "検出ルール": rule_name(candidate.reason),
            "候補文字列": candidate.original,
            "正規化文字列": candidate.normalized,
            "座標": rect_text(candidate.rect),
            "検出理由": candidate.reason,
        }
        for index, candidate in enumerate(active_candidates, start=1)
    ]
    write_csv(CANDIDATES_CSV, candidate_rows)

    failed_detected = quality.verdict == QUALITY_FAILED
    initial_reasons_text = display_reasons(initial_reasons, page_number)
    after_manual_reasons_text = display_reasons(after_manual_reasons, page_number)
    auto_output_stopped = not initial_can_output and "FAILED_UNRESOLVED" in initial_reasons_text
    manual_required = initial_page_review_state[0] == PAGE_FAILED_UNRESOLVED
    not_completed_before_review = initial_page_review_state[0] != PAGE_COMPLETED
    workflow_rows = [
        {"項目": "評価対象ページ", "値": str(page_number), "備考": "OCR品質FAILED検証用ページ。通常検証recallから分離"},
        {"項目": "OCR品質判定がFAILEDになるか", "値": "PASS" if failed_detected else "FAIL", "備考": quality.verdict},
        {"項目": "FAILEDページを正しく検出できるか", "値": "PASS" if failed_detected else "FAIL", "備考": quality.warning_reason},
        {"項目": "最終PDFの自動出力を停止できるか", "値": "PASS" if auto_output_stopped else "FAIL", "備考": initial_reasons_text},
        {"項目": "ユーザーへ手動確認を要求できるか", "値": "PASS" if manual_required else "FAIL", "備考": initial_page_review_state[0]},
        {"項目": "確認が終わるまで処理を完了扱いにしないか", "値": "PASS" if not_completed_before_review else "FAIL", "備考": initial_page_review_state[0]},
        {"項目": "手動確認後の出力制御状態", "値": "PASS" if after_manual_can_output else "FAIL", "備考": after_manual_reasons_text if after_manual_reasons else "出力条件は満たす"},
        {"項目": "通常検証7ページrecallへ含めるか", "値": "NO", "備考": "ページ22は別評価"},
    ]
    write_csv(WORKFLOW_CSV, workflow_rows)

    state = {
        "schema": "pdf31_page22_failed_review_state_v1",
        "source_pdf": str(SOURCE),
        "source_sha256": sha256_file(SOURCE),
        "page_number": page_number,
        "quality": quality_rows[0],
        "initial_page_review_state": initial_page_review_state[0],
        "initial_verification_state": VERIFICATION_NOT_EVALUATED,
        "manual_review_required": manual_required,
        "after_manual_page_review_state": after_manual_page_review_state[0],
        "after_manual_verification_state": after_manual_verification_state[0],
        "completed_page_review_state_if_output_runs": completed_page_review_state[0],
        "candidate_count": len(active_candidates),
        "normal_validation_recall_inclusion": False,
    }
    REVIEW_STATE_JSON.write_text(json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8")

    lines = [
        "PDF31ページ22 OCR品質FAILED評価報告書",
        f"対象ページ: {page_number}",
        "通常検証用7ページ、開発用6ページ、最終テスト用17ページは処理していません。",
        "候補生成ルールは変更していません。",
        "",
        f"OCR品質判定: {quality.verdict}",
        f"OCR取得文字数: {quality.ocr_text_chars}",
        f"座標付き単語数: {quality.coordinate_word_count}",
        f"候補数: {quality.candidate_count}",
        f"警告理由: {quality.warning_reason}",
        "",
        "評価:",
        *[f"- {row['項目']}: {row['値']} ({row['備考']})" for row in workflow_rows],
        "",
        f"処理秒数: {time.perf_counter() - start:.3f}",
        "",
        "出力:",
        f"- {QUALITY_CSV}",
        f"- {CANDIDATES_CSV}",
        f"- {WORKFLOW_CSV}",
        f"- {REVIEW_STATE_JSON}",
        f"- {PAGE_IMAGE}",
        f"- {CANDIDATE_IMAGE}",
    ]
    REPORT_TXT.write_text("\n".join(lines), encoding="utf-8")

    print("pdf31_page22_failed_evaluation=completed")
    print(f"page={page_number}")
    print(f"quality_verdict={quality.verdict}")
    print(f"ocr_text_chars={quality.ocr_text_chars}")
    print(f"coordinate_word_count={quality.coordinate_word_count}")
    print(f"candidate_count={quality.candidate_count}")
    print(f"initial_final_output_allowed={initial_can_output}")
    print("initial_block_reasons=" + initial_reasons_text)
    print(f"manual_review_required={manual_required}")
    print(f"after_manual_final_output_allowed={after_manual_can_output}")
    print(f"quality_csv={QUALITY_CSV}")
    print(f"candidate_csv={CANDIDATES_CSV}")
    print(f"workflow_csv={WORKFLOW_CSV}")
    print(f"review_state_json={REVIEW_STATE_JSON}")
    print(f"report={REPORT_TXT}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
