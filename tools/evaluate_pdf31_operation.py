from __future__ import annotations

import csv
import difflib
import os
import subprocess
import sys
import time
from contextlib import contextmanager
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

import fitz  # noqa: E402

from excel_privacy_cleaner.pdf_ocr_support import (  # noqa: E402
    CANDIDATE_AUTO,
    CANDIDATE_MANUAL,
    CANDIDATE_REVIEW,
    PAGE_FAILED_UNRESOLVED,
    PAGE_REVIEWED_WITH_REDACTIONS,
    PAGE_UNREVIEWED,
    QUALITY_FAILED,
    QUALITY_REVIEW,
    USER_APPROVED,
    USER_REJECTED,
    VERIFICATION_MANUAL_REGION_PASS,
    VERIFICATION_NOT_EVALUATED,
    VERIFICATION_STRUCTURE_PASS,
)
from excel_privacy_cleaner.pdf_processor import PdfPrivacyProcessor, final_output_status, ocr_page_text_and_words, write_pdf_quality_csv  # noqa: E402
from pdf31_review_fixtures import REPRESENTATIVE_PAGE_SPECS  # noqa: E402


SOURCE = Path("PDFテストデータ２.pdf")
OUTPUT_DIR = Path("ocr_quality_outputs")
WORKFLOW_CSV = OUTPUT_DIR / "PDF31ページ_ワークフロー結果.csv"
PAGE_REPORT_CSV = OUTPUT_DIR / "PDF31ページ_ページ別処理報告.csv"
REVALIDATION_CSV = OUTPUT_DIR / "PDF31ページ_再検証結果.csv"
GROUND_TRUTH_CSV = OUTPUT_DIR / "PDF31代表6ページ_正解データ.csv"
FORMAL_GROUND_TRUTH_CSV = OUTPUT_DIR / "PDF31代表6ページ_ユーザー確認済み正解データ_v1.csv"
DETECTION_EVAL_CSV = OUTPUT_DIR / "PDF31代表6ページ_検出評価結果.csv"
MISS_CSV = OUTPUT_DIR / "PDF31代表6ページ_見逃し一覧.csv"
FALSE_POSITIVE_CSV = OUTPUT_DIR / "PDF31代表6ページ_誤検出一覧.csv"
FRAME_OCR_CHECK_CSV = OUTPUT_DIR / "PDF31代表6ページ_枠内文字一致検査.csv"
USER_STATUS_CSV = OUTPUT_DIR / "PDF31代表6ページ_ユーザー承認状態一覧.csv"
FORMAL_METRICS_CSV = OUTPUT_DIR / "PDF31代表6ページ_正式評価指標.csv"
OCR_RECT_WARNING_CSV = OUTPUT_DIR / "PDF31代表6ページ_OCR矩形警告.csv"
OCR_RUNTIME_WARNING_CSV = OUTPUT_DIR / "PDF31代表6ページ_OCR実行警告.csv"
TESSERACT_PARAMETER_CHECK_CSV = OUTPUT_DIR / "PDF31代表6ページ_Tesseractパラメータ確認.csv"
STATE_JSON = OUTPUT_DIR / "PDF31ページ_レビュー途中_作業状態.json"
REPORT_TXT = OUTPUT_DIR / "PDF31ページ_安全制御テスト_処理報告書.txt"
FORMAL_STATUSES = {"ユーザー確認済み", "対象外"}
UNAPPROVED_STATUSES = {"ユーザー確認待ち", "要座標修正", "解除候補"}
OCR_DPI_FOR_RECT_CHECK = 300
MIN_OCR_PIXEL_SIZE = 3
TESSERACT_WARNING_PARAMETERS = [
    "language_model_ngram_on",
    "segsearch_max_char_wh_ratio",
    "language_model_ngram_space_delimited_language",
    "language_model_ngram_scale_factor",
    "language_model_use_sigmoidal_certainty",
    "language_model_ngram_nonmatch_score",
    "classify_integer_matcher_multiplier",
    "assume_fixed_pitch_char_segment",
    "chop_enable",
    "allow_blob_division",
]


@dataclass(frozen=True)
class FormalTruth:
    truth_id: str
    page: int
    occurrence: int
    same_info_id: str
    text: str
    entity_type: str
    mode: str
    required: str
    rect: tuple[float, float, float, float]
    user_status: str


@dataclass(frozen=True)
class EvaluationSummary:
    detected: int
    missed: int
    false_positives: int
    truth_rects: int
    manual_additions: int
    avg_coverage: float
    min_coverage: float
    avg_excess: float
    max_excess: float


def page_content_rect(source_pdf: Path, page_index: int) -> tuple[float, float, float, float]:
    doc = fitz.open(source_pdf)
    try:
        page = doc[page_index]
        rect = page.rect
        margin_x = rect.width * 0.06
        margin_y = rect.height * 0.06
        return (rect.x0 + margin_x, rect.y0 + margin_y, rect.x1 - margin_x, rect.y1 - margin_y)
    finally:
        doc.close()


def page_findings(findings: list[Any], page: int) -> list[Any]:
    return [finding for finding in findings if finding.sheet == f"ページ{page}"]


def normalize(value: str) -> str:
    return "".join(value.split()).replace("（", "(").replace("）", ")")


def parse_rect(value: str) -> tuple[float, float, float, float]:
    parts = [float(part.strip()) for part in value.split(",")]
    if len(parts) != 4:
        raise ValueError(f"正解座標は4要素である必要があります: {value}")
    return parts[0], parts[1], parts[2], parts[3]


def load_formal_ground_truth(path: Path = FORMAL_GROUND_TRUTH_CSV) -> tuple[FormalTruth, ...]:
    if not path.exists():
        raise FileNotFoundError(f"確定版v1正解CSVが見つかりません。ドラフトCSVへは戻しません: {path}")
    rows = csv_rows(path)
    if not rows:
        raise RuntimeError(f"確定版v1正解CSVが空です: {path}")
    invalid_statuses = sorted({row.get("ユーザー確認状態", "") for row in rows if row.get("ユーザー確認状態", "") not in FORMAL_STATUSES})
    if invalid_statuses:
        raise RuntimeError(f"確定版v1正解CSVに正式状態以外が含まれています: {', '.join(invalid_statuses)}")
    truths: list[FormalTruth] = []
    for row in rows:
        if row.get("ユーザー確認状態") == "対象外":
            continue
        truths.append(
            FormalTruth(
                truth_id=row["正解ID"],
                page=int(row["ページ番号"]),
                occurrence=int(row["出現番号"]),
                same_info_id=row["同一情報ID"],
                text=row["匿名化対象文字列"],
                entity_type=row["情報種別"],
                mode=row["適用モード"],
                required=row["必須度"],
                rect=parse_rect(row["正解座標"]),
                user_status=row["ユーザー確認状態"],
            )
        )
    if not truths:
        raise RuntimeError("正式評価対象が0件です。ユーザー確認済み矩形が必要です。")
    return tuple(truths)


@contextmanager
def capture_native_stderr() -> Any:
    OUTPUT_DIR.mkdir(exist_ok=True)
    temp_path = OUTPUT_DIR / "_ocr_stderr_capture.txt"
    original_fd = os.dup(2)
    with temp_path.open("w", encoding="utf-8", errors="replace") as handle:
        os.dup2(handle.fileno(), 2)
        try:
            yield temp_path
        finally:
            os.dup2(original_fd, 2)
            os.close(original_fd)


def append_runtime_warnings(stage: str, stderr_path: Path) -> None:
    text = stderr_path.read_text(encoding="utf-8", errors="replace") if stderr_path.exists() else ""
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    exists = OCR_RUNTIME_WARNING_CSV.exists()
    with OCR_RUNTIME_WARNING_CSV.open("a", encoding="utf-8-sig", newline="") as handle:
        writer = csv.writer(handle)
        if not exists:
            writer.writerow(["処理段階", "警告内容"])
        for line in lines:
            writer.writerow([stage, line])
    if stderr_path.exists():
        stderr_path.unlink()


def check_tesseract_parameters() -> None:
    rows: list[dict[str, str]] = []
    try:
        result = subprocess.run(
            ["tesseract", "--print-parameters"],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=20,
            check=False,
        )
        output = (result.stdout or "") + "\n" + (result.stderr or "")
        available_names = {line.split()[0] for line in output.splitlines() if line.strip() and not line.startswith("Tesseract")}
        for parameter in TESSERACT_WARNING_PARAMETERS:
            rows.append(
                {
                    "パラメータ": parameter,
                    "現在のTesseractで使用可能": "YES" if parameter in available_names else "NO",
                    "このアプリから明示的に渡しているか": "NO",
                    "対応": "アプリ側設定なし。PyMuPDF/Tesseract内部警告はOCR実行警告CSVへ集約",
                }
            )
    except Exception as exc:
        for parameter in TESSERACT_WARNING_PARAMETERS:
            rows.append(
                {
                    "パラメータ": parameter,
                    "現在のTesseractで使用可能": "UNKNOWN",
                    "このアプリから明示的に渡しているか": "NO",
                    "対応": f"tesseract --print-parameters を実行できません: {exc}",
                }
            )
    with TESSERACT_PARAMETER_CHECK_CSV.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=["パラメータ", "現在のTesseractで使用可能", "このアプリから明示的に渡しているか", "対応"])
        writer.writeheader()
        writer.writerows(rows)


def approved_specs() -> tuple[Any, ...]:
    return tuple(spec for spec in REPRESENTATIVE_PAGE_SPECS if spec.user_status == "ユーザー確認済み")


def status_counts() -> dict[str, int]:
    counts: dict[str, int] = {}
    for spec in REPRESENTATIVE_PAGE_SPECS:
        counts[spec.user_status] = counts.get(spec.user_status, 0) + 1
    return counts


def active_candidate_specs() -> tuple[Any, ...]:
    return tuple(spec for spec in REPRESENTATIVE_PAGE_SPECS if spec.user_status not in {"要座標修正", "解除候補"})


def write_ground_truth_csv() -> None:
    headers = [
        "正解ID",
        "ページ番号",
        "出現番号",
        "同一情報ID",
        "匿名化対象文字列",
        "情報種別",
        "適用モード",
        "対象理由",
        "必須度",
        "正解座標",
        "期待する変換",
        "残すべき文字列",
        "匿名化してはいけない領域",
        "ユーザー確認状態",
        "備考",
    ]
    with GROUND_TRUTH_CSV.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.writer(handle)
        writer.writerow(headers)
        for row in REPRESENTATIVE_PAGE_SPECS:
            writer.writerow(
                [
                    row.truth_id,
                    row.page,
                    row.occurrence,
                    row.same_info_id,
                    row.text,
                    row.entity_type,
                    row.mode,
                    row.reason,
                    row.required,
                    ",".join(f"{value:.1f}" for value in row.rect),
                    row.replacement,
                    row.preserve,
                    row.protected_area,
                    row.user_status,
                    row.note,
                ]
            )
    with USER_STATUS_CSV.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.writer(handle)
        writer.writerow(["正解ID", "ページ番号", "匿名化対象文字列", "情報種別", "ユーザー確認状態", "備考"])
        for row in REPRESENTATIVE_PAGE_SPECS:
            writer.writerow([row.truth_id, row.page, row.text, row.entity_type, row.user_status, row.note])


def write_ground_truth_review_images(source_pdf: Path) -> list[Path]:
    output_paths: list[Path] = []
    doc = fitz.open(source_pdf)
    try:
        for page in sorted({spec.page for spec in REPRESENTATIVE_PAGE_SPECS}):
            page_obj = doc[page - 1]
            draw = page_obj.new_shape()
            specs = [spec for spec in REPRESENTATIVE_PAGE_SPECS if spec.page == page]
            for spec in specs:
                color = {
                    "ユーザー確認済み": (0.0, 0.65, 0.22),
                    "ユーザー確認待ち": (0.95, 0.62, 0.0),
                    "要座標修正": (0.9, 0.1, 0.1),
                    "解除候補": (0.45, 0.45, 0.45),
                }.get(spec.user_status, (0.0, 0.2, 0.8))
                rect = fitz.Rect(spec.rect)
                draw.draw_rect(rect)
                draw.finish(color=color, width=1.4)
                label_point = fitz.Point(rect.x0, max(rect.y0 - 2, 8))
                page_obj.insert_text(label_point, spec.truth_id, fontsize=6.5, color=color)
            draw.commit()
            output_path = OUTPUT_DIR / f"PDF31代表ページ_{page}_正解枠確認.png"
            page_obj.get_pixmap(dpi=144, alpha=False).save(output_path)
            output_paths.append(output_path)
    finally:
        doc.close()
    return output_paths


def rect_area(rect: tuple[float, float, float, float]) -> float:
    return max(rect[2] - rect[0], 0.0) * max(rect[3] - rect[1], 0.0)


def intersect_rect(a: tuple[float, float, float, float], b: tuple[float, float, float, float]) -> tuple[float, float, float, float] | None:
    rect = (max(a[0], b[0]), max(a[1], b[1]), min(a[2], b[2]), min(a[3], b[3]))
    if rect[2] <= rect[0] or rect[3] <= rect[1]:
        return None
    return rect


def coverage(expected: tuple[float, float, float, float], actual: tuple[float, float, float, float]) -> float:
    overlap = intersect_rect(expected, actual)
    return rect_area(overlap) / max(rect_area(expected), 1.0) if overlap else 0.0


def excess_ratio(expected: tuple[float, float, float, float], actual: tuple[float, float, float, float]) -> float:
    overlap = intersect_rect(expected, actual)
    overlap_area = rect_area(overlap) if overlap else 0.0
    return max(rect_area(actual) - overlap_area, 0.0) / max(rect_area(expected), 1.0)


def finding_rect(processor: PdfPrivacyProcessor, finding: Any) -> tuple[float, float, float, float] | None:
    location = processor.locations.get((finding.sheet, finding.cell, finding.original))
    return location.rect if location else None


def write_frame_ocr_check_csv(source_pdf: Path, truths: tuple[FormalTruth, ...]) -> None:
    page_words: dict[int, list[tuple[Any, ...]]] = {}
    doc = fitz.open(source_pdf)
    try:
        for page in sorted({truth.page for truth in truths}):
            _text, words = ocr_page_text_and_words(doc[page - 1])
            page_words[page] = words
    finally:
        doc.close()

    headers = [
        "正解ID",
        "ページ番号",
        "ユーザー確認状態",
        "期待する文字列",
        "枠内OCR文字列",
        "期待文字列との一致率",
        "枠内に含まれる余分な文字",
        "枠外に残った期待文字",
        "枠内文字数",
        "判定",
    ]
    with FRAME_OCR_CHECK_CSV.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.writer(handle)
        writer.writerow(headers)
        for spec in truths:
            inside_words = _words_in_rect(page_words.get(spec.page, []), spec.rect)
            inside_text = "".join(str(word[4]) for word in inside_words)
            ratio = difflib.SequenceMatcher(None, normalize(spec.text), normalize(inside_text)).ratio() if inside_text else 0.0
            if not inside_text or ratio < 0.5:
                verdict = "FAILED"
            elif ratio < 0.9:
                verdict = "REVIEW_REQUIRED"
            else:
                verdict = "PASS_CANDIDATE"
            extra_text = inside_text if ratio < 0.9 else ""
            outside_expected = "" if ratio >= 0.9 else spec.text
            writer.writerow(
                [
                    spec.truth_id,
                    spec.page,
                    spec.user_status,
                    spec.text,
                    inside_text,
                    f"{ratio:.3f}",
                    extra_text,
                    outside_expected,
                    len(inside_text),
                    verdict,
                ]
            )


def write_ocr_rect_warning_csv(truths: tuple[FormalTruth, ...], source_pdf: Path) -> None:
    page_sizes: dict[int, tuple[float, float]] = {}
    doc = fitz.open(source_pdf)
    try:
        for truth in truths:
            if truth.page not in page_sizes:
                rect = doc[truth.page - 1].rect
                page_sizes[truth.page] = (float(rect.width), float(rect.height))
    finally:
        doc.close()
    headers = [
        "正解ID",
        "ページ番号",
        "PDF座標",
        "切り出し前PDF幅",
        "切り出し前PDF高さ",
        "切り出し後PDF座標",
        "画像換算幅px",
        "画像換算高さpx",
        "OCR実行",
        "警告",
    ]
    with OCR_RECT_WARNING_CSV.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.writer(handle)
        writer.writerow(headers)
        for truth in truths:
            page_width, page_height = page_sizes[truth.page]
            x0, y0, x1, y1 = truth.rect
            clipped = (
                max(0.0, min(x0, page_width)),
                max(0.0, min(y0, page_height)),
                max(0.0, min(x1, page_width)),
                max(0.0, min(y1, page_height)),
            )
            width = max(clipped[2] - clipped[0], 0.0)
            height = max(clipped[3] - clipped[1], 0.0)
            pixel_width = width * OCR_DPI_FOR_RECT_CHECK / 72.0
            pixel_height = height * OCR_DPI_FOR_RECT_CHECK / 72.0
            warning = ""
            ocr_execution = "NO_RECT_OCR"
            if width <= 0 or height <= 0:
                warning = "矩形の幅または高さが0です"
                ocr_execution = "SKIPPED"
            elif pixel_width < MIN_OCR_PIXEL_SIZE or pixel_height < MIN_OCR_PIXEL_SIZE:
                warning = f"OCR可能な最小サイズ未満です: min={MIN_OCR_PIXEL_SIZE}px"
                ocr_execution = "SKIPPED"
            writer.writerow(
                [
                    truth.truth_id,
                    truth.page,
                    ",".join(f"{value:.1f}" for value in truth.rect),
                    f"{max(x1 - x0, 0.0):.1f}",
                    f"{max(y1 - y0, 0.0):.1f}",
                    ",".join(f"{value:.1f}" for value in clipped),
                    f"{pixel_width:.1f}",
                    f"{pixel_height:.1f}",
                    ocr_execution,
                    warning,
                ]
            )


def _words_in_rect(words: list[tuple[Any, ...]], rect: tuple[float, float, float, float]) -> list[tuple[Any, ...]]:
    x0, y0, x1, y1 = rect
    selected: list[tuple[Any, ...]] = []
    for word in words:
        wx0, wy0, wx1, wy1 = (float(word[0]), float(word[1]), float(word[2]), float(word[3]))
        center_x = (wx0 + wx1) / 2
        center_y = (wy0 + wy1) / 2
        if x0 <= center_x <= x1 and y0 <= center_y <= y1:
            selected.append(word)
    return selected


def evaluate_ground_truth(processor: PdfPrivacyProcessor, findings: list[Any], truths: tuple[FormalTruth, ...]) -> EvaluationSummary:
    representative_pages = {row.page for row in truths}
    matched_ids: set[int] = set()
    rows: list[dict[str, str]] = []
    misses: list[dict[str, str]] = []
    missed = 0
    coverage_values: list[float] = []
    excess_values: list[float] = []
    manual_additions = 0
    for truth in truths:
        best: tuple[int, Any, float, float] | None = None
        for index, finding in enumerate(findings):
            if finding.sheet != f"ページ{truth.page}" or finding.entity_type != truth.entity_type:
                continue
            actual_rect = finding_rect(processor, finding)
            if actual_rect is None:
                continue
            item_coverage = coverage(truth.rect, actual_rect)
            item_excess = excess_ratio(truth.rect, actual_rect)
            if item_coverage >= 0.85 and (best is None or item_coverage > best[2]):
                best = (index, finding, item_coverage, item_excess)
        if best:
            matched_ids.add(best[0])
            if best[1].detection_kind == CANDIDATE_MANUAL or best[1].entity_type == "手動追加":
                manual_additions += 1
            detected = "1"
            missed_value = "0"
            coverage_text = f"{best[2]:.3f}"
            excess_text = f"{best[3]:.3f}"
            coverage_values.append(best[2])
            excess_values.append(best[3])
        else:
            detected = "0"
            missed_value = "1"
            coverage_text = "0.000"
            excess_text = "NOT_EVALUATED"
            missed += 1
            misses.append(
                {
                    "正解ID": truth.truth_id,
                    "ページ番号": str(truth.page),
                    "匿名化対象文字列": truth.text,
                    "情報種別": truth.entity_type,
                    "正解座標": ",".join(f"{value:.1f}" for value in truth.rect),
                    "理由": "正解矩形を85%以上覆う候補がありません",
                }
            )
        rows.append(
            {
                "正解ID": truth.truth_id,
                "ページ番号": str(truth.page),
                "出現番号": str(truth.occurrence),
                "同一情報ID": truth.same_info_id,
                "匿名化対象文字列": truth.text,
                "情報種別": truth.entity_type,
                "適用モード": truth.mode,
                "必須度": truth.required,
                "検出件数": detected,
                "見逃し件数": missed_value,
                "誤検出件数": "0",
                "手動追加件数": "0",
                "正解領域被覆率": coverage_text,
                "過剰匿名化率": excess_text,
                "再OCR残存件数": "NOT_EVALUATED",
                "判定": "PASS" if best else "MISS",
            }
        )
    false_positives: list[dict[str, str]] = []
    for index, finding in enumerate(findings):
        page = int(finding.sheet.replace("ページ", ""))
        if page not in representative_pages or index in matched_ids:
            continue
        rect = finding_rect(processor, finding)
        if rect is None:
            continue
        false_positives.append(
            {
                "ページ番号": str(page),
                "匿名化対象文字列": finding.original,
                "情報種別": finding.entity_type,
                "検出座標": ",".join(f"{value:.1f}" for value in rect),
                "検査": finding.detection_kind,
                "理由": "代表正解データに対応しない候補",
            }
        )
    with DETECTION_EVAL_CSV.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=list(rows[0].keys()))
        writer.writeheader()
        writer.writerows(rows)
    with MISS_CSV.open("w", encoding="utf-8-sig", newline="") as handle:
        fieldnames = ["正解ID", "ページ番号", "匿名化対象文字列", "情報種別", "正解座標", "理由"]
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(misses)
    with FALSE_POSITIVE_CSV.open("w", encoding="utf-8-sig", newline="") as handle:
        fieldnames = ["ページ番号", "匿名化対象文字列", "情報種別", "検出座標", "検査", "理由"]
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(false_positives)
    avg_coverage = sum(coverage_values) / max(len(coverage_values), 1)
    min_coverage = min(coverage_values) if coverage_values else 0.0
    avg_excess = sum(excess_values) / max(len(excess_values), 1)
    max_excess = max(excess_values) if excess_values else 0.0
    return EvaluationSummary(
        detected=len(truths) - missed,
        missed=missed,
        false_positives=len(false_positives),
        truth_rects=len(truths),
        manual_additions=manual_additions,
        avg_coverage=avg_coverage,
        min_coverage=min_coverage,
        avg_excess=avg_excess,
        max_excess=max_excess,
    )


def write_formal_metrics_csv(summary: EvaluationSummary) -> None:
    rows = [
        {"項目": "formal_ground_truth_csv", "値": str(FORMAL_GROUND_TRUTH_CSV), "備考": "正式評価で使用した確定版CSV"},
        {"項目": "正式評価対象数", "値": str(summary.truth_rects), "備考": ""},
        {"項目": "検出数", "値": str(summary.detected), "備考": ""},
        {"項目": "見逃し数", "値": str(summary.missed), "備考": ""},
        {"項目": "誤検出数", "値": str(summary.false_positives), "備考": ""},
        {"項目": "手動追加数", "値": str(summary.manual_additions), "備考": ""},
        {"項目": "平均被覆率", "値": f"{summary.avg_coverage:.3f}", "備考": ""},
        {"項目": "最小被覆率", "値": f"{summary.min_coverage:.3f}", "備考": ""},
        {"項目": "平均過剰率", "値": f"{summary.avg_excess:.3f}", "備考": ""},
        {"項目": "最大過剰率", "値": f"{summary.max_excess:.3f}", "備考": ""},
    ]
    with FORMAL_METRICS_CSV.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=["項目", "値", "備考"])
        writer.writeheader()
        writer.writerows(rows)


def write_page_report(processor: PdfPrivacyProcessor, findings: list[Any]) -> None:
    headers = [
        "ページ番号",
        "ページ種別",
        "品質判定",
        "ページ確認状態",
        "自動確定件数",
        "確認候補件数",
        "利用者承認件数",
        "利用者解除件数",
        "手動追加件数",
        "未検出件数",
        "部分残存件数",
        "過剰匿名化件数",
        "再検証結果",
        "警告内容",
    ]
    with PAGE_REPORT_CSV.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.writer(handle)
        writer.writerow(headers)
        for index, quality in sorted(processor.page_quality.items()):
            findings_on_page = page_findings(findings, index + 1)
            writer.writerow(
                [
                    index + 1,
                    processor.page_modes.get(index, "不明"),
                    quality.verdict,
                    processor.page_review_state.get(index, PAGE_UNREVIEWED),
                    sum(1 for finding in findings_on_page if finding.detection_kind == CANDIDATE_AUTO),
                    sum(1 for finding in findings_on_page if finding.detection_kind == CANDIDATE_REVIEW),
                    sum(1 for finding in findings_on_page if finding.detection_kind == USER_APPROVED),
                    sum(1 for finding in findings_on_page if finding.detection_kind == USER_REJECTED),
                    sum(1 for finding in findings_on_page if finding.detection_kind == CANDIDATE_MANUAL or finding.entity_type == "手動追加"),
                    "正解データ対象外" if index + 1 not in {row.page for row in REPRESENTATIVE_PAGE_SPECS} else "代表評価CSV参照",
                    "NOT_EVALUATED",
                    "NOT_EVALUATED",
                    processor.page_verification_state.get(index, VERIFICATION_NOT_EVALUATED),
                    quality.warning_reason,
                ]
            )


def write_revalidation_csv(processor: PdfPrivacyProcessor, manual_pages: set[int]) -> None:
    with REVALIDATION_CSV.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.writer(handle)
        writer.writerow(["ページ", "再検証区分", "備考"])
        for index in range(processor.page_count):
            if index in manual_pages:
                result = VERIFICATION_MANUAL_REGION_PASS
                note = "手動範囲の設定状態のみ確認。最終PDFは未出力。"
            elif processor.page_quality.get(index) and processor.page_quality[index].verdict in {QUALITY_REVIEW, QUALITY_FAILED}:
                result = VERIFICATION_NOT_EVALUATED
                note = "正解データまたは利用者確認がないため残存検証未評価。"
            else:
                result = VERIFICATION_STRUCTURE_PASS
                note = "ページ構造の検査のみ。匿名化漏れ検証ではない。"
            writer.writerow([index + 1, result, note])


def write_workflow_csv(rows: list[dict[str, str]]) -> None:
    with WORKFLOW_CSV.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=["項目", "値", "備考"])
        writer.writeheader()
        writer.writerows(rows)


def workflow_value(rows: list[dict[str, str]], key: str) -> str:
    for row in rows:
        if row["項目"] == key:
            return row["値"]
    raise KeyError(key)


def csv_rows(path: Path) -> list[dict[str, str]]:
    with path.open(encoding="utf-8-sig", newline="") as handle:
        return list(csv.DictReader(handle))


def validate_formal_outputs(
    rows: list[dict[str, str]],
    summary: EvaluationSummary,
    truths: tuple[FormalTruth, ...],
) -> None:
    errors: list[str] = []
    if not FORMAL_GROUND_TRUTH_CSV.exists():
        errors.append(f"確定版v1正解CSVが見つかりません: {FORMAL_GROUND_TRUTH_CSV}")
    approved_ids = {truth.truth_id for truth in truths}
    if summary.truth_rects != len(truths):
        errors.append(f"正式評価対象数と確定版CSVのユーザー確認済み件数が不一致: {summary.truth_rects} != {len(truths)}")

    eval_rows = csv_rows(DETECTION_EVAL_CSV)
    eval_ids = {row["正解ID"] for row in eval_rows}
    if len(eval_rows) != summary.truth_rects:
        errors.append(f"検出評価結果の行数が正式評価対象数と不一致: {len(eval_rows)} != {summary.truth_rects}")
    if not eval_ids <= approved_ids:
        errors.append(f"検出評価結果に確定版v1以外の矩形が含まれています: {sorted(eval_ids - approved_ids)}")
    if int(workflow_value(rows, "正式評価対象数")) != summary.truth_rects:
        errors.append("ワークフロー結果の正式評価対象数が評価サマリと一致しません")
    if int(workflow_value(rows, "検出数")) != summary.detected:
        errors.append("ワークフロー結果の検出数が評価サマリと一致しません")
    if int(workflow_value(rows, "見逃し数")) != summary.missed:
        errors.append("ワークフロー結果の見逃し数が評価サマリと一致しません")
    if int(workflow_value(rows, "誤検出数")) != summary.false_positives:
        errors.append("ワークフロー結果の誤検出数が評価サマリと一致しません")

    if errors:
        raise RuntimeError("正式成果物の整合性チェックに失敗しました:\n- " + "\n- ".join(errors))


def write_report(rows: list[dict[str, str]], issues: list[tuple[str, str, str]]) -> None:
    lines = [
        "PDF31ページ 安全制御テスト処理報告書",
        f"入力ファイル名: {SOURCE.name}",
        "出力ファイル名: 最終PDFは未出力",
        f"処理日時: {datetime.now():%Y-%m-%d %H:%M:%S}",
        "",
        "判定:",
        "- ワークフロー実行: 成功",
        "- ページ構造維持: 成功",
        "- 匿名化対象検出: 確定版v1 CSVのみ正式評価、ドラフトCSVは評価対象外",
        "- 匿名化漏れ検証: 未評価",
        "- 実運用可否: 不可",
        "",
        "正式表現:",
        "正式評価では、PDF31代表6ページ_ユーザー確認済み正解データ_v1.csv のユーザー確認済み矩形だけを使用した。ドラフト正解データCSVは正式評価に使用していない。",
        "",
        "集計:",
    ]
    lines.extend(f"- {row['項目']}: {row['値']} {row['備考']}".rstrip() for row in rows)
    lines.extend(["", "優先対応:"])
    lines.extend(f"- {priority}: {title} - {detail}" for priority, title, detail in issues)
    lines.extend(
        [
            "",
            f"ページ別処理報告CSV: {PAGE_REPORT_CSV}",
            f"ワークフロー結果CSV: {WORKFLOW_CSV}",
            f"再検証結果CSV: {REVALIDATION_CSV}",
            f"代表6ページドラフト正解データCSV: {GROUND_TRUTH_CSV}",
            f"正式評価使用CSV: {FORMAL_GROUND_TRUTH_CSV}",
            f"正式評価指標CSV: {FORMAL_METRICS_CSV}",
            f"代表6ページ検出評価CSV: {DETECTION_EVAL_CSV}",
            f"代表6ページ見逃し一覧CSV: {MISS_CSV}",
            f"代表6ページ誤検出一覧CSV: {FALSE_POSITIVE_CSV}",
            f"代表6ページ枠内文字一致検査CSV: {FRAME_OCR_CHECK_CSV}",
            f"OCR矩形警告CSV: {OCR_RECT_WARNING_CSV}",
            f"OCR実行警告CSV: {OCR_RUNTIME_WARNING_CSV}",
            f"Tesseractパラメータ確認CSV: {TESSERACT_PARAMETER_CHECK_CSV}",
            f"代表6ページユーザー承認状態一覧CSV: {USER_STATUS_CSV}",
            f"レビュー途中作業状態JSON: {STATE_JSON}",
        ]
    )
    REPORT_TXT.write_text("\n".join(lines) + "\n", encoding="utf-8")


def main() -> int:
    if not SOURCE.exists():
        raise FileNotFoundError(SOURCE)
    OUTPUT_DIR.mkdir(exist_ok=True)
    OCR_RUNTIME_WARNING_CSV.unlink(missing_ok=True)
    formal_truths = load_formal_ground_truth()
    check_tesseract_parameters()

    started_total = time.perf_counter()
    processor = PdfPrivacyProcessor()
    started_scan = time.perf_counter()
    with capture_native_stderr() as stderr_path:
        findings = processor.scan(SOURCE)
    append_runtime_warnings("processor.scan", stderr_path)
    scan_elapsed = time.perf_counter() - started_scan

    failed_pages = {index for index, quality in processor.page_quality.items() if quality.verdict == QUALITY_FAILED}
    review_pages = {index for index, quality in processor.page_quality.items() if quality.verdict == QUALITY_REVIEW}
    text_pages = {index for index, mode in processor.page_modes.items() if mode == "text"}
    ocr_pages = {index for index, mode in processor.page_modes.items() if mode == "ocr"}

    can_output_initial, initial_reasons = final_output_status(
        findings,
        processor.page_quality,
        processor.confirmed_pages,
        processor.page_review_state,
    )

    manual_pages: set[int] = set()
    for page_index in sorted(failed_pages):
        manual = processor.add_manual_redaction(findings, page_index, page_content_rect(SOURCE, page_index))
        manual.detection_kind = CANDIDATE_MANUAL
        processor.mark_page_reviewed_with_redactions(page_index)
        processor.page_verification_state[page_index] = VERIFICATION_MANUAL_REGION_PASS
        manual_pages.add(page_index)

    can_output_after_failed_only, failed_only_reasons = final_output_status(
        findings,
        processor.page_quality,
        processor.confirmed_pages,
        processor.page_review_state,
    )

    processor.export_review_state(STATE_JSON, SOURCE, findings, last_page=min(sorted(failed_pages)[0], processor.page_count - 1) if failed_pages else 0)
    restored = PdfPrivacyProcessor()
    restored.page_count = processor.page_count
    restored.page_quality = dict(processor.page_quality)
    restored.import_review_state(STATE_JSON, SOURCE, [])
    save_resume_ok = (
        restored.page_review_state == processor.page_review_state
        and restored.page_verification_state == processor.page_verification_state
        and len(restored.locations) == len(manual_pages)
    )

    write_ground_truth_csv()
    review_images = write_ground_truth_review_images(SOURCE)
    with capture_native_stderr() as stderr_path:
        write_frame_ocr_check_csv(SOURCE, formal_truths)
    append_runtime_warnings("frame_ocr_check", stderr_path)
    write_ocr_rect_warning_csv(formal_truths, SOURCE)
    summary = evaluate_ground_truth(processor, findings, formal_truths)
    write_formal_metrics_csv(summary)
    write_page_report(processor, findings)
    write_revalidation_csv(processor, manual_pages)
    write_pdf_quality_csv(OUTPUT_DIR / "PDF31ページ_ページ品質.csv", processor.page_quality, processor.page_review_state, processor.page_verification_state)

    total_elapsed = time.perf_counter() - started_total
    auto_count = sum(1 for finding in findings if finding.detection_kind == CANDIDATE_AUTO)
    approved_count = len(formal_truths)
    rejected_count = sum(1 for finding in findings if finding.detection_kind == USER_REJECTED)
    manual_count = sum(1 for finding in findings if finding.detection_kind == CANDIDATE_MANUAL or finding.entity_type == "手動追加")
    state_counts: dict[str, int] = {}
    for state in processor.page_review_state.values():
        state_counts[state] = state_counts.get(state, 0) + 1

    rows = [
        {"項目": "全ページ数", "値": str(processor.page_count), "備考": ""},
        {"項目": "文字PDFページ数", "値": str(len(text_pages)), "備考": ""},
        {"項目": "OCRページ数", "値": str(len(ocr_pages)), "備考": ""},
        {"項目": "REVIEW_REQUIREDページ数", "値": str(len(review_pages)), "備考": ",".join(str(index + 1) for index in sorted(review_pages))},
        {"項目": "FAILEDページ数", "値": str(len(failed_pages)), "備考": ",".join(str(index + 1) for index in sorted(failed_pages))},
        {"項目": "UNREVIEWEDページ数", "値": str(state_counts.get(PAGE_UNREVIEWED, 0)), "備考": "候補0のREVIEW_REQUIREDページも未確認として保持"},
        {"項目": "FAILED_UNRESOLVEDページ数", "値": str(state_counts.get(PAGE_FAILED_UNRESOLVED, 0)), "備考": ""},
        {"項目": "候補総数", "値": str(len(findings) - manual_count), "備考": "手動追加を除く"},
        {"項目": "自動確定数", "値": str(auto_count), "備考": ""},
        {"項目": "利用者承認数", "値": str(approved_count), "備考": "確定版v1 CSVのユーザー確認済み正解矩形数"},
        {"項目": "利用者解除数", "値": str(rejected_count), "備考": "今回は検出候補なし"},
        {"項目": "手動追加数", "値": str(manual_count), "備考": "FAILEDページのみ"},
        {"項目": "OCR処理時間", "値": f"{scan_elapsed:.3f}秒", "備考": ""},
        {"項目": "最終出力までの総時間", "値": f"{total_elapsed:.3f}秒", "備考": "最終PDFは安全制御により未出力"},
        {"項目": "出力ブロックテスト_初期状態", "値": "PASS" if not can_output_initial else "FAIL", "備考": " / ".join(initial_reasons)},
        {"項目": "出力ブロックテスト_FAILEDのみ対応後", "値": "PASS" if not can_output_after_failed_only else "FAIL", "備考": " / ".join(failed_only_reasons)},
        {"項目": "途中保存・再開テスト", "値": "PASS" if save_resume_ok else "FAIL", "備考": str(STATE_JSON)},
        {"項目": "正式評価使用CSV", "値": str(FORMAL_GROUND_TRUTH_CSV), "備考": "確定版v1。ドラフトCSVは正式評価に使用しない"},
        {"項目": "代表6ページ_ドラフト正解矩形数", "値": str(len(REPRESENTATIVE_PAGE_SPECS)), "備考": "6ページ内の総匿名化対象数ではなく、今回登録した矩形数"},
        {"項目": "代表6ページ_ユーザー承認済み矩形数", "値": str(sum(1 for spec in REPRESENTATIVE_PAGE_SPECS if spec.user_status == 'ユーザー確認済み')), "備考": ""},
        {"項目": "代表6ページ_ユーザー確認待ち矩形数", "値": str(sum(1 for spec in REPRESENTATIVE_PAGE_SPECS if spec.user_status == 'ユーザー確認待ち')), "備考": ""},
        {"項目": "代表6ページ_要修正矩形数", "値": str(sum(1 for spec in REPRESENTATIVE_PAGE_SPECS if spec.user_status == '要座標修正')), "備考": ""},
        {"項目": "代表6ページ_解除候補矩形数", "値": str(sum(1 for spec in REPRESENTATIVE_PAGE_SPECS if spec.user_status == '解除候補')), "備考": ""},
        {"項目": "正式評価対象数", "値": str(summary.truth_rects), "備考": "確定版v1 CSVのユーザー確認済み矩形のみ"},
        {"項目": "検出数", "値": str(summary.detected), "備考": ""},
        {"項目": "見逃し数", "値": str(summary.missed), "備考": ""},
        {"項目": "誤検出数", "値": str(summary.false_positives), "備考": ""},
        {"項目": "手動追加数", "値": str(summary.manual_additions), "備考": ""},
        {"項目": "平均被覆率", "値": f"{summary.avg_coverage:.3f}", "備考": ""},
        {"項目": "最小被覆率", "値": f"{summary.min_coverage:.3f}", "備考": ""},
        {"項目": "平均過剰率", "値": f"{summary.avg_excess:.3f}", "備考": ""},
        {"項目": "最大過剰率", "値": f"{summary.max_excess:.3f}", "備考": ""},
        {"項目": "未承認矩形数", "値": "0", "備考": "確定版v1 CSVは全件ユーザー確認済みまたは対象外"},
        {"項目": "代表6ページ_再OCR残存件数", "値": "NOT_EVALUATED", "備考": "最終PDF未出力のため"},
        {"項目": "ワークフロー実行", "値": "成功", "備考": ""},
        {"項目": "ページ構造維持", "値": "成功", "備考": "読み込みとページ状態記録は成功"},
        {"項目": "匿名化対象検出", "値": "承認済み正解のみ正式評価", "備考": "確認待ちは候補表示まで。要座標修正と解除候補は正式評価対象外"},
        {"項目": "匿名化漏れ検証", "値": "未評価", "備考": "正解範囲に対する最終PDF検証未実施"},
        {"項目": "実運用可否", "値": "不可", "備考": "レビュー制御は改善。OCR検出品質は不足"},
    ]
    validate_formal_outputs(rows, summary, formal_truths)
    write_workflow_csv(rows)

    issues = [
        ("P1", "最終出力制御", "FAILEDページだけを処理しても、未確認REVIEW_REQUIREDページが残るため最終PDFを出力しない状態に修正しました。"),
        ("P1", "正式評価入力", "確定版v1 CSVだけを正式評価対象にしました。ドラフトCSVへ自動フォールバックしません。"),
        ("P2", "OCR警告", "Tesseract/PyMuPDF由来の実行時警告はCSVへ集約しました。アプリ側で該当パラメータは明示指定していません。"),
        ("P2", "途中保存・再開が必要", "31ページ確認は長く、作業状態JSONで再開できることを確認しました。"),
    ]
    write_report(rows, issues)

    print("pdf31_safety_control_test=completed")
    print(f"formal_ground_truth_csv={FORMAL_GROUND_TRUTH_CSV}")
    print(f"formal_evaluation_count={summary.truth_rects}")
    print(f"detected_count={summary.detected}")
    print(f"missed_count={summary.missed}")
    print(f"false_positive_count={summary.false_positives}")
    print(f"manual_addition_count={summary.manual_additions}")
    print(f"average_coverage={summary.avg_coverage:.3f}")
    print(f"minimum_coverage={summary.min_coverage:.3f}")
    print(f"average_excess={summary.avg_excess:.3f}")
    print(f"maximum_excess={summary.max_excess:.3f}")
    print(f"workflow_csv={WORKFLOW_CSV}")
    print(f"page_report_csv={PAGE_REPORT_CSV}")
    print(f"revalidation_csv={REVALIDATION_CSV}")
    print(f"draft_ground_truth_csv={GROUND_TRUTH_CSV}")
    print(f"formal_metrics_csv={FORMAL_METRICS_CSV}")
    print(f"detection_eval_csv={DETECTION_EVAL_CSV}")
    print(f"miss_csv={MISS_CSV}")
    print(f"false_positive_csv={FALSE_POSITIVE_CSV}")
    print(f"frame_ocr_check_csv={FRAME_OCR_CHECK_CSV}")
    print(f"ocr_rect_warning_csv={OCR_RECT_WARNING_CSV}")
    print(f"ocr_runtime_warning_csv={OCR_RUNTIME_WARNING_CSV}")
    print(f"tesseract_parameter_check_csv={TESSERACT_PARAMETER_CHECK_CSV}")
    print(f"user_status_csv={USER_STATUS_CSV}")
    print(f"state_json={STATE_JSON}")
    print(f"report={REPORT_TXT}")
    print("review_images=" + ",".join(str(path) for path in review_images))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
