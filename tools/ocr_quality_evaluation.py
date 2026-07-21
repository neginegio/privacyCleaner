from __future__ import annotations

import csv
import hashlib
import html
import math
import shutil
import sys
import tempfile
import time
import unicodedata
from dataclasses import dataclass
from difflib import SequenceMatcher
from pathlib import Path
from typing import Any

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

import fitz  # noqa: E402

from excel_privacy_cleaner.pdf_processor import (  # noqa: E402
    PdfPrivacyProcessor,
    bundled_tessdata_dir,
    write_pdf_debug_image,
)
from excel_privacy_cleaner.presidio_japanese import JapanesePresidioDetector, entity_label  # noqa: E402
from excel_privacy_cleaner.pdf_processor import _pdf_extra_results  # noqa: E402


OUTPUT_DIR = Path("ocr_quality_outputs")
SCAN_PDF = OUTPUT_DIR / "正解付きスキャンPDF_1ページ.pdf"
TEXT_SOURCE_PDF = OUTPUT_DIR / "正解付きテキスト源PDF_1ページ.pdf"
EXPECTED_CSV = OUTPUT_DIR / "正解付きスキャンPDF_期待値.csv"
OCR_COMPARISON_CSV = OUTPUT_DIR / "OCR条件別比較.csv"
DEBUG_IMAGE = OUTPUT_DIR / "正解付きスキャンPDF_候補枠デバッグ.png"
ANONYMIZED_PDF = OUTPUT_DIR / "正解付きスキャンPDF_匿名化済み.pdf"
REOCR_CSV = OUTPUT_DIR / "匿名化後_再OCR結果.csv"
ZERO_CLASSIFICATION_CSV = OUTPUT_DIR / "PDFテストデータ2_文字取得0ページ分類.csv"
PAGE22_IMAGE = OUTPUT_DIR / "PDFテストデータ2_page22_render.png"


EXPECTED_ROWS = [
    ("山田 太郎", "氏名", "変換", "個人001", "FALSE", 1),
    ("山田太郎", "氏名", "変換", "個人001", "FALSE", 1),
    ("ヤマダ タロウ", "氏名カナ", "変換", "個人001", "FALSE", 1),
    ("株式会社青空製作所", "会社名", "変換", "法人001", "FALSE", 1),
    ("東京都千代田区千代田1-1", "住所", "変換", "東京都千代田区", "FALSE", 1),
    ("090-1111-2222", "電話番号", "変換", "090-****-2222", "FALSE", 1),
    ("taro.yamada@example.com", "メールアドレス", "変換", "mail001@example.invalid", "FALSE", 1),
    ("1234567", "銀行口座番号", "完全マスキング", "*******", "FALSE", 1),
    ("sk-test-1234567890abcdef", "APIキー", "完全マスキング", "[SECRET]", "FALSE", 1),
    ("森さん", "氏名候補", "要確認", "個人候補", "FALSE", 1),
    ("森を抜ける", "一般文", "対象外", "変換しない", "TRUE", 1),
    ("南さん", "氏名候補", "要確認", "個人候補", "FALSE", 1),
    ("南側の入口", "一般文", "対象外", "変換しない", "TRUE", 1),
]

PAGE_LINES = [
    "氏名: 山田 太郎",
    "氏名: 山田太郎",
    "氏名カナ: ヤマダ タロウ",
    "会社名: 株式会社青空製作所",
    "住所: 東京都千代田区千代田1-1",
    "電話: 090-1111-2222",
    "メール: taro.yamada@example.com",
    "口座: 架空銀行 東京支店 普通 1234567",
    "APIキー: sk-test-1234567890abcdef",
    "森さん",
    "森を抜ける",
    "南さん",
    "南側の入口",
]


@dataclass(frozen=True)
class Condition:
    key: str
    language: str
    dpi: int
    tessdata: Path


def normalize(value: str) -> str:
    return "".join(unicodedata.normalize("NFKC", value).split()).lower()


def font_path() -> Path:
    for candidate in (
        Path("C:/Windows/Fonts/meiryo.ttc"),
        Path("C:/Windows/Fonts/msgothic.ttc"),
        Path("C:/Windows/Fonts/BIZ-UDGothicR.ttc"),
    ):
        if candidate.exists():
            return candidate
    raise RuntimeError("Japanese font was not found.")


def mirror_tessdata(path: Path) -> Path:
    path = path.resolve()
    try:
        str(path).encode("ascii")
        return path
    except UnicodeEncodeError:
        digest = hashlib.sha1(str(path).encode("utf-8")).hexdigest()[:12]
        target = Path(tempfile.gettempdir()) / "PrivacyCleanerEvalTessdata" / digest
        target.mkdir(parents=True, exist_ok=True)
        for name in ("jpn.traineddata", "eng.traineddata", "osd.traineddata"):
            source_file = path / name
            target_file = target / name
            if not target_file.exists() or target_file.stat().st_size != source_file.stat().st_size:
                shutil.copy2(source_file, target_file)
        return target


def create_scan_pdf() -> dict[str, tuple[float, float, float, float]]:
    OUTPUT_DIR.mkdir(exist_ok=True)
    line_rects: dict[str, tuple[float, float, float, float]] = {}
    doc = fitz.open()
    page = doc.new_page(width=595, height=842)
    font = font_path()
    font_css_path = str(font).replace("\\", "/")
    css = (
        f"@font-face {{font-family: evaljp; src: url({font_css_path});}} "
        "body, div {font-family: evaljp; font-size: 17px; color: #000000;}"
    )
    y = 72
    page.insert_htmlbox(
        fitz.Rect(54, 24, 540, 54),
        "<div style='font-size:18px;'>OCR品質評価用スキャンPDF</div>",
        css=css,
    )
    for line in PAGE_LINES:
        page.insert_htmlbox(fitz.Rect(64, y - 24, 560, y + 10), f"<div>{html.escape(line)}</div>", css=css)
        line_rects[line] = (58, y - 24, 560, y + 12)
        y += 42
    doc.save(TEXT_SOURCE_PDF, garbage=4, clean=True, deflate=True)
    expected_rects = expected_rects_from_text_page(page)
    pix = doc[0].get_pixmap(dpi=240, alpha=False)
    doc.close()

    scan_doc = fitz.open()
    scan_page = scan_doc.new_page(width=595, height=842)
    scan_page.insert_image(scan_page.rect, pixmap=pix)
    scan_doc.save(SCAN_PDF, garbage=4, clean=True, deflate=True)
    scan_doc.close()
    write_expected_csv(expected_rects)
    return line_rects


def expected_rects_from_text_page(page: Any) -> dict[str, tuple[float, float, float, float]]:
    words = list(page.get_text("words"))

    def union(block_no: int, word_numbers: tuple[int, ...]) -> tuple[float, float, float, float]:
        selected = [word for word in words if int(word[5]) == block_no and int(word[7]) in word_numbers]
        if not selected:
            raise RuntimeError(f"Expected text source word was not found: block={block_no}, words={word_numbers}")
        return (
            min(float(word[0]) for word in selected),
            min(float(word[1]) for word in selected),
            max(float(word[2]) for word in selected),
            max(float(word[3]) for word in selected),
        )

    return {
        "OCR品質評価用スキャンPDF": union(0, (0,)),
        "山田 太郎": union(1, (1, 2)),
        "山田太郎": union(2, (1,)),
        "ヤマダ タロウ": union(3, (1, 2)),
        "株式会社青空製作所": union(4, (1,)),
        "東京都千代田区千代田1-1": union(5, (1,)),
        "090-1111-2222": union(6, (1,)),
        "taro.yamada@example.com": union(7, (1,)),
        "1234567": union(8, (4,)),
        "sk-test-1234567890abcdef": union(9, (1,)),
        "森さん": union(10, (0,)),
        "森を抜ける": union(11, (0,)),
        "南さん": union(12, (0,)),
        "南側の入口": union(13, (0,)),
    }


def load_expected_rects() -> dict[str, tuple[float, float, float, float]]:
    if not TEXT_SOURCE_PDF.exists():
        create_scan_pdf()
    doc = fitz.open(TEXT_SOURCE_PDF)
    try:
        return expected_rects_from_text_page(doc[0])
    finally:
        doc.close()


def write_expected_csv(expected_rects: dict[str, tuple[float, float, float, float]] | None = None) -> None:
    expected_rects = expected_rects or load_expected_rects()
    with EXPECTED_CSV.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.writer(handle)
        writer.writerow(
            [
                "原文",
                "情報種別",
                "期待する判定",
                "期待する変換",
                "誤検出してはいけないか",
                "ページ番号",
                "期待X0",
                "期待Y0",
                "期待X1",
                "期待Y1",
            ]
        )
        for row in EXPECTED_ROWS:
            rect = expected_rects[row[0]]
            writer.writerow([*row, *[f"{value:.3f}" for value in rect]])


def ocr_page(pdf_path: Path, condition: Condition) -> tuple[float, str, list[tuple[Any, ...]]]:
    doc = fitz.open(pdf_path)
    try:
        page = doc[0]
        start = time.perf_counter()
        textpage = page.get_textpage_ocr(
            language=condition.language,
            dpi=condition.dpi,
            full=True,
            tessdata=str(mirror_tessdata(condition.tessdata)),
        )
        elapsed = time.perf_counter() - start
        return elapsed, textpage.extractTEXT(), list(textpage.extractWORDS())
    finally:
        doc.close()


def detector_findings(text: str) -> list[tuple[str, str]]:
    detector = JapanesePresidioDetector()
    findings: list[tuple[str, str]] = []
    for result in detector.analyze(text):
        findings.append((text[result.start : result.end], entity_label(result.entity_type)))
    for start, end, entity_type in _pdf_extra_results(text):
        findings.append((text[start:end], entity_label(entity_type)))
    return findings


def japanese_ratio(text: str) -> float:
    chars = [ch for ch in text if not ch.isspace()]
    if not chars:
        return 0.0
    japanese = sum(1 for ch in chars if "\u3040" <= ch <= "\u30ff" or "\u4e00" <= ch <= "\u9fff")
    return japanese / len(chars)


def compare_conditions() -> str:
    best_dir = OUTPUT_DIR / "tessdata_best"
    conditions = [
        Condition("A", "jpn+eng", 300, bundled_tessdata_dir()),
        Condition("B", "jpn+eng", 400, bundled_tessdata_dir()),
        Condition("C", "jpn", 400, bundled_tessdata_dir()),
        Condition("D", "jpn+eng", 400, best_dir),
    ]
    positive_rows = [row for row in EXPECTED_ROWS if row[2] != "対象外"]
    false_guard_rows = [row for row in EXPECTED_ROWS if row[4] == "TRUE"]
    rows: list[dict[str, Any]] = []
    for condition in conditions:
        elapsed, text, words = ocr_page(SCAN_PDF, condition)
        text_norm = normalize(text)
        exact_hits = 0
        partial_hits = 0
        correct_chars = 0
        misses = 0
        for original, *_rest in positive_rows:
            original_norm = normalize(original)
            if original_norm in text_norm:
                exact_hits += 1
                correct_chars += len(original_norm)
                continue
            ratio = SequenceMatcher(None, original_norm, text_norm).quick_ratio()
            if ratio >= 0.40:
                partial_hits += 1
            else:
                misses += 1
        findings = detector_findings(text)
        false_positives = 0
        for original, *_rest in false_guard_rows:
            guard = normalize(original)
            if any(guard and guard in normalize(value) for value, _label in findings):
                false_positives += 1
        coordinate_words = [word for word in words if len(word) >= 5 and word[2] > word[0] and word[3] > word[1]]
        rows.append(
            {
                "条件": condition.key,
                "language": condition.language,
                "dpi": condition.dpi,
                "tessdata": "tessdata_best" if condition.key == "D" else "現行tessdata",
                "処理時間": f"{elapsed:.3f}",
                "取得文字数": len(text),
                "単語数": len(words),
                "日本語文字割合": f"{japanese_ratio(text):.3f}",
                "正しく認識した期待文字数": correct_chars,
                "見逃し数": misses,
                "誤認識数": partial_hits,
                "個人情報候補検出数": len(findings),
                "誤検出数": false_positives,
                "文字座標の取得数": len(coordinate_words),
                "完全一致期待項目数": exact_hits,
                "OCR先頭200文字": text[:200].replace("\n", "\\n"),
            }
        )
    with OCR_COMPARISON_CSV.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=list(rows[0].keys()))
        writer.writeheader()
        writer.writerows(rows)
    rows.sort(
        key=lambda row: (
            int(row["完全一致期待項目数"]),
            int(row["正しく認識した期待文字数"]),
            -int(row["誤検出数"]),
            -float(row["処理時間"]),
        ),
        reverse=True,
    )
    return str(rows[0]["条件"])


def debug_actual_candidates() -> tuple[int, int]:
    processor = PdfPrivacyProcessor()
    try:
        findings = processor.scan(SCAN_PDF)
        write_pdf_debug_image(SCAN_PDF, findings, processor.locations, DEBUG_IMAGE, page_index=0)
        return len(findings), sum(1 for finding in findings if finding.detection_kind == "確認候補")
    finally:
        processor.cleanup()


def redact_known_sensitive_regions(line_rects: dict[str, tuple[float, float, float, float]]) -> None:
    sensitive_prefixes = (
        "氏名:",
        "氏名カナ:",
        "会社名:",
        "住所:",
        "電話:",
        "メール:",
        "口座:",
        "APIキー:",
    )
    doc = fitz.open(SCAN_PDF)
    try:
        page = doc[0]
        for line, rect in line_rects.items():
            if line.startswith(sensitive_prefixes):
                page.add_redact_annot(fitz.Rect(rect), fill=(0, 0, 0))
        page.apply_redactions(
            images=fitz.PDF_REDACT_IMAGE_PIXELS,
            graphics=fitz.PDF_REDACT_LINE_ART_REMOVE_IF_COVERED,
            text=fitz.PDF_REDACT_TEXT_REMOVE,
        )
        doc.save(ANONYMIZED_PDF, garbage=4, clean=True, deflate=True, encryption=fitz.PDF_ENCRYPT_NONE)
    finally:
        doc.close()


def reocr_anonymized(best_key: str) -> None:
    best_dir = OUTPUT_DIR / "tessdata_best"
    condition_map = {
        "A": Condition("A", "jpn+eng", 300, bundled_tessdata_dir()),
        "B": Condition("B", "jpn+eng", 400, bundled_tessdata_dir()),
        "C": Condition("C", "jpn", 400, bundled_tessdata_dir()),
        "D": Condition("D", "jpn+eng", 400, best_dir),
    }
    elapsed, text, words = ocr_page(ANONYMIZED_PDF, condition_map[best_key])
    text_norm = normalize(text)
    sensitive = [
        "山田 太郎",
        "山田太郎",
        "ヤマダ タロウ",
        "090-1111-2222",
        "taro.yamada@example.com",
        "1234567",
        "sk-test-1234567890abcdef",
    ]
    readable = ["森を抜ける", "南側の入口"]
    with REOCR_CSV.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.writer(handle)
        writer.writerow(["種別", "原文", "結果", "備考"])
        for value in sensitive:
            writer.writerow(["漏えい確認", value, "残存なし" if normalize(value) not in text_norm else "残存あり", ""])
        for value in readable:
            exact = normalize(value) in text_norm
            ratio = SequenceMatcher(None, normalize(value), text_norm).quick_ratio()
            writer.writerow(["可読性確認", value, "読める" if exact or ratio >= 0.40 else "確認不可", f"similarity={ratio:.3f}"])
        writer.writerow(["再OCR", "処理時間", f"{elapsed:.3f}", "秒"])
        writer.writerow(["再OCR", "取得文字数", str(len(text)), ""])
        writer.writerow(["再OCR", "単語数", str(len(words)), ""])
        writer.writerow(["再OCR", "OCR全文", text.replace("\n", "\\n"), ""])


def classify_zero_pages(pdf_path: Path) -> dict[str, int]:
    counters = {
        "全ページ数": 0,
        "OCR関数呼び出し完了ページ数": 0,
        "OCR例外ページ数": 0,
        "1文字以上取得できたページ数": 0,
        "文字取得0ページ数": 0,
        "品質基準を満たしたページ数": 0,
        "個人情報候補検出ページ数": 0,
        "匿名化処理ページ数": 0,
        "再OCR検証成功ページ数": 0,
    }
    doc = fitz.open(pdf_path)
    rows: list[dict[str, Any]] = []
    try:
        counters["全ページ数"] = doc.page_count
        condition = Condition("A", "jpn+eng", 300, bundled_tessdata_dir())
        for page_index in range(doc.page_count):
            page = doc[page_index]
            exception = ""
            text = ""
            words: list[tuple[Any, ...]] = []
            elapsed = 0.0
            try:
                start = time.perf_counter()
                textpage = page.get_textpage_ocr(
                    language=condition.language,
                    dpi=condition.dpi,
                    full=True,
                    tessdata=str(mirror_tessdata(condition.tessdata)),
                )
                elapsed = time.perf_counter() - start
                text = textpage.extractTEXT()
                words = list(textpage.extractWORDS())
                counters["OCR関数呼び出し完了ページ数"] += 1
            except Exception as exc:
                exception = str(exc)
                counters["OCR例外ページ数"] += 1
            text_chars = len(text.strip())
            coordinate_words = [word for word in words if len(word) >= 5 and word[2] > word[0] and word[3] > word[1]]
            if text_chars > 0:
                counters["1文字以上取得できたページ数"] += 1
            else:
                counters["文字取得0ページ数"] += 1
            if text_chars >= 20 and coordinate_words and japanese_ratio(text) >= 0.05:
                counters["品質基準を満たしたページ数"] += 1
            if detector_findings(text):
                counters["個人情報候補検出ページ数"] += 1

            classification = ""
            reason = ""
            if text_chars == 0:
                classification, reason = classify_page_image(page, exception)
            rows.append(
                {
                    "ページ番号": page_index + 1,
                    "OCR関数呼び出し完了": "TRUE" if not exception else "FALSE",
                    "OCR例外": exception,
                    "処理時間秒": f"{elapsed:.3f}" if elapsed else "",
                    "取得文字数": text_chars,
                    "単語数": len(words),
                    "座標付き単語数": len(coordinate_words),
                    "品質基準": "TRUE" if text_chars >= 20 and coordinate_words and japanese_ratio(text) >= 0.05 else "FALSE",
                    "分類": classification,
                    "理由": reason,
                    "OCR先頭80文字": text[:80].replace("\n", "\\n"),
                }
            )
        pix = doc[21].get_pixmap(dpi=150, alpha=False)
        pix.save(PAGE22_IMAGE)
    finally:
        doc.close()
    with ZERO_CLASSIFICATION_CSV.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=list(rows[0].keys()))
        writer.writeheader()
        writer.writerows(rows)
    return counters


def classify_page_image(page: Any, exception: str) -> tuple[str, str]:
    if page.rotation:
        return "ページが回転している", f"rotation={page.rotation}"
    drawings = len(page.get_drawings())
    image_count = len(page.get_images(full=True))
    if image_count == 0 and drawings == 0:
        return "ページ上に文字が存在しない", "画像と描画要素がありません。"
    pix = page.get_pixmap(dpi=72, alpha=False)
    samples = pix.samples
    if not samples:
        return "画像品質が低い", "画像サンプルを取得できません。"
    total = pix.width * pix.height
    dark = 0
    nonwhite = 0
    edge_dark = 0
    edge_total = 0
    for y in range(pix.height):
        for x in range(pix.width):
            offset = (y * pix.width + x) * pix.n
            r, g, b = samples[offset], samples[offset + 1], samples[offset + 2]
            lum = (r + g + b) / 3
            if lum < 60:
                dark += 1
            if lum < 245:
                nonwhite += 1
            if x < 8 or y < 8 or x >= pix.width - 8 or y >= pix.height - 8:
                edge_total += 1
                if lum < 80:
                    edge_dark += 1
    nonwhite_ratio = nonwhite / max(total, 1)
    dark_ratio = dark / max(total, 1)
    edge_dark_ratio = edge_dark / max(edge_total, 1)
    if nonwhite_ratio < 0.01:
        return "ページ上に文字が存在しない", f"nonwhite_ratio={nonwhite_ratio:.4f}"
    if edge_dark_ratio > 0.10:
        return "大きな余白または黒い枠がある", f"edge_dark_ratio={edge_dark_ratio:.4f}"
    if dark_ratio < 0.01:
        return "画像品質が低い", f"低コントラストまたは薄い文字: dark_ratio={dark_ratio:.4f}"
    if exception:
        return "文字は見えるが原因不明", exception[:120]
    return "画像品質が低い", f"nonwhite_ratio={nonwhite_ratio:.4f}, dark_ratio={dark_ratio:.4f}"


def main() -> int:
    OUTPUT_DIR.mkdir(exist_ok=True)
    line_rects = create_scan_pdf()
    best_key = compare_conditions()
    actual_count, review_count = debug_actual_candidates()
    redact_known_sensitive_regions(line_rects)
    reocr_anonymized(best_key)
    counters = classify_zero_pages(Path("PDFテストデータ２.pdf"))
    summary = OUTPUT_DIR / "OCR品質評価サマリー.txt"
    summary.write_text(
        "\n".join(
            [
                f"best_condition={best_key}",
                f"actual_candidate_count={actual_count}",
                f"actual_review_count={review_count}",
                *[f"{key}={value}" for key, value in counters.items()],
                f"scan_pdf={SCAN_PDF}",
                f"expected_csv={EXPECTED_CSV}",
                f"ocr_comparison_csv={OCR_COMPARISON_CSV}",
                f"debug_image={DEBUG_IMAGE}",
                f"anonymized_pdf={ANONYMIZED_PDF}",
                f"reocr_csv={REOCR_CSV}",
                f"zero_classification_csv={ZERO_CLASSIFICATION_CSV}",
                f"page22_image={PAGE22_IMAGE}",
            ]
        )
        + "\n",
        encoding="utf-8",
    )
    print(summary.read_text(encoding="utf-8"))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
