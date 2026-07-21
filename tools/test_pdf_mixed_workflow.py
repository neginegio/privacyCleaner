from __future__ import annotations

import csv
import html
import sys
import time
import unicodedata
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
    QUALITY_FAILED,
    QUALITY_REVIEW,
    USER_APPROVED,
    USER_REJECTED,
)
from excel_privacy_cleaner.pdf_processor import (  # noqa: E402
    PdfPrivacyProcessor,
    final_output_status,
    write_pdf_findings_csv,
    write_pdf_quality_csv,
)


OUTPUT_DIR = Path("ocr_quality_outputs")
MIXED_PDF = OUTPUT_DIR / "混在3ページPDF_テスト.pdf"
TEXT_SOURCE_PDF = OUTPUT_DIR / "混在3ページPDF_正解座標用テキスト源.pdf"
RESULT_CSV = OUTPUT_DIR / "混在3ページPDF_ワークフロー結果.csv"
REOCR_CSV = OUTPUT_DIR / "混在3ページPDF_再検証結果.csv"
COORDINATE_CSV = OUTPUT_DIR / "混在3ページPDF_座標比較.csv"
PAGE_REPORT_CSV = OUTPUT_DIR / "混在3ページPDF_ページ別処理報告.csv"


@dataclass(frozen=True)
class ExpectedItem:
    page: int
    original: str
    entity_type: str
    rect: tuple[float, float, float, float]
    required: bool = True


@dataclass(frozen=True)
class DirectPdfResult:
    pdf_path: Path
    csv_path: Path
    quality_csv_path: Path
    report_path: Path


def font_path() -> Path:
    for candidate in (
        Path("C:/Windows/Fonts/meiryo.ttc"),
        Path("C:/Windows/Fonts/msgothic.ttc"),
        Path("C:/Windows/Fonts/BIZ-UDGothicR.ttc"),
    ):
        if candidate.exists():
            return candidate
    raise RuntimeError("Japanese font was not found.")


def normalize(value: str) -> str:
    return "".join(unicodedata.normalize("NFKC", value).split()).lower()


def rect_area(rect: tuple[float, float, float, float] | None) -> float:
    if rect is None:
        return 0.0
    return max(rect[2] - rect[0], 0.0) * max(rect[3] - rect[1], 0.0)


def intersect_rect(
    a: tuple[float, float, float, float],
    b: tuple[float, float, float, float],
) -> tuple[float, float, float, float] | None:
    rect = (max(a[0], b[0]), max(a[1], b[1]), min(a[2], b[2]), min(a[3], b[3]))
    return rect if rect[2] > rect[0] and rect[3] > rect[1] else None


def coverage(expected: tuple[float, float, float, float], actual: tuple[float, float, float, float]) -> float:
    return rect_area(intersect_rect(expected, actual)) / max(rect_area(expected), 1.0)


def excess_ratio(expected: tuple[float, float, float, float], actual: tuple[float, float, float, float]) -> float:
    return max(rect_area(actual) - rect_area(intersect_rect(expected, actual)), 0.0) / max(rect_area(expected), 1.0)


def rect_text(rect: tuple[float, float, float, float]) -> str:
    return ",".join(f"{value:.3f}" for value in rect)


def insert_html_lines(page: Any, lines: list[str], *, color: str = "#000000", size: int = 17, start_y: int = 72) -> None:
    font_uri = str(font_path()).replace("\\", "/")
    css = (
        f"@font-face {{font-family: evaljp; src: url({font_uri});}} "
        f"div {{font-family: evaljp; font-size: {size}px; color: {color};}}"
    )
    y = start_y
    for line in lines:
        page.insert_htmlbox(fitz.Rect(64, y - 24, 560, y + 10), f"<div>{html.escape(line)}</div>", css=css)
        y += 42 if size >= 14 else 30


def insert_low_quality_lines(page: Any, lines: list[str]) -> None:
    font_uri = str(font_path()).replace("\\", "/")
    css = (
        f"@font-face {{font-family: evaljp; src: url({font_uri});}} "
        "div {font-family: evaljp; font-size: 7px; color: #777777;}"
    )
    y = 72
    for index, line in enumerate(lines):
        x = 64 + index * 2
        page.insert_htmlbox(fitz.Rect(x, y - 22, 360, y + 6), f"<div>{html.escape(line)}</div>", css=css)
        y += 24


def word_rect(page: Any, block_no: int, word_numbers: tuple[int, ...]) -> tuple[float, float, float, float]:
    words = [
        word
        for word in page.get_text("words")
        if int(word[5]) == block_no and int(word[7]) in word_numbers
    ]
    if not words:
        raise RuntimeError(f"Expected words not found: block={block_no}, words={word_numbers}")
    return (
        min(float(word[0]) for word in words),
        min(float(word[1]) for word in words),
        max(float(word[2]) for word in words),
        max(float(word[3]) for word in words),
    )


def create_text_source_pdf() -> tuple[fitz.Document, list[ExpectedItem]]:
    page1_lines = [
        "氏名: 山田 太郎",
        "住所: 東京都千代田区千代田1-1",
        "電話: 090-1111-2222",
        "メール: taro.yamada@example.com",
        "口座: 架空銀行 東京支店 普通 1234567",
        "APIキー: sk-test-1234567890abcdef",
    ]
    page2_lines = [
        "低品質 佐藤花子",
    ]
    doc = fitz.open()
    page1 = doc.new_page(width=595, height=842)
    insert_html_lines(page1, page1_lines)
    page2 = doc.new_page(width=595, height=842)
    insert_low_quality_lines(page2, page2_lines)
    page3 = doc.new_page(width=595, height=842)
    page3.insert_text((72, 72), "Name: Yamada Taro", fontsize=12)
    page3.insert_text((72, 96), "Phone: 090-9999-8888", fontsize=12)
    page3.insert_text((72, 120), "Email: page3.yamada@example.com", fontsize=12)

    page1 = doc[0]
    page3 = doc[2]
    expected = [
        ExpectedItem(1, "山田 太郎", "氏名", word_rect(page1, 0, (1, 2))),
        ExpectedItem(1, "東京都千代田区千代田1-1", "住所", word_rect(page1, 1, (1,))),
        ExpectedItem(1, "090-1111-2222", "電話番号", word_rect(page1, 2, (1,))),
        ExpectedItem(1, "taro.yamada@example.com", "メールアドレス", word_rect(page1, 3, (1,))),
        ExpectedItem(1, "1234567", "銀行口座番号", word_rect(page1, 4, (4,))),
        ExpectedItem(1, "sk-test-1234567890abcdef", "APIキー", word_rect(page1, 5, (1,))),
    ]

    for value, entity_type in (
        ("Yamada Taro", "氏名"),
        ("090-9999-8888", "電話番号"),
        ("page3.yamada@example.com", "メールアドレス"),
    ):
        rects = page3.search_for(value)
        if rects:
            rect = rects[0]
            expected.append(ExpectedItem(3, value, entity_type, (rect.x0, rect.y0, rect.x1, rect.y1)))

    return doc, expected


def create_mixed_pdf() -> list[ExpectedItem]:
    OUTPUT_DIR.mkdir(exist_ok=True)
    source_doc, expected = create_text_source_pdf()
    source_doc.save(TEXT_SOURCE_PDF, garbage=4, clean=True, deflate=True)
    mixed = fitz.open()
    try:
        page1 = mixed.new_page(width=595, height=842)
        page1_pix = source_doc[0].get_pixmap(dpi=240, alpha=False)
        page1.insert_image(page1.rect, pixmap=page1_pix)

        page2 = mixed.new_page(width=595, height=842)
        page2_pix = source_doc[1].get_pixmap(dpi=140, alpha=False)
        page2.insert_image(page2.rect, pixmap=page2_pix)

        mixed.insert_pdf(source_doc, from_page=2, to_page=2)
        mixed.save(MIXED_PDF, garbage=4, clean=True, deflate=True)
    finally:
        mixed.close()
        source_doc.close()
    return expected


def assert_true(value: bool, message: str) -> None:
    if not value:
        raise AssertionError(message)


def entity_matches(expected: str, actual: str) -> bool:
    if expected == actual:
        return True
    return expected == "銀行口座番号" and actual == "銀行口座"


def find_matching_finding(
    item: ExpectedItem,
    findings: list[Any],
    locations: dict[tuple[str, str, str], Any],
    used: set[int],
) -> tuple[int, Any] | None:
    matches: list[tuple[float, int, Any]] = []
    for index, finding in enumerate(findings):
        if index in used or finding.sheet != f"ページ{item.page}":
            continue
        if not entity_matches(item.entity_type, finding.entity_type):
            continue
        if item.entity_type not in {"住所", "会社名"} and normalize(finding.original) != normalize(item.original):
            continue
        location = locations.get((finding.sheet, finding.cell, finding.original))
        if location is None:
            continue
        rect = location.rect
        y_distance = abs(((rect[1] + rect[3]) / 2) - ((item.rect[1] + item.rect[3]) / 2))
        matches.append((y_distance, index, finding))
    if not matches:
        return None
    matches.sort(key=lambda value: value[0])
    return matches[0][1], matches[0][2]


def black_ratio(pdf_path: Path, page_index: int, rect: tuple[float, float, float, float], dpi: int = 72) -> float:
    doc = fitz.open(pdf_path)
    try:
        page = doc[page_index]
        scale = dpi / 72.0
        pix = page.get_pixmap(dpi=dpi, alpha=False)
        x0 = max(int(rect[0] * scale), 0)
        y0 = max(int(rect[1] * scale), 0)
        x1 = min(int(rect[2] * scale), pix.width)
        y1 = min(int(rect[3] * scale), pix.height)
        if x1 <= x0 or y1 <= y0:
            return 0.0
        black = 0
        total = 0
        for y in range(y0, y1):
            for x in range(x0, x1):
                offset = (y * pix.width + x) * pix.n
                red = pix.samples[offset]
                green = pix.samples[offset + 1]
                blue = pix.samples[offset + 2]
                if red < 35 and green < 35 and blue < 35:
                    black += 1
                total += 1
        return black / max(total, 1)
    finally:
        doc.close()


def selected_location(finding: Any, locations: dict[tuple[str, str, str], Any]) -> tuple[float, float, float, float] | None:
    location = locations.get((finding.sheet, finding.cell, finding.original))
    return location.rect if location else None


def write_coordinate_csv(
    expected: list[ExpectedItem],
    findings: list[Any],
    locations: dict[tuple[str, str, str], Any],
    output_pdf: Path,
) -> tuple[int, int, dict[int, list[float]]]:
    rows: list[dict[str, str]] = []
    used: set[int] = set()
    missing = 0
    partial = 0
    excess_by_page: dict[int, list[float]] = {1: [], 2: [], 3: []}
    for item in expected:
        matched = find_matching_finding(item, findings, locations, used)
        if matched is None:
            missing += 1
            rows.append(
                {
                    "ページ": str(item.page),
                    "原文": item.original,
                    "情報種別": item.entity_type,
                    "期待座標": rect_text(item.rect),
                    "検出座標": "",
                    "匿名化座標": "",
                    "正解領域の被覆率": "0.000",
                    "匿名化領域の過剰率": "",
                    "黒塗り率": "",
                    "判定": "FAIL",
                    "理由": "候補が未検出",
                }
            )
            continue
        index, finding = matched
        used.add(index)
        actual = selected_location(finding, locations)
        if actual is None:
            missing += 1
            continue
        item_coverage = coverage(item.rect, actual)
        item_excess = excess_ratio(item.rect, actual)
        black = black_ratio(output_pdf, item.page - 1, item.rect)
        if item.page == 3:
            ok = item_coverage >= 0.95 and item_excess <= 0.75
        else:
            ok = item_coverage >= 0.99 and item_excess <= 0.75 and black >= 0.99
        if not ok and (item_coverage < 0.99 or (item.page != 3 and black < 0.99)):
            partial += 1
        if item_excess > 0.75:
            excess_by_page[item.page].append(item_excess)
        else:
            excess_by_page[item.page].append(item_excess)
        rows.append(
            {
                "ページ": str(item.page),
                "原文": item.original,
                "情報種別": item.entity_type,
                "期待座標": rect_text(item.rect),
                "検出座標": rect_text(actual),
                "匿名化座標": rect_text(actual),
                "正解領域の被覆率": f"{item_coverage:.3f}",
                "匿名化領域の過剰率": f"{item_excess:.3f}",
                "黒塗り率": f"{black:.3f}",
                "判定": "PASS" if ok else "FAIL",
                "理由": "" if ok else "座標または画素検証不合格",
            }
        )
    with COORDINATE_CSV.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=list(rows[0].keys()))
        writer.writeheader()
        writer.writerows(rows)
    return missing, partial, excess_by_page


def write_page_report(
    path: Path,
    findings: list[Any],
    quality_snapshot: dict[int, Any],
    page_modes: dict[int, str],
    revalidation: dict[int, str],
    missing_by_page: dict[int, int],
    partial_by_page: dict[int, int],
    excessive_by_page: dict[int, int],
    excess_values: dict[int, list[float]],
) -> None:
    rows: list[dict[str, str]] = []
    for index in range(3):
        page_findings = [finding for finding in findings if finding.sheet == f"ページ{index + 1}"]
        user_approved = sum(1 for finding in page_findings if finding.detection_kind == USER_APPROVED)
        user_rejected = sum(1 for finding in page_findings if finding.detection_kind == USER_REJECTED)
        manual = sum(1 for finding in page_findings if finding.entity_type == "手動追加" or finding.detection_kind == CANDIDATE_MANUAL)
        auto = sum(1 for finding in page_findings if finding.detection_kind == CANDIDATE_AUTO)
        review_initial = sum(
            1
            for finding in page_findings
            if finding.detection_kind in {USER_APPROVED, USER_REJECTED, CANDIDATE_REVIEW}
        )
        page_excess = excess_values.get(index + 1, [])
        rows.append(
            {
                "ページ番号": str(index + 1),
                "ページ種別": page_modes.get(index, "不明"),
                "品質判定": quality_snapshot[index].verdict,
                "自動確定件数": str(auto),
                "確認候補件数": str(review_initial),
                "利用者承認件数": str(user_approved),
                "利用者解除件数": str(user_rejected),
                "手動追加件数": str(manual),
                "未検出件数": str(missing_by_page.get(index + 1, 0)),
                "部分残存件数": str(partial_by_page.get(index + 1, 0)),
                "過剰匿名化件数": str(excessive_by_page.get(index + 1, 0)),
                "平均過剰率": f"{sum(page_excess) / len(page_excess):.3f}" if page_excess else "0.000",
                "最大過剰率": f"{max(page_excess):.3f}" if page_excess else "0.000",
                "再検証結果": revalidation.get(index + 1, "未実施"),
                "警告内容": quality_snapshot[index].warning_reason,
            }
        )
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=list(rows[0].keys()))
        writer.writeheader()
        writer.writerows(rows)


def append_page_report_to_text(report_path: Path, page_report_csv: Path) -> None:
    with report_path.open("a", encoding="utf-8") as handle:
        handle.write("\nページ別検証:\n")
        with page_report_csv.open("r", encoding="utf-8-sig", newline="") as source:
            for row in csv.DictReader(source):
                handle.write(
                    f"ページ{row['ページ番号']}: 種別={row['ページ種別']}, 品質={row['品質判定']}, "
                    f"自動確定={row['自動確定件数']}, 確認候補={row['確認候補件数']}, "
                    f"承認={row['利用者承認件数']}, 解除={row['利用者解除件数']}, 手動={row['手動追加件数']}, "
                    f"未検出={row['未検出件数']}, 部分残存={row['部分残存件数']}, "
                    f"過剰匿名化={row['過剰匿名化件数']}, 平均過剰率={row['平均過剰率']}, "
                    f"最大過剰率={row['最大過剰率']}, 再検証={row['再検証結果']}, 警告={row['警告内容']}\n"
                )
        handle.write(f"ページ別品質CSV: {report_path.with_name(report_path.stem.replace('_処理報告書', '_ページ品質') + '.csv')}\n")
        handle.write(f"座標比較CSV: {COORDINATE_CSV}\n")
        handle.write(f"再検証CSV: {REOCR_CSV}\n")
        handle.write(f"ページ別処理報告CSV: {page_report_csv}\n")


def convert_directly(
    source_pdf: Path,
    findings: list[Any],
    locations: dict[tuple[str, str, str], Any],
    quality_snapshot: dict[int, Any],
) -> DirectPdfResult:
    output_pdf = OUTPUT_DIR / f"{source_pdf.stem}_匿名化PDF_{datetime.now():%Y%m%d_%H%M%S}.pdf"
    doc = fitz.open(source_pdf)
    try:
        for finding in findings:
            if not finding.enabled:
                continue
            location = locations.get((finding.sheet, finding.cell, finding.original))
            if location is None:
                continue
            page = doc[location.page_index]
            page.add_redact_annot(fitz.Rect(location.rect), fill=(0, 0, 0))
        for page in doc:
            page.apply_redactions(
                images=fitz.PDF_REDACT_IMAGE_PIXELS,
                graphics=fitz.PDF_REDACT_LINE_ART_REMOVE_IF_COVERED,
                text=fitz.PDF_REDACT_TEXT_REMOVE,
            )
        doc.set_metadata({})
        try:
            doc.del_xml_metadata()
        except Exception:
            pass
        doc.save(output_pdf, garbage=4, clean=True, deflate=True, encryption=fitz.PDF_ENCRYPT_NONE)
    finally:
        doc.close()

    csv_path = output_pdf.with_name(f"{output_pdf.stem}_検出変換結果.csv")
    quality_csv_path = output_pdf.with_name(f"{output_pdf.stem}_ページ品質.csv")
    report_path = output_pdf.with_name(f"{output_pdf.stem}_処理報告書.txt")
    write_pdf_findings_csv(csv_path, findings)
    write_pdf_quality_csv(quality_csv_path, quality_snapshot)
    report_path.write_text(
        "\n".join(
            [
                "PDF匿名化 処理報告書",
                f"入力ファイル名: {source_pdf.name}",
                f"出力ファイル名: {output_pdf.name}",
                f"処理日時: {datetime.now():%Y-%m-%d %H:%M:%S}",
                "ページ数: 3",
                "匿名化方法: 黒塗り",
                f"検出件数: {len(findings)}",
                f"変換件数: {sum(1 for finding in findings if finding.enabled)}",
                f"解除件数: {sum(1 for finding in findings if not finding.enabled)}",
                "検証結果: ページ別検証を参照",
                "警告内容:",
                "- なし",
            ]
        )
        + "\n",
        encoding="utf-8",
    )
    return DirectPdfResult(output_pdf, csv_path, quality_csv_path, report_path)


def main() -> int:
    expected = create_mixed_pdf()
    rows: list[dict[str, str]] = []

    def record(item: str, ok: bool, detail: str = "") -> None:
        rows.append({"項目": item, "結果": "PASS" if ok else "FAIL", "詳細": detail})
        assert_true(ok, f"{item}: {detail}")

    processor = PdfPrivacyProcessor()
    start = time.perf_counter()
    findings = processor.scan(MIXED_PDF)
    scan_elapsed = time.perf_counter() - start

    record("3ページ構成", processor.page_count == 3)
    record("1ページ目OCR処理", processor.page_modes.get(0) == "ocr")
    record("2ページ目OCR処理", processor.page_modes.get(1) == "ocr")
    record("3ページ目文字PDF処理", processor.page_modes.get(2) == "text")
    record("1ページ目OCR候補あり", any(finding.sheet == "ページ1" for finding in findings))
    record("2ページ目品質不足", processor.page_quality[1].verdict in {QUALITY_REVIEW, QUALITY_FAILED}, processor.page_quality[1].warning_reason)
    record("3ページ目文字PDF候補あり", any(finding.sheet == "ページ3" for finding in findings))

    can_output, reasons = final_output_status(findings, processor.page_quality, processor.confirmed_pages, processor.page_review_state)
    record("未確認状態では最終出力不可", not can_output, " / ".join(reasons))

    first_review = next(finding for finding in findings if finding.detection_kind == CANDIDATE_REVIEW and finding.sheet == "ページ1")
    original_location = processor.locations[(first_review.sheet, first_review.cell, first_review.original)]
    shifted = (
        original_location.rect[0] + 2,
        original_location.rect[1] + 2,
        original_location.rect[2] + 4,
        original_location.rect[3] + 4,
    )
    processor.locations[(first_review.sheet, first_review.cell, first_review.original)] = type(original_location)(
        original_location.page_index,
        shifted,
    )
    record("ページ移動後の枠変更保持", processor.locations[(first_review.sheet, first_review.cell, first_review.original)].rect == shifted)
    processor.locations[(first_review.sheet, first_review.cell, first_review.original)] = original_location

    quality_snapshot = dict(processor.page_quality)
    page_modes_snapshot = dict(processor.page_modes)
    location_snapshot_before_convert: dict[tuple[str, str, str], Any]

    for finding in findings:
        if finding.detection_kind == CANDIDATE_REVIEW:
            if finding.sheet == "ページ2":
                finding.enabled = False
                finding.detection_kind = USER_REJECTED
            else:
                finding.enabled = True
                finding.detection_kind = USER_APPROVED

    manual = processor.add_manual_redaction(findings, 1, (56.0, 46.0, 390.0, 96.0), entity_type="手動追加")
    manual.detection_kind = "MANUAL_CONFIRMED"
    record("2ページ目に手動匿名化範囲追加", manual.sheet == "ページ2")
    processor.mark_page_reviewed_with_redactions(0)
    processor.mark_page_reviewed_with_redactions(1)
    processor.mark_page_reviewed_with_redactions(2)

    can_output, reasons = final_output_status(findings, processor.page_quality, processor.confirmed_pages, processor.page_review_state)
    record("全ページ確認後は出力可能", can_output, " / ".join(reasons))

    location_snapshot_before_convert = dict(processor.locations)
    start = time.perf_counter()
    result = convert_directly(MIXED_PDF, findings, location_snapshot_before_convert, quality_snapshot)
    total_elapsed = time.perf_counter() - start
    record("匿名化済みPDF出力", result.pdf_path.exists(), str(result.pdf_path))

    missing, partial, excess_values = write_coordinate_csv(expected, findings, location_snapshot_before_convert, result.pdf_path)
    excessive_by_page = {
        page: sum(1 for value in values if value > 0.75)
        for page, values in excess_values.items()
    }
    missing_by_page = {page: 0 for page in (1, 2, 3)}
    partial_by_page = {page: 0 for page in (1, 2, 3)}
    for row in csv.DictReader(COORDINATE_CSV.open("r", encoding="utf-8-sig")):
        page = int(row["ページ"])
        if row["判定"] != "PASS":
            if row["理由"] == "候補が未検出":
                missing_by_page[page] += 1
            else:
                partial_by_page[page] += 1

    record("座標比較で未検出なし", missing == 0, f"missing={missing}")
    record("座標比較で部分残存なし", partial == 0, f"partial={partial}")

    source = fitz.open(MIXED_PDF)
    output = fitz.open(result.pdf_path)
    revalidation: dict[int, str] = {}
    try:
        record("ページ数維持", source.page_count == output.page_count == 3)
        for index in range(3):
            record(f"ページ{index + 1}サイズ維持", source[index].rect == output[index].rect)
            record(f"ページ{index + 1}方向維持", source[index].rotation == output[index].rotation)
        reocr_text: list[str] = []
        for index in range(output.page_count):
            text = output[index].get_text("text")
            if index == 1:
                manual_black = black_ratio(result.pdf_path, 1, location_snapshot_before_convert[(manual.sheet, manual.cell, manual.original)].rect)
                revalidation[index + 1] = "PASS" if manual_black >= 0.985 else "FAIL"
            elif page_modes_snapshot.get(index) == "text":
                revalidation[index + 1] = "PASS"
            else:
                page_items = [item for item in expected if item.page == index + 1]
                black_ok = all(black_ratio(result.pdf_path, index, item.rect) >= 0.985 for item in page_items)
                revalidation[index + 1] = "PASS" if black_ok else "FAIL"
            reocr_text.append(text)
    finally:
        source.close()
        output.close()

    sensitive_by_page = {
        1: ["山田太郎", "090-1111-2222", "taro.yamada@example.com", "1234567", "sk-test-1234567890abcdef"],
        2: ["佐藤花子"],
        3: ["Yamada Taro", "090-9999-8888", "page3.yamada@example.com"],
    }
    for page, values in sensitive_by_page.items():
        normalized_text = normalize(reocr_text[page - 1])
        for value in values:
            if normalize(value) in normalized_text:
                revalidation[page] = "FAIL"
            record(f"ページ{page}再検証で残らない: {value}", normalize(value) not in normalized_text)

    write_page_report(
        PAGE_REPORT_CSV,
        findings,
        quality_snapshot,
        page_modes_snapshot,
        revalidation,
        missing_by_page,
        partial_by_page,
        excessive_by_page,
        excess_values,
    )
    append_page_report_to_text(result.report_path, PAGE_REPORT_CSV)

    with RESULT_CSV.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=["項目", "結果", "詳細"])
        writer.writeheader()
        writer.writerows(rows)
        writer.writerow({"項目": "OCR処理時間", "結果": "INFO", "詳細": f"{scan_elapsed:.3f}秒"})
        writer.writerow({"項目": "最終出力処理時間", "結果": "INFO", "詳細": f"{total_elapsed:.3f}秒"})
        writer.writerow({"項目": "候補総数", "結果": "INFO", "詳細": str(len(findings))})

    with REOCR_CSV.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.writer(handle)
        writer.writerow(["ページ", "再検証結果", "再OCRまたは文字抽出テキスト"])
        for index, text in enumerate(reocr_text, start=1):
            writer.writerow([index, revalidation[index], text.replace("\n", "\\n")])

    print("pdf_mixed_workflow_tests=passed")
    print(f"mixed_pdf={MIXED_PDF}")
    print(f"anonymized_pdf={result.pdf_path}")
    print(f"quality_csv={result.quality_csv_path}")
    print(f"coordinate_csv={COORDINATE_CSV}")
    print(f"revalidation_csv={REOCR_CSV}")
    print(f"page_report_csv={PAGE_REPORT_CSV}")
    print(f"report={result.report_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
