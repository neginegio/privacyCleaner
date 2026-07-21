from __future__ import annotations

import csv
import sys
import unicodedata
from pathlib import Path
from typing import Any

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "tools"))

import fitz  # noqa: E402

from excel_privacy_cleaner.pdf_processor import PdfPrivacyProcessor, ocr_page_text_and_words, write_pdf_debug_image  # noqa: E402
from excel_privacy_cleaner.pdf_ocr_support import CANDIDATE_AUTO, CANDIDATE_REVIEW, QUALITY_REVIEW, USER_APPROVED  # noqa: E402
from ocr_quality_evaluation import DEBUG_IMAGE, EXPECTED_ROWS, OUTPUT_DIR, SCAN_PDF, create_scan_pdf, load_expected_rects  # noqa: E402


RESULT_CSV = OUTPUT_DIR / "OCR支援機能テスト結果.csv"
REOCR_CSV = OUTPUT_DIR / "OCR支援機能_匿名化後再OCR.csv"
COORDINATE_CSV = OUTPUT_DIR / "OCR支援機能_座標比較.csv"
MANUAL_CSV = OUTPUT_DIR / "OCR支援機能_手動操作テスト.csv"


def assert_true(condition: bool, message: str) -> None:
    if not condition:
        raise AssertionError(message)


def normalize(value: str) -> str:
    return "".join(unicodedata.normalize("NFKC", value).split()).lower()


def entity_matches(expected: str, actual: str) -> bool:
    if expected == actual:
        return True
    return expected == "銀行口座番号" and actual == "銀行口座"


def rect_area(rect: tuple[float, float, float, float]) -> float:
    return max(rect[2] - rect[0], 0.0) * max(rect[3] - rect[1], 0.0)


def intersect_rect(a: tuple[float, float, float, float], b: tuple[float, float, float, float]) -> tuple[float, float, float, float] | None:
    rect = (max(a[0], b[0]), max(a[1], b[1]), min(a[2], b[2]), min(a[3], b[3]))
    return rect if rect[2] > rect[0] and rect[3] > rect[1] else None


def coverage(expected: tuple[float, float, float, float], actual: tuple[float, float, float, float]) -> float:
    intersection = intersect_rect(expected, actual)
    return rect_area(intersection) / max(rect_area(expected), 1.0) if intersection else 0.0


def excess_ratio(expected: tuple[float, float, float, float], actual: tuple[float, float, float, float]) -> float:
    intersection = intersect_rect(expected, actual)
    intersection_area = rect_area(intersection) if intersection else 0.0
    return max(rect_area(actual) - intersection_area, 0.0) / max(rect_area(expected), 1.0)


def rect_text(rect: tuple[float, float, float, float]) -> str:
    return ",".join(f"{value:.3f}" for value in rect)


def find_matching_finding(
    expected_text: str,
    expected_type: str,
    expected_rect: tuple[float, float, float, float],
    findings: list[Any],
    locations: dict[tuple[str, str, str], Any],
    used: set[int],
) -> tuple[int, Any] | None:
    expected_norm = normalize(expected_text)
    matches: list[tuple[float, int, Any]] = []
    for index, finding in enumerate(findings):
        if index in used:
            continue
        if not entity_matches(expected_type, finding.entity_type):
            continue
        if normalize(finding.original) != expected_norm and expected_type not in {"住所", "会社名"}:
            continue
        location = locations.get((finding.sheet, finding.cell, finding.original))
        if location is None:
            continue
        rect = location.rect
        y_distance = abs(((rect[1] + rect[3]) / 2) - ((expected_rect[1] + expected_rect[3]) / 2))
        matches.append((y_distance, index, finding))
    if not matches:
        return None
    matches.sort(key=lambda item: item[0])
    return matches[0][1], matches[0][2]


def black_ratio(pdf_path: Path, rect: tuple[float, float, float, float], dpi: int = 180) -> float:
    doc = fitz.open(pdf_path)
    try:
        page = doc[0]
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
        samples = pix.samples
        for y in range(y0, y1):
            for x in range(x0, x1):
                offset = (y * pix.width + x) * pix.n
                red = samples[offset]
                green = samples[offset + 1]
                blue = samples[offset + 2]
                if red < 35 and green < 35 and blue < 35:
                    black += 1
                total += 1
        return black / max(total, 1)
    finally:
        doc.close()


def main() -> int:
    OUTPUT_DIR.mkdir(exist_ok=True)
    create_scan_pdf()
    expected_rects = load_expected_rects()

    processor = PdfPrivacyProcessor()
    findings = processor.scan(SCAN_PDF)
    quality = processor.page_quality[0]
    write_pdf_debug_image(SCAN_PDF, findings, processor.locations, DEBUG_IMAGE, page_index=0)

    rows: list[dict[str, str]] = []

    def record(name: str, result: bool, detail: str = "") -> None:
        rows.append({"項目": name, "結果": "PASS" if result else "FAIL", "詳細": detail})
        assert_true(result, f"{name}: {detail}")

    by_type = {(finding.entity_type, finding.original): finding for finding in findings}
    entity_types = {finding.entity_type for finding in findings}
    auto = {finding.entity_type for finding in findings if finding.detection_kind == CANDIDATE_AUTO}
    review = {finding.entity_type for finding in findings if finding.detection_kind == CANDIDATE_REVIEW}
    initial_review_count = sum(1 for finding in findings if finding.detection_kind == CANDIDATE_REVIEW)

    record("品質判定がREVIEW_REQUIRED", quality.verdict == QUALITY_REVIEW, quality.warning_reason)
    record("メール検出", "メールアドレス" in auto)
    record("APIキー検出", "APIキー" in auto)
    record("電話番号検出", "電話番号" in auto)
    record("口座番号検出", "銀行口座" in review)
    record("氏名REVIEW_REQUIRED", "氏名" in review)
    record("氏名カナREVIEW_REQUIRED", "氏名カナ" in review)
    record("住所REVIEW_REQUIRED", "住所" in review)
    record("会社名REVIEW_REQUIRED", "会社名" in review)
    record("森さん人名候補", ("氏名候補", "森さん") in by_type)
    record("南さん人名候補", ("氏名候補", "南さん") in by_type)
    record("森を抜けるを自動確定しない", all(finding.original != "森を抜ける" or finding.detection_kind != CANDIDATE_AUTO for finding in findings))
    record("南側の入口を自動確定しない", all(finding.original != "南側の入口" or finding.detection_kind != CANDIDATE_AUTO for finding in findings))

    try:
        processor.convert_with_artifacts(SCAN_PDF, findings, output_dir=OUTPUT_DIR, redaction_mode="black", write_artifacts=False)
        blocked = False
    except RuntimeError as exc:
        blocked = "未確認候補" in str(exc)
    record("未確認候補が残ると出力しない", blocked)

    for finding in findings:
        if finding.detection_kind == CANDIDATE_REVIEW:
            finding.enabled = True
            finding.detection_kind = USER_APPROVED
    processor.mark_page_reviewed_with_redactions(0)

    coordinate_rows: list[dict[str, str]] = []
    used: set[int] = set()
    missing = 0
    partial = 0
    excessive = 0
    protected_overlaps = 0
    protected_rects = {
        "タイトル": expected_rects["OCR品質評価用スキャンPDF"],
        "森を抜ける": expected_rects["森を抜ける"],
        "南側の入口": expected_rects["南側の入口"],
    }
    positive_expected = [row for row in EXPECTED_ROWS if row[4] == "FALSE"]
    for expected_text, expected_type, *_rest in positive_expected:
        expected_rect = expected_rects[expected_text]
        matched = find_matching_finding(expected_text, expected_type, expected_rect, findings, processor.locations, used)
        if matched is None:
            missing += 1
            coordinate_rows.append(
                {
                    "原文": expected_text,
                    "情報種別": expected_type,
                    "期待座標": rect_text(expected_rect),
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
        actual_rect = processor.locations[(finding.sheet, finding.cell, finding.original)].rect
        item_coverage = coverage(expected_rect, actual_rect)
        item_excess = excess_ratio(expected_rect, actual_rect)
        protected_hits = [
            name
            for name, rect in protected_rects.items()
            if (rect_area(intersect_rect(actual_rect, rect)) / max(rect_area(rect), 1.0) if intersect_rect(actual_rect, rect) else 0.0) > 0.05
        ]
        ok = item_coverage >= 0.999 and item_excess <= 0.75 and not protected_hits
        if item_coverage < 0.999:
            partial += 1
        if item_excess > 0.75:
            excessive += 1
        if protected_hits:
            protected_overlaps += 1
        coordinate_rows.append(
            {
                "原文": expected_text,
                "情報種別": expected_type,
                "期待座標": rect_text(expected_rect),
                "検出座標": rect_text(actual_rect),
                "匿名化座標": rect_text(actual_rect),
                "正解領域の被覆率": f"{item_coverage:.3f}",
                "匿名化領域の過剰率": f"{item_excess:.3f}",
                "黒塗り率": "",
                "判定": "PASS" if ok else "FAIL",
                "理由": "" if ok else f"protected={','.join(protected_hits)}",
            }
        )

    record("期待する11項目をすべて検出", missing == 0, f"missing={missing}")
    record("正解領域の被覆率100%", partial == 0, f"partial={partial}")
    record("過剰匿名化0件", excessive == 0, f"excessive={excessive}")
    record("タイトル/対象外文章に重ならない", protected_overlaps == 0, f"protected_overlaps={protected_overlaps}")

    result = processor.convert_with_artifacts(SCAN_PDF, findings, output_dir=OUTPUT_DIR, redaction_mode="black", write_artifacts=True)
    record("匿名化PDFを出力", result.pdf_path.exists(), str(result.pdf_path))

    for row in coordinate_rows:
        if row["判定"] != "PASS":
            continue
        ratio = black_ratio(result.pdf_path, tuple(float(value) for value in row["期待座標"].split(",")))
        row["黒塗り率"] = f"{ratio:.3f}"
        if ratio < 0.985:
            row["判定"] = "FAIL"
            row["理由"] = "正解領域内に非黒画素が残っています"
    pixel_failures = sum(1 for row in coordinate_rows if row["判定"] != "PASS")
    record("画像上にも部分文字が残らない", pixel_failures == 0, f"pixel_failures={pixel_failures}")

    doc = fitz.open(result.pdf_path)
    try:
        text, _words = ocr_page_text_and_words(doc[0])
    finally:
        doc.close()
    normalized_text = normalize(text)
    sensitive_values = [
        "山田太郎",
        "ヤマダタロウ",
        "090-1111-2222",
        "taro.yamada@example.com",
        "1234567",
        "sk-test-1234567890abcdef",
    ]
    for value in sensitive_values:
        record(f"再OCRで残らない: {value}", normalize(value) not in normalized_text, text.replace("\n", "\\n"))
    record("対象外文章が読める: 森を抜ける", normalize("森を抜ける") in normalized_text, text.replace("\n", "\\n"))
    record("対象外文章が読める: 南側の入口", normalize("南側の入口") in normalized_text, text.replace("\n", "\\n"))

    title_rect = expected_rects["OCR品質評価用スキャンPDF"]
    record("タイトルを匿名化しない", black_ratio(result.pdf_path, title_rect) < 0.50)

    manual_rows: list[dict[str, str]] = []
    manual_processor = PdfPrivacyProcessor()
    manual_findings = manual_processor.scan(SCAN_PDF)
    manual = manual_processor.add_manual_redaction(manual_findings, 0, (420.0, 700.0, 460.0, 730.0), entity_type="手動追加", replacement="[REDACTED]")
    old_location = manual_processor.locations[(manual.sheet, manual.cell, manual.original)]
    moved_rect = (old_location.rect[0] + 5, old_location.rect[1] + 4, old_location.rect[2] + 25, old_location.rect[3] + 12)
    manual_processor.locations[(manual.sheet, manual.cell, manual.original)] = type(old_location)(old_location.page_index, moved_rect)
    manual_rows.append({"項目": "手動追加", "結果": "PASS" if manual in manual_findings else "FAIL", "詳細": manual.cell})
    manual_rows.append(
        {
            "項目": "移動/拡大縮小状態保持",
            "結果": "PASS" if manual_processor.locations[(manual.sheet, manual.cell, manual.original)].rect == moved_rect else "FAIL",
            "詳細": rect_text(moved_rect),
        }
    )
    manual_processor.cleanup()

    with RESULT_CSV.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=["項目", "結果", "詳細"])
        writer.writeheader()
        writer.writerows(rows)

    with COORDINATE_CSV.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(
            handle,
            fieldnames=[
                "原文",
                "情報種別",
                "期待座標",
                "検出座標",
                "匿名化座標",
                "正解領域の被覆率",
                "匿名化領域の過剰率",
                "黒塗り率",
                "判定",
                "理由",
            ],
        )
        writer.writeheader()
        writer.writerows(coordinate_rows)

    with MANUAL_CSV.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=["項目", "結果", "詳細"])
        writer.writeheader()
        writer.writerows(manual_rows)

    with REOCR_CSV.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.writer(handle)
        writer.writerow(["項目", "値"])
        writer.writerow(["再OCR全文", text.replace("\n", "\\n")])
        writer.writerow(["匿名化PDF", str(result.pdf_path)])
        writer.writerow(["検出候補数", len(findings)])
        writer.writerow(["品質判定", quality.verdict])

    auto_count = sum(1 for finding in findings if finding.detection_kind == CANDIDATE_AUTO)
    approved_count = sum(1 for finding in findings if finding.detection_kind == USER_APPROVED)
    unresolved_count = sum(1 for finding in findings if finding.detection_kind == CANDIDATE_REVIEW)
    with result.report_path.open("a", encoding="utf-8") as handle:
        handle.write("\n画像座標検証:\n")
        handle.write(f"自動確定件数: {auto_count}\n")
        handle.write(f"確認候補件数_初期: {initial_review_count}\n")
        handle.write(f"利用者承認件数: {approved_count}\n")
        handle.write("手動追加件数: 0\n")
        handle.write(f"未検出件数: {missing}\n")
        handle.write(f"部分残存件数: {pixel_failures}\n")
        handle.write(f"過剰匿名化件数: {excessive}\n")
        handle.write(f"残りの要確認件数: {unresolved_count}\n")
        handle.write(f"座標比較CSV: {COORDINATE_CSV}\n")

    print("pdf_ocr_assist_tests=passed")
    print(f"findings={len(findings)}")
    print(f"quality={quality.verdict}")
    print(f"result_csv={RESULT_CSV}")
    print(f"coordinate_csv={COORDINATE_CSV}")
    print(f"reocr_csv={REOCR_CSV}")
    print(f"anonymized_pdf={result.pdf_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
