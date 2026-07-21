from __future__ import annotations

import csv
import json
import os
import shutil
import sys
import time
from contextlib import contextmanager
from pathlib import Path
from typing import Any

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

import fitz  # noqa: E402

from excel_privacy_cleaner.pdf_processor import PdfPrivacyProcessor, final_output_status  # noqa: E402


ROOT = Path(__file__).resolve().parents[1]
SOURCE = ROOT / "PDFテストデータ２.pdf"
OUTPUT_DIR = ROOT / "ocr_quality_outputs"
DATASET_SPLIT_JSON = ROOT / "config" / "evaluation" / "pdf31_dataset_split_v1.json"
FORMAL_GROUND_TRUTH_CSV = OUTPUT_DIR / "PDF31代表6ページ_ユーザー確認済み正解データ_v1.csv"
REORDERED_SOURCE_PAGES = (15, 5, 13, 11, 6, 10)


def load_dataset_split() -> dict[str, Any]:
    return json.loads(DATASET_SPLIT_JSON.read_text(encoding="utf-8"))


def pages_from_split(key: str) -> tuple[int, ...]:
    split = load_dataset_split()
    return tuple(int(page) for page in split[key])


@contextmanager
def capture_native_stderr(path: Path):
    path.parent.mkdir(parents=True, exist_ok=True)
    saved_fd = os.dup(2)
    log_fd = os.open(path, os.O_WRONLY | os.O_CREAT | os.O_TRUNC, 0o600)
    try:
        os.dup2(log_fd, 2)
        yield
    finally:
        os.dup2(saved_fd, 2)
        os.close(saved_fd)
        os.close(log_fd)


def write_csv(path: Path, rows: list[dict[str, Any]], fieldnames: list[str] | None = None) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    if fieldnames is None:
        keys: list[str] = []
        for row in rows:
            for key in row:
                if key not in keys:
                    keys.append(key)
        fieldnames = keys
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)


def audit_detection_independence() -> dict[str, str]:
    src_dir = ROOT / "src" / "excel_privacy_cleaner"
    context_path = src_dir / "pdf_context_rules.py"
    processor_path = src_dir / "pdf_processor.py"
    context_text = context_path.read_text(encoding="utf-8")
    detection_text = "\n".join(path.read_text(encoding="utf-8", errors="ignore") for path in (context_path, processor_path))
    formal_name = FORMAL_GROUND_TRUTH_CSV.name
    checks = [
        {
            "検査項目": "候補生成コードが正解CSVを参照していない",
            "判定": "PASS" if formal_name not in detection_text and "DictReader" not in context_text and ".open(" not in context_text else "FAIL",
            "根拠": str(context_path.relative_to(ROOT)),
        },
        {
            "検査項目": "候補生成コードがtruth_idを参照していない",
            "判定": "PASS" if "truth_id" not in detection_text else "FAIL",
            "根拠": "src/excel_privacy_cleaner/pdf_context_rules.py + pdf_processor.py",
        },
        {
            "検査項目": "候補生成コードが正解座標を参照していない",
            "判定": "PASS" if "正解座標" not in detection_text and "ground_truth" not in detection_text else "FAIL",
            "根拠": "src/excel_privacy_cleaner/pdf_context_rules.py + pdf_processor.py",
        },
        {
            "検査項目": "候補生成コードが固定ページ番号を使用していない",
            "判定": "PASS" if not any(token in detection_text for token in ("P10", "P11", "P13", "P15", "page == 10", "page == 11", "page == 13", "page == 15", "page_number == 10", "page_number == 11", "page_number == 13", "page_number == 15")) else "FAIL",
            "根拠": "検出処理内の固定ページID/ページ番号条件",
        },
        {
            "検査項目": "候補生成コードが固定座標を使用していない",
            "判定": "PASS" if "ContextCandidateSpec(" not in context_text else "FAIL",
            "根拠": str(context_path.relative_to(ROOT)),
        },
        {
            "検査項目": "候補生成コードが評価用データセット設定を参照していない",
            "判定": "PASS" if DATASET_SPLIT_JSON.name not in detection_text and "development_pages" not in detection_text and "validation_ocr_eligible_pages" not in detection_text else "FAIL",
            "根拠": "src/excel_privacy_cleaner/pdf_context_rules.py + pdf_processor.py",
        },
    ]
    write_csv(OUTPUT_DIR / "PDF31評価データ混入検査.csv", checks, ["検査項目", "判定", "根拠"])
    return {row["検査項目"]: row["判定"] for row in checks}


def scan_pdf(pdf_path: Path, label: str) -> tuple[list[Any], float, PdfPrivacyProcessor]:
    processor = PdfPrivacyProcessor()
    start = time.perf_counter()
    with capture_native_stderr(OUTPUT_DIR / f"{label}_stderr.log"):
        findings = processor.scan(pdf_path)
    elapsed = time.perf_counter() - start
    return findings, elapsed, processor


def page_number_from_finding(finding: Any) -> int | None:
    if not str(finding.sheet).startswith("ページ"):
        return None
    try:
        return int(str(finding.sheet).replace("ページ", ""))
    except ValueError:
        return None


def count_pages(findings: list[Any], pages: tuple[int, ...]) -> int:
    return sum(1 for finding in findings if page_number_from_finding(finding) in pages)


def findings_rows(findings: list[Any], pages: tuple[int, ...], page_map: dict[int, int] | None = None) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    for finding in findings:
        page = page_number_from_finding(finding)
        if page is None or page not in pages:
            continue
        source_page = page_map.get(page, page) if page_map else page
        rows.append(
            {
                "ページ番号": page,
                "元ページ番号": source_page,
                "情報種別": finding.entity_type,
                "検出種別": finding.detection_kind,
                "候補文字列": finding.original,
                "変換候補": finding.replacement,
                "理由": finding.reason,
            }
        )
    return rows


def scan_without_ground_truth() -> dict[str, Any]:
    backup = FORMAL_GROUND_TRUTH_CSV.with_suffix(".csv.generalization_backup")
    moved = False
    if FORMAL_GROUND_TRUTH_CSV.exists():
        if backup.exists():
            backup.unlink()
        FORMAL_GROUND_TRUTH_CSV.replace(backup)
        moved = True
    try:
        findings, elapsed, _processor = scan_pdf(SOURCE, "正解CSVなし候補生成")
        representative_pages = pages_from_split("development_pages")
        count = count_pages(findings, representative_pages)
        rows = findings_rows(findings, representative_pages)
        write_csv(OUTPUT_DIR / "PDF31代表6ページ_正解CSVなし候補一覧.csv", rows)
        result = {
            "検査": "正解CSVなし候補生成",
            "formal_csv_removed": "YES" if moved else "NOT_FOUND",
            "candidate_generation_completed": "YES",
            "generated_candidate_count": count,
            "expected_candidate_count": 49,
            "result": "PASS" if count == 49 else "FAIL",
            "elapsed_seconds": f"{elapsed:.3f}",
            "note": "正解CSVと固定座標を使わない候補生成結果",
        }
    finally:
        if moved:
            backup.replace(FORMAL_GROUND_TRUTH_CSV)
    write_csv(OUTPUT_DIR / "PDF31代表6ページ_正解CSVなし候補生成検査.csv", [result])
    return result


def create_reordered_pdf(path: Path) -> dict[int, int]:
    path.parent.mkdir(parents=True, exist_ok=True)
    page_map: dict[int, int] = {}
    with fitz.open(SOURCE) as source_doc, fitz.open() as output_doc:
        for new_index, source_page in enumerate(REORDERED_SOURCE_PAGES, start=1):
            output_doc.insert_pdf(source_doc, from_page=source_page - 1, to_page=source_page - 1)
            page_map[new_index] = source_page
        output_doc.save(path)
    return page_map


def create_scaled_pdf(path: Path, scale: float = 0.95) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    representative_pages = pages_from_split("development_pages")
    with fitz.open(SOURCE) as source_doc, fitz.open() as output_doc:
        for source_page in representative_pages:
            page = source_doc[source_page - 1]
            new_page = output_doc.new_page(width=page.rect.width, height=page.rect.height)
            pix = page.get_pixmap(matrix=fitz.Matrix(2, 2), alpha=False)
            target_width = page.rect.width * scale
            target_height = page.rect.height * scale
            x0 = (page.rect.width - target_width) / 2
            y0 = (page.rect.height - target_height) / 2
            new_page.insert_image(fitz.Rect(x0, y0, x0 + target_width, y0 + target_height), pixmap=pix)
        output_doc.save(path)


def scan_reordered_pdf() -> dict[str, Any]:
    pdf_path = OUTPUT_DIR / "PDF31代表6ページ_順序変更テスト.pdf"
    page_map = create_reordered_pdf(pdf_path)
    findings, elapsed, _processor = scan_pdf(pdf_path, "順序変更候補生成")
    count = count_pages(findings, tuple(page_map.keys()))
    rows = findings_rows(findings, tuple(page_map.keys()), page_map=page_map)
    write_csv(OUTPUT_DIR / "PDF31代表6ページ_順序変更候補一覧.csv", rows)
    result = {
        "検査": "ページ順序変更",
        "test_pdf": str(pdf_path.relative_to(ROOT)),
        "page_order": ",".join(str(page) for page in REORDERED_SOURCE_PAGES),
        "generated_candidate_count": count,
        "expected_candidate_count": 49,
        "result": "PASS" if count == 49 else "FAIL",
        "elapsed_seconds": f"{elapsed:.3f}",
        "note": "固定ページ番号に依存しない候補生成の確認",
    }
    write_csv(OUTPUT_DIR / "PDF31代表6ページ_順序変更候補生成検査.csv", [result])
    return result


def scan_scaled_pdf() -> dict[str, Any]:
    pdf_path = OUTPUT_DIR / "PDF31代表6ページ_縮小95テスト.pdf"
    create_scaled_pdf(pdf_path)
    findings, elapsed, _processor = scan_pdf(pdf_path, "縮小95候補生成")
    representative_page_indexes = tuple(range(1, len(pages_from_split("development_pages")) + 1))
    count = count_pages(findings, representative_page_indexes)
    rows = findings_rows(findings, representative_page_indexes)
    write_csv(OUTPUT_DIR / "PDF31代表6ページ_レイアウト変更候補一覧.csv", rows)
    result = {
        "検査": "軽微なレイアウト変更_95%縮小",
        "test_pdf": str(pdf_path.relative_to(ROOT)),
        "generated_candidate_count": count,
        "expected_candidate_count": 49,
        "result": "PASS" if count == 49 else "FAIL",
        "elapsed_seconds": f"{elapsed:.3f}",
        "note": "正解座標を使わない候補生成の確認",
    }
    write_csv(OUTPUT_DIR / "PDF31代表6ページ_レイアウト変更候補生成検査.csv", [result])
    return result


def validation_initial_candidates() -> dict[str, Any]:
    findings, elapsed, _processor = scan_pdf(SOURCE, "ホールドアウト初回候補")
    validation_pages = pages_from_split("validation_ocr_eligible_pages")
    failed_pages = pages_from_split("validation_ocr_failed_pages")
    rows = findings_rows(findings, validation_pages)
    failed_rows = findings_rows(findings, failed_pages)
    selection_rows = [
        {"ページ番号": 1, "選定理由": "表紙/概要系の可能性があり代表6ページと構造が異なる"},
        {"ページ番号": 2, "選定理由": "目次または導入系として候補密度が低い可能性を確認"},
        {"ページ番号": 3, "選定理由": "本文中心ページとして文脈検出の汎化を見る"},
        {"ページ番号": 7, "選定理由": "役員/株主ページ以外の表構造を確認"},
        {"ページ番号": 8, "選定理由": "代表ページ外の表または説明文を確認"},
        {"ページ番号": 16, "選定理由": "販売先ページ近傍だが開発対象外の未知ページ"},
        {"ページ番号": 31, "選定理由": "末尾ページの構造差とOCR安定性を確認"},
    ]
    failed_selection_rows = [
        {"ページ番号": 22, "選定理由": "OCR品質FAILED検証用。通常検証用recallから分離"},
    ]
    write_csv(OUTPUT_DIR / "PDF31通常検証7ページ_ページ選定.csv", selection_rows)
    write_csv(OUTPUT_DIR / "PDF31通常検証7ページ_初回候補.csv", rows)
    write_csv(OUTPUT_DIR / "PDF31ページ22_FAILED検証_ページ選定.csv", failed_selection_rows)
    write_csv(OUTPUT_DIR / "PDF31ページ22_FAILED検証_初回候補.csv", failed_rows)
    result = {
        "検査": "通常検証7ページ初回候補",
        "selected_pages": ",".join(str(page) for page in validation_pages),
        "candidate_count": len(rows),
        "failed_page22_candidate_count": len(failed_rows),
        "elapsed_seconds": f"{elapsed:.3f}",
        "note": "この候補出力後、ユーザー確認済み正解データが確定するまで検出ルールは変更しない",
    }
    page22_quality_rows: list[dict[str, Any]] = []
    for page in failed_pages:
        quality = _processor.page_quality.get(page - 1)
        can_output, reasons = final_output_status(
            findings,
            _processor.page_quality,
            _processor.confirmed_pages,
            _processor.page_review_state,
        )
        page22_quality_rows.append(
            {
                "ページ番号": page,
                "OCR品質判定": quality.verdict if quality else "NOT_EVALUATED",
                "FAILEDページ検出": "YES" if quality and quality.verdict == "FAILED" else "NO",
                "最終PDF自動出力停止": "YES" if not can_output else "NO",
                "ユーザー手動確認要求": "YES" if any("FAILED" in reason or "UNREVIEWED" in reason for reason in reasons) else "NO",
                "完了扱い抑止": "YES" if not can_output else "NO",
                "候補数": len(failed_rows),
                "警告理由": quality.warning_reason if quality else "",
            }
        )
    write_csv(OUTPUT_DIR / "PDF31ページ22_OCR品質FAILED検証.csv", page22_quality_rows)
    return result


def main() -> None:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    audit = audit_detection_independence()
    no_truth = scan_without_ground_truth()
    reordered = scan_reordered_pdf()
    scaled = scan_scaled_pdf()
    validation = validation_initial_candidates()
    summary_rows = [
        {"項目": "混入検査_PASS数", "値": sum(1 for value in audit.values() if value == "PASS"), "備考": f"{len(audit)}項目中"},
        {"項目": "正解CSVなし候補数", "値": no_truth["generated_candidate_count"], "備考": no_truth["result"]},
        {"項目": "順序変更候補数", "値": reordered["generated_candidate_count"], "備考": reordered["result"]},
        {"項目": "95%縮小候補数", "値": scaled["generated_candidate_count"], "備考": scaled["result"]},
        {"項目": "通常検証7ページ初回候補数", "値": validation["candidate_count"], "備考": validation["selected_pages"]},
        {"項目": "ページ22初回候補数", "値": validation["failed_page22_candidate_count"], "備考": "OCR品質FAILED検証用。通常recallから分離"},
    ]
    write_csv(OUTPUT_DIR / "PDF31汎化確認サマリー.csv", summary_rows, ["項目", "値", "備考"])
    print("independence_audit_pass", f"{summary_rows[0]['値']}/{len(audit)}")
    print("candidate_generation_without_ground_truth", no_truth["result"], no_truth["generated_candidate_count"])
    print("reordered_page_test", reordered["result"], reordered["generated_candidate_count"])
    print("layout_95_scale_test", scaled["result"], scaled["generated_candidate_count"])
    print("validation_ocr_eligible_initial_candidate_count", validation["candidate_count"])
    print("validation_ocr_failed_page22_initial_candidate_count", validation["failed_page22_candidate_count"])


if __name__ == "__main__":
    main()
