from __future__ import annotations

import argparse
import os
import sys
import tempfile
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

import fitz  # noqa: E402

from excel_privacy_cleaner.pdf_processor import (  # noqa: E402
    PdfPrivacyProcessor,
    bundled_tessdata_dir,
    ocr_page_text_and_words,
    ocr_tessdata_dir,
    validate_ocr_environment,
    write_pdf_debug_image,
)


def make_subset_pdf(source: Path, output: Path, page_count: int) -> Path:
    src_doc = fitz.open(source)
    out_doc = fitz.open()
    try:
        out_doc.insert_pdf(src_doc, from_page=0, to_page=min(page_count, src_doc.page_count) - 1)
        out_doc.save(output, garbage=4, clean=True, deflate=True)
    finally:
        out_doc.close()
        src_doc.close()
    return output


def japanese_ratio(text: str) -> float:
    chars = [ch for ch in text if not ch.isspace()]
    if not chars:
        return 0.0
    japanese = sum(1 for ch in chars if "\u3040" <= ch <= "\u30ff" or "\u4e00" <= ch <= "\u9fff")
    return japanese / len(chars)


def inspect_first_page(pdf_path: Path, output_dir: Path) -> None:
    errors = validate_ocr_environment()
    print(f"pymupdf_version={fitz.version[0]}")
    print(f"mupdf_version={fitz.version[1]}")
    print(f"bundled_tessdata={bundled_tessdata_dir()}")
    print(f"runtime_tessdata={ocr_tessdata_dir()}")
    print(f"loaded_language_files={','.join(path.name for path in sorted(ocr_tessdata_dir().glob('*.traineddata')))}")
    if errors:
        print(f"ocr_environment_errors={' / '.join(errors)}")
        return

    doc = fitz.open(pdf_path)
    try:
        start = time.perf_counter()
        text, words = ocr_page_text_and_words(doc[0])
        elapsed = time.perf_counter() - start
    finally:
        doc.close()

    coordinate_words = [word for word in words if len(word) >= 5 and word[2] > word[0] and word[3] > word[1]]
    print("ocr_success=True")
    print(f"ocr_elapsed_sec={elapsed:.2f}")
    print(f"ocr_text_chars={len(text)}")
    print("ocr_first_500=" + repr(text[:500]))
    print(f"japanese_char_ratio={japanese_ratio(text):.3f}")
    print(f"ocr_word_count={len(words)}")
    print(f"coordinate_word_count={len(coordinate_words)}")
    print("ocr_confidence_available=False")

    processor = PdfPrivacyProcessor()
    try:
        findings = processor.scan(pdf_path)
        debug_image = output_dir / f"{pdf_path.stem}_page1_candidates.png"
        write_pdf_debug_image(pdf_path, findings, processor.locations, debug_image, page_index=0)
        print(f"candidate_count={len(findings)}")
        print(f"debug_image={debug_image}")
        page1_findings = [finding for finding in findings if finding.sheet == "ページ1"]
        if page1_findings:
            for page_index in range(processor.page_count):
                page_has_redactions = any(finding.enabled and finding.sheet == f"ページ{page_index + 1}" for finding in findings)
                if page_has_redactions:
                    processor.mark_page_reviewed_with_redactions(page_index)
                else:
                    processor.mark_page_no_sensitive_data(page_index)
            result = processor.convert_with_artifacts(pdf_path, findings, output_dir=output_dir, redaction_mode="black")
            print(f"anonymized_pdf={result.pdf_path}")
            print(f"converted_count={result.converted_count}")
        else:
            processor.cleanup()
            print("anonymized_pdf=skipped_no_page1_candidates")
    except Exception as exc:
        processor.cleanup()
        print(f"scan_or_anonymize_error={exc}")


def scan_pages(source_pdf: Path, page_count: int, output_dir: Path) -> None:
    with tempfile.TemporaryDirectory(prefix="pdf_ocr_subset_") as tmp:
        subset = Path(tmp) / f"{source_pdf.stem}_{page_count}p.pdf"
        make_subset_pdf(source_pdf, subset, page_count)
        processor = PdfPrivacyProcessor()
        start = time.perf_counter()
        findings = []
        error = ""
        try:
            findings = processor.scan(subset)
        except Exception as exc:
            error = str(exc)
        finally:
            elapsed = time.perf_counter() - start
            print(f"scan_pages={page_count}")
            print(f"page_count={processor.page_count}")
            print(f"text_pdf_pages={sum(1 for mode in processor.page_modes.values() if mode == 'text')}")
            print(f"ocr_target_pages={sum(1 for mode in processor.page_modes.values() if mode == 'ocr')}")
            print(f"ocr_success_pages={len(processor.ocr_success_pages)}")
            print(f"ocr_failed_pages={','.join(str(index + 1) for index in sorted(processor.ocr_failed_pages)) or 'none'}")
            print("low_confidence_pages=unavailable")
            print(f"elapsed_sec={elapsed:.2f}")
            print(f"avg_sec_per_page={elapsed / max(processor.page_count, 1):.2f}")
            print(f"candidate_count={len(findings)}")
            print(f"unreviewed_count={sum(1 for finding in findings if finding.detection_kind == '確認候補')}")
            if error:
                print(f"scan_error={error}")
            processor.cleanup()


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("pdf", type=Path)
    parser.add_argument("--pages", choices=["1", "3", "31"], default="1")
    args = parser.parse_args()

    os.environ.setdefault("PYTHONIOENCODING", "utf-8")
    output_dir = args.pdf.parent / "pdf_ocr_test_outputs"
    output_dir.mkdir(exist_ok=True)

    if args.pages == "1":
        one_page = output_dir / f"{args.pdf.stem}_page1.pdf"
        make_subset_pdf(args.pdf, one_page, 1)
        inspect_first_page(one_page, output_dir)
    else:
        scan_pages(args.pdf, int(args.pages), output_dir)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
