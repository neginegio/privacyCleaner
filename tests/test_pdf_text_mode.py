from __future__ import annotations

import sys
import tempfile
from pathlib import Path

import fitz

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from excel_privacy_cleaner.excel_processor import ProcessingOptions  # noqa: E402
from excel_privacy_cleaner.pdf_ocr_support import CANDIDATE_REVIEW, USER_REJECTED  # noqa: E402
from excel_privacy_cleaner.pdf_processor import PDF_ASSISTANCE_NOTICE, PdfPrivacyProcessor  # noqa: E402


def assert_true(condition: bool, message: str) -> None:
    if not condition:
        raise AssertionError(message)


def create_text_pdf(path: Path) -> None:
    doc = fitz.open()
    page1 = doc.new_page()
    page1.insert_text((72, 72), "Name: Yamada Taro", fontsize=12)
    page1.insert_text((72, 96), "Phone: 090-1111-2222", fontsize=12)
    page1.insert_text((72, 120), "Email: taro.yamada@example.com", fontsize=12)
    page2 = doc.new_page()
    page2.insert_text((72, 72), "Contact: Yamada Taro", fontsize=12)
    page2.insert_text((72, 96), "Phone again: 090-1111-2222", fontsize=12)
    doc.save(path)
    doc.close()


def main() -> int:
    with tempfile.TemporaryDirectory(prefix="pdf_privacy_test_") as tmp:
        source = Path(tmp) / "sample.pdf"
        create_text_pdf(source)

        processor = PdfPrivacyProcessor()
        findings = processor.scan(source, options=ProcessingOptions(mode="analysis"))
        originals = {finding.original for finding in findings}
        assert_true("Yamada Taro" in originals, "Name should be detected")
        assert_true("090-1111-2222" in originals, "Phone should be detected")
        assert_true("taro.yamada@example.com" in originals, "Email should be detected")

        name_replacements = {finding.replacement for finding in findings if finding.original == "Yamada Taro"}
        assert_true(len(name_replacements) == 1, "Same person should receive the same pseudonym")

        # GiNZA (run on this English-labeled fixture) misreads the "Phone"/
        # "Email" field labels as organization names -- reject them the way
        # the review dialog's "解除" action would, rather than approving
        # obvious false positives just to unblock the test.
        for finding in findings:
            if finding.detection_kind == CANDIDATE_REVIEW:
                finding.enabled = False
                finding.detection_kind = USER_REJECTED

        processor.mark_page_reviewed_with_redactions(0)
        processor.mark_page_reviewed_with_redactions(1)

        result = processor.convert_with_artifacts(
            source,
            findings,
            Path(tmp),
            options=ProcessingOptions(mode="analysis"),
            redaction_mode="pseudonym",
        )
        assert_true(result.pdf_path.exists(), "Anonymized PDF should be written")
        assert_true(result.csv_path.exists(), "PDF findings CSV should be written")
        assert_true(result.report_path.exists(), "PDF processing report should be written")
        report_text = result.report_path.read_text(encoding="utf-8")
        assert_true("全ページ確認前提のPDF匿名化支援機能" in report_text, "Report should identify PDF assist mode")
        assert_true("完全自動匿名化: いいえ" in report_text, "Report should not describe PDF as fully automatic")
        assert_true(PDF_ASSISTANCE_NOTICE in report_text, "Report should include PDF assistance notice")

        output_doc = fitz.open(result.pdf_path)
        try:
            assert_true(output_doc.page_count == 2, "Page count should be preserved")
            output_text = "\n".join(output_doc[index].get_text("text") for index in range(output_doc.page_count))
        finally:
            output_doc.close()

        for original in ("Yamada Taro", "090-1111-2222", "taro.yamada@example.com"):
            assert_true(original not in output_text, f"Original text should not remain: {original}")

        blank = Path(tmp) / "blank.pdf"
        blank_doc = fitz.open()
        blank_doc.new_page()
        blank_doc.save(blank)
        blank_doc.close()
        unsupported_processor = PdfPrivacyProcessor()
        blank_findings = unsupported_processor.scan(blank)
        assert_true(blank_findings == [], "Blank PDF should not produce findings")
        assert_true(0 in unsupported_processor.page_quality, "Blank PDF should still receive page quality")
        try:
            unsupported_processor.convert_with_artifacts(blank, blank_findings, Path(tmp))
        except RuntimeError as exc:
            assert_true("出力できません" in str(exc) or "変換対象" in str(exc), "Blank PDF should not be output as completed anonymization")
            assert_true(PDF_ASSISTANCE_NOTICE in str(exc), "Output block should explain PDF assistance limitations")
        else:
            raise AssertionError("Blank PDF should not be output")

    print("pdf_text_mode_tests=passed")
    return 0


def test_pdf_text_mode() -> None:
    assert main() == 0


if __name__ == "__main__":
    raise SystemExit(main())
