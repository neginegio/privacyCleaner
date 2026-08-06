from __future__ import annotations

import sys
import tempfile
from pathlib import Path

import fitz

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from excel_privacy_cleaner.excel_processor import ProcessingOptions  # noqa: E402
from excel_privacy_cleaner.pdf_ocr_support import CANDIDATE_REVIEW, USER_REJECTED  # noqa: E402
from excel_privacy_cleaner.pdf_processor import (  # noqa: E402
    PDF_ASSISTANCE_NOTICE,
    PdfPrivacyProcessor,
    pdf_review_state_path,
)


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

        # Simulate the user pressing "解除" on a legitimate, high-confidence
        # candidate (not just a false-positive review item). Its original
        # text will correctly remain in the output PDF, and that must not
        # be treated as a post-output validation failure -- only findings
        # the user actually chose to convert should be checked for leftover
        # text. See the "比良タイヤ工業所" report: a rejected candidate's
        # surviving text blocked CSV/report/audit artifact generation even
        # though the anonymized PDF itself was written successfully.
        rejected_phone_findings = [finding for finding in findings if finding.original == "090-1111-2222"]
        assert_true(len(rejected_phone_findings) > 0, "Phone finding should exist to reject")
        for finding in rejected_phone_findings:
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

        for original in ("Yamada Taro", "taro.yamada@example.com"):
            assert_true(original not in output_text, f"Original text should not remain: {original}")
        assert_true(
            "090-1111-2222" in output_text,
            "Rejected candidate's original text should remain untouched in the output",
        )

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


def test_pdf_review_state_roundtrip() -> None:
    # Reproduces the "毎回毎回ページ確認するのが大変です" request: re-scanning
    # the exact same source PDF should restore previously approved/rejected/
    # edited decisions instead of starting the review from scratch. This also
    # guards the _finding_state_id fix -- before it, the id included the
    # post-review detection_kind, so a saved USER_REJECTED finding could never
    # match the CANDIDATE_REVIEW/auto-approved finding produced by a fresh scan.
    with tempfile.TemporaryDirectory(prefix="pdf_review_state_test_") as tmp:
        source = Path(tmp) / "sample.pdf"
        create_text_pdf(source)

        first_processor = PdfPrivacyProcessor()
        first_findings = first_processor.scan(source, options=ProcessingOptions(mode="analysis"))

        email_findings = [finding for finding in first_findings if finding.original == "taro.yamada@example.com"]
        assert_true(len(email_findings) > 0, "Email finding should exist")
        for finding in email_findings:
            finding.enabled = False
            finding.detection_kind = USER_REJECTED

        phone_findings = [finding for finding in first_findings if finding.original == "090-1111-2222"]
        assert_true(len(phone_findings) > 0, "Phone finding should exist")
        for finding in phone_findings:
            finding.replacement = "CUSTOM_REPLACEMENT"

        state_path = pdf_review_state_path(source)
        first_processor.export_review_state(state_path, source, first_findings, last_page=1)
        assert_true(state_path.exists(), "Review state file should be written")

        second_processor = PdfPrivacyProcessor()
        second_findings = second_processor.scan(source, options=ProcessingOptions(mode="analysis"))
        second_processor.import_review_state(state_path, source, second_findings)

        restored_email = [finding for finding in second_findings if finding.original == "taro.yamada@example.com"]
        assert_true(len(restored_email) > 0, "Email finding should still exist after re-scan")
        for finding in restored_email:
            assert_true(not finding.enabled, "Rejected email decision should carry over")
            assert_true(finding.detection_kind == USER_REJECTED, "Rejected status should carry over")

        restored_phone = [finding for finding in second_findings if finding.original == "090-1111-2222"]
        assert_true(len(restored_phone) > 0, "Phone finding should still exist after re-scan")
        for finding in restored_phone:
            assert_true(finding.replacement == "CUSTOM_REPLACEMENT", "Edited replacement text should carry over")

        assert_true(second_processor.last_review_page == 1, "Last reviewed page should carry over")

        mismatched_source = Path(tmp) / "sample_copy.pdf"
        create_text_pdf(mismatched_source)
        third_processor = PdfPrivacyProcessor()
        third_findings = third_processor.scan(mismatched_source, options=ProcessingOptions(mode="analysis"))
        try:
            third_processor.import_review_state(state_path, mismatched_source, third_findings)
        except RuntimeError as exc:
            assert_true("一致しません" in str(exc), "Mismatched source PDF should be rejected clearly")
        else:
            raise AssertionError("Importing a review state saved for a different PDF should fail")

    print("pdf_review_state_roundtrip_tests=passed")


def test_pdf_cross_page_literal_propagation() -> None:
    # Reproduces the "比良タイヤ工業所" report: a name/company is confidently
    # detected once (via GiNZA, which is context-dependent and has no
    # cross-page memory), but the identical string also appears verbatim on
    # other pages that never independently trigger any detection rule.
    # Approving only the one detected occurrence then leaves the others
    # unredacted, which the post-output residual-text check correctly flags.
    with tempfile.TemporaryDirectory(prefix="pdf_propagation_test_") as tmp:
        source = Path(tmp) / "sample.pdf"
        doc = fitz.open()
        page1 = doc.new_page()
        page1.insert_text((72, 72), "Contact: Yamada Taro", fontsize=12)
        page2 = doc.new_page()
        # Needs to be >= OCR_MIN_TEXT_CHARS (20) or the page is classified as
        # an OCR page instead of a text-layer page, which the propagation
        # pass intentionally skips (OCR pages run their own pipeline).
        page2.insert_text((72, 72), "Reference notes: Yamada Taro entry", fontsize=12)
        doc.save(source)
        doc.close()

        processor = PdfPrivacyProcessor()
        findings = processor.scan(source, options=ProcessingOptions(mode="analysis"))
        matches = [finding for finding in findings if finding.original == "Yamada Taro"]
        assert_true(len(matches) >= 1, "Name should be detected at least once")

        # Simulate the exact real-world gap: page 2's occurrence was never
        # independently detected, by removing it and re-running the
        # cross-page propagation pass directly against a fresh document
        # handle, the same way scan() does internally.
        page2_matches = [finding for finding in matches if finding.sheet == "ページ2"]
        for finding in page2_matches:
            findings.remove(finding)
            processor.locations.pop((finding.sheet, finding.cell, finding.original), None)
        assert_true(len(findings) > 0, "Page 1's occurrence should remain as the propagation source")

        seen = {finding.dedupe_key for finding in findings}
        fresh_doc = fitz.open(processor.temp_pdf)
        try:
            processor._propagate_known_literals_across_pages(fresh_doc, findings, seen)
        finally:
            fresh_doc.close()

        restored = [
            finding for finding in findings if finding.original == "Yamada Taro" and finding.sheet == "ページ2"
        ]
        assert_true(len(restored) == 1, "Propagation should recreate the missing page-2 finding")
        assert_true(restored[0].detection_kind == CANDIDATE_REVIEW, "Propagated finding should require review")
        assert_true(not restored[0].enabled, "Propagated finding should not be auto-enabled")
        assert_true(
            restored[0].replacement == matches[0].replacement,
            "Propagated finding should reuse the same pseudonym as the original detection",
        )

    print("pdf_cross_page_literal_propagation_tests=passed")


def test_pdf_output_validation_scoped_to_each_findings_own_location() -> None:
    # Reproduces the "LCS" report: the same short string was approved for
    # redaction at one location but deliberately rejected (kept as-is) at
    # many other locations across the document. A document-wide or
    # page-wide "does this text still appear anywhere" check would always
    # fail in that situation -- it finds the deliberately-kept occurrences,
    # not a leak at the approved spot. Validation must be scoped to each
    # enabled finding's own rect.
    with tempfile.TemporaryDirectory(prefix="pdf_scoped_validation_test_") as tmp:
        source = Path(tmp) / "sample.pdf"
        create_text_pdf(source)

        processor = PdfPrivacyProcessor()
        findings = processor.scan(source, options=ProcessingOptions(mode="analysis"))

        phone_findings = [finding for finding in findings if finding.original == "090-1111-2222"]
        assert_true(len(phone_findings) == 2, "Phone number should be detected on both page 1 and page 2")
        page1_finding = next(finding for finding in phone_findings if finding.sheet == "ページ1")
        page2_finding = next(finding for finding in phone_findings if finding.sheet == "ページ2")
        # page1: approve (should be redacted). page2: reject (should stay).
        page1_finding.enabled = True
        page2_finding.enabled = False
        page2_finding.detection_kind = USER_REJECTED

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
        assert_true(result.warnings == (), f"Conversion should succeed with no warnings, got: {result.warnings}")
        assert_true(result.csv_path.exists(), "CSV should be written")
        assert_true(result.report_path.exists(), "Report should be written")

        output_doc = fitz.open(result.pdf_path)
        try:
            page1_text = output_doc[0].get_text("text")
            page2_text = output_doc[1].get_text("text")
        finally:
            output_doc.close()
        assert_true("090-1111-2222" not in page1_text, "Approved occurrence on page 1 should be redacted")
        assert_true("090-1111-2222" in page2_text, "Rejected occurrence on page 2 should remain untouched")

    print("pdf_output_validation_scoped_to_each_findings_own_location_tests=passed")


if __name__ == "__main__":
    raise SystemExit(main())
