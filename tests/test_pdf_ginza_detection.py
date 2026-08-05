from __future__ import annotations

import sys
import tempfile
from pathlib import Path

import fitz

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from excel_privacy_cleaner.excel_processor import ProcessingOptions  # noqa: E402
from excel_privacy_cleaner.ginza_japanese import WORD_NLP_DETECTION_RULE  # noqa: E402
from excel_privacy_cleaner.pdf_ocr_support import CANDIDATE_REVIEW  # noqa: E402
from excel_privacy_cleaner.pdf_processor import PdfPrivacyProcessor  # noqa: E402


def assert_true(condition: bool, message: str) -> None:
    if not condition:
        raise AssertionError(message)


def create_bare_company_name_pdf(path: Path) -> None:
    # Reproduces the same real-world recall gap fixed in Word/Excel: a
    # company's short form used without its legal-entity suffix (株式会社
    # etc.) is invisible to Presidio's pattern rules, since there is no
    # attached structural cue for them to key off of.
    doc = fitz.open()
    page = doc.new_page()
    # OCR_MIN_TEXT_CHARS (20) gates text-mode vs. OCR-mode page handling, so
    # the fixture needs enough running text to land on the text-mode path.
    page.insert_text((72, 72), "本日のミーティング内容について", fontsize=12, fontname="japan")
    page.insert_text((72, 96), "場所　アルプススチール様1階応接室", fontsize=12, fontname="japan")
    doc.save(path)
    doc.close()


def test_ginza_catches_bare_company_name_on_text_mode_page() -> None:
    with tempfile.TemporaryDirectory(prefix="pdf_ginza_") as tmpdir:
        tmp = Path(tmpdir)
        source = tmp / "fixture.pdf"
        create_bare_company_name_pdf(source)

        processor = PdfPrivacyProcessor()
        findings = processor.scan(source, options=ProcessingOptions(mode="analysis"))

        assert_true(processor.page_modes.get(0) == "text", "Fixture should be treated as a text-layer page, not OCR")

        ginza_findings = [f for f in findings if f.detection_kind == CANDIDATE_REVIEW]
        assert_true(
            bool(ginza_findings),
            "GiNZA should surface a candidate for the bare company name that Presidio structurally cannot see",
        )
        assert_true(
            any(f.original == "アルプススチール" for f in ginza_findings),
            "GiNZA candidate should cover the bare company name",
        )
        for finding in ginza_findings:
            assert_true(not finding.enabled, "AI-sourced PDF findings must always default to review-required")
            assert_true(
                WORD_NLP_DETECTION_RULE in finding.reason,
                "Reason text must clearly identify this as an AI/NLP judgment",
            )
