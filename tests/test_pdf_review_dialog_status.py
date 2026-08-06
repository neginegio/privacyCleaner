from __future__ import annotations

import os
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

import fitz  # noqa: E402
from PySide6.QtWidgets import QApplication  # noqa: E402

from excel_privacy_cleaner.excel_processor import ProcessingOptions  # noqa: E402
from excel_privacy_cleaner.pdf_ocr_support import CANDIDATE_MANUAL, USER_APPROVED, USER_REJECTED  # noqa: E402
from excel_privacy_cleaner.pdf_processor import PdfPrivacyProcessor  # noqa: E402
from excel_privacy_cleaner.pdf_review_dialog import PdfCandidateReviewDialog  # noqa: E402


def assert_true(condition: bool, message: str) -> None:
    if not condition:
        raise AssertionError(message)


def test_rejecting_a_high_confidence_pattern_candidate_updates_status() -> None:
    # Reproduces the "検査欄にPDFテキストって書いてあるのはどういう意味？... 却下し
    # てもPDFテキストのまま" report: the 検査 column must always reflect the
    # current review decision, regardless of what detection method originally
    # found the candidate. Presidio/pattern matches on the PDF text layer used
    # to be created with detection_kind stuck at "PDFテキスト" forever; they
    # now start as CANDIDATE_AUTO and must flip to USER_APPROVED/USER_REJECTED
    # like every other non-manual candidate once a human acts on them.
    app = QApplication.instance() or QApplication([])
    with tempfile.TemporaryDirectory(prefix="pdf_review_status_test_") as tmp:
        source = Path(tmp) / "sample.pdf"
        doc = fitz.open()
        page = doc.new_page()
        page.insert_text((72, 72), "Phone: 090-1111-2222", fontsize=12)
        doc.save(source)
        doc.close()

        processor = PdfPrivacyProcessor()
        findings = processor.scan(source, options=ProcessingOptions(mode="analysis"))
        phone_findings = [finding for finding in findings if finding.original == "090-1111-2222"]
        assert_true(len(phone_findings) == 1, "Phone number should be detected once")
        finding = phone_findings[0]
        assert_true(finding.enabled, "High-confidence pattern matches should start enabled/auto-approved")

        dialog = PdfCandidateReviewDialog(processor, findings)
        try:
            item = next(item for item, item_finding in dialog.items.items() if item_finding is finding)
            dialog.scene.clearSelection()
            item.setSelected(True)

            dialog.set_selected_enabled(False)
            assert_true(
                finding.detection_kind == USER_REJECTED,
                f"Rejecting an auto-detected candidate should set USER_REJECTED, got {finding.detection_kind!r}",
            )

            item.setSelected(True)
            dialog.set_selected_enabled(True)
            assert_true(
                finding.detection_kind == USER_APPROVED,
                f"Re-approving should set USER_APPROVED, got {finding.detection_kind!r}",
            )
        finally:
            dialog.close()


def test_manual_redaction_status_is_not_overwritten_by_toggling() -> None:
    app = QApplication.instance() or QApplication([])
    with tempfile.TemporaryDirectory(prefix="pdf_review_status_manual_test_") as tmp:
        source = Path(tmp) / "sample.pdf"
        doc = fitz.open()
        page = doc.new_page()
        page.insert_text((72, 72), "Filler text so the page counts as a text-layer page.", fontsize=12)
        doc.save(source)
        doc.close()

        processor = PdfPrivacyProcessor()
        findings = processor.scan(source, options=ProcessingOptions(mode="analysis"))

        dialog = PdfCandidateReviewDialog(processor, findings)
        try:
            dialog.add_manual_item()
            manual_finding = next(finding for finding in dialog.findings if finding.detection_kind == CANDIDATE_MANUAL)
            item = next(item for item, item_finding in dialog.items.items() if item_finding is manual_finding)
            dialog.scene.clearSelection()
            item.setSelected(True)

            dialog.set_selected_enabled(False)
            assert_true(
                manual_finding.detection_kind == CANDIDATE_MANUAL,
                f"Toggling a manual redaction should keep its own status label, got {manual_finding.detection_kind!r}",
            )
        finally:
            dialog.close()
