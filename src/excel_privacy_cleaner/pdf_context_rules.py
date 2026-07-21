from __future__ import annotations

from typing import Any

from .pdf_ocr_support import OcrCandidate


def context_candidates_for_page(page_number: int, page_rect: Any) -> list[OcrCandidate]:
    """Return structure-based PDF candidates.

    Ground-truth CSVs and reviewed rectangles must never be read here. This
    hook is intentionally data-independent; evaluation fixtures live under
    tools/ and are loaded only by scoring/review utilities.
    """
    return []
