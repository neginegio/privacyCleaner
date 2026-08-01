"""Thin pytest wrappers around tools/test_*.py scripts that need local-only fixtures.

These tests are NOT part of the default CI run. They depend on real PII-bearing
sample files (real PDFs / Excel workbooks) that are intentionally excluded from
git via .gitignore, on a Windows Japanese font being installed, or on a local
Tesseract OCR install. They stay in tools/ as their original scripts so they
can still be run directly with `python tools/test_xxx.py`; this file just lets
pytest discover and skip them cleanly instead of contributors hitting a raw
FileNotFoundError.

Run locally once you have the private fixtures in place:

    RUN_LOCAL_DATA_TESTS=1 pytest tests/test_local_data.py -v
"""

from __future__ import annotations

import os
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
TOOLS = ROOT / "tools"
sys.path.insert(0, str(TOOLS))
sys.path.insert(0, str(ROOT / "src"))

RUN_LOCAL_DATA_TESTS = os.environ.get("RUN_LOCAL_DATA_TESTS") == "1"

pytestmark = pytest.mark.local_data

requires_local_data = pytest.mark.skipif(
    not RUN_LOCAL_DATA_TESTS,
    reason="requires the private local test corpus; set RUN_LOCAL_DATA_TESTS=1 to run",
)


def _skip_if_missing(*paths: Path) -> None:
    missing = [str(path) for path in paths if not path.exists()]
    if missing:
        pytest.skip(f"missing local fixture(s): {missing}")


@requires_local_data
def test_analysis_mode() -> None:
    _skip_if_missing(ROOT / "Excel匿名化アプリ_テストデータ.xlsx")
    import test_analysis_mode as mod

    assert mod.main() == 0


@requires_local_data
def test_pdf_baseline_v1() -> None:
    _skip_if_missing(ROOT / "PDFテストデータ２.pdf")
    import test_pdf_baseline_v1 as mod

    assert mod.main() == 0


@requires_local_data
def test_pdf_generic_detection() -> None:
    _skip_if_missing(ROOT / "PDFテストデータ２.pdf")
    import test_pdf_generic_detection as mod

    assert mod.main() == 0


@requires_local_data
def test_pdf_mixed_workflow() -> None:
    import test_pdf_mixed_workflow as mod

    try:
        mod.font_path()
    except RuntimeError as exc:
        pytest.skip(str(exc))
    assert mod.main() == 0


@requires_local_data
def test_pdf_ocr_assist() -> None:
    import test_pdf_ocr_assist as mod

    assert mod.main() == 0


@requires_local_data
def test_pdf31_ground_truth_review() -> None:
    _skip_if_missing(ROOT / "PDFテストデータ２.pdf")
    os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
    import test_pdf31_ground_truth_review as mod

    assert mod.main() == 0
