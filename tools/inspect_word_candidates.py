from __future__ import annotations

import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from excel_privacy_cleaner.word_processor import detect_word_candidates  # noqa: E402


def main() -> int:
    parser = argparse.ArgumentParser(description="Inspect Word Phase 2 candidates without modifying the file.")
    parser.add_argument("path", type=Path)
    args = parser.parse_args()

    _inventory, candidates = detect_word_candidates(args.path)
    for candidate in candidates:
        print(
            f"{candidate.category} source={candidate.source} story={candidate.story_type} "
            f"location={candidate.location_id} chars={candidate.char_start}:{candidate.char_end} "
            f"runs={candidate.affected_run_indices} text={candidate.text!r} rule={candidate.detection_rule}"
        )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
