from __future__ import annotations

import argparse
import csv
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from excel_privacy_cleaner.word_processor import candidates_for_inventory, extract_word_structure  # noqa: E402


FIELDS = (
    "category",
    "location_id",
    "char_start",
    "char_end",
    "text",
    "story_type",
    "container_type",
    "detection_rule",
    "confidence",
    "source",
)


def main() -> int:
    parser = argparse.ArgumentParser(description="Freeze Word candidate detection results to a CSV for holdout evaluation.")
    parser.add_argument("--docx", required=True, type=Path)
    parser.add_argument("--out", required=True, type=Path)
    args = parser.parse_args()

    inventory = extract_word_structure(args.docx)
    candidates = candidates_for_inventory(inventory)

    args.out.parent.mkdir(parents=True, exist_ok=True)
    with args.out.open("w", encoding="utf-8-sig", newline="") as output_file:
        writer = csv.DictWriter(output_file, fieldnames=FIELDS)
        writer.writeheader()
        for candidate in candidates:
            writer.writerow(
                {
                    "category": candidate.category,
                    "location_id": candidate.location_id,
                    "char_start": candidate.char_start,
                    "char_end": candidate.char_end,
                    "text": candidate.text,
                    "story_type": candidate.story_type,
                    "container_type": candidate.container_type,
                    "detection_rule": candidate.detection_rule,
                    "confidence": round(candidate.confidence, 3),
                    "source": candidate.source,
                }
            )
    print(f"wrote {len(candidates)} candidates to {args.out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
