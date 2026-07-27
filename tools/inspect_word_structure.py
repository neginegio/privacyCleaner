from __future__ import annotations

import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from excel_privacy_cleaner.word_processor import extract_word_structure  # noqa: E402


def main() -> int:
    parser = argparse.ArgumentParser(description="Inspect Word .docx structure without modifying the file.")
    parser.add_argument("path", type=Path)
    args = parser.parse_args()

    inventory = extract_word_structure(args.path)
    for run in inventory.runs:
        location = [
            run.story_type.upper(),
            f"container={run.container_type}",
            f"section={run.section_index}",
            f"paragraph={run.paragraph_index}",
            f"run={run.run_index}",
            f"chars={run.char_start}:{run.char_end}",
        ]
        if run.header_footer_type:
            location.append(f"kind={run.header_footer_type}")
        if run.table_index is not None:
            location.extend([f"table={run.table_index}", f"row={run.row_index}", f"cell={run.cell_index}"])
        if run.hyperlink_url:
            location.append(f"url={run.hyperlink_url}")
        print(" ".join(location) + f" text={run.text!r}")

    if inventory.hyperlinks:
        print("HYPERLINKS")
        for hyperlink in inventory.hyperlinks:
            print(f"{hyperlink.story_type.upper()} paragraph={hyperlink.paragraph_index} text={hyperlink.text!r} url={hyperlink.url}")

    if inventory.unsupported_features:
        print("UNSUPPORTED")
        for feature in inventory.unsupported_features:
            print(f"{feature.feature_type} part={feature.part_name} count={feature.count}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
