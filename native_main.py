import sys
from pathlib import Path

from excel_privacy_cleaner.qt_app import main


if __name__ == "__main__":
    if "--smoke" in sys.argv:
        raise SystemExit(0)
    if "--ocr-check" in sys.argv:
        output_path = Path("ocr_check_result.txt")
        if "--ocr-check-output" in sys.argv:
            index = sys.argv.index("--ocr-check-output")
            if index + 1 < len(sys.argv):
                output_path = Path(sys.argv[index + 1])
        try:
            from excel_privacy_cleaner.pdf_processor import ocr_page_text_and_words, ocr_tessdata_dir, validate_ocr_environment

            import fitz

            pdf_path = Path(sys.argv[sys.argv.index("--ocr-check") + 1])
            errors = validate_ocr_environment()
            lines: list[str] = []
            if errors:
                lines.append("tessdata=unavailable")
                lines.append("ok=False")
                lines.extend(f"error={error}" for error in errors)
                output_path.write_text("\n".join(lines) + "\n", encoding="utf-8")
                raise SystemExit(1)
            lines.append(f"tessdata={ocr_tessdata_dir()}")
            doc = fitz.open(pdf_path)
            try:
                text, words = ocr_page_text_and_words(doc[0])
            finally:
                doc.close()
            lines.extend(["ok=True", f"text_chars={len(text)}", f"word_count={len(words)}", f"first_120={text[:120]!r}"])
            output_path.write_text("\n".join(lines) + "\n", encoding="utf-8")
            raise SystemExit(0)
        except SystemExit:
            raise
        except Exception as exc:
            output_path.write_text(f"ok=False\nfatal_error={exc}\n", encoding="utf-8")
            raise SystemExit(1)
    raise SystemExit(main())
