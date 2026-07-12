import sys

from excel_privacy_cleaner.web_app import render_page, run


if __name__ == "__main__":
    if "--smoke" in sys.argv:
        render_page(status="smoke")
        raise SystemExit(0)
    run()
