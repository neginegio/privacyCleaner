import sys

from excel_privacy_cleaner.qt_app import main


if __name__ == "__main__":
    if "--smoke" in sys.argv:
        raise SystemExit(0)
    raise SystemExit(main())
