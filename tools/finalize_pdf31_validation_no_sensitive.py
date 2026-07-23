from __future__ import annotations

import csv
import json
import sys
from datetime import datetime
from pathlib import Path
from typing import Any

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))


ROOT = Path(__file__).resolve().parents[1]
OUTPUT_DIR = ROOT / "ocr_quality_outputs" / "PDF31検証用7ページ_初回候補レビュー"
REVIEW_STATE_JSON = OUTPUT_DIR / "PDF31検証用7ページ_正解レビュー準備_state.json"
CANDIDATES_CSV = OUTPUT_DIR / "PDF31検証用7ページ_初回候補一覧.csv"
FORMAL_GT_CSV = OUTPUT_DIR / "PDF31検証用7ページ_ユーザー確認済み正解データ_v1.csv"
PAGE_REVIEW_CSV = OUTPUT_DIR / "PDF31検証用7ページ_対象なしレビュー状態.csv"
EVALUATION_CSV = OUTPUT_DIR / "PDF31検証用7ページ_正式評価結果.csv"
REPORT_TXT = OUTPUT_DIR / "PDF31検証用7ページ_対象なし確定報告書.txt"

VALIDATION_PAGES = [1, 2, 3, 7, 8, 16, 31]
FORBIDDEN_PAGES = {22, 4, 9, 12, 14, 17, 18, 19, 20, 21, 23, 24, 25, 26, 27, 28, 29, 30}

GROUND_TRUTH_FIELDS = [
    "正解ID",
    "ページ番号",
    "出現番号",
    "同一情報ID",
    "匿名化対象文字列",
    "情報種別",
    "適用モード",
    "対象理由",
    "必須度",
    "正解座標",
    "期待する変換",
    "残すべき文字列",
    "匿名化してはいけない領域",
    "ユーザー確認状態",
    "備考",
]


def read_csv_rows(path: Path) -> list[dict[str, Any]]:
    with path.open(encoding="utf-8-sig", newline="") as handle:
        return list(csv.DictReader(handle))


def write_csv(path: Path, rows: list[dict[str, Any]], fieldnames: list[str] | None = None) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    if fieldnames is None:
        fieldnames = []
        for row in rows:
            for key in row:
                if key not in fieldnames:
                    fieldnames.append(key)
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)


def validate_state(state: dict[str, Any]) -> None:
    pages = [int(page) for page in state.get("pages", [])]
    if pages != VALIDATION_PAGES:
        raise RuntimeError(f"Unexpected validation pages: {pages}")
    if set(pages) & FORBIDDEN_PAGES:
        raise RuntimeError("Forbidden pages are included in the review state")


def load_state() -> dict[str, Any]:
    if not REVIEW_STATE_JSON.exists():
        raise FileNotFoundError(REVIEW_STATE_JSON)
    state = json.loads(REVIEW_STATE_JSON.read_text(encoding="utf-8"))
    validate_state(state)
    return state


def candidate_count() -> int:
    if not CANDIDATES_CSV.exists():
        raise FileNotFoundError(CANDIDATES_CSV)
    return len(read_csv_rows(CANDIDATES_CSV))


def write_formal_ground_truth() -> None:
    write_csv(FORMAL_GT_CSV, [], GROUND_TRUTH_FIELDS)


def update_review_state(state: dict[str, Any]) -> dict[str, Any]:
    confirmed_at = datetime.now().isoformat(timespec="seconds")
    state["review_status"] = "completed_no_sensitive_data"
    state["confirmed_at"] = confirmed_at
    state["formal_ground_truth_csv"] = str(FORMAL_GT_CSV)
    state["page_review_state"] = {
        str(page): "REVIEWED_NO_SENSITIVE_DATA" for page in VALIDATION_PAGES
    }
    state["formal_truth_count"] = 0
    state["note"] = "User confirmed that all validation pages have no anonymization targets."
    REVIEW_STATE_JSON.write_text(json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8")
    return state


def write_page_review_csv(state: dict[str, Any]) -> None:
    rows = [
        {
            "ページ番号": page,
            "レビュー状態": state["page_review_state"][str(page)],
            "確認内容": "匿名化対象なし",
        }
        for page in VALIDATION_PAGES
    ]
    write_csv(PAGE_REVIEW_CSV, rows)


def write_evaluation_csv(candidates: int) -> None:
    rows = [
        {"項目": "対象ページ", "値": ",".join(str(page) for page in VALIDATION_PAGES), "備考": ""},
        {"項目": "正解矩形数", "値": "0", "備考": "ユーザー確認により全ページ対象なし"},
        {"項目": "候補数", "値": str(candidates), "備考": "現在の第2段階ルールによる初回候補"},
        {"項目": "検出数", "値": "0", "備考": "正解矩形なし"},
        {"項目": "見逃し数", "値": "0", "備考": "正解矩形なし"},
        {"項目": "誤検出数", "値": str(candidates), "備考": "正解矩形なしのため候補があれば誤検出"},
        {"項目": "recall", "値": "N/A", "備考": "正解矩形数0のため評価対象外"},
        {"項目": "precision", "値": "N/A", "備考": "正解矩形数0のため評価対象外"},
        {"項目": "判定", "値": "対象なし確認済み", "備考": "正式評価用正解データは0件"},
    ]
    write_csv(EVALUATION_CSV, rows)


def write_report(candidates: int) -> None:
    lines = [
        "PDF31検証用7ページ 対象なし確定報告書",
        f"対象ページ: {', '.join(str(page) for page in VALIDATION_PAGES)}",
        "ユーザー確認結果: 全ページで匿名化対象なし",
        "ページ22と最終テスト用17ページは処理していません。",
        "候補生成ルールは変更していません。",
        "",
        "評価結果:",
        "正解矩形数: 0",
        f"候補数: {candidates}",
        "検出数: 0",
        "見逃し数: 0",
        f"誤検出数: {candidates}",
        "recall: N/A",
        "precision: N/A",
        "",
        "出力:",
        f"- {FORMAL_GT_CSV}",
        f"- {PAGE_REVIEW_CSV}",
        f"- {EVALUATION_CSV}",
        f"- {REVIEW_STATE_JSON}",
    ]
    REPORT_TXT.write_text("\n".join(lines), encoding="utf-8")


def main() -> int:
    state = load_state()
    candidates = candidate_count()
    write_formal_ground_truth()
    state = update_review_state(state)
    write_page_review_csv(state)
    write_evaluation_csv(candidates)
    write_report(candidates)

    print("pdf31_validation_no_sensitive_finalization=completed")
    print("pages=" + ",".join(str(page) for page in VALIDATION_PAGES))
    print("formal_truth_count=0")
    print(f"candidate_count={candidates}")
    print("detected_count=0")
    print("missed_count=0")
    print(f"false_positive_count={candidates}")
    print("recall=N/A")
    print("precision=N/A")
    print(f"formal_ground_truth_csv={FORMAL_GT_CSV}")
    print(f"page_review_csv={PAGE_REVIEW_CSV}")
    print(f"evaluation_csv={EVALUATION_CSV}")
    print(f"review_state_json={REVIEW_STATE_JSON}")
    print(f"report={REPORT_TXT}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
