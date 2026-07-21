from __future__ import annotations

import csv
import json
import sys
from datetime import datetime
from pathlib import Path
from typing import Any

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))


ROOT = Path(__file__).resolve().parents[1]
OUTPUT_DIR = ROOT / "ocr_quality_outputs" / "PDF31最終テスト17ページ_初回候補レビュー"
REVIEW_STATE_JSON = OUTPUT_DIR / "PDF31最終テスト17ページ_正解レビュー準備_state.json"
CANDIDATES_CSV = OUTPUT_DIR / "PDF31最終テスト17ページ_初回候補一覧.csv"
FORMAL_GT_CSV = OUTPUT_DIR / "PDF31最終テスト17ページ_ユーザー確認済み正解データ_v1.csv"
PAGE_REVIEW_CSV = OUTPUT_DIR / "PDF31最終テスト17ページ_対象なしレビュー状態.csv"
EXCLUDED_CANDIDATES_CSV = OUTPUT_DIR / "PDF31最終テスト17ページ_対象外候補一覧.csv"
EVALUATION_CSV = OUTPUT_DIR / "PDF31最終テスト17ページ_正式評価結果.csv"
REPORT_TXT = OUTPUT_DIR / "PDF31最終テスト17ページ_対象なし確定報告書.txt"
FINAL_EVALUATION_REPORT_TXT = OUTPUT_DIR / "PDF31最終テスト17ページ_初回評価報告書.txt"

FINAL_TEST_PAGES = [4, 9, 12, 14, 17, 18, 19, 20, 21, 23, 24, 25, 26, 27, 28, 29, 30]
FORBIDDEN_PAGES = {1, 2, 3, 5, 6, 7, 8, 10, 11, 13, 15, 16, 22, 31}

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
    if pages != FINAL_TEST_PAGES:
        raise RuntimeError(f"Unexpected final test pages: {pages}")
    if set(pages) & FORBIDDEN_PAGES:
        raise RuntimeError("Forbidden non-final pages are included in the review state")


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


def candidate_rows() -> list[dict[str, Any]]:
    if not CANDIDATES_CSV.exists():
        raise FileNotFoundError(CANDIDATES_CSV)
    return read_csv_rows(CANDIDATES_CSV)


def write_formal_ground_truth() -> None:
    write_csv(FORMAL_GT_CSV, [], GROUND_TRUTH_FIELDS)


def update_review_state(state: dict[str, Any]) -> dict[str, Any]:
    confirmed_at = datetime.now().isoformat(timespec="seconds")
    state["review_status"] = "completed_no_sensitive_data"
    state["confirmed_at"] = confirmed_at
    state["formal_ground_truth_csv"] = str(FORMAL_GT_CSV)
    state["page_review_state"] = {
        str(page): "REVIEWED_NO_SENSITIVE_DATA" for page in FINAL_TEST_PAGES
    }
    state["formal_truth_count"] = 0
    state["note"] = (
        "User confirmed that all final-test pages have no anonymization targets. "
        "Financial statements on pages 18 and 19 are not anonymization targets."
    )
    for candidate in state.get("candidates", []):
        candidate["user_review_status"] = "対象外"
        candidate["evaluation_result"] = "FALSE_POSITIVE"
        candidate["user_note"] = "ユーザー確認により匿名化対象外"
    REVIEW_STATE_JSON.write_text(json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8")
    return state


def write_page_review_csv(state: dict[str, Any]) -> None:
    rows = []
    for page in FINAL_TEST_PAGES:
        note = ""
        if page == 18:
            note = "貸借対照表。財務情報あり。匿名化対象ではない。"
        elif page == 19:
            note = "損益計算書。財務情報あり。匿名化対象ではない。"
        rows.append(
            {
                "ページ番号": page,
                "レビュー状態": state["page_review_state"][str(page)],
                "確認内容": "匿名化対象なし",
                "備考": note,
            }
        )
    write_csv(PAGE_REVIEW_CSV, rows)


def write_excluded_candidates_csv(candidates: list[dict[str, Any]]) -> None:
    rows = []
    for row in candidates:
        rows.append(
            {
                **row,
                "ユーザー判断": "対象外",
                "評価区分": "誤検出",
                "正式評価への扱い": "正解矩形0件のため誤検出として集計",
            }
        )
    write_csv(EXCLUDED_CANDIDATES_CSV, rows)


def write_evaluation_csv(candidates: int) -> None:
    rows = [
        {"項目": "対象ページ", "値": ",".join(str(page) for page in FINAL_TEST_PAGES), "備考": ""},
        {"項目": "正解矩形数", "値": "0", "備考": "ユーザー確認により全ページ対象なし"},
        {"項目": "候補数", "値": str(candidates), "備考": "現在の第2段階ルールによる初回候補"},
        {"項目": "検出数", "値": "0", "備考": "正解矩形なし"},
        {"項目": "見逃し数", "値": "0", "備考": "正解矩形なし"},
        {"項目": "誤検出数", "値": str(candidates), "備考": "正解矩形なしのため候補があれば誤検出"},
        {"項目": "recall", "値": "N/A", "備考": "正解矩形数0のため評価対象外"},
        {"項目": "precision", "値": "0.000" if candidates else "N/A", "備考": "候補がすべて対象外"},
        {"項目": "判定", "値": "対象なし確認済み", "備考": "正式評価用正解データは0件"},
    ]
    write_csv(EVALUATION_CSV, rows)


def write_report(candidates: int) -> None:
    lines = [
        "PDF31最終テスト17ページ 対象なし確定報告書",
        f"対象ページ: {', '.join(str(page) for page in FINAL_TEST_PAGES)}",
        "ユーザー確認結果: 全ページで匿名化対象なし",
        "ページ18: 貸借対照表の財務情報は存在するが、匿名化対象ではない。",
        "ページ19: 損益計算書の財務情報は存在するが、匿名化対象ではない。",
        "開発用6ページ、通常検証用7ページ、ページ22は処理していません。",
        "候補生成ルールは変更していません。",
        "",
        "評価結果:",
        "正解矩形数: 0",
        f"候補数: {candidates}",
        "検出数: 0",
        "見逃し数: 0",
        f"誤検出数: {candidates}",
        "recall: N/A",
        f"precision: {'0.000' if candidates else 'N/A'}",
        "",
        "出力:",
        f"- {FORMAL_GT_CSV}",
        f"- {PAGE_REVIEW_CSV}",
        f"- {EXCLUDED_CANDIDATES_CSV}",
        f"- {EVALUATION_CSV}",
        f"- {REVIEW_STATE_JSON}",
    ]
    REPORT_TXT.write_text("\n".join(lines), encoding="utf-8")
    FINAL_EVALUATION_REPORT_TXT.write_text("\n".join(lines), encoding="utf-8")


def main() -> int:
    state = load_state()
    candidate_data = candidate_rows()
    candidates = len(candidate_data)
    write_formal_ground_truth()
    state = update_review_state(state)
    write_page_review_csv(state)
    write_excluded_candidates_csv(candidate_data)
    write_evaluation_csv(candidates)
    write_report(candidates)

    print("pdf31_final_no_sensitive_finalization=completed")
    print("pages=" + ",".join(str(page) for page in FINAL_TEST_PAGES))
    print("formal_truth_count=0")
    print(f"candidate_count={candidates}")
    print("detected_count=0")
    print("missed_count=0")
    print(f"false_positive_count={candidates}")
    print("recall=N/A")
    print(f"precision={'0.000' if candidates else 'N/A'}")
    print(f"formal_ground_truth_csv={FORMAL_GT_CSV}")
    print(f"page_review_csv={PAGE_REVIEW_CSV}")
    print(f"excluded_candidates_csv={EXCLUDED_CANDIDATES_CSV}")
    print(f"evaluation_csv={EVALUATION_CSV}")
    print(f"review_state_json={REVIEW_STATE_JSON}")
    print(f"report={REPORT_TXT}")
    print(f"final_evaluation_report={FINAL_EVALUATION_REPORT_TXT}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
