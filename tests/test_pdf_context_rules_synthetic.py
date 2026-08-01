from __future__ import annotations

import sys
from dataclasses import dataclass
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from excel_privacy_cleaner.pdf_context_rules import context_candidates_for_page  # noqa: E402


@dataclass(frozen=True)
class Rect:
    x0: float = 0.0
    y0: float = 0.0
    x1: float = 600.0
    y1: float = 800.0


def word(x0: float, y0: float, x1: float, y1: float, text: str, block: int, line: int, index: int) -> tuple[float, float, float, float, str, int, int, int]:
    return (x0, y0, x1, y1, text, block, line, index)


def by_type(candidates, entity_type: str):
    return [candidate for candidate in candidates if candidate.entity_type == entity_type]


def assert_candidate(candidates, entity_type: str, contains: str) -> None:
    matches = [candidate for candidate in by_type(candidates, entity_type) if contains in candidate.normalized]
    assert matches, f"{entity_type} candidate containing {contains!r} was not generated: {candidates}"


def test_split_table_company() -> None:
    words = [
        word(40, 100, 70, 118, "No", 0, 0, 0),
        word(110, 100, 170, 118, "販売先", 0, 0, 1),
        word(350, 100, 410, 118, "売上高", 0, 0, 2),
        word(112, 135, 210, 150, "株式会社花園", 1, 1, 0),
        word(112, 153, 165, 168, "テック", 1, 2, 0),
        word(360, 135, 430, 150, "1,200", 1, 1, 1),
    ]
    candidates = context_candidates_for_page(1, Rect(), words)
    assert_candidate(candidates, "会社名", "株式会社花園テック")
    company = by_type(candidates, "会社名")[0]
    assert "結合情報" in company.reason
    assert company.rect[1] < 135 and company.rect[3] > 168
    assert company.rect[2] < 350


def test_split_address_column() -> None:
    words = [
        word(45, 100, 75, 118, "No", 0, 0, 0),
        word(120, 100, 175, 118, "所在地", 0, 0, 1),
        word(420, 100, 480, 118, "面積", 0, 0, 2),
        word(122, 135, 230, 150, "東京都中央区", 1, 1, 0),
        word(122, 153, 222, 168, "銀座1-2-3", 1, 2, 0),
        word(430, 135, 475, 150, "120.5", 1, 1, 1),
    ]
    candidates = context_candidates_for_page(1, Rect(), words)
    assert_candidate(candidates, "住所", "東京都中央区銀座1-2-3")
    address = by_type(candidates, "住所")[0]
    assert address.rect[2] < 420


def test_person_column_values() -> None:
    words = [
        word(40, 100, 70, 118, "No", 0, 0, 0),
        word(120, 100, 165, 118, "氏名", 0, 0, 1),
        word(280, 100, 330, 118, "比率", 0, 0, 2),
        word(122, 135, 160, 150, "佐藤", 1, 1, 0),
        word(166, 135, 205, 150, "一郎", 1, 1, 1),
        word(285, 135, 320, 150, "10%", 1, 1, 2),
        word(122, 160, 160, 175, "田中", 1, 2, 0),
        word(166, 160, 205, 175, "花子", 1, 2, 1),
        word(285, 160, 320, 175, "20%", 1, 2, 2),
    ]
    candidates = context_candidates_for_page(1, Rect(), words)
    names = [candidate.normalized for candidate in by_type(candidates, "氏名")]
    assert "佐藤一郎" in names
    assert "田中花子" in names


def test_label_right_multiple_word_value() -> None:
    words = [
        word(60, 100, 115, 118, "会社名", 0, 0, 0),
        word(150, 100, 220, 118, "株式会社青空", 0, 0, 1),
        word(225, 100, 280, 118, "製作所", 0, 0, 2),
    ]
    candidates = context_candidates_for_page(1, Rect(), words)
    assert_candidate(candidates, "会社名", "株式会社青空製作所")


def test_label_below_value() -> None:
    words = [
        word(60, 100, 115, 118, "代表者", 0, 0, 0),
        word(62, 138, 100, 156, "山本", 1, 1, 0),
        word(105, 138, 142, 156, "一郎", 1, 1, 1),
    ]
    candidates = context_candidates_for_page(1, Rect(), words)
    assert_candidate(candidates, "氏名", "山本一郎")


def test_general_context_suppression() -> None:
    words = [
        word(50, 50, 150, 68, "貸借対照表", 0, 0, 0),
        word(60, 100, 170, 118, "有形固定資産", 1, 1, 0),
        word(200, 100, 260, 118, "1,000", 1, 1, 1),
        word(60, 150, 230, 168, "銀行について相談できます", 2, 2, 0),
        word(60, 180, 260, 198, "47都道府県に制度を設置しています", 2, 3, 0),
    ]
    candidates = context_candidates_for_page(1, Rect(), words)
    assert not by_type(candidates, "会社名")
    assert not by_type(candidates, "銀行名")
    assert not by_type(candidates, "住所")


def test_section_general_word_is_not_company() -> None:
    words = [
        word(60, 100, 140, 118, "主な販売先", 0, 0, 0),
        word(70, 140, 135, 158, "製品設計", 1, 1, 0),
        word(70, 165, 135, 183, "物流", 1, 2, 0),
    ]
    candidates = context_candidates_for_page(1, Rect(), words)
    assert not by_type(candidates, "会社名")


def test_clear_company_column_value_is_candidate() -> None:
    words = [
        word(50, 100, 90, 118, "No", 0, 0, 0),
        word(120, 100, 180, 118, "取引先", 0, 0, 1),
        word(320, 100, 380, 118, "割合", 0, 0, 2),
        word(122, 135, 190, 150, "合同会社", 1, 1, 0),
        word(195, 135, 250, 150, "緑橋", 1, 1, 1),
        word(325, 135, 360, 150, "35%", 1, 1, 2),
    ]
    candidates = context_candidates_for_page(1, Rect(), words)
    assert_candidate(candidates, "会社名", "合同会社緑橋")


def test_real_bank_value_is_kept() -> None:
    words = [
        word(60, 100, 130, 118, "取引銀行", 0, 0, 0),
        word(165, 100, 260, 118, "青空銀行", 0, 0, 1),
        word(265, 100, 330, 118, "本店", 0, 0, 2),
    ]
    candidates = context_candidates_for_page(1, Rect(), words)
    assert_candidate(candidates, "銀行名", "青空銀行")


def main() -> int:
    test_split_table_company()
    test_split_address_column()
    test_person_column_values()
    test_label_right_multiple_word_value()
    test_label_below_value()
    test_general_context_suppression()
    test_section_general_word_is_not_company()
    test_clear_company_column_value_is_candidate()
    test_real_bank_value_is_kept()
    print("pdf_context_rules_synthetic_tests=passed")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
