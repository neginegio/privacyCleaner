from __future__ import annotations

import sys
import tempfile
from pathlib import Path

from docx import Document

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from excel_privacy_cleaner.excel_processor import ProcessingOptions  # noqa: E402
from excel_privacy_cleaner.word_processor import (  # noqa: E402
    WordPrivacyProcessor,
    WordReplacementDecision,
    candidates_for_inventory,
    extract_word_structure,
)


def assert_true(condition: bool, message: str) -> None:
    if not condition:
        raise AssertionError(message)


def _enable_all(decisions: list[WordReplacementDecision]) -> list[WordReplacementDecision]:
    for decision in decisions:
        if decision.candidate.source != "hyperlink_target":
            decision.enabled = True
    return decisions


def create_bank_details_fixture(path: Path) -> None:
    # Mirrors a real "振込口座" (bank transfer details) table found in a
    # holdout review: header row of column labels, a value row with a plain
    # bank name/branch/account-type/account-number, then a katakana
    # "account holder name" row that's merged across all 4 columns (a real
    # extraction quirk of this kind of form, not something artificially
    # constructed for the test).
    document = Document()
    document.add_paragraph("住    所：愛知県名古屋市熱田区六野1-2-19")
    document.add_paragraph("センチュリースクエア神宮306")

    table = document.add_table(rows=4, cols=4)
    table.cell(0, 0).text = "金融機関名"
    table.cell(0, 1).text = "支店名"
    table.cell(0, 2).text = "種別"
    table.cell(0, 3).text = "口座番号"
    table.cell(1, 0).text = "三菱ＵＦＪ銀行"
    table.cell(1, 1).text = "柳橋支店"
    table.cell(1, 2).text = "普通"
    table.cell(1, 3).text = "1452640"
    table.cell(2, 0).text = "口座名義人（カタカナ）"
    table.cell(2, 1).text = "口座名義人（カタカナ）"
    table.cell(2, 2).text = "口座名義人（カタカナ）"
    table.cell(2, 3).text = "口座名義人（カタカナ）"
    merged = table.cell(3, 0).merge(table.cell(3, 3))
    merged.text = "ﾎｿﾔ　ﾋﾛｼ"
    document.save(path)


def test_merged_table_cell_is_not_processed_multiple_times() -> None:
    # Reproduces a real holdout-review finding: a table cell merged across
    # 4 columns (the katakana account-holder name) was extracted 4 times
    # (once per apparent column, since table.rows[i].cells does not
    # deduplicate a merged cell), and applying its redaction decision 4
    # times to what's actually one physical run corrupted the replacement
    # text into "個人004040404" instead of "個人004".
    with tempfile.TemporaryDirectory(prefix="word_merged_cell_test_") as tmp:
        source = Path(tmp) / "fixture.docx"
        create_bank_details_fixture(source)

        inventory = extract_word_structure(source)
        candidates = candidates_for_inventory(inventory)
        name_candidates = [c for c in candidates if c.text == "ﾎｿﾔ　ﾋﾛｼ"]
        assert_true(len(name_candidates) == 1, f"Merged cell should produce exactly one candidate, got {len(name_candidates)}")

        processor = WordPrivacyProcessor()
        decisions = processor.scan(source, options=ProcessingOptions(mode="analysis"))
        _enable_all(decisions)
        result = processor.convert(source, decisions, output_dir=Path(tmp))

        output_table = Document(result.output_path).tables[0]
        replaced = output_table.cell(3, 0).text
        assert_true("ﾎｿﾔ" not in replaced and "ﾋﾛｼ" not in replaced, "Original katakana name must not remain")
        assert_true(replaced.count("個人") == 1, f"Replacement must not be duplicated/corrupted, got {replaced!r}")


def test_katakana_full_name_with_space_detected_as_one_candidate() -> None:
    # Reproduces the same holdout finding from the detection side: GiNZA
    # only recognized the 3-character surname portion of "ﾎｿﾔ　ﾋﾛｼ" as a
    # person, leaving the given name exposed right next to the redacted
    # surname. A deterministic regex rule (mirroring the existing kanji
    # "word_japanese_full_name_space" rule, but for half-width katakana)
    # must catch the whole name as one match regardless of what GiNZA does.
    with tempfile.TemporaryDirectory(prefix="word_katakana_name_test_") as tmp:
        source = Path(tmp) / "fixture.docx"
        document = Document()
        document.add_paragraph("口座名義人：ﾎｿﾔ　ﾋﾛｼ")
        document.save(source)

        inventory = extract_word_structure(source)
        candidates = candidates_for_inventory(inventory)
        assert_true(
            any(c.category == "氏名" and c.text == "ﾎｿﾔ　ﾋﾛｼ" for c in candidates),
            "The full katakana name (surname + given name) should be detected as a single candidate",
        )

    print("word_katakana_full_name_with_space_tests=passed")


def test_bank_name_with_fullwidth_latin_letters_detected() -> None:
    # Reproduces the other half of the holdout finding: "三菱ＵＦＪ銀行"
    # uses full-width Latin letters (common in real Japanese business/bank
    # names), which the bank-name regex's character class silently excluded
    # -- the match just broke partway through the name instead of erroring,
    # so the bank name went completely undetected.
    with tempfile.TemporaryDirectory(prefix="word_bank_name_test_") as tmp:
        source = Path(tmp) / "fixture.docx"
        document = Document()
        document.add_paragraph("三菱ＵＦＪ銀行のご案内です。")
        document.save(source)

        inventory = extract_word_structure(source)
        candidates = candidates_for_inventory(inventory)
        assert_true(
            any(c.category == "銀行名" and c.text == "三菱ＵＦＪ銀行" for c in candidates),
            "A bank name spelled with full-width Latin letters should still be detected",
        )

    print("word_bank_name_fullwidth_latin_tests=passed")


def test_bank_account_number_detected_via_table_column_header() -> None:
    # Reproduces the most severe part of the holdout finding: the account
    # number itself was never detected at all -- there was no detection
    # rule for a bare account-number table cell with no label of its own
    # (the label "口座番号" sits in a separate header-row cell). The
    # structural table-column detector must recover it, and convert() must
    # mask it in the output.
    with tempfile.TemporaryDirectory(prefix="word_account_number_test_") as tmp:
        source = Path(tmp) / "fixture.docx"
        create_bank_details_fixture(source)

        inventory = extract_word_structure(source)
        candidates = candidates_for_inventory(inventory)
        account_candidates = [c for c in candidates if c.category == "銀行口座"]
        assert_true(len(account_candidates) == 1, f"Expected exactly one bank-account candidate, got {len(account_candidates)}")
        assert_true(account_candidates[0].text == "1452640", "The bare account-number cell should be the detected text")

        # A column that has no "口座番号" header anywhere must not have its
        # plain-looking numeric cells treated as account numbers -- this
        # rule is deliberately scoped to real column headers, not a blanket
        # any-digit-string match.
        assert_true(
            not any(c.category == "銀行口座" and c.text == "1452640" and c.detection_rule != "word_account_number_column" for c in candidates),
            "Only the header-driven rule should have produced this candidate",
        )

        processor = WordPrivacyProcessor()
        decisions = processor.scan(source, options=ProcessingOptions(mode="analysis"))
        _enable_all(decisions)
        result = processor.convert(source, decisions, output_dir=Path(tmp))

        output_table = Document(result.output_path).tables[0]
        assert_true("1452640" not in output_table.cell(1, 3).text, "The account number must not remain in the output")

    print("word_bank_account_number_table_column_tests=passed")


def test_address_continuation_paragraph_detected() -> None:
    # Reproduces the third part of the holdout finding: a form split an
    # address across two paragraphs -- "住所：<prefecture/city/street>" then,
    # on the very next paragraph, just the building name and room number
    # ("センチュリースクエア神宮306") with no address-typical keyword at all.
    # The label paragraph got redacted; the continuation, having nothing an
    # address detector could key off of on its own, did not.
    with tempfile.TemporaryDirectory(prefix="word_address_continuation_test_") as tmp:
        source = Path(tmp) / "fixture.docx"
        create_bank_details_fixture(source)

        inventory = extract_word_structure(source)
        candidates = candidates_for_inventory(inventory)
        continuation = [c for c in candidates if c.text == "センチュリースクエア神宮306"]
        assert_true(len(continuation) == 1, "The address continuation paragraph should be recovered as its own candidate")
        assert_true(continuation[0].category == "住所", "It should be categorized as an address")
        assert_true(
            continuation[0].confidence < 0.75,
            "A structurally-inferred (not directly matched) continuation should require review, not auto-apply",
        )

    print("word_address_continuation_tests=passed")


if __name__ == "__main__":
    test_merged_table_cell_is_not_processed_multiple_times()
    test_katakana_full_name_with_space_detected_as_one_candidate()
    test_bank_name_with_fullwidth_latin_letters_detected()
    test_bank_account_number_detected_via_table_column_header()
    test_address_continuation_paragraph_detected()
    print("all tests passed")
