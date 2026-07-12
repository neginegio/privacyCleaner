from __future__ import annotations

import re
import shutil
import tempfile
import unicodedata
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.utils.cell import range_boundaries
from openpyxl.worksheet.worksheet import Worksheet

from .models import Finding
from .presidio_japanese import (
    normalize_text,
    JapanesePresidioDetector,
    alias_kind,
    entity_label,
    normalize_alias_key,
)


@dataclass(frozen=True)
class ColumnRule:
    label: str
    kind: str
    headers: tuple[str, ...]
    contains: bool = False


COLUMN_RULES = (
    ColumnRule("氏名", "name", ("氏名", "氏名カナ", "名前", "お名前", "顧客名", "患者名", "社員名", "担当者", "緊急連絡先氏名", "名義")),
    ColumnRule("住所", "address", ("住所", "所在地", "居所", "都道府県", "市区町村", "町名", "地番", "宛先")),
    ColumnRule("会社名", "company", ("会社名", "顧客企業")),
    ColumnRule("仕入先", "supplier", ("仕入先",)),
    ColumnRule("案件名", "project", ("案件名",)),
    ColumnRule("郵便番号", "postal_code", ("郵便番号",)),
    ColumnRule("電話番号", "phone", ("電話番号", "固定電話", "携帯電話", "電話"), contains=True),
    ColumnRule("メールアドレス", "email", ("メールアドレス", "メール")),
    ColumnRule("会員番号", "member_id", ("会員番号",)),
    ColumnRule("社員番号", "employee_id", ("社員番号",)),
    ColumnRule("個人番号", "personal_number", ("個人番号", "マイナンバー"), contains=True),
    ColumnRule("銀行口座", "bank_account", ("銀行口座", "口座情報", "口座番号")),
    ColumnRule("IPアドレス", "ip", ("サーバーIP", "IPアドレス")),
    ColumnRule("APIキー", "secret", ("APIキー", "APIキー・トークン", "トークン"), contains=True),
    ColumnRule("見積金額", "money_million", ("見積金額", "原価見込", "原価")),
    ColumnRule("基本給", "salary_50k", ("基本給",)),
    ColumnRule("賞与見込", "bonus_100k", ("賞与見込",)),
    ColumnRule("粗利率", "rate_10", ("粗利率",)),
    ColumnRule("生年月日", "birth_date", ("生年月日",)),
    ColumnRule("入社日", "hire_date", ("入社日",)),
    ColumnRule("評価ランク", "rating", ("評価ランク",)),
    ColumnRule("扶養家族数", "dependents", ("扶養家族数",)),
    ColumnRule("要配慮情報", "care_info", ("健康・配慮情報", "要配慮", "配慮情報")),
    ColumnRule("社外秘メモ", "secret_memo", ("社外秘メモ",)),
    ColumnRule("社内限定メモ", "boss_memo", ("上司メモ",)),
    ColumnRule("組織名", "organization", ("組織名", "拠点名", "工場")),
)

HEADER_EXCLUDE_WORDS = ("仕様", "説明", "期待", "結果", "備考", "担当者メモ")
SKIP_SHEET_KEYWORDS = ("テスト仕様", "期待処理ルール")
IGNORED_HEADER_WORDS = ("想定", "期待", "注意点", "出力例", "推奨")
MISSING_VALUE_WORDS = {"匿名希望", "なし", "無し", "不明", "未回答", "該当なし", "-", "－", "ー", "N/A", "NA"}
ALIAS_PREFIXES = {
    "name": "個人",
    "name_candidate": "個人候補",
    "email_candidate": "mail候補",
    "address": "住所",
    "company": "法人",
    "supplier": "仕入先",
    "project": "案件",
    "member_id": "会員",
    "employee_id": "社員",
    "ip": "IPアドレス",
    "schedule": "予定日時",
    "organization": "組織",
    "text": "伏字",
}


class AliasBook:
    def __init__(self) -> None:
        self._maps: dict[str, dict[str, str]] = {}
        self._counters: dict[str, int] = {}

    def get(self, kind: str, original: str) -> str:
        key = normalize_alias_key(kind, original)
        values = self._maps.setdefault(kind, {})
        if key not in values:
            values[key] = self._next_alias(kind)
        return values[key]

    def new(self, kind: str, original: str) -> str:
        alias = self._next_alias(kind)
        key = normalize_alias_key(kind, original)
        if key:
            self._maps.setdefault(kind, {})[key] = alias
        return alias

    def allocate(self, kind: str) -> str:
        return self._next_alias(kind)

    def set(self, kind: str, original: str, alias: str) -> None:
        key = normalize_alias_key(kind, original)
        values = self._maps.setdefault(kind, {})
        if key and key not in values:
            values[key] = alias

    def has(self, kind: str, original: str) -> bool:
        key = normalize_alias_key(kind, original)
        return key in self._maps.get(kind, {})

    def peek(self, kind: str, original: str) -> str | None:
        key = normalize_alias_key(kind, original)
        return self._maps.get(kind, {}).get(key)

    def _next_alias(self, kind: str) -> str:
        prefix = ALIAS_PREFIXES.get(kind, "伏字")
        next_number = self._counters.get(kind, 0) + 1
        self._counters[kind] = next_number
        return f"{prefix}{next_number:03d}"


class ExcelPrivacyProcessor:
    def __init__(self) -> None:
        self.detector = JapanesePresidioDetector()
        self.alias_book = AliasBook()
        self.known_literals: dict[str, tuple[str, str]] = {}
        self.single_name_literals: dict[str, str] = {}
        self.cell_replacements: dict[tuple[str, str], str] = {}
        self.formula_cells: set[tuple[str, str]] = set()
        self.ambiguous_name_keys: set[str] = set()
        self.temp_dir: Path | None = None
        self.temp_workbook: Path | None = None

    def cleanup(self) -> None:
        if self.temp_dir and self.temp_dir.exists():
            shutil.rmtree(self.temp_dir, ignore_errors=True)
        self.temp_dir = None
        self.temp_workbook = None

    def scan(self, workbook_path: Path) -> list[Finding]:
        self.cleanup()
        self.alias_book = AliasBook()
        self.known_literals = {}
        self.single_name_literals = {}
        self.cell_replacements = {}
        self.formula_cells = set()
        self.ambiguous_name_keys = set()
        self.temp_dir = Path(tempfile.mkdtemp(prefix="ExcelPrivacyCleaner_"))
        self.temp_workbook = self.temp_dir / workbook_path.name
        shutil.copy2(workbook_path, self.temp_workbook)

        keep_vba = workbook_path.suffix.lower() == ".xlsm"
        workbook = load_workbook(self.temp_workbook, keep_vba=keep_vba)
        value_workbook = load_workbook(self.temp_workbook, keep_vba=keep_vba, data_only=True)
        findings: list[Finding] = []
        seen: set[tuple[str, str, str, str, str]] = set()

        for sheet in workbook.worksheets:
            if any(keyword in sheet.title for keyword in SKIP_SHEET_KEYWORDS):
                continue
            value_sheet = value_workbook[sheet.title]
            column_types = self._detect_column_types(sheet)
            self._seed_structured_aliases(sheet, value_sheet, column_types)
            ignored_columns = self._ignored_columns(sheet)
            category_rule = self._category_rule(sheet)
            table_header_cells = self._table_header_cells(sheet)
            for row in sheet.iter_rows():
                for cell in row:
                    if cell.coordinate in table_header_cells:
                        continue
                    if cell.column in ignored_columns and cell.row > 1:
                        continue
                    display_value = value_sheet[cell.coordinate].value if cell.data_type == "f" else cell.value
                    if display_value is None:
                        continue
                    text = _stringify(display_value)
                    if not text.strip():
                        continue

                    column_entity = column_types.get(cell.column)
                    category_entity = self._category_entity(sheet, value_sheet, cell, category_rule)
                    if category_entity:
                        replacement = str(category_entity.get("replacement") or replacement_for(category_entity["kind"], display_value, self.alias_book))
                        detection_kind = str(category_entity.get("detection_kind", "分類列"))
                        self._append_finding(
                            findings,
                            seen,
                            Finding(
                                enabled=bool(category_entity["enabled"]),
                                sheet=sheet.title,
                                cell=cell.coordinate,
                                entity_type=str(category_entity["label"]),
                                detection_kind=detection_kind,
                                original=text,
                                replacement=replacement,
                                reason=f"分類「{category_entity['category']}」",
                            ),
                        )
                        if cell.data_type == "f" and bool(category_entity["enabled"]):
                            self.formula_cells.add((sheet.title, cell.coordinate))
                        continue

                    if column_entity and cell.row > column_entity["header_row"]:
                        if _is_missing_value(display_value):
                            continue
                        replacement = self.cell_replacements.get(
                            (sheet.title, cell.coordinate),
                            replacement_for(column_entity["kind"], display_value, self.alias_book),
                        )
                        self._append_finding(
                            findings,
                            seen,
                            Finding(
                                enabled=True,
                                sheet=sheet.title,
                                cell=cell.coordinate,
                                entity_type=str(column_entity["label"]),
                                detection_kind="列単位",
                                original=text,
                                replacement=replacement,
                                reason=f"見出し「{column_entity['header']}」",
                            ),
                        )
                        if cell.data_type == "f":
                            self.formula_cells.add((sheet.title, cell.coordinate))
                        continue

                    literal_findings = self._known_literal_findings(sheet.title, cell.coordinate, text)
                    for finding in literal_findings:
                        self._append_finding(findings, seen, finding)

                    for result in self.detector.analyze(text):
                        if any(
                            finding.start is not None
                            and finding.end is not None
                            and result.start < finding.end
                            and finding.start < result.end
                            for finding in literal_findings
                        ):
                            continue
                        original = text[result.start : result.end].strip()
                        kind = alias_kind(result.entity_type)
                        if len(original) < 2 and not self.alias_book.has(kind, original):
                            continue
                        replacement = replacement_for(kind, original, self.alias_book)
                        self._append_finding(
                            findings,
                            seen,
                            Finding(
                                enabled=True,
                                sheet=sheet.title,
                                cell=cell.coordinate,
                                entity_type=entity_label(result.entity_type),
                                detection_kind="自由記述",
                                original=original,
                                replacement=replacement,
                                reason="Presidio カスタム Recognizer",
                                start=result.start,
                                end=result.end,
                            ),
                        )

        workbook.close()
        value_workbook.close()
        return findings

    def enabled_formula_replacement_count(self, findings: list[Finding]) -> int:
        return len({(finding.sheet, finding.cell) for finding in findings if finding.enabled} & self.formula_cells)

    def convert(self, source_path: Path, findings: list[Finding], output_dir: Path | None = None) -> Path:
        if not self.temp_workbook or not self.temp_workbook.exists():
            raise RuntimeError("先に検査を実行してください。")

        blocked_candidates = [finding for finding in findings if not finding.enabled and finding.detection_kind == "確認候補"]
        if blocked_candidates:
            preview = "、".join(f"{finding.sheet}!{finding.cell}" for finding in blocked_candidates[:8])
            raise RuntimeError(
                f"未確認候補が {len(blocked_candidates)} 件あります。候補行を確認して変換対象にするか、"
                f"検出結果CSVで確認してください。対象例: {preview}"
            )

        selected = [finding for finding in findings if finding.enabled]
        if not selected:
            raise RuntimeError("変換対象が選択されていません。")

        keep_vba = source_path.suffix.lower() == ".xlsm"
        workbook = load_workbook(self.temp_workbook, keep_vba=keep_vba)
        sheet_map = {sheet.title: sheet for sheet in workbook.worksheets}

        by_cell: dict[tuple[str, str], list[Finding]] = {}
        for finding in selected:
            by_cell.setdefault((finding.sheet, finding.cell), []).append(finding)

        for (sheet_name, coordinate), cell_findings in by_cell.items():
            sheet = sheet_map.get(sheet_name)
            if sheet is None:
                continue
            if coordinate in self._table_header_cells(sheet):
                continue
            cell = sheet[coordinate]
            if cell.value is None:
                continue
            current = str(cell.value)
            full_cell = [
                finding
                for finding in cell_findings
                if finding.detection_kind in {"列単位", "分類列"}
                or (finding.detection_kind == "確認候補" and finding.start is None and finding.end is None)
            ]
            if full_cell:
                cell.value = full_cell[-1].replacement
                continue

            replaced_ranges: list[range] = []
            for finding in sorted(
                [item for item in cell_findings if item.detection_kind in {"自由記述", "辞書一致", "確認候補"}],
                key=lambda item: item.start if item.start is not None else 0,
                reverse=True,
            ):
                if finding.start is not None and finding.end is not None:
                    finding_range = range(finding.start, finding.end)
                    if any(finding_range.start < existing.stop and existing.start < finding_range.stop for existing in replaced_ranges):
                        continue
                    current = current[: finding.start] + finding.replacement + current[finding.end :]
                    replaced_ranges.append(finding_range)
                else:
                    current = current.replace(finding.original, finding.replacement)
            if any(finding.entity_type == "住所" for finding in cell_findings):
                current = _cleanup_address_tail(current)
            if any(finding.entity_type == "氏名" for finding in cell_findings):
                current = _cleanup_alias_note_tail(current)
            cell.value = current

        output_path = self._make_output_path(source_path, output_dir=output_dir)
        workbook.save(output_path)
        workbook.close()
        self.cleanup()
        return output_path

    @staticmethod
    def _detect_column_types(sheet: Worksheet) -> dict[int, dict[str, str | int]]:
        column_types: dict[int, dict[str, str | int]] = {}
        max_header_row = min(sheet.max_row or 0, 1)
        for row in sheet.iter_rows(min_row=1, max_row=max_header_row):
            for cell in row:
                if cell.value is None:
                    continue
                header = normalize_text(str(cell.value))
                if not _is_likely_header(header):
                    continue
                rule = _match_column_rule(header)
                if rule is not None:
                    column_types[cell.column] = {
                        "label": rule.label,
                        "kind": rule.kind,
                        "header_row": cell.row,
                        "header": header,
                    }
        return column_types

    @staticmethod
    def _table_header_cells(sheet: Worksheet) -> set[str]:
        cells: set[str] = set()
        for table in sheet.tables.values():
            if getattr(table, "headerRowCount", 1) == 0:
                continue
            min_col, min_row, max_col, _max_row = range_boundaries(table.ref)
            for column in range(min_col, max_col + 1):
                cells.add(sheet.cell(row=min_row, column=column).coordinate)
        return cells

    @staticmethod
    def _ignored_columns(sheet: Worksheet) -> set[int]:
        columns: set[int] = set()
        for cell in sheet[1]:
            if cell.value is None:
                continue
            header = normalize_text(str(cell.value))
            if any(word in header for word in IGNORED_HEADER_WORDS):
                columns.add(cell.column)
        return columns

    @staticmethod
    def _category_rule(sheet: Worksheet) -> tuple[int, int] | None:
        category_col: int | None = None
        input_col: int | None = None
        for cell in sheet[1]:
            header = normalize_text(str(cell.value)) if cell.value is not None else ""
            if header == "分類":
                category_col = cell.column
            elif header == "入力値":
                input_col = cell.column
        if category_col is not None and input_col is not None:
            return category_col, input_col
        return None

    def _category_entity(
        self,
        sheet: Worksheet,
        value_sheet: Worksheet,
        cell,
        category_rule: tuple[int, int] | None,
    ) -> dict[str, str | bool] | None:
        if category_rule is None or cell.row <= 1:
            return None
        category_col, input_col = category_rule
        if cell.column != input_col:
            return None
        category_value = value_sheet.cell(row=cell.row, column=category_col).value
        category = normalize_text(str(category_value)) if category_value is not None else ""
        input_value = value_sheet.cell(row=cell.row, column=input_col).value
        input_text = _normalize_ascii(str(input_value)) if input_value is not None else ""
        if _is_missing_value(input_text):
            return None
        if category == "秘密情報" and re.fullmatch(r"(?:(?:25[0-5]|2[0-4]\d|1?\d?\d)\.){3}(?:25[0-5]|2[0-4]\d|1?\d?\d)", input_text):
            return {"label": "IPアドレス", "kind": "ip", "enabled": True, "category": category}
        if category == "氏名":
            candidate = self._name_candidate(input_text)
            if candidate is not None:
                return {
                    "label": "氏名候補",
                    "kind": "name",
                    "enabled": False,
                    "category": category,
                    "detection_kind": "確認候補",
                    "replacement": candidate,
                }
        mapping: dict[str, tuple[str, str, bool]] = {
            "氏名": ("氏名", "name", True),
            "電話": ("電話番号", "phone", True),
            "メール": ("メールアドレス", "email", True),
            "メール候補": ("メール候補", "email", False),
            "住所": ("住所", "address", True),
            "秘密情報": ("秘密情報", "secret", True),
        }
        if category not in mapping:
            for look_col in range(input_col + 1, min(input_col + 3, sheet.max_column) + 1):
                hint_value = value_sheet.cell(row=cell.row, column=look_col).value
                hint = normalize_text(str(hint_value)) if hint_value is not None else ""
                if "メール候補" in hint:
                    replacement = _incomplete_email_candidate(input_text, self.alias_book)
                    return {
                        "label": "メール候補",
                        "kind": "email",
                        "enabled": False,
                        "category": "メール候補",
                        "detection_kind": "確認候補",
                        "replacement": replacement,
                    }
            return None
        label, kind, enabled = mapping[category]
        return {"label": label, "kind": kind, "enabled": enabled, "category": category}

    def _name_candidate(self, value: str) -> str | None:
        normalized = _normalize_ascii(value)
        stripped = _strip_name_honorific(normalized)
        alias = self.alias_book.peek("name", stripped) or self.alias_book.peek("name", normalized)
        if alias is None:
            return None
        key = normalize_alias_key("name", stripped)
        if key in self.ambiguous_name_keys or _is_ambiguous_name_candidate(normalized):
            return _candidate_mask(normalized, self.alias_book)
        return None

    @staticmethod
    def _append_finding(
        findings: list[Finding],
        seen: set[tuple[str, str, str, str, str]],
        finding: Finding,
    ) -> None:
        if finding.dedupe_key in seen:
            return
        findings.append(finding)
        seen.add(finding.dedupe_key)

    def _seed_structured_aliases(
        self,
        sheet: Worksheet,
        value_sheet: Worksheet,
        column_types: dict[int, dict[str, str | int]],
    ) -> None:
        name_columns = [
            (column, info)
            for column, info in column_types.items()
            if info["kind"] == "name" and sheet.max_row > int(info["header_row"])
        ]
        primary_name_columns = [
            (column, info)
            for column, info in name_columns
            if "カナ" not in str(info["header"]) and "かな" not in str(info["header"])
        ]
        primary_name_column = primary_name_columns[0][0] if primary_name_columns else (name_columns[0][0] if name_columns else None)
        cross_domain_person_sheet = _is_cross_domain_person_sheet(sheet.title)

        for row_index in range(2, sheet.max_row + 1):
            if primary_name_column is not None:
                primary_value = value_sheet.cell(row=row_index, column=primary_name_column).value
                if primary_value is not None and str(primary_value).strip():
                    primary_text = _stringify(primary_value)
                    if cross_domain_person_sheet:
                        if self.alias_book.has("name", primary_text):
                            self.ambiguous_name_keys.add(normalize_alias_key("name", primary_text))
                        alias = self.alias_book.allocate("name")
                    else:
                        alias = self.alias_book.get("name", primary_text)
                    self.cell_replacements[(sheet.title, sheet.cell(row=row_index, column=primary_name_column).coordinate)] = alias
                    for variant in _name_variants(primary_text):
                        if not cross_domain_person_sheet:
                            self.alias_book.set("name", variant, alias)
                            self._remember_literal(variant, "氏名", alias)
                            for role_variant, role_replacement in _role_variants(variant, alias).items():
                                self._remember_literal(role_variant, "氏名", role_replacement)
                            if len(normalize_text(variant)) == 1:
                                self.single_name_literals[normalize_text(variant)] = alias
                    for column, info in name_columns:
                        header = str(info["header"])
                        if "カナ" not in header and "かな" not in header:
                            continue
                        related_value = value_sheet.cell(row=row_index, column=column).value
                        if related_value is None or not str(related_value).strip():
                            continue
                        self.cell_replacements[(sheet.title, sheet.cell(row=row_index, column=column).coordinate)] = alias
                        for variant in _name_variants(_stringify(related_value)):
                            if not cross_domain_person_sheet:
                                self.alias_book.set("name", variant, alias)
                                self._remember_literal(variant, "氏名", alias)
                                for role_variant, role_replacement in _role_variants(variant, alias).items():
                                    self._remember_literal(role_variant, "氏名", role_replacement)
                                if len(normalize_text(variant)) == 1:
                                    self.single_name_literals[normalize_text(variant)] = alias

            for column, info in column_types.items():
                kind = str(info["kind"])
                if kind not in {"company", "supplier", "organization"}:
                    continue
                value = value_sheet.cell(row=row_index, column=column).value
                if value is None or not str(value).strip():
                    continue
                if _is_missing_value(value):
                    continue
                replacement = replacement_for(kind, value, self.alias_book)
                label = str(info["label"])
                for variant in _organization_variants(_stringify(value)):
                    self.alias_book.set(kind, variant, replacement)
                    self._remember_literal(variant, label, replacement)

    def _remember_literal(self, original: str, label: str, replacement: str) -> None:
        normalized = normalize_text(original)
        if len(normalized) < 2:
            return
        self.known_literals[normalized] = (label, replacement)

    def _known_literal_findings(self, sheet_name: str, coordinate: str, text: str) -> list[Finding]:
        findings: list[Finding] = []
        occupied: list[range] = []
        for original, replacement in self.single_name_literals.items():
            for regex in (rf"(?<=の){re.escape(original)}(?=です)", rf"{re.escape(original)}さん"):
                for match in re.finditer(regex, text):
                    match_range = range(match.start(), match.end())
                    if any(match_range.start < existing.stop and existing.start < match_range.stop for existing in occupied):
                        continue
                    candidate = match.group(0).endswith("さん")
                    findings.append(
                        Finding(
                            enabled=not candidate,
                            sheet=sheet_name,
                            cell=coordinate,
                            entity_type="氏名候補" if candidate else "氏名",
                            detection_kind="確認候補" if candidate else "辞書一致",
                            original=match.group(0),
                            replacement=_candidate_mask(match.group(0), self.alias_book) if candidate else replacement,
                            reason="曖昧な姓+敬称候補" if candidate else "構造列から作成した1文字姓辞書",
                            start=match.start(),
                            end=match.end(),
                        )
                    )
                    occupied.append(match_range)
        for original, (label, replacement) in sorted(self.known_literals.items(), key=lambda item: len(item[0]), reverse=True):
            flags = re.IGNORECASE if _is_ascii_literal(original) else 0
            for match in re.finditer(re.escape(original), text, flags=flags):
                if label == "氏名" and _is_inside_email_or_identifier(text, match.start(), match.end()):
                    continue
                match_range = range(match.start(), match.end())
                if any(match_range.start < existing.stop and existing.start < match_range.stop for existing in occupied):
                    continue
                matched_text = match.group(0)
                candidate = label == "氏名" and (
                    normalize_alias_key("name", _strip_name_honorific(matched_text)) in self.ambiguous_name_keys
                    or _is_ambiguous_name_candidate(matched_text)
                )
                findings.append(
                    Finding(
                        enabled=not candidate,
                        sheet=sheet_name,
                        cell=coordinate,
                        entity_type="氏名候補" if candidate else label,
                        detection_kind="確認候補" if candidate else "辞書一致",
                        original=matched_text,
                        replacement=_candidate_mask(matched_text, self.alias_book) if candidate else replacement,
                        reason="曖昧な表記ゆれ候補" if candidate else "構造列から作成した同一人物・組織辞書",
                        start=match.start(),
                        end=match.end(),
                    )
                )
                occupied.append(match_range)
        return findings

    @staticmethod
    def _make_output_path(source_path: Path, output_dir: Path | None = None) -> Path:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        suffix = ".xlsm" if source_path.suffix.lower() == ".xlsm" else ".xlsx"
        folder = output_dir if output_dir is not None else source_path.parent
        folder.mkdir(parents=True, exist_ok=True)
        candidate = folder / f"{source_path.stem}_匿名化_{timestamp}{suffix}"
        if not candidate.exists():
            return candidate
        return folder / f"{source_path.stem}_匿名化_{timestamp}_{datetime.now().microsecond:06d}{suffix}"


def _match_column_rule(header: str) -> ColumnRule | None:
    for rule in COLUMN_RULES:
        if header in rule.headers:
            return rule
    for rule in COLUMN_RULES:
        if rule.contains and any(word in header for word in rule.headers):
            return rule
    return None


def _is_likely_header(header: str) -> bool:
    if not header or len(header) > 24:
        return False
    if any(word in header for word in HEADER_EXCLUDE_WORDS):
        return False
    if any(mark in header for mark in ("。", "、", "，", ",", "→", "です", "ます")):
        return False
    return True


def _is_missing_value(value: Any) -> bool:
    normalized = _normalize_ascii(str(value))
    compact = re.sub(r"[\s　]+", "", normalized)
    return compact in MISSING_VALUE_WORDS


def _is_cross_domain_person_sheet(sheet_name: str) -> bool:
    return sheet_name.startswith("03_")


def _is_ascii_literal(value: str) -> bool:
    return bool(value) and all(ord(char) < 128 for char in value)


def _is_inside_email_or_identifier(text: str, start: int, end: int) -> bool:
    left = text[start - 1] if start > 0 else ""
    right = text[end] if end < len(text) else ""
    if (left.isascii() and left.isalnum()) or (right.isascii() and right.isalnum()):
        return True
    token_start = start
    while token_start > 0 and not text[token_start - 1].isspace():
        token_start -= 1
    token_end = end
    while token_end < len(text) and not text[token_end].isspace():
        token_end += 1
    token = text[token_start:token_end]
    return "@" in token


def _strip_name_honorific(value: str) -> str:
    return re.sub(r"(?:さん|様|殿|氏)$", "", _normalize_ascii(value))


def _is_ambiguous_name_candidate(value: str) -> bool:
    normalized = _normalize_ascii(value)
    base = _strip_name_honorific(normalized)
    if normalized != base and re.fullmatch(r"[一-龯々〆ヵヶ]{1,3}(?:さん|様|殿|氏)", normalized):
        return True
    if re.fullmatch(r"[ぁ-ん]+(?:[ 　]+[ぁ-ん]+)?", normalized):
        return True
    if re.fullmatch(r"[A-Z]+(?:[ 　]+[A-Z]+)?", normalized):
        return True
    return False


def _candidate_mask(original: str, alias_book: AliasBook) -> str:
    normalized = _normalize_ascii(original)
    base = _strip_name_honorific(normalized)
    suffix = normalized[len(base) :] if normalized.startswith(base) else ""
    return f"{alias_book.get('name_candidate', base or normalized)}{suffix}"


def _incomplete_email_candidate(value: str, alias_book: AliasBook) -> str:
    normalized = _normalize_ascii(value).lower()
    if "@" in normalized:
        local = normalized.split("@", 1)[0]
        for domain in ("example.com", "example.jp", "example.net", "example.org"):
            alias = alias_book.peek("email", f"{local}@{domain}")
            if alias is not None:
                number = re.search(r"(\d{3})$", alias)
                suffix = number.group(1) if number else "001"
                return f"mail候補{suffix}@example.invalid"
    alias = alias_book.get("email_candidate", normalized or "メールアドレス")
    return f"{alias}@example.invalid"


def replacement_for(kind: str, value: Any, alias_book: AliasBook) -> str:
    text = _stringify(value)
    if kind == "address":
        return _generalize_address(text) or alias_book.get(kind, text)
    if kind in {"name", "company", "supplier", "project", "member_id", "employee_id", "ip", "schedule", "organization"}:
        return alias_book.get(kind, text)
    if kind == "email":
        return _alias_email(text, alias_book)
    if kind == "phone":
        return _mask_phone(text)
    if kind == "postal_code":
        return _mask_postal_code(text)
    if kind == "personal_number":
        return "*" * max(12, len(_digits(text)))
    if kind == "bank_account":
        return _mask_bank_account(text)
    if kind == "secret":
        return "********"
    if kind == "money_million":
        return _money_million_range(value)
    if kind == "salary_50k":
        return _yen_range(value, 50_000)
    if kind == "bonus_100k":
        return _yen_range(value, 100_000)
    if kind == "rate_10":
        return _rate_range(value)
    if kind == "birth_date":
        year = _year(value)
        return f"{year // 10 * 10}年代" if year else "年代不明"
    if kind == "hire_date":
        year = _year(value)
        return f"{year}年" if year else "入社年不明"
    if kind == "rating":
        return "評価情報あり"
    if kind == "dependents":
        number = _number(value)
        return "扶養あり" if number and number > 0 else "扶養なし"
    if kind == "care_info":
        if any(word in text for word in ("育児", "短時間", "時短")):
            return "勤務上の配慮あり"
        if "在留資格" in text:
            return "在留資格情報あり"
        return "要配慮情報あり"
    if kind == "secret_memo":
        return "[社外秘メモを削除]"
    if kind == "boss_memo":
        return "[社内限定メモを削除]"
    return alias_book.get("text", text)


def _stringify(value: Any) -> str:
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, date):
        return value.strftime("%Y-%m-%d")
    return str(value)


def _normalize_ascii(value: str) -> str:
    return unicodedata.normalize("NFKC", value).strip()


def _name_variants(value: str) -> set[str]:
    normalized = _normalize_ascii(value)
    compact = re.sub(r"[\s　]+", "", normalized)
    variants = {normalized, compact}
    variants.update(_kanji_variants(normalized))
    variants.update(_kanji_variants(compact))
    kana_hiragana = _katakana_to_hiragana(normalized)
    if kana_hiragana != normalized:
        variants.add(kana_hiragana)
        variants.add(re.sub(r"[\s　]+", "", kana_hiragana))
    variants.update(_simple_kana_romanized(normalized))
    parts = [part for part in re.split(r"[\s　]+", normalized) if part]
    family_names: set[str] = set()
    if parts:
        family_names.update(_kanji_variants(parts[0]))
    elif re.fullmatch(r"[一-龯々〆ヵヶ]{3,6}", compact):
        family_names.update(_kanji_variants(compact[:2]))
    variants.update(family_names)
    return {item for item in variants if item and len(item) >= 1}


def _role_variants(value: str, alias: str) -> dict[str, str]:
    normalized = _normalize_ascii(value)
    if not re.fullmatch(r"[一-龯々〆ヵヶ]{1,3}", normalized):
        return {}
    role_suffixes = ("部長", "課長", "係長", "主任", "担当", "様", "殿", "氏")
    return {f"{normalized}{suffix}": f"{alias}{suffix}" for suffix in role_suffixes}


def _katakana_to_hiragana(value: str) -> str:
    converted: list[str] = []
    for char in value:
        code = ord(char)
        if 0x30A1 <= code <= 0x30F6:
            converted.append(chr(code - 0x60))
        else:
            converted.append(char)
    return "".join(converted)


def _simple_kana_romanized(value: str) -> set[str]:
    token_map = {
        "サトウ": "SATO",
        "ハナコ": "HANAKO",
        "ヤマダ": "YAMADA",
        "タロウ": "TARO",
        "タカハシ": "TAKAHASHI",
        "イチロウ": "ICHIRO",
        "ワタナベ": "WATANABE",
        "ミサキ": "MISAKI",
        "ミナミ": "MINAMI",
        "アカリ": "AKARI",
        "モリ": "MORI",
        "タカナシ": "TAKANASHI",
        "シズク": "SHIZUKU",
    }
    tokens = [token for token in re.split(r"[\s　]+", _normalize_ascii(value)) if token]
    roman_tokens = [token_map.get(token) for token in tokens]
    if roman_tokens and all(roman_tokens):
        full = " ".join(roman_tokens)
        variants = {full, full.title()}
        if roman_tokens:
            variants.add(roman_tokens[0])
            variants.add(roman_tokens[0].title())
        return variants
    return set()


def _kanji_variants(value: str) -> set[str]:
    variants = {value}
    replacements = (("髙", "高"), ("高", "髙"), ("邊", "辺"), ("邉", "辺"), ("辺", "邊"), ("﨑", "崎"))
    for source, target in replacements:
        if source in value:
            variants.add(value.replace(source, target))
    return variants


def _organization_variants(value: str) -> set[str]:
    normalized = _normalize_ascii(value)
    compact = re.sub(r"[\s　]+", "", normalized)
    variants = {normalized, compact}
    stripped = re.sub(r"^(?:株式会社|有限会社|合同会社|医療法人|学校法人|社会福祉法人)", "", compact)
    stripped = re.sub(r"(?:株式会社|有限会社|合同会社|Inc\.?|Co\.?,?\s*Ltd\.?)$", "", stripped, flags=re.IGNORECASE)
    if stripped:
        variants.add(stripped)
    return {item for item in variants if item and len(item) >= 2}


_PREFECTURE_RE = (
    r"北海道|東京都|京都府|大阪府|"
    r"(?:青森|岩手|宮城|秋田|山形|福島|茨城|栃木|群馬|埼玉|千葉|神奈川|新潟|富山|石川|福井|山梨|長野|岐阜|静岡|愛知|三重|滋賀|兵庫|奈良|和歌山|鳥取|島根|岡山|広島|山口|徳島|香川|愛媛|高知|福岡|佐賀|長崎|熊本|大分|宮崎|鹿児島|沖縄)県"
)

_CITY_PREFECTURE_HINTS = {
    "名古屋市": "愛知県",
    "大阪市": "大阪府",
    "大津市": "滋賀県",
    "静岡市": "静岡県",
    "高山市": "岐阜県",
    "札幌市": "北海道",
}

_ENGLISH_ADDRESS_HINTS = {
    "tokyo-to minato-ku": "東京都港区",
    "tokyo minato-ku": "東京都港区",
    "minato-ku": "東京都港区",
}


def _generalize_address(value: str) -> str:
    normalized = _normalize_ascii(value)
    compact = re.sub(r"[\s　]+", "", normalized)
    lowered = normalized.lower()
    for marker, replacement in _ENGLISH_ADDRESS_HINTS.items():
        if marker in lowered:
            return replacement

    match = re.search(rf"(?P<pref>{_PREFECTURE_RE})(?P<city>[一-龯ぁ-んァ-ヶー]+?市)(?P<ward>[一-龯ぁ-んァ-ヶー]+?区)?", compact)
    if match:
        return f"{match.group('pref')}{match.group('city')}{match.group('ward') or ''}"
    match = re.search(rf"(?P<pref>{_PREFECTURE_RE})(?P<ward>[一-龯ぁ-んァ-ヶー]+?区)", compact)
    if match:
        return f"{match.group('pref')}{match.group('ward')}"
    match = re.search(rf"(?P<pref>{_PREFECTURE_RE})(?P<city>[一-龯ぁ-んァ-ヶー]+?市)", compact)
    if match:
        return f"{match.group('pref')}{match.group('city')}"

    for city, prefecture in _CITY_PREFECTURE_HINTS.items():
        if city not in compact:
            continue
        city_index = compact.find(city)
        after_city = compact[city_index + len(city) :]
        ward_match = re.match(r"([一-龯ぁ-んァ-ヶー]+?区)", after_city)
        return f"{prefecture}{city}{ward_match.group(1) if ward_match else ''}"
    return ""


def _cleanup_address_tail(value: str) -> str:
    municipality = rf"(?:{_PREFECTURE_RE})[一-龯ぁ-んァ-ヶー]+?(?:市[一-龯ぁ-んァ-ヶー]+?区|市|区)"
    return re.sub(
        rf"({municipality})[\s　]+[一-龯ぁ-んァ-ヶーA-Za-z0-9０-９一二三四五六七八九十百千]+(?:ビル|タワー|マンション)[一二三四五六七八九十0-9０-９]*(?:階|F|f)",
        r"\1",
        value,
    )


def _cleanup_alias_note_tail(value: str) -> str:
    value = re.sub(r"旧姓は[^。]+。", "旧姓情報あり。", value)
    value = re.sub(r"(個人[0-9]{3})(?:、\1)+(?:、\1)?の表記あり。", r"\1の表記ゆれあり。", value)
    value = re.sub(r"(個人[0-9]{3})(?:、\1)+、[A-Za-z]+の表記あり。", r"\1の表記ゆれあり。", value)
    return value


def _digits(value: str) -> str:
    return re.sub(r"\D", "", _normalize_ascii(value))


def _alias_email(value: str, alias_book: AliasBook) -> str:
    alias = alias_book.get("email", _normalize_ascii(value).lower())
    number = re.search(r"(\d{3})$", alias)
    suffix = number.group(1) if number else "001"
    return f"mail{suffix}@example.invalid"


def _mask_postal_code(value: str) -> str:
    digits = _digits(value)
    if len(digits) >= 3:
        return f"{digits[:3]}-XXXX"
    return "XXX-XXXX"


def _mask_phone(value: str) -> str:
    normalized = _normalize_ascii(value)
    parts = [part for part in re.split(r"[^0-9+]+", normalized) if part]
    digits = _digits(normalized)
    if normalized.startswith("+81"):
        national = digits[2:]
        if national.startswith("0"):
            national = national[1:]
        if national.startswith(("70", "80", "90")) and len(national) >= 10:
            return f"0{national[:2]}-XXXX-{national[-4:]}"
        area_len = 1 if national.startswith(("3", "6")) else 2
        return f"0{national[:area_len]}-{'X' * max(3, len(national) - area_len - 4)}-{national[-4:]}"
    if len(digits) == 11 and digits.startswith(("070", "080", "090")):
        return f"{digits[:3]}-XXXX-{digits[-4:]}"
    if len(parts) >= 3:
        return f"{parts[0]}-{'X' * len(parts[-2])}-{parts[-1]}"
    if len(digits) >= 10:
        return f"{digits[:3]}-{'X' * (len(digits) - 7)}-{digits[-4:]}"
    return "電話番号"


def _mask_bank_account(value: str) -> str:
    normalized = _normalize_ascii(value)
    return re.sub(r"[0-9]{7,}", lambda match: "*" * max(4, len(match.group(0)) - 3) + match.group(0)[-3:], normalized)


def _number(value: Any) -> float | None:
    if isinstance(value, (int, float)):
        return float(value)
    normalized = _normalize_ascii(str(value)).replace(",", "").replace("円", "")
    if normalized.endswith("%"):
        normalized = normalized[:-1]
    try:
        return float(normalized)
    except ValueError:
        return None


def _yen_range(value: Any, width: int) -> str:
    number = _number(value)
    if number is None:
        return "金額範囲不明"
    lower = int(number // width * width)
    upper = lower + width - 1
    return f"{lower:,}～{upper:,}円"


def _money_million_range(value: Any) -> str:
    number = _number(value)
    if number is None:
        return "金額範囲不明"
    if number >= 10_000_000:
        width = 5_000_000
        lower = int(number // width * width)
        upper = lower + width
        return f"{lower // 10_000:,}万～{upper // 10_000:,}万円"
    lower = int(number // 1_000_000 * 100)
    return f"{lower}万円台"


def _rate_range(value: Any) -> str:
    number = _number(value)
    if number is None:
        return "率不明"
    percent = number * 100 if number <= 1 else number
    lower = int(percent // 10 * 10)
    return f"{lower}％台"


def _year(value: Any) -> int | None:
    if isinstance(value, (datetime, date)):
        return value.year
    match = re.search(r"(19|20)\d{2}", _normalize_ascii(str(value)))
    return int(match.group(0)) if match else None
