from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass
from statistics import median
from typing import Any, Iterable

from .pdf_ocr_support import CANDIDATE_REVIEW, OcrCandidate


RULE_LABEL_VALUE = "label_value"
RULE_TABLE_COLUMN = "table_column"
RULE_SECTION_BLOCK = "section_block"
RULE_PATTERN = "pattern"

DEFAULT_ENABLED_RULES = frozenset(
    {
        RULE_LABEL_VALUE,
        RULE_TABLE_COLUMN,
        RULE_SECTION_BLOCK,
        RULE_PATTERN,
    }
)


@dataclass(frozen=True)
class _WordBox:
    text: str
    normalized: str
    x0: float
    y0: float
    x1: float
    y1: float
    block: int
    line: int
    index: int

    @property
    def cx(self) -> float:
        return (self.x0 + self.x1) / 2

    @property
    def cy(self) -> float:
        return (self.y0 + self.y1) / 2


@dataclass(frozen=True)
class _LineBox:
    words: tuple[_WordBox, ...]
    text: str
    normalized: str
    rect: tuple[float, float, float, float]
    block: int
    line: int

    @property
    def x0(self) -> float:
        return self.rect[0]

    @property
    def y0(self) -> float:
        return self.rect[1]

    @property
    def x1(self) -> float:
        return self.rect[2]

    @property
    def y1(self) -> float:
        return self.rect[3]

    @property
    def cx(self) -> float:
        return (self.x0 + self.x1) / 2

    @property
    def cy(self) -> float:
        return (self.y0 + self.y1) / 2


@dataclass(frozen=True)
class _DocumentContext:
    financial_statement: bool = False
    explanatory_text: bool = False
    bank_guidance: bool = False


_LABEL_RULES = (
    ("企業名称", "会社名", "法人001", "label_value:corporate_name"),
    ("会社名", "会社名", "法人001", "label_value:corporate_name"),
    ("法人名", "会社名", "法人001", "label_value:corporate_name"),
    ("商号", "会社名", "法人001", "label_value:corporate_name"),
    ("所在地", "住所", "市区町村まで", "label_value:address"),
    ("住所", "住所", "市区町村まで", "label_value:address"),
    ("代表者", "氏名", "個人001", "label_value:representative"),
    ("代表取締役", "氏名", "個人001", "label_value:representative"),
    ("取引銀行", "銀行名", "金融機関001", "label_value:bank"),
    ("金融機関", "銀行名", "金融機関001", "label_value:bank"),
)

_TABLE_HEADER_RULES = (
    ("氏名", "氏名", "個人001", "table_column:person_name"),
    ("株主", "氏名", "個人001", "table_column:shareholder_name"),
    ("株主名", "氏名", "個人001", "table_column:shareholder_name"),
    ("所在地", "住所", "市区町村まで", "table_column:address"),
    ("住所", "住所", "市区町村まで", "table_column:address"),
    ("販売先", "会社名", "法人001", "table_column:customer"),
    ("仕入先", "会社名", "法人001", "table_column:supplier"),
    ("外注先", "会社名", "法人001", "table_column:subcontractor"),
    ("取引先", "会社名", "法人001", "table_column:business_partner"),
)

_SECTION_HEADERS = (
    "主な仕入",
    "仕入先",
    "外注先",
    "主な販売先",
    "販売先",
    "取引先",
    "株主一覧",
    "役員構成",
)

_NON_TARGET_SUMMARY_WORDS = (
    "合計",
    "小計",
    "その他",
    "値引",
    "売上高",
    "構成比",
    "単位",
    "年度",
)

_BUSINESS_PROCESS_WORDS = (
    "製品企画",
    "製品設計",
    "仕上げ",
    "鋼材仕入",
    "部品仕入",
    "塗料仕入",
    "物流",
    "納品",
)

_NON_VALUE_CONTEXT_WORDS = (
    "項目",
    "摘要",
    "対象会社",
    "役職",
    "続柄",
    "株式数",
    "議決権",
    "面積",
    "時価",
    "利用区分",
    "建物",
    "敷地",
    "売上高",
    "構成比",
    "年度",
    "比率",
    "合計",
    "相談",
    "説明",
    "制度",
    "協議会",
)

_HEADER_OR_LABEL_ONLY_WORDS = (
    "企業名称",
    "会社名",
    "法人名",
    "商号",
    "所在地",
    "住所",
    "代表者",
    "代表取締役",
    "取引銀行",
    "金融機関",
    "氏名",
    "株主",
    "株主名",
    "販売先",
    "仕入先",
    "外注先",
    "取引先",
)

_FINANCIAL_STATEMENT_WORDS = (
    "貸借対照表",
    "損益計算書",
    "勘定科目",
    "資産の部",
    "負債の部",
    "純資産",
    "流動資産",
    "固定資産",
    "流動負債",
    "固定負債",
    "売上高",
    "売上原価",
    "販売費",
    "営業利益",
)

_EXPLANATORY_TEXT_WORDS = (
    "制度",
    "説明",
    "相談",
    "協議会",
    "設置",
    "活動",
    "ご案内",
    "お問い合わせ",
    "都道府県",
    "市区町村",
)

_BANK_GUIDANCE_WORDS = (
    "銀行取引",
    "銀行協会",
    "相談室",
    "あっせん",
    "苦情",
    "活動",
)

_STRICT_CORPORATE_MARKER_RE = re.compile(
    r"(株式会社|有限会社|合同会社|[（(]株[）)]?|[（(]有[）)]?|㈱|㈲|株\)|有\))"
)
_FUZZY_CORPORATE_MARKER_RE = re.compile(r"[帳幅帥師梯椎常]")
_ADDRESS_HINT_RE = re.compile(r"(都|道|府|県|市|区|町|村|字|丁目|番|号|-|ー|−|－)")
_FINANCIAL_RE = re.compile(r"(銀行|信用金庫|信金|金融公庫|公庫|信用組合)")
_NUMERIC_HEAVY_RE = re.compile(r"^[0-9０-９,.\-ー−－%％①-⑳\s]+$")


def context_candidates_for_page(
    page_number: int,
    page_rect: Any,
    words: list[tuple[Any, ...]] | None = None,
    enabled_rules: set[str] | frozenset[str] | None = None,
) -> list[OcrCandidate]:
    """Generate generic structure-based OCR candidates.

    This function must remain independent from ground-truth data, reviewed
    rectangles, dataset split files, page-specific branches, and fixed page
    coordinates. The page number is accepted only for API compatibility and is
    not used by the rules below.
    """
    del page_number
    word_boxes = _word_boxes(words or [])
    if not word_boxes:
        return []
    lines = _line_boxes(word_boxes)
    visual_rows = _visual_rows(lines)
    context = _document_context(lines)
    active_rules = DEFAULT_ENABLED_RULES if enabled_rules is None else frozenset(enabled_rules)
    candidates: list[OcrCandidate] = []
    if RULE_LABEL_VALUE in active_rules:
        candidates.extend(_label_value_candidates(visual_rows, page_rect, context))
    if RULE_TABLE_COLUMN in active_rules:
        candidates.extend(_table_column_candidates(visual_rows, page_rect, context))
    if RULE_SECTION_BLOCK in active_rules:
        candidates.extend(_section_block_candidates(lines, page_rect, context))
    if RULE_PATTERN in active_rules:
        candidates.extend(_standalone_pattern_candidates(lines, page_rect, context))
    return _dedupe(candidates)


def _word_boxes(words: list[tuple[Any, ...]]) -> list[_WordBox]:
    boxes: list[_WordBox] = []
    for fallback_index, word in enumerate(words):
        if len(word) < 5:
            continue
        text = str(word[4]).strip()
        if not text:
            continue
        block = int(word[5]) if len(word) > 5 else 0
        line = int(word[6]) if len(word) > 6 else fallback_index
        index = int(word[7]) if len(word) > 7 else fallback_index
        boxes.append(
            _WordBox(
                text=text,
                normalized=_normalize(text),
                x0=float(word[0]),
                y0=float(word[1]),
                x1=float(word[2]),
                y1=float(word[3]),
                block=block,
                line=line,
                index=index,
            )
        )
    return sorted(boxes, key=lambda item: (item.block, item.line, item.index, item.y0, item.x0))


def _line_boxes(words: list[_WordBox]) -> list[_LineBox]:
    groups: dict[tuple[int, int], list[_WordBox]] = {}
    for word in words:
        groups.setdefault((word.block, word.line), []).append(word)
    lines: list[_LineBox] = []
    for (block, line), group in groups.items():
        ordered = tuple(sorted(group, key=lambda item: item.x0))
        text = " ".join(word.text for word in ordered)
        normalized = _normalize("".join(word.normalized for word in ordered))
        lines.append(
            _LineBox(
                words=ordered,
                text=text,
                normalized=normalized,
                rect=(
                    min(word.x0 for word in ordered),
                    min(word.y0 for word in ordered),
                    max(word.x1 for word in ordered),
                    max(word.y1 for word in ordered),
                ),
                block=block,
                line=line,
            )
        )
    return sorted(lines, key=lambda item: (item.y0, item.x0, item.block, item.line))


def _visual_rows(lines: list[_LineBox]) -> list[_LineBox]:
    rows: list[list[_LineBox]] = []
    for line in sorted(lines, key=lambda item: (item.y0, item.x0)):
        placed = False
        for row in rows:
            row_center = median(item.cy for item in row)
            row_height = median(max(item.y1 - item.y0, 1.0) for item in row)
            line_height = max(line.y1 - line.y0, 1.0)
            if abs(line.cy - row_center) <= max(row_height, line_height) * 0.55:
                row.append(line)
                placed = True
                break
        if not placed:
            rows.append([line])
    visual: list[_LineBox] = []
    for row_index, row in enumerate(rows):
        words = tuple(sorted((word for line in row for word in line.words), key=lambda item: item.x0))
        if not words:
            continue
        visual.append(
            _LineBox(
                words=words,
                text=" ".join(word.text for word in words),
                normalized=_normalize("".join(word.normalized for word in words)),
                rect=(
                    min(word.x0 for word in words),
                    min(word.y0 for word in words),
                    max(word.x1 for word in words),
                    max(word.y1 for word in words),
                ),
                block=0,
                line=row_index,
            )
        )
    return visual


def _document_context(lines: list[_LineBox]) -> _DocumentContext:
    text = _normalize("".join(line.normalized for line in lines))
    financial_hits = sum(1 for word in _FINANCIAL_STATEMENT_WORDS if word in text)
    explanation_hits = sum(1 for word in _EXPLANATORY_TEXT_WORDS if word in text)
    bank_guidance_hits = sum(1 for word in _BANK_GUIDANCE_WORDS if word in text)
    return _DocumentContext(
        financial_statement=financial_hits >= 2,
        explanatory_text=explanation_hits >= 2,
        bank_guidance=bank_guidance_hits >= 2,
    )


def _label_value_candidates(lines: list[_LineBox], page_rect: Any, context: _DocumentContext) -> list[OcrCandidate]:
    candidates: list[OcrCandidate] = []
    for line in lines:
        for label, entity_type, replacement, rule_name in _LABEL_RULES:
            if label not in line.normalized:
                continue
            label_words = [word for word in line.words if label in word.normalized or word.normalized in label]
            label_right = max((word.x1 for word in label_words), default=line.x0)
            right_words = [word for word in line.words if word.x0 > label_right + 20]
            value_source = f"{label}ラベルの右側にある値を{entity_type}候補として検出"
            referenced_header = label
            rect_override = None
            join_info = ""
            if not right_words:
                right_words = _value_words_below_label(lines, line, label_words, entity_type)
                value_source = f"{label}ラベルの直下にある値を{entity_type}候補として検出"
                rect_override = _label_below_value_rect(right_words, line, page_rect) if right_words else None
                join_info = _join_info(right_words, "ラベル直下の近接OCR行") if right_words else ""
            if not right_words:
                continue
            if _looks_like_mixed_table_value(right_words):
                continue
            value_words = _trim_value_words(right_words, entity_type)
            if not value_words:
                continue
            if not _value_length_ok(value_words, entity_type):
                continue
            if _should_suppress_candidate(value_words, entity_type, context, has_structure=True):
                continue
            candidate = _candidate_from_words(
                value_words,
                page_rect,
                entity_type,
                replacement,
                rule_name,
                value_source,
                referenced_header=referenced_header,
                rect_override=rect_override,
                confidence=_confidence_for(entity_type, has_structure=True, has_format=_has_entity_format(value_words, entity_type)),
                join_info=join_info,
            )
            if candidate:
                candidates.append(candidate)
    return candidates


def _value_words_below_label(
    lines: list[_LineBox],
    label_line: _LineBox,
    label_words: list[_WordBox],
    entity_type: str,
) -> list[_WordBox]:
    if not label_words:
        return []
    label_left = min(word.x0 for word in label_words)
    label_right = max(word.x1 for word in label_words)
    label_width = max(label_right - label_left, 1.0)
    candidates: list[_WordBox] = []
    for row in lines:
        if row.y0 <= label_line.y1:
            continue
        gap = row.y0 - label_line.y1
        row_height = max(row.y1 - row.y0, 1.0)
        if gap > row_height * 3.0:
            break
        if _section_heading(row) or _is_header_or_label_only(row.normalized):
            break
        if _contains_non_value_context(row.normalized):
            continue
        aligned = abs(row.x0 - label_left) <= max(label_width, row_height * 4.0)
        overlaps = row.x0 <= label_right + label_width and row.x1 >= label_left - label_width
        if not (aligned or overlaps):
            continue
        row_words = _trim_value_words(list(row.words), entity_type)
        if not row_words:
            continue
        value = _normalize("".join(word.normalized for word in row_words))
        if entity_type == "会社名" and not _looks_like_company_text(value):
            continue
        if entity_type == "住所" and not _looks_like_address(value):
            continue
        if entity_type == "氏名" and not _looks_like_person_text(value):
            continue
        if entity_type == "銀行名" and not _looks_like_financial_institution_name(value):
            continue
        candidates.extend(row_words)
        break
    return candidates


def _column_value_rows(
    lines: list[_LineBox],
    header: _LineBox,
    header_word: _WordBox,
    left_boundary: float,
    right_boundary: float,
    entity_type: str,
    page_rect: Any,
) -> list[tuple[_LineBox, list[_WordBox], str]]:
    base_rows = [*_embedded_rows_below_header(header, header_word), *_rows_below(lines, header)]
    results: list[tuple[_LineBox, list[_WordBox], str]] = []
    consumed: set[int] = set()
    row_infos: list[tuple[int, _LineBox, list[_WordBox]]] = []
    for index, row in enumerate(base_rows):
        row_words = [word for word in row.words if left_boundary <= word.cx <= right_boundary]
        if not row_words and entity_type in {"会社名", "住所"}:
            row_words = [word for word in row.words if word.x0 < right_boundary and word.x1 > left_boundary]
        row_infos.append((index, row, row_words))

    for index, row, row_words in row_infos:
        if index in consumed:
            continue
        if not row_words:
            continue
        group_rows = [(row, row_words)]
        consumed.add(index)
        current_rect = _rect_for_words(row_words)
        if entity_type == "氏名":
            combined_row = _line_from_words(row_words, row.line)
            results.append((combined_row, row_words, "単一表行"))
            continue
        for next_index, next_row, next_words in row_infos[index + 1 : index + 4]:
            if next_index in consumed:
                continue
            if not next_words:
                continue
            if _section_heading(next_row):
                break
            gap = next_row.y0 - current_rect[3]
            row_height = max(current_rect[3] - current_rect[1], next_row.y1 - next_row.y0, 1.0)
            if gap < -row_height * 0.5:
                continue
            if gap > row_height * 1.65:
                break
            next_rect = _rect_for_words(next_words)
            horizontal_overlap = _rect_horizontal_overlap_ratio(current_rect, next_rect)
            same_column = horizontal_overlap >= 0.28 or abs(_rect_center_x(current_rect) - _rect_center_x(next_rect)) <= max(
                _rect_width(current_rect), _rect_width(next_rect), 1.0
            ) * 0.55
            combined_words = [word for _row, words in group_rows for word in words] + next_words
            if not same_column:
                continue
            if not _can_merge_table_words(combined_words, entity_type):
                continue
            group_rows.append((next_row, next_words))
            consumed.add(next_index)
            current_rect = _rect_for_words(combined_words)

        group_words = [word for _row, words in group_rows for word in words]
        combined_row = _line_from_words(group_words, row.line)
        reason = "同一列内の近接OCR行を結合" if len(group_rows) > 1 else "単一表行"
        results.append((combined_row, group_words, reason))
    return results


def _table_column_candidates(lines: list[_LineBox], page_rect: Any, context: _DocumentContext) -> list[OcrCandidate]:
    candidates: list[OcrCandidate] = []
    for header in lines:
        headers = _header_words(header)
        if len(headers) < 2:
            continue
        for label, entity_type, replacement, rule_name in _TABLE_HEADER_RULES:
            matching = [word for word in headers if label in word.normalized]
            if not matching:
                continue
            header_word = matching[0]
            left_boundary, right_boundary = _column_boundaries(headers, header_word, page_rect)
            rows = _column_value_rows(lines, header, header_word, left_boundary, right_boundary, entity_type, page_rect)
            for row, row_words, join_reason in rows:
                if not row_words and entity_type in {"会社名", "住所"}:
                    row_words = [word for word in row.words if word.x0 < right_boundary and word.x1 > left_boundary]
                value_words = _trim_value_words(row_words, entity_type)
                if not value_words:
                    continue
                if not _value_length_ok(value_words, entity_type):
                    continue
                if _should_suppress_candidate(value_words, entity_type, context, has_structure=True):
                    continue
                rect_override = _bounded_value_rect(value_words, left_boundary, right_boundary, row, page_rect)
                candidate = _candidate_from_words(
                    value_words,
                    page_rect,
                    entity_type,
                    replacement,
                    rule_name,
                    f"{label}列の下にある行値を{entity_type}候補として検出",
                    referenced_header=label,
                    rect_override=rect_override,
                    confidence=_confidence_for(entity_type, has_structure=True, has_format=_has_entity_format(value_words, entity_type)),
                    join_info=_join_info(value_words, join_reason),
                )
                if candidate:
                    candidates.append(candidate)
    return candidates


def _fragmented_section_candidates(
    lines: list[_LineBox],
    heading_line: _LineBox,
    active_heading: str,
    zone: tuple[float, float],
    active_until_y: float,
    page_rect: Any,
    context: _DocumentContext,
) -> list[OcrCandidate]:
    if active_heading in {"役員構成", "株主一覧"}:
        entity_type = "氏名"
        replacement = "個人001"
        rule_name = "section_block:person_list"
    else:
        entity_type = "会社名"
        replacement = "法人001"
        rule_name = "section_block:business_partner"

    zone_lines = [
        line
        for line in lines
        if line.y0 > heading_line.y1 + 4
        and line.y0 <= active_until_y
        and not _section_heading(line)
        and _line_in_zone(line, zone)
        and not _is_header_or_label_only(line.normalized)
        and not _contains_non_value_context(line.normalized)
    ]
    candidates: list[OcrCandidate] = []
    used: set[int] = set()
    for index, line in enumerate(zone_lines):
        if index in used:
            continue
        if entity_type == "会社名" and not _fragment_can_start_business(line):
            continue
        if entity_type == "氏名" and not _fragment_can_start_person(line):
            continue
        group = [line]
        used.add(index)
        current_rect = line.rect
        for next_index, next_line in enumerate(zone_lines[index + 1 : index + 4], start=index + 1):
            if next_index in used:
                continue
            gap = next_line.y0 - current_rect[3]
            height = max(_rect_height(current_rect), next_line.y1 - next_line.y0, 1.0)
            if gap < -height * 0.4:
                continue
            if gap > height * 1.45:
                break
            if not _fragment_lines_belong_together(group[-1], next_line):
                continue
            combined_words = [word for item in group for word in item.words] + list(next_line.words)
            if not _can_merge_section_words(combined_words, entity_type):
                continue
            group.append(next_line)
            used.add(next_index)
            current_rect = _rect_for_words(combined_words)

        if len(group) < 2:
            continue
        value_words = _trim_value_words([word for item in group for word in item.words], entity_type)
        if not value_words:
            continue
        if not _value_length_ok(value_words, entity_type):
            continue
        if _should_suppress_candidate(value_words, entity_type, context, has_structure=True):
            continue
        if entity_type == "会社名" and not _looks_like_company_text(_normalize("".join(word.normalized for word in value_words))):
            continue
        if entity_type == "氏名" and not _looks_like_person_text(_normalize("".join(word.normalized for word in value_words))):
            continue
        candidate = _candidate_from_words(
            value_words,
            page_rect,
            entity_type,
            replacement,
            rule_name,
            f"{active_heading}見出し配下で複数OCRブロックに分割された値を{entity_type}候補として検出",
            referenced_header=active_heading,
            rect_override=_section_value_rect(value_words, zone, page_rect, entity_type),
            confidence=_confidence_for(entity_type, has_structure=True, has_format=_has_entity_format(value_words, entity_type)),
            join_info=_join_info(value_words, "同一見出しゾーン内の近接ブロックを結合"),
        )
        if candidate:
            candidates.append(candidate)
    return candidates


def _section_block_candidates(lines: list[_LineBox], page_rect: Any, context: _DocumentContext) -> list[OcrCandidate]:
    candidates: list[OcrCandidate] = []
    heading_lines = [(line, _section_heading(line)) for line in lines if _section_heading(line)]
    for heading_line, active_heading in heading_lines:
        zone = _section_zone(heading_line, [line for line, _heading in heading_lines], page_rect)
        active_until_y = heading_line.y1 + 260
        for line in lines:
            if line.y0 <= heading_line.y1 + 4 or line.y0 > active_until_y:
                continue
            if _section_heading(line):
                continue
            if not _line_in_zone(line, zone):
                continue
            if _is_header_or_label_only(line.normalized):
                continue
            if _contains_non_value_context(line.normalized):
                continue
            if active_heading in {"役員構成", "株主一覧"}:
                entity_type = "氏名"
                replacement = "個人001"
                rule_name = "section_block:person_list"
            else:
                entity_type = "会社名"
                replacement = "法人001"
                rule_name = "section_block:business_partner"
            value_words = _trim_value_words(list(line.words), entity_type)
            if not value_words:
                continue
            if not _value_length_ok(value_words, entity_type):
                continue
            if _should_suppress_candidate(value_words, entity_type, context, has_structure=True):
                continue
            if entity_type == "会社名" and not _looks_like_section_business_candidate(line):
                continue
            if entity_type == "氏名" and not _looks_like_person_row(line):
                continue
            candidate = _candidate_from_words(
                value_words,
                page_rect,
                entity_type,
                replacement,
                rule_name,
                f"{active_heading}見出し配下の文字ブロックを{entity_type}候補として検出",
                referenced_header=active_heading,
                rect_override=_section_value_rect(value_words, zone, page_rect, entity_type),
                confidence=_confidence_for(entity_type, has_structure=True, has_format=_has_entity_format(value_words, entity_type)),
                join_info=_join_info(value_words, "単一OCR行"),
            )
            if candidate:
                candidates.append(candidate)
        candidates.extend(_fragmented_section_candidates(lines, heading_line, active_heading, zone, active_until_y, page_rect, context))
    return candidates


def _standalone_pattern_candidates(lines: list[_LineBox], page_rect: Any, context: _DocumentContext) -> list[OcrCandidate]:
    candidates: list[OcrCandidate] = []
    for line in lines:
        if _is_noise_or_summary(line.normalized):
            continue
        if _is_header_or_label_only(line.normalized):
            continue
        if context.financial_statement and not _has_explicit_value_context(line.normalized):
            continue
        if _has_corporate_marker(line.normalized) and _looks_like_business_name_line(line):
            value_words = _trim_value_words(list(line.words), "会社名")
            if not _value_length_ok(value_words, "会社名"):
                continue
            if _should_suppress_candidate(value_words, "会社名", context, has_structure=False):
                continue
            candidate = _candidate_from_words(
                value_words,
                page_rect,
                "会社名",
                "法人001",
                "pattern:corporate_marker",
                "法人表記またはそのOCR揺れを含む文字ブロックを会社名候補として検出",
                referenced_header="法人表記",
                confidence=_confidence_for("会社名", has_structure=False, has_format=True),
            )
            if candidate:
                candidates.append(candidate)
        elif _looks_like_standalone_address(line.normalized):
            value_words = _trim_value_words(list(line.words), "住所")
            if not _value_length_ok(value_words, "住所"):
                continue
            if _should_suppress_candidate(value_words, "住所", context, has_structure=False):
                continue
            candidate = _candidate_from_words(
                value_words,
                page_rect,
                "住所",
                "市区町村まで",
                "pattern:address",
                "都道府県、市区町村、字、番地などの住所構成を含む文字ブロックを住所候補として検出",
                referenced_header="住所構成",
                confidence=_confidence_for("住所", has_structure=False, has_format=True),
            )
            if candidate:
                candidates.append(candidate)
        elif _looks_like_financial_institution_name(line.normalized):
            value_words = _trim_value_words(list(line.words), "銀行名")
            if not _value_length_ok(value_words, "銀行名"):
                continue
            if _should_suppress_candidate(value_words, "銀行名", context, has_structure=False):
                continue
            candidate = _candidate_from_words(
                value_words,
                page_rect,
                "銀行名",
                "金融機関001",
                "pattern:financial_institution",
                "金融機関を示す一般語を含む文字ブロックを銀行名候補として検出",
                referenced_header="金融機関表記",
                confidence=_confidence_for("銀行名", has_structure=False, has_format=True),
            )
            if candidate:
                candidates.append(candidate)
    return candidates


def _header_words(line: _LineBox) -> list[_WordBox]:
    headers: list[_WordBox] = []
    for word in line.words:
        if any(label in word.normalized for label, *_rest in _TABLE_HEADER_RULES):
            headers.append(word)
        elif re.search(r"(No|種類|区分|面積|評価|金額|比率|年度|年.*月.*期|株式数|議決権)", word.normalized, re.IGNORECASE):
            headers.append(word)
    return sorted(headers, key=lambda item: item.cx)


def _column_boundaries(headers: list[_WordBox], header: _WordBox, page_rect: Any) -> tuple[float, float]:
    ordered = sorted(headers, key=lambda item: item.cx)
    index = ordered.index(header)
    previous_right = max(float(page_rect.x0), header.x0 - max(header.x1 - header.x0, 40) * 2.6)
    next_left = float(page_rect.x1)
    if index > 0:
        previous_right = ordered[index - 1].x1
    if index + 1 < len(ordered):
        next_left = ordered[index + 1].x0
    left = max(float(page_rect.x0), previous_right)
    right = min(float(page_rect.x1), next_left)
    if right - left < max(header.x1 - header.x0, 30):
        padding = max(header.x1 - header.x0, 40)
        left = max(float(page_rect.x0), header.cx - padding * 1.8)
        right = min(float(page_rect.x1), header.cx + padding * 1.8)
    return left, right


def _embedded_rows_below_header(header_line: _LineBox, header_word: _WordBox) -> list[_LineBox]:
    header_height = max(header_word.y1 - header_word.y0, 1.0)
    embedded_words = [
        word for word in header_line.words if word.cy > header_word.cy + header_height * 1.3
    ]
    if not embedded_words:
        return []
    return _visual_rows(_line_boxes(embedded_words))


def _rows_below(lines: list[_LineBox], header: _LineBox) -> list[_LineBox]:
    rows: list[_LineBox] = []
    header_height = max(header.y1 - header.y0, 1.0)
    for line in lines:
        if line is header:
            continue
        if line.cy <= header.cy + header_height * 0.40:
            continue
        if line.y0 - header.y1 > 380:
            break
        if _section_heading(line) and line.y0 > header.y1 + 25:
            break
        rows.append(line)
    return rows


def _trim_value_words(words: list[_WordBox], entity_type: str) -> list[_WordBox]:
    words = [word for word in words if _normalize(word.text)]
    if not words:
        return []
    if entity_type in {"会社名", "銀行名"}:
        words = [word for word in words if not _is_numeric_token(word.normalized)]
        while words and _is_leading_marker(words[0].normalized):
            words = words[1:]
        filtered: list[_WordBox] = []
        for word in words:
            if _is_non_target_token(word.normalized):
                break
            filtered.append(word)
        words = filtered
    elif entity_type == "住所":
        words = [word for word in words if not _is_non_target_token(word.normalized)]
    elif entity_type == "氏名":
        words = [word for word in words if not _is_non_target_token(word.normalized) and not _is_numeric_token(word.normalized)]
    if not words:
        return []
    normalized = _normalize("".join(word.normalized for word in words))
    if _is_noise_or_summary(normalized):
        return []
    if _is_header_or_label_only(normalized):
        return []
    if _contains_non_value_context(normalized):
        return []
    if entity_type == "会社名" and not (_has_corporate_marker(normalized) or _looks_like_company_text(normalized)):
        return []
    if entity_type == "住所" and not _looks_like_address(normalized):
        return []
    if entity_type == "銀行名" and not _FINANCIAL_RE.search(normalized):
        return []
    if entity_type == "氏名" and not _looks_like_person_text(normalized):
        return []
    return words


def _bounded_value_rect(
    words: list[_WordBox],
    left_boundary: float,
    right_boundary: float,
    row: _LineBox,
    page_rect: Any,
) -> tuple[float, float, float, float]:
    rect = _rect_for_words(words)
    height = max(rect[3] - rect[1], row.y1 - row.y0, 1.0)
    x_margin = height * 0.35
    y_margin = height * 0.20
    return _safe_rect(
        (
            max(left_boundary, rect[0] - x_margin),
            min(row.y0, rect[1]) - y_margin,
            min(right_boundary, rect[2] + x_margin),
            max(row.y1, rect[3]) + y_margin,
        ),
        page_rect,
    )


def _label_below_value_rect(
    words: list[_WordBox],
    label_line: _LineBox,
    page_rect: Any,
) -> tuple[float, float, float, float]:
    rect = _rect_for_words(words)
    height = max(_rect_height(rect), label_line.y1 - label_line.y0, 1.0)
    return _safe_rect(
        (
            rect[0] - height * 0.35,
            rect[1] - height * 0.20,
            rect[2] + height * 0.35,
            rect[3] + height * 0.20,
        ),
        page_rect,
    )


def _section_value_rect(
    words: list[_WordBox],
    zone: tuple[float, float],
    page_rect: Any,
    entity_type: str,
) -> tuple[float, float, float, float]:
    rect = _rect_for_words(words)
    height = _rect_height(rect)
    if entity_type == "会社名":
        x_margin = height * 0.80
        y_margin = height * 0.60
    elif entity_type == "氏名":
        x_margin = height * 0.75
        y_margin = height * 0.24
    else:
        x_margin = height * 0.25
        y_margin = height * 0.18
    left, right = zone
    return _safe_rect(
        (
            max(left, rect[0] - x_margin),
            rect[1] - y_margin,
            min(right, rect[2] + x_margin),
            rect[3] + y_margin,
        ),
        page_rect,
    )


_LIGATURE_KANJI_RANGE = "一-龠々"
_LIGATURE_KANA_RANGE = "ァ-ヶーぁ-んー"
_LIGATURE_PREFIX_RE = re.compile(rf"^[{_LIGATURE_KANJI_RANGE}][{_LIGATURE_KANA_RANGE}]{{2,}}")
_LIGATURE_SUFFIX_RE = re.compile(rf"[{_LIGATURE_KANA_RANGE}]{{2,}}[{_LIGATURE_KANJI_RANGE}]$")
_LIGATURE_STRIP_CHARS = "・-*•. 　"

LIGATURE_OCR_NOTE = (
    "先頭または末尾の1文字はOCRが「㈱」等の合字を別の漢字に誤認識した可能性があります。原本で確認してください"
)


def _possible_ligature_note(original: str, entity_type: str) -> str:
    """会社名候補で、単独の漢字1文字がカタカナ/ひらがな連続に隣接している場合に注記を付ける。

    Tesseract は「㈱」のような囲み文字を安定して読めず、DPIを変えても
    別の漢字(師/梯/常/犀/帳など)にランダムに化けるため、テキストを推測補正
    せずに「要確認」の理由欄で注意喚起するだけに留める。
    """
    if entity_type != "会社名":
        return ""
    stripped = original.strip(_LIGATURE_STRIP_CHARS)
    if not stripped:
        return ""
    if _LIGATURE_PREFIX_RE.search(stripped) or _LIGATURE_SUFFIX_RE.search(stripped):
        return LIGATURE_OCR_NOTE
    return ""


def _join_info(words: list[_WordBox], reason: str) -> str:
    blocks = {word.block for word in words}
    lines = {(word.block, word.line) for word in words}
    normalized = _normalize("".join(word.normalized for word in words))
    rect = _rect_for_words(words)
    rect_text = ",".join(f"{value:.1f}" for value in rect)
    return f"結合情報: 元ブロック数={len(blocks)} / 元行数={len(lines)} / 結合理由={reason} / 結合後文字列={normalized} / 結合矩形={rect_text}"


def _candidate_from_words(
    words: Iterable[_WordBox],
    page_rect: Any,
    entity_type: str,
    replacement: str,
    rule_name: str,
    reason: str,
    referenced_header: str,
    rect_override: tuple[float, float, float, float] | None = None,
    confidence: str = "MEDIUM",
    join_info: str = "",
) -> OcrCandidate | None:
    word_list = list(words)
    if not word_list:
        return None
    original = " ".join(word.text for word in word_list).strip()
    normalized = _normalize("".join(word.normalized for word in word_list))
    if not normalized:
        return None
    rect = rect_override or _rect_for_words(word_list)
    rect = _expand_rect(rect, entity_type)
    rect = _safe_rect(rect, page_rect)
    if rect[2] - rect[0] <= 1 or rect[3] - rect[1] <= 1:
        return None
    ligature_note = _possible_ligature_note(original, entity_type)
    return OcrCandidate(
        original=original,
        normalized=normalized,
        entity_type=entity_type,
        status=CANDIDATE_REVIEW,
        replacement=replacement,
        rect=rect,
        reason=f"{rule_name}: {reason}。参照見出し/列={referenced_header}。信頼度={confidence}"
        + (f"。{join_info}" if join_info else "")
        + (f"。{ligature_note}" if ligature_note else ""),
    )


def _rect_for_words(words: Iterable[_WordBox]) -> tuple[float, float, float, float]:
    word_list = list(words)
    return (
        min(word.x0 for word in word_list),
        min(word.y0 for word in word_list),
        max(word.x1 for word in word_list),
        max(word.y1 for word in word_list),
    )


def _line_from_words(words: list[_WordBox], line_index: int = 0) -> _LineBox:
    ordered = tuple(sorted(words, key=lambda item: (item.y0, item.x0)))
    return _LineBox(
        words=ordered,
        text=" ".join(word.text for word in ordered),
        normalized=_normalize("".join(word.normalized for word in ordered)),
        rect=_rect_for_words(ordered),
        block=ordered[0].block if ordered else 0,
        line=line_index,
    )


def _expand_rect(rect: tuple[float, float, float, float], entity_type: str) -> tuple[float, float, float, float]:
    x0, y0, x1, y1 = rect
    height = max(y1 - y0, 1.0)
    if entity_type == "氏名":
        x_padding = height * 0.85
        y_padding = height * 0.12
    elif entity_type == "住所":
        x_padding = height * 0.30
        y_padding = height * 0.16
    elif entity_type == "会社名":
        x_padding = height * 0.20
        y_padding = height * 0.18
    else:
        x_padding = height * 0.15
        y_padding = height * 0.12
    return (x0 - x_padding, y0 - y_padding, x1 + x_padding, y1 + y_padding)


def _clip_rect(
    rect: tuple[float, float, float, float],
    bounds: tuple[float, float, float, float],
    page_rect: Any,
) -> tuple[float, float, float, float]:
    x0, y0, x1, y1 = rect
    bx0, by0, bx1, by1 = bounds
    return _safe_rect((max(x0, bx0), max(y0, by0), min(x1, bx1), min(y1, by1)), page_rect)


def _safe_rect(rect: tuple[float, float, float, float], page_rect: Any) -> tuple[float, float, float, float]:
    x0, y0, x1, y1 = rect
    return (
        max(float(page_rect.x0), min(x0, float(page_rect.x1))),
        max(float(page_rect.y0), min(y0, float(page_rect.y1))),
        max(float(page_rect.x0), min(x1, float(page_rect.x1))),
        max(float(page_rect.y0), min(y1, float(page_rect.y1))),
    )


def _section_heading(line: _LineBox) -> str:
    for heading in _SECTION_HEADERS:
        if heading in line.normalized:
            return heading
    return ""


def _section_zone(
    heading: _LineBox,
    headings: list[_LineBox],
    page_rect: Any,
) -> tuple[float, float]:
    same_band = [
        item
        for item in headings
        if item is not heading and abs(item.cy - heading.cy) <= max(heading.y1 - heading.y0, 1.0) * 3.0
    ]
    left = float(page_rect.x0)
    right = float(page_rect.x1)
    left_neighbors = [item for item in same_band if item.cx < heading.cx]
    right_neighbors = [item for item in same_band if item.cx > heading.cx]
    if left_neighbors:
        left = (max(item.cx for item in left_neighbors) + heading.cx) / 2
    else:
        left = max(float(page_rect.x0), heading.cx - 150)
    if right_neighbors:
        right = (min(item.cx for item in right_neighbors) + heading.cx) / 2
    else:
        right = min(float(page_rect.x1), heading.cx + 150)
    return left, right


def _line_in_zone(line: _LineBox, zone: tuple[float, float]) -> bool:
    left, right = zone
    overlap = min(line.x1, right) - max(line.x0, left)
    if overlap <= 0:
        return False
    return overlap / max(line.x1 - line.x0, 1.0) >= 0.45 or left <= line.cx <= right


def _can_merge_table_words(words: list[_WordBox], entity_type: str) -> bool:
    value = _normalize("".join(word.normalized for word in words))
    if len(words) > 8:
        return False
    if entity_type == "氏名":
        return 2 <= sum(1 for char in value if _is_japanese(char)) <= 10 and not _contains_non_value_context(value)
    if entity_type == "住所":
        return len(value) <= 80 and _looks_like_address(value)
    if entity_type == "会社名":
        return len(value) <= 36 and _looks_like_company_text(value)
    return len(value) <= 60


def _can_merge_section_words(words: list[_WordBox], entity_type: str) -> bool:
    value = _normalize("".join(word.normalized for word in words))
    if len(words) > 7 or _contains_non_value_context(value) or _is_noise_or_summary(value):
        return False
    if entity_type == "会社名":
        return 4 <= len(value) <= 30 and _looks_like_company_text(value)
    if entity_type == "氏名":
        return 2 <= len(value) <= 10 and _looks_like_person_text(value)
    return False


def _fragment_lines_belong_together(previous: _LineBox, current: _LineBox) -> bool:
    previous_rect = previous.rect
    current_rect = current.rect
    horizontal_overlap = _rect_horizontal_overlap_ratio(previous_rect, current_rect)
    center_distance = abs(_rect_center_x(previous_rect) - _rect_center_x(current_rect))
    max_width = max(_rect_width(previous_rect), _rect_width(current_rect), 1.0)
    return horizontal_overlap >= 0.18 or center_distance <= max_width * 0.75


def _fragment_can_start_business(line: _LineBox) -> bool:
    value = line.normalized
    if _is_noise_or_summary(value) or _contains_non_value_context(value):
        return False
    if any(word in value for word in _BUSINESS_PROCESS_WORDS):
        return False
    japanese = sum(1 for char in value if _is_japanese(char))
    return 1 <= len(line.words) <= 4 and 2 <= japanese <= 18 and not _NUMERIC_HEAVY_RE.match(value)


def _fragment_can_start_person(line: _LineBox) -> bool:
    value = line.normalized
    if _is_noise_or_summary(value) or _contains_non_value_context(value):
        return False
    japanese = sum(1 for char in value if _is_japanese(char))
    return 1 <= len(line.words) <= 4 and 1 <= japanese <= 8 and not _NUMERIC_HEAVY_RE.match(value)


def _has_corporate_marker(value: str) -> bool:
    return bool(_STRICT_CORPORATE_MARKER_RE.search(value))


def _has_fuzzy_corporate_marker(value: str) -> bool:
    return bool(_FUZZY_CORPORATE_MARKER_RE.search(value))


def _looks_like_business_name_line(line: _LineBox) -> bool:
    text = line.normalized
    if _is_noise_or_summary(text):
        return False
    if _has_corporate_marker(text):
        return True
    japanese = sum(1 for char in text if _is_japanese(char))
    return japanese >= 3 and len(text) <= 28 and not _NUMERIC_HEAVY_RE.match(text)


def _looks_like_section_business_candidate(line: _LineBox) -> bool:
    text = line.normalized
    if _looks_like_ocr_artifact_business(text):
        return False
    if _has_corporate_marker(text):
        return True
    if any(word in text for word in _BUSINESS_PROCESS_WORDS):
        return False
    if _has_fuzzy_corporate_marker(text) and 4 <= len(text) <= 24:
        return True
    bullet_or_short_block = line.text.strip().startswith(("・", "●", "-")) or len(line.words) <= 3
    japanese = sum(1 for char in text if _is_japanese(char))
    return bullet_or_short_block and 4 <= japanese <= 18 and not _NUMERIC_HEAVY_RE.match(text)


def _looks_like_company_text(value: str) -> bool:
    if _is_noise_or_summary(value):
        return False
    if _looks_like_ocr_artifact_business(value):
        return False
    if _has_corporate_marker(value):
        return True
    if _has_fuzzy_corporate_marker(value) and 4 <= len(value) <= 24:
        return True
    japanese = sum(1 for char in value if _is_japanese(char))
    return 3 <= japanese <= 24 and not _NUMERIC_HEAVY_RE.match(value) and not _looks_like_general_sentence(value)


def _looks_like_address(value: str) -> bool:
    if _is_noise_or_summary(value):
        return False
    if _looks_like_general_sentence(value):
        return False
    return bool(_ADDRESS_HINT_RE.search(value)) and sum(1 for char in value if _is_japanese(char)) >= 3


def _looks_like_standalone_address(value: str) -> bool:
    if not _looks_like_address(value):
        return False
    if any(token in value for token in ("設置", "制度", "説明", "協議会", "相談", "活動")):
        return False
    return bool(re.search(r"(東京都|北海道|京都府|大阪府|[一-龯]{2,3}県)", value)) and bool(re.search(r"(市|区|町|村|字|丁目|番|号)", value))


def _looks_like_financial_institution_name(value: str) -> bool:
    if not _FINANCIAL_RE.search(value):
        return False
    if value in {"取引銀行", "金融機関", "銀行"}:
        return False
    if any(token in value for token in ("相談", "案内", "活動", "取引", "協会", "あっせん", "苦情", "お受け", "知りたい", "委員会")):
        return False
    return bool(re.search(r"[一-龯ぁ-んァ-ヶA-Za-z]{2,}(銀行|信用金庫|信金|信用組合|農業協同組合|金融公庫|公庫)", value))


def _looks_like_general_sentence(value: str) -> bool:
    if len(value) < 18:
        return False
    sentence_markers = ("です", "ます", "された", "している", "について", "に関する", "の場合", "ため", "こと", "もの", "こちら", "ください")
    return any(marker in value for marker in sentence_markers)


def _looks_like_ocr_artifact_business(value: str) -> bool:
    if _has_corporate_marker(value) or _has_fuzzy_corporate_marker(value):
        return False
    japanese = sum(1 for char in value if _is_japanese(char))
    ascii_alnum = sum(1 for char in value if char.isascii() and char.isalnum())
    symbols = sum(1 for char in value if not char.isalnum() and not _is_japanese(char))
    if japanese < 3 and ascii_alnum + symbols >= japanese + 2:
        return True
    if symbols >= 4 and japanese < 5:
        return True
    return False


def _has_explicit_value_context(value: str) -> bool:
    return any(token in value for token in ("取引銀行", "金融機関", "所在地", "住所", "販売先", "仕入先", "外注先", "会社名", "法人名"))


def _looks_like_person_row(line: _LineBox) -> bool:
    value = line.normalized
    return _looks_like_person_text(value) or ("役職" in value and sum(1 for char in value if _is_japanese(char)) >= 2)


def _looks_like_person_text(value: str) -> bool:
    if _is_noise_or_summary(value):
        return False
    if any(token in value for token in ("会社", "銀行", "所在地", "販売", "仕入", "外注", "面積", "比率")):
        return False
    japanese = sum(1 for char in value if _is_japanese(char))
    return 2 <= japanese <= 8 and not _NUMERIC_HEAVY_RE.match(value)


def _has_entity_format(words: list[_WordBox], entity_type: str) -> bool:
    value = _normalize("".join(word.normalized for word in words))
    if entity_type == "会社名":
        return _has_corporate_marker(value) or _has_fuzzy_corporate_marker(value)
    if entity_type == "住所":
        return _looks_like_address(value)
    if entity_type == "銀行名":
        return _looks_like_financial_institution_name(value) or bool(_FINANCIAL_RE.search(value))
    if entity_type == "氏名":
        return _looks_like_person_text(value)
    return False


def _should_suppress_candidate(
    words: list[_WordBox],
    entity_type: str,
    context: _DocumentContext,
    *,
    has_structure: bool,
) -> bool:
    value = _normalize("".join(word.normalized for word in words))
    if not value:
        return True
    if entity_type == "会社名":
        if context.financial_statement and not has_structure and not _has_corporate_marker(value):
            return True
        if _looks_like_accounting_term(value):
            return True
        if _looks_like_general_sentence(value) and not _has_corporate_marker(value):
            return True
    if entity_type == "住所":
        if context.explanatory_text and not has_structure:
            return True
        if any(token in value for token in ("制度", "協議会", "設置", "説明", "相談")) and not re.search(r"(丁目|番|号|字|[0-9０-９]-)", value):
            return True
    if entity_type == "銀行名":
        if not has_structure and not _looks_like_financial_institution_name(value):
            return True
        if context.bank_guidance and any(token in value for token in ("相談", "案内", "協会", "取引", "活動", "あっせん")):
            return True
    return False


def _looks_like_accounting_term(value: str) -> bool:
    accounting = (
        "現金",
        "預金",
        "売掛金",
        "棚卸資産",
        "有形固定資産",
        "無形固定資産",
        "投資",
        "買掛金",
        "借入金",
        "資本金",
        "利益剰余金",
        "売上高",
        "売上原価",
        "営業利益",
        "経常利益",
        "当期純利益",
    )
    return any(term in value for term in accounting)


def _confidence_for(entity_type: str, *, has_structure: bool, has_format: bool) -> str:
    if has_structure and has_format:
        return "HIGH"
    if has_structure:
        return "MEDIUM"
    if has_format and entity_type in {"住所", "銀行名", "会社名"}:
        return "LOW"
    return "LOW"


def _value_length_ok(words: list[_WordBox], entity_type: str) -> bool:
    if not words:
        return False
    value = _normalize("".join(word.normalized for word in words))
    if entity_type == "氏名":
        return 2 <= len(value) <= 10
    if entity_type == "会社名":
        return 4 <= len(value) <= 36
    if entity_type == "銀行名":
        return 3 <= len(value) <= 40
    if entity_type == "住所":
        return 6 <= len(value) <= 80
    return True


def _looks_like_mixed_table_value(words: list[_WordBox]) -> bool:
    value = _normalize("".join(word.normalized for word in words))
    if len(value) > 90:
        return True
    context_hits = sum(1 for token in _NON_VALUE_CONTEXT_WORDS if token in value)
    return context_hits >= 2


def _is_header_or_label_only(value: str) -> bool:
    if not value:
        return True
    stripped = value.strip(" :：|/・-ー−－")
    if stripped in _HEADER_OR_LABEL_ONLY_WORDS:
        return True
    if len(stripped) <= 6 and any(word == stripped for word in _NON_VALUE_CONTEXT_WORDS):
        return True
    return False


def _contains_non_value_context(value: str) -> bool:
    context_hits = sum(1 for token in _NON_VALUE_CONTEXT_WORDS if token in value)
    return context_hits >= 2


def _is_noise_or_summary(value: str) -> bool:
    if not value:
        return True
    if any(word in value for word in _NON_TARGET_SUMMARY_WORDS):
        return True
    if _NUMERIC_HEAVY_RE.match(value):
        return True
    alnum = sum(1 for char in value if char.isascii() and char.isalnum())
    japanese = sum(1 for char in value if _is_japanese(char))
    return alnum > 12 and japanese == 0


def _is_non_target_token(value: str) -> bool:
    return not value or any(word in value for word in _NON_TARGET_SUMMARY_WORDS)


def _is_numeric_token(value: str) -> bool:
    return bool(_NUMERIC_HEAVY_RE.match(value))


def _is_leading_marker(value: str) -> bool:
    return value in {"・", "●", "■", "□", "団", "園", "圖", "と", ">", "|", "/", "\\"}


def _normalize(value: str) -> str:
    normalized = unicodedata.normalize("NFKC", value)
    return re.sub(r"\s+", "", normalized)


def _is_japanese(char: str) -> bool:
    return (
        "\u3040" <= char <= "\u30ff"
        or "\u3400" <= char <= "\u9fff"
        or "\uf900" <= char <= "\ufaff"
    )


def _dedupe(candidates: list[OcrCandidate]) -> list[OcrCandidate]:
    deduped: list[OcrCandidate] = []
    seen: set[tuple[str, str, int, int, int, int]] = set()
    for candidate in candidates:
        x0, y0, x1, y1 = candidate.rect
        key = (
            candidate.entity_type,
            candidate.normalized,
            round(x0),
            round(y0),
            round(x1),
            round(y1),
        )
        if key in seen:
            continue
        seen.add(key)
        duplicate_index = _find_overlapping_duplicate(deduped, candidate)
        if duplicate_index is not None:
            if _candidate_priority(candidate) > _candidate_priority(deduped[duplicate_index]):
                deduped[duplicate_index] = candidate
            continue
        deduped.append(candidate)
    return deduped


def _find_overlapping_duplicate(candidates: list[OcrCandidate], candidate: OcrCandidate) -> int | None:
    for index, existing in enumerate(candidates):
        if existing.entity_type != candidate.entity_type:
            continue
        if _rect_overlap_over_smaller(existing.rect, candidate.rect) < 0.78:
            continue
        existing_text = existing.normalized
        candidate_text = candidate.normalized
        if existing_text in candidate_text or candidate_text in existing_text:
            return index
        if _rect_center_distance(existing.rect, candidate.rect) <= max(_rect_height(existing.rect), _rect_height(candidate.rect)):
            return index
    return None


def _candidate_priority(candidate: OcrCandidate) -> tuple[int, float]:
    reason = candidate.reason
    if reason.startswith(RULE_LABEL_VALUE):
        rule_rank = 4
    elif reason.startswith(RULE_TABLE_COLUMN):
        rule_rank = 3
    elif reason.startswith(RULE_PATTERN):
        rule_rank = 2
    elif reason.startswith(RULE_SECTION_BLOCK):
        rule_rank = 1
    else:
        rule_rank = 0
    return rule_rank, -_rect_area(candidate.rect)


def _rect_overlap_over_smaller(
    first: tuple[float, float, float, float],
    second: tuple[float, float, float, float],
) -> float:
    x0 = max(first[0], second[0])
    y0 = max(first[1], second[1])
    x1 = min(first[2], second[2])
    y1 = min(first[3], second[3])
    overlap = max(x1 - x0, 0.0) * max(y1 - y0, 0.0)
    smaller = min(_rect_area(first), _rect_area(second))
    return overlap / max(smaller, 1.0)


def _rect_area(rect: tuple[float, float, float, float]) -> float:
    return max(rect[2] - rect[0], 0.0) * max(rect[3] - rect[1], 0.0)


def _rect_height(rect: tuple[float, float, float, float]) -> float:
    return max(rect[3] - rect[1], 1.0)


def _rect_width(rect: tuple[float, float, float, float]) -> float:
    return max(rect[2] - rect[0], 1.0)


def _rect_center_x(rect: tuple[float, float, float, float]) -> float:
    return (rect[0] + rect[2]) / 2


def _rect_horizontal_overlap_ratio(
    first: tuple[float, float, float, float],
    second: tuple[float, float, float, float],
) -> float:
    overlap = max(min(first[2], second[2]) - max(first[0], second[0]), 0.0)
    return overlap / max(min(_rect_width(first), _rect_width(second)), 1.0)


def _rect_center_distance(
    first: tuple[float, float, float, float],
    second: tuple[float, float, float, float],
) -> float:
    first_cx = (first[0] + first[2]) / 2
    first_cy = (first[1] + first[3]) / 2
    second_cx = (second[0] + second[2]) / 2
    second_cy = (second[1] + second[3]) / 2
    return ((first_cx - second_cx) ** 2 + (first_cy - second_cy) ** 2) ** 0.5
