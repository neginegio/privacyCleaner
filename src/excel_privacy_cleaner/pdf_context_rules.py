from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass
from typing import Any, Iterable

from .pdf_ocr_support import CANDIDATE_REVIEW, OcrCandidate


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

_CORPORATE_MARKER_RE = re.compile(
    r"(株式会社|有限会社|合同会社|[（(]株[）)]?|[（(]有[）)]?|㈱|㈲|株\)|有\)|[帳幅帥師梯椎常][一-龯ぁ-んァ-ヶーA-Za-z0-9]{2,})"
)
_ADDRESS_HINT_RE = re.compile(r"(都|道|府|県|市|区|町|村|字|丁目|番|号|-|ー|−|－)")
_FINANCIAL_RE = re.compile(r"(銀行|信用金庫|信金|金融公庫|公庫|信用組合)")
_NUMERIC_HEAVY_RE = re.compile(r"^[0-9０-９,.\-ー−－%％①-⑳\s]+$")


def context_candidates_for_page(
    page_number: int,
    page_rect: Any,
    words: list[tuple[Any, ...]] | None = None,
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
    candidates: list[OcrCandidate] = []
    candidates.extend(_label_value_candidates(visual_rows, page_rect))
    candidates.extend(_table_column_candidates(visual_rows, page_rect))
    candidates.extend(_section_block_candidates(lines, page_rect))
    candidates.extend(_standalone_pattern_candidates(lines, page_rect))
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
            row_y0 = min(item.y0 for item in row)
            row_y1 = max(item.y1 for item in row)
            median_height = max((row_y1 - row_y0), 1.0)
            overlaps = min(row_y1, line.y1) - max(row_y0, line.y0)
            if overlaps >= -median_height * 0.35:
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


def _label_value_candidates(lines: list[_LineBox], page_rect: Any) -> list[OcrCandidate]:
    candidates: list[OcrCandidate] = []
    for line in lines:
        for label, entity_type, replacement, rule_name in _LABEL_RULES:
            if label not in line.normalized:
                continue
            label_words = [word for word in line.words if label in word.normalized or word.normalized in label]
            label_right = max((word.x1 for word in label_words), default=line.x0)
            right_words = [word for word in line.words if word.x0 > label_right + 20]
            if not right_words:
                continue
            value_words = _trim_value_words(right_words, entity_type)
            if not value_words:
                continue
            candidate = _candidate_from_words(
                value_words,
                page_rect,
                entity_type,
                replacement,
                rule_name,
                f"{label}ラベルの右側にある値を{entity_type}候補として検出",
                referenced_header=label,
            )
            if candidate:
                candidates.append(candidate)
    return candidates


def _table_column_candidates(lines: list[_LineBox], page_rect: Any) -> list[OcrCandidate]:
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
            rows = _rows_below(lines, header)
            for row in rows:
                row_words = [word for word in row.words if left_boundary <= word.cx <= right_boundary]
                if not row_words and entity_type in {"会社名", "住所"}:
                    row_words = [word for word in row.words if word.x0 < right_boundary and word.x1 > left_boundary]
                value_words = _trim_value_words(row_words, entity_type)
                if not value_words:
                    continue
                rect_override = None
                if entity_type == "住所":
                    rect_override = _clip_rect(_rect_for_words(value_words), (left_boundary, row.y0, right_boundary, row.y1), page_rect)
                candidate = _candidate_from_words(
                    value_words,
                    page_rect,
                    entity_type,
                    replacement,
                    rule_name,
                    f"{label}列の下にある行値を{entity_type}候補として検出",
                    referenced_header=label,
                    rect_override=rect_override,
                )
                if candidate:
                    candidates.append(candidate)
    return candidates


def _section_block_candidates(lines: list[_LineBox], page_rect: Any) -> list[OcrCandidate]:
    candidates: list[OcrCandidate] = []
    active_heading = ""
    active_until_y = -1.0
    for line in lines:
        heading = _section_heading(line)
        if heading:
            active_heading = heading
            active_until_y = line.y1 + 460
            continue
        if not active_heading or line.y0 > active_until_y:
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
        if entity_type == "会社名" and not (_has_corporate_marker(line.normalized) or _looks_like_business_name_line(line)):
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
        )
        if candidate:
            candidates.append(candidate)
    return candidates


def _standalone_pattern_candidates(lines: list[_LineBox], page_rect: Any) -> list[OcrCandidate]:
    candidates: list[OcrCandidate] = []
    for line in lines:
        if _is_noise_or_summary(line.normalized):
            continue
        if _has_corporate_marker(line.normalized) and _looks_like_business_name_line(line):
            candidate = _candidate_from_words(
                _trim_value_words(list(line.words), "会社名"),
                page_rect,
                "会社名",
                "法人001",
                "pattern:corporate_marker",
                "法人表記またはそのOCR揺れを含む文字ブロックを会社名候補として検出",
                referenced_header="法人表記",
            )
            if candidate:
                candidates.append(candidate)
        elif _looks_like_standalone_address(line.normalized):
            candidate = _candidate_from_words(
                _trim_value_words(list(line.words), "住所"),
                page_rect,
                "住所",
                "市区町村まで",
                "pattern:address",
                "都道府県、市区町村、字、番地などの住所構成を含む文字ブロックを住所候補として検出",
                referenced_header="住所構成",
            )
            if candidate:
                candidates.append(candidate)
        elif _FINANCIAL_RE.search(line.normalized) and line.normalized not in {"取引銀行", "金融機関", "銀行"}:
            candidate = _candidate_from_words(
                _trim_value_words(list(line.words), "銀行名"),
                page_rect,
                "銀行名",
                "金融機関001",
                "pattern:financial_institution",
                "金融機関を示す一般語を含む文字ブロックを銀行名候補として検出",
                referenced_header="金融機関表記",
            )
            if candidate:
                candidates.append(candidate)
    return candidates


def _header_words(line: _LineBox) -> list[_WordBox]:
    headers: list[_WordBox] = []
    for word in line.words:
        if any(label in word.normalized for label, *_rest in _TABLE_HEADER_RULES):
            headers.append(word)
        elif re.search(r"(No|種類|区分|面積|評価|金額|比率|年度|株式数|議決権)", word.normalized, re.IGNORECASE):
            headers.append(word)
    return sorted(headers, key=lambda item: item.cx)


def _column_boundaries(headers: list[_WordBox], header: _WordBox, page_rect: Any) -> tuple[float, float]:
    ordered = sorted(headers, key=lambda item: item.cx)
    index = ordered.index(header)
    previous_right = float(page_rect.x0)
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


def _rows_below(lines: list[_LineBox], header: _LineBox) -> list[_LineBox]:
    rows: list[_LineBox] = []
    for line in lines:
        if line.y0 <= header.y1 + 8:
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
    if entity_type == "会社名" and not (_has_corporate_marker(normalized) or _looks_like_company_text(normalized)):
        return []
    if entity_type == "住所" and not _looks_like_address(normalized):
        return []
    if entity_type == "銀行名" and not _FINANCIAL_RE.search(normalized):
        return []
    if entity_type == "氏名" and not _looks_like_person_text(normalized):
        return []
    return words


def _candidate_from_words(
    words: Iterable[_WordBox],
    page_rect: Any,
    entity_type: str,
    replacement: str,
    rule_name: str,
    reason: str,
    referenced_header: str,
    rect_override: tuple[float, float, float, float] | None = None,
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
    return OcrCandidate(
        original=original,
        normalized=normalized,
        entity_type=entity_type,
        status=CANDIDATE_REVIEW,
        replacement=replacement,
        rect=rect,
        reason=f"{rule_name}: {reason}。参照見出し/列={referenced_header}。信頼度=0.60",
    )


def _rect_for_words(words: Iterable[_WordBox]) -> tuple[float, float, float, float]:
    word_list = list(words)
    return (
        min(word.x0 for word in word_list),
        min(word.y0 for word in word_list),
        max(word.x1 for word in word_list),
        max(word.y1 for word in word_list),
    )


def _expand_rect(rect: tuple[float, float, float, float], entity_type: str) -> tuple[float, float, float, float]:
    x0, y0, x1, y1 = rect
    height = max(y1 - y0, 1.0)
    x_padding = height * (0.85 if entity_type == "氏名" else 0.15)
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


def _has_corporate_marker(value: str) -> bool:
    return bool(_CORPORATE_MARKER_RE.search(value))


def _looks_like_business_name_line(line: _LineBox) -> bool:
    text = line.normalized
    if _is_noise_or_summary(text):
        return False
    if _has_corporate_marker(text):
        return True
    japanese = sum(1 for char in text if _is_japanese(char))
    return japanese >= 3 and len(text) <= 28 and not _NUMERIC_HEAVY_RE.match(text)


def _looks_like_company_text(value: str) -> bool:
    if _is_noise_or_summary(value):
        return False
    if _has_corporate_marker(value):
        return True
    japanese = sum(1 for char in value if _is_japanese(char))
    return 3 <= japanese <= 24 and not _NUMERIC_HEAVY_RE.match(value)


def _looks_like_address(value: str) -> bool:
    if _is_noise_or_summary(value):
        return False
    return bool(_ADDRESS_HINT_RE.search(value)) and sum(1 for char in value if _is_japanese(char)) >= 3


def _looks_like_standalone_address(value: str) -> bool:
    if not _looks_like_address(value):
        return False
    return bool(re.search(r"(東京都|北海道|京都府|大阪府|[一-龯]{2,3}県)", value))


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
        deduped.append(candidate)
    return deduped
