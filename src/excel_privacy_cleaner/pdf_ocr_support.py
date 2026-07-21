from __future__ import annotations

import json
import re
import unicodedata
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from .resources import resource_path


QUALITY_PASS = "PASS"
QUALITY_REVIEW = "REVIEW_REQUIRED"
QUALITY_FAILED = "FAILED"

CANDIDATE_AUTO = "AUTO_CONFIRMED"
CANDIDATE_REVIEW = "REVIEW_REQUIRED"
CANDIDATE_REJECTED = "REJECTED"
CANDIDATE_MANUAL = "MANUAL"
USER_APPROVED = "USER_APPROVED"
USER_REJECTED = "USER_REJECTED"

PAGE_UNREVIEWED = "UNREVIEWED"
PAGE_REVIEWED_NO_SENSITIVE_DATA = "REVIEWED_NO_SENSITIVE_DATA"
PAGE_REVIEWED_WITH_REDACTIONS = "REVIEWED_WITH_REDACTIONS"
PAGE_FAILED_UNRESOLVED = "FAILED_UNRESOLVED"
PAGE_VERIFICATION_FAILED = "VERIFICATION_FAILED"
PAGE_COMPLETED = "COMPLETED"

VERIFICATION_STRUCTURE_PASS = "STRUCTURE_PASS"
VERIFICATION_MANUAL_REGION_PASS = "MANUAL_REGION_PASS"
VERIFICATION_TEXT_RESIDUAL_PASS = "TEXT_RESIDUAL_PASS"
VERIFICATION_IMAGE_RESIDUAL_PASS = "IMAGE_RESIDUAL_PASS"
VERIFICATION_NOT_EVALUATED = "NOT_EVALUATED"
VERIFICATION_FAILED = "FAILED"


@dataclass(frozen=True)
class OcrQualityConfig:
    min_ocr_chars_pass: int = 40
    min_ocr_chars_review: int = 10
    max_garble_ratio_pass: float = 0.08
    max_garble_ratio_review: float = 0.18
    min_text_area_ratio_pass: float = 0.0005
    min_text_area_ratio_review: float = 0.0001
    candidate_free_review_chars: int = 80
    visual_dark_ratio_threshold: float = 0.001


@dataclass(frozen=True)
class OcrCandidate:
    original: str
    normalized: str
    entity_type: str
    status: str
    replacement: str
    rect: tuple[float, float, float, float]
    reason: str


@dataclass(frozen=True)
class PdfPageQuality:
    page_number: int
    ocr_text_chars: int
    coordinate_word_count: int
    japanese_char_count: int
    alnum_char_count: int
    symbol_char_count: int
    garble_candidate_count: int
    text_area_ratio: float
    candidate_count: int
    verdict: str
    warning_reason: str


@dataclass(frozen=True)
class _CharBox:
    source: str
    normalized: str
    rect: tuple[float, float, float, float]
    word_index: int


@dataclass(frozen=True)
class _NormalizedOcr:
    text: str
    chars: tuple[_CharBox, ...]


def load_quality_config() -> OcrQualityConfig:
    path = resource_path("config/pdf_ocr_quality.json")
    if not path.exists():
        return OcrQualityConfig()
    data = json.loads(path.read_text(encoding="utf-8"))
    allowed = {field for field in OcrQualityConfig.__dataclass_fields__}
    return OcrQualityConfig(**{key: value for key, value in data.items() if key in allowed})


def detect_ocr_candidates(
    words: list[tuple[Any, ...]],
    page_rect: Any,
    replacements: dict[str, str] | None = None,
) -> list[OcrCandidate]:
    normalized = _build_normalized_ocr(words)
    text = normalized.text
    replacements = replacements or {}
    candidates: list[OcrCandidate] = []
    boundary = r"(?=氏名|氏名カナ|カナ|会社名|法人名|住所|所在地|電話|TEL|メール|口座|API|パスワード|森|南|$)"

    def add(pattern: str, entity_type: str, status: str, replacement: str, reason: str, flags: int = 0, group: int = 1) -> None:
        for match in re.finditer(pattern, text, flags):
            start, end = match.span(group if match.lastindex else 0)
            value = text[start:end]
            rect = _rect_for_range(normalized.chars, start, end)
            if rect is None:
                continue
            rect = _safe_candidate_rect(rect, page_rect, entity_type)
            if entity_type == "氏名候補":
                rect = _tighten_short_japanese_candidate_rect(rect, value, page_rect)
            elif entity_type == "銀行口座":
                rect = _tighten_account_number_rect(rect, value, page_rect)
            elif entity_type == "電話番号":
                rect = _tighten_phone_number_rect(rect, value, page_rect)
            candidates.append(
                OcrCandidate(
                    original=_display_original(entity_type, _source_for_range(normalized.chars, start, end) or value, value),
                    normalized=value,
                    entity_type=entity_type,
                    status=status,
                    replacement=replacements.get(entity_type, replacement),
                    rect=rect,
                    reason=reason,
                )
            )

    add(
        r"([A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,})",
        "メールアドレス",
        CANDIDATE_AUTO,
        "mail001@example.invalid",
        "OCR正規化後に明確なメール形式として検出",
    )
    add(
        r"(sk-[A-Za-z0-9_-]{10,}|api[_-]?[A-Za-z0-9_-]{8,}|secret[_-]?[A-Za-z0-9_-]{8,}|tok[_-]?[A-Za-z0-9_-]{8,})",
        "APIキー",
        CANDIDATE_AUTO,
        "[SECRET]",
        "OCR正規化後に明確な秘密情報形式として検出",
        flags=re.IGNORECASE,
    )
    add(
        r"(?:電話|TEL|Tel|tel)[:：]?([0O][0-9OIl]{1,3}[-ー−－]?[0-9OIl]{3,4}[-ー−－]?[0-9OIl]{4})",
        "電話番号",
        CANDIDATE_AUTO,
        "090-****-2222",
        "電話番号候補範囲だけでO/0/I/l/1を正規化して検出",
    )
    add(
        r"(?:普通|当座)([0-9OIl]{7})",
        "銀行口座",
        CANDIDATE_REVIEW,
        "*******",
        "口座種別に続く7桁番号として検出。OCR誤認識確認が必要",
    )
    add(
        rf"(?:氏名|名前)[:：]?([一-龯々〆ヵヶ]{{2,8}}?){boundary}",
        "氏名",
        CANDIDATE_REVIEW,
        "個人001",
        "氏名ラベルに続く日本語名候補。人による確認が必要",
    )
    add(
        rf"(?:氏名カナ|カナ)[:：]?([ァ-ヶー]{{4,16}}?){boundary}",
        "氏名カナ",
        CANDIDATE_REVIEW,
        "個人001",
        "氏名カナ候補。氏名との対応確認が必要",
    )
    add(
        rf"(?:会社名|法人名)[:：]?((?:株式会社|有限会社|合同会社)[一-龯々〆ヵヶぁ-んァ-ヶーA-Za-z0-9]{{2,24}}?){boundary}",
        "会社名",
        CANDIDATE_REVIEW,
        "法人001",
        "会社名ラベルに続く法人名候補。人による確認が必要",
    )
    add(
        rf"(?:住所|所在地)[:：]?((?:東京都|北海道|京都府|大阪府|[一-龯]{{2,3}}県)[一-龯々〆ヵヶぁ-んァ-ヶー0-9\-ー−－]{{4,40}}?){boundary}",
        "住所",
        CANDIDATE_REVIEW,
        "東京都千代田区",
        "住所ラベルに続く住所候補。OCR揺れ確認が必要",
    )
    add(
        r"(?:パスワード|password|Password)[:：]?([A-Za-z0-9!#$%&@_\-]{6,})",
        "パスワード",
        CANDIDATE_REVIEW,
        "[SECRET]",
        "パスワードラベルに続く秘密情報候補。人による確認が必要",
    )
    add(
        r"([一-龯々〆ヵヶ]{1,3}さん)",
        "氏名候補",
        CANDIDATE_REVIEW,
        "個人候補",
        "姓のみの敬称表現のため要確認",
    )

    rejected_patterns = (
        (r"森を抜ける", "一般文として扱い、人名候補から除外"),
        (r"南側の入口", "一般文として扱い、人名候補から除外"),
    )
    for pattern, reason in rejected_patterns:
        for match in re.finditer(pattern, text):
            rect = _rect_for_range(normalized.chars, match.start(), match.end())
            if rect is None:
                continue
            candidates.append(
                OcrCandidate(
                    original=_source_for_range(normalized.chars, match.start(), match.end()) or match.group(0),
                    normalized=match.group(0),
                    entity_type="対象外",
                    status=CANDIDATE_REJECTED,
                    replacement="変換しない",
                    rect=rect,
                    reason=reason,
                )
            )

    return _dedupe_candidates(candidates, page_rect)


def evaluate_page_quality(
    page: Any,
    text: str,
    words: list[tuple[Any, ...]],
    candidates: list[OcrCandidate],
    exception: str = "",
    config: OcrQualityConfig | None = None,
) -> PdfPageQuality:
    config = config or load_quality_config()
    stripped = text.strip()
    coordinate_words = [word for word in words if len(word) >= 5 and float(word[2]) > float(word[0]) and float(word[3]) > float(word[1])]
    japanese_count = sum(1 for char in stripped if _is_japanese(char))
    alnum_count = sum(1 for char in stripped if char.isalnum() and not _is_japanese(char))
    symbol_count = sum(1 for char in stripped if not char.isspace() and not char.isalnum() and not _is_japanese(char))
    garble_count = sum(1 for char in stripped if _is_garble_candidate(char))
    text_area_ratio = _text_area_ratio(page, coordinate_words)
    reasons: list[str] = []
    verdict = QUALITY_PASS

    if exception:
        verdict = QUALITY_FAILED
        reasons.append(f"OCR例外: {exception}")
    elif not _page_image_available(page):
        verdict = QUALITY_FAILED
        reasons.append("ページ画像を取得できません。")
    elif not coordinate_words:
        if _page_has_visible_marks(page, config):
            verdict = QUALITY_FAILED
            reasons.append("目視上は文字があると推定されますがOCR文字数または座標付き単語が0です。")
        else:
            verdict = QUALITY_REVIEW
            reasons.append("OCR文字と座標付き単語がありません。")
    else:
        if len(stripped) < config.min_ocr_chars_review:
            verdict = QUALITY_REVIEW
            reasons.append("OCR文字数が少なすぎます。")
        elif len(stripped) < config.min_ocr_chars_pass:
            verdict = QUALITY_REVIEW
            reasons.append("OCR文字数が少ないため確認が必要です。")

        garble_ratio = garble_count / max(len(stripped), 1)
        if garble_ratio > config.max_garble_ratio_review:
            verdict = QUALITY_REVIEW
            reasons.append("文字化け候補の割合が高いです。")
        elif garble_ratio > config.max_garble_ratio_pass and verdict == QUALITY_PASS:
            verdict = QUALITY_REVIEW
            reasons.append("文字化け候補が含まれます。")

        if text_area_ratio < config.min_text_area_ratio_review:
            verdict = QUALITY_REVIEW
            reasons.append("ページ面積に対する文字領域が小さすぎます。")
        elif text_area_ratio < config.min_text_area_ratio_pass and verdict == QUALITY_PASS:
            verdict = QUALITY_REVIEW
            reasons.append("ページ面積に対する文字領域が小さいため確認が必要です。")

        active_candidates = [candidate for candidate in candidates if candidate.status != CANDIDATE_REJECTED]
        if len(stripped) >= config.candidate_free_review_chars and not active_candidates:
            verdict = QUALITY_REVIEW
            reasons.append("一定量の文字がありますが検出候補が0件です。")
        if any(candidate.status == CANDIDATE_REVIEW for candidate in candidates):
            verdict = QUALITY_REVIEW
            reasons.append("氏名、住所、会社名、口座番号などの要確認候補があります。")

    return PdfPageQuality(
        page_number=int(getattr(page, "number", 0)) + 1,
        ocr_text_chars=len(stripped),
        coordinate_word_count=len(coordinate_words),
        japanese_char_count=japanese_count,
        alnum_char_count=alnum_count,
        symbol_char_count=symbol_count,
        garble_candidate_count=garble_count,
        text_area_ratio=text_area_ratio,
        candidate_count=len([candidate for candidate in candidates if candidate.status != CANDIDATE_REJECTED]),
        verdict=verdict,
        warning_reason=" / ".join(reasons) if reasons else "なし",
    )


def _build_normalized_ocr(words: list[tuple[Any, ...]]) -> _NormalizedOcr:
    chars: list[_CharBox] = []
    parts: list[str] = []
    sorted_words = sorted(words, key=lambda word: (int(word[5]), int(word[6]), int(word[7]) if len(word) > 7 else 0))
    for word_index, word in enumerate(sorted_words):
        raw = str(word[4])
        normalized = _normalize_general(raw)
        raw_chars = list(raw)
        norm_chars = list(normalized)
        if not norm_chars:
            continue
        x0, y0, x1, y1 = (float(word[0]), float(word[1]), float(word[2]), float(word[3]))
        count = max(len(norm_chars), 1)
        for index, norm_char in enumerate(norm_chars):
            left = x0 + (x1 - x0) * index / count
            right = x0 + (x1 - x0) * (index + 1) / count
            source_char = raw_chars[min(index, len(raw_chars) - 1)] if raw_chars else norm_char
            chars.append(_CharBox(source=source_char, normalized=norm_char, rect=(left, y0, right, y1), word_index=word_index))
            parts.append(norm_char)
    return _NormalizedOcr(text="".join(parts), chars=tuple(chars))


def _normalize_general(value: str) -> str:
    text = unicodedata.normalize("NFKC", value)
    text = text.replace("＠", "@").replace("．", ".")
    text = re.sub(r"[‐‑‒–—―−－ー]", "-", text)
    return "".join(text.split())


def _rect_for_range(chars: tuple[_CharBox, ...], start: int, end: int) -> tuple[float, float, float, float] | None:
    selected = chars[start:end]
    if not selected:
        return None
    x0 = min(char.rect[0] for char in selected)
    y0 = min(char.rect[1] for char in selected)
    x1 = max(char.rect[2] for char in selected)
    y1 = max(char.rect[3] for char in selected)
    if start > 0:
        previous = chars[start - 1]
        first = selected[0]
        if previous.normalized in {":", "："}:
            if previous.word_index == first.word_index:
                x0 = min(x0, previous.rect[2] - (y1 - y0) * 0.25)
            else:
                x0 = min(x0, previous.rect[2] - min((y1 - y0) * 0.08, 2.0))
    return (x0, y0, x1, y1)


def _safe_candidate_rect(rect: tuple[float, float, float, float], page_rect: Any, entity_type: str) -> tuple[float, float, float, float]:
    x0, y0, x1, y1 = rect
    height = max(y1 - y0, 1.0)
    pad_left = height * 0.08
    pad_right = height * 0.20
    if entity_type in {"メールアドレス", "APIキー", "パスワード"}:
        pad_y = max(height * 0.35, 7.0)
    elif entity_type in {"電話番号", "銀行口座"}:
        pad_y = max(height * 0.20, 4.0)
    else:
        pad_y = max(height * 0.12, 3.0)
    return (
        max(float(page_rect.x0), x0 - pad_left),
        max(float(page_rect.y0), y0 - pad_y),
        min(float(page_rect.x1), x1 + pad_right),
        min(float(page_rect.y1), y1 + pad_y),
    )


def _tighten_short_japanese_candidate_rect(
    rect: tuple[float, float, float, float],
    value: str,
    page_rect: Any,
) -> tuple[float, float, float, float]:
    if not re.fullmatch(r"[一-龯々〆ヵヶ]{1,3}さん", value):
        return rect
    x0, y0, x1, y1 = rect
    height = max(y1 - y0, 1.0)
    text_height = height / 1.7
    expected_width = text_height * max(len(value), 2) * 1.10
    if x1 - x0 <= expected_width:
        return rect
    return (x0, y0, min(float(page_rect.x1), x0 + expected_width), y1)


def _tighten_account_number_rect(
    rect: tuple[float, float, float, float],
    value: str,
    page_rect: Any,
) -> tuple[float, float, float, float]:
    if not re.fullmatch(r"[0-9OIl]{7}", value):
        return rect
    x0, y0, x1, y1 = rect
    height = max(y1 - y0, 1.0)
    text_height = height / 1.7
    expected_width = text_height * len(value) * 0.65
    if x1 - x0 <= expected_width:
        return rect
    return (max(float(page_rect.x0), x1 - expected_width), y0, x1, y1)


def _tighten_phone_number_rect(
    rect: tuple[float, float, float, float],
    value: str,
    page_rect: Any,
) -> tuple[float, float, float, float]:
    if not re.fullmatch(r"[0O][0-9OIl]{1,3}[-ー−－]?[0-9OIl]{3,4}[-ー−－]?[0-9OIl]{4}", value):
        return rect
    x0, y0, x1, y1 = rect
    height = max(y1 - y0, 1.0)
    text_height = height / 1.4
    expected_width = text_height * len(value) * 0.52
    if x1 - x0 <= expected_width:
        return rect
    return (max(float(page_rect.x0), x1 - expected_width), y0, x1, y1)


def _source_for_range(chars: tuple[_CharBox, ...], start: int, end: int) -> str:
    return "".join(char.source for char in chars[start:end])


def _dedupe_candidates(candidates: list[OcrCandidate], page_rect: Any) -> list[OcrCandidate]:
    unique: list[OcrCandidate] = []
    seen: set[tuple[str, str, str]] = set()
    for candidate in candidates:
        if not _rect_inside_page(candidate.rect, page_rect):
            continue
        rect_key = tuple(round(value, 1) for value in candidate.rect)
        key = (candidate.entity_type, candidate.status, candidate.normalized, str(rect_key))
        if key in seen:
            continue
        seen.add(key)
        unique.append(candidate)
    return unique


def _display_original(entity_type: str, source: str, normalized: str) -> str:
    if entity_type in {"電話番号", "銀行口座", "メールアドレス", "APIキー"}:
        return normalized
    return source


def _rect_inside_page(rect: tuple[float, float, float, float], page_rect: Any) -> bool:
    x0, y0, x1, y1 = rect
    return x1 > x0 and y1 > y0 and x0 >= page_rect.x0 - 1 and y0 >= page_rect.y0 - 1 and x1 <= page_rect.x1 + 1 and y1 <= page_rect.y1 + 1


def _text_area_ratio(page: Any, words: list[tuple[Any, ...]]) -> float:
    area = 0.0
    for word in words:
        area += max(float(word[2]) - float(word[0]), 0) * max(float(word[3]) - float(word[1]), 0)
    page_area = max(float(page.rect.width) * float(page.rect.height), 1.0)
    return min(area / page_area, 1.0)


def _page_image_available(page: Any) -> bool:
    try:
        pix = page.get_pixmap(dpi=48, alpha=False)
        return pix.width > 0 and pix.height > 0 and bool(pix.samples)
    except Exception:
        return False


def _page_has_visible_marks(page: Any, config: OcrQualityConfig) -> bool:
    try:
        pix = page.get_pixmap(dpi=48, alpha=False)
    except Exception:
        return False
    samples = pix.samples
    if not samples:
        return False
    dark = 0
    total = pix.width * pix.height
    channels = pix.n
    for offset in range(0, len(samples), channels):
        red = samples[offset]
        green = samples[offset + 1] if channels > 1 else red
        blue = samples[offset + 2] if channels > 2 else red
        if (red + green + blue) / 3 < 210:
            dark += 1
    return dark / max(total, 1) >= config.visual_dark_ratio_threshold


def _is_japanese(char: str) -> bool:
    return "\u3040" <= char <= "\u30ff" or "\u4e00" <= char <= "\u9fff"


def _is_garble_candidate(char: str) -> bool:
    if char in "\ufffd□■●○◆◇�":
        return True
    category = unicodedata.category(char)
    return category == "Co"
