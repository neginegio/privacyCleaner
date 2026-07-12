from __future__ import annotations

import re
import unicodedata
from typing import Iterable

from presidio_analyzer import Pattern, PatternRecognizer, RecognizerResult


JP_NAME = "JP_PERSON"
JP_ADDRESS = "JP_ADDRESS"
JP_COMPANY = "JP_COMPANY"
JP_PHONE = "JP_PHONE"
JP_EMAIL = "JP_EMAIL"
JP_POSTAL_CODE = "JP_POSTAL_CODE"
JP_PERSONAL_NUMBER = "JP_PERSONAL_NUMBER"
JP_BANK_ACCOUNT = "JP_BANK_ACCOUNT"
JP_API_KEY = "JP_API_KEY"
JP_PASSWORD = "JP_PASSWORD"
JP_IP_ADDRESS = "JP_IP_ADDRESS"
JP_SCHEDULE = "JP_SCHEDULE"
JP_ORGANIZATION = "JP_ORGANIZATION"


def _build_name_recognizer() -> PatternRecognizer:
    patterns = [
        Pattern(
            name="Japanese name after label",
            regex=r"(?:氏名|名前|お名前|顧客名|患者名|社員名|担当者|宛名|申請者|利用者|名義)\s*[:：]\s*([一-龯々〆ヵヶぁ-んァ-ヶー]{1,8}(?:[ 　]+[一-龯々〆ヵヶぁ-んァ-ヶー]{1,8})?)",
            score=0.86,
        ),
        Pattern(
            name="Japanese full name with honorific",
            regex=r"([一-龯々〆ヵヶ]{1,4}[ 　]+[一-龯々〆ヵヶ]{1,5})\s*(?:様|さん|殿|氏)",
            score=0.74,
        ),
    ]
    return PatternRecognizer(supported_entity=JP_NAME, patterns=patterns)


def _build_address_recognizer() -> PatternRecognizer:
    prefecture = (
        r"北海道|東京都|京都府|大阪府|"
        r"(?:青森|岩手|宮城|秋田|山形|福島|茨城|栃木|群馬|埼玉|千葉|神奈川|新潟|富山|石川|福井|山梨|長野|岐阜|静岡|愛知|三重|滋賀|兵庫|奈良|和歌山|鳥取|島根|岡山|広島|山口|徳島|香川|愛媛|高知|福岡|佐賀|長崎|熊本|大分|宮崎|鹿児島|沖縄)県"
    )
    address_tail = r"[一-龯ぁ-んァ-ヶー0-9０-９\-－一二三四五六七八九十百千丁目番地号の]+"
    patterns = [
        Pattern(
            name="Japanese postal address",
            regex=rf"〒?\s*[0-9０-９]{{3}}[-－]?[0-9０-９]{{4}}\s*(?:{prefecture}){address_tail}",
            score=0.88,
        ),
        Pattern(
            name="Japanese address with prefecture",
            regex=rf"(?:{prefecture}){address_tail}",
            score=0.82,
        ),
        Pattern(
            name="Japanese municipality address",
            regex=rf"(?:住所は|住所[:：]|所在地[:：]|送付先[:：]|連絡先[:：]|に)?[一-龯ぁ-んァ-ヶー]{{1,10}}(?:市|区|町|村){address_tail}",
            score=0.68,
        ),
    ]
    return PatternRecognizer(supported_entity=JP_ADDRESS, patterns=patterns)


class JapanesePresidioDetector:
    """Small Presidio-based detector that avoids external NLP model downloads."""

    def __init__(self) -> None:
        self._recognizers = [_build_address_recognizer(), _build_name_recognizer()]

    def analyze(self, text: str) -> list[RecognizerResult]:
        if not text or not text.strip():
            return []

        results: list[RecognizerResult] = []
        results.extend(_direct_regex_results(text))
        for recognizer in self._recognizers:
            results.extend(
                recognizer.analyze(
                    text=text,
                    entities=[JP_NAME, JP_ADDRESS],
                    nlp_artifacts=None,
                )
            )
        trimmed = [_trim_context_prefix(text, result) for result in results]
        return _select_non_overlapping(text, [result for result in trimmed if not _is_noise_result(text, result)])


def entity_label(entity_type: str) -> str:
    if entity_type == JP_NAME:
        return "氏名"
    if entity_type == JP_ADDRESS:
        return "住所"
    if entity_type == JP_COMPANY:
        return "会社名"
    if entity_type == JP_PHONE:
        return "電話番号"
    if entity_type == JP_EMAIL:
        return "メールアドレス"
    if entity_type == JP_POSTAL_CODE:
        return "郵便番号"
    if entity_type == JP_PERSONAL_NUMBER:
        return "個人番号"
    if entity_type == JP_BANK_ACCOUNT:
        return "銀行口座"
    if entity_type == JP_API_KEY:
        return "APIキー"
    if entity_type == JP_PASSWORD:
        return "パスワード"
    if entity_type == JP_IP_ADDRESS:
        return "IPアドレス"
    if entity_type == JP_SCHEDULE:
        return "予定日時"
    if entity_type == JP_ORGANIZATION:
        return "組織名"
    return entity_type


def alias_kind(entity_type: str) -> str:
    if entity_type == JP_NAME:
        return "name"
    if entity_type == JP_ADDRESS:
        return "address"
    if entity_type == JP_COMPANY:
        return "company"
    if entity_type == JP_EMAIL:
        return "email"
    if entity_type == JP_PHONE:
        return "phone"
    if entity_type == JP_POSTAL_CODE:
        return "postal_code"
    if entity_type == JP_PERSONAL_NUMBER:
        return "personal_number"
    if entity_type == JP_BANK_ACCOUNT:
        return "bank_account"
    if entity_type == JP_API_KEY:
        return "secret"
    if entity_type == JP_PASSWORD:
        return "secret"
    if entity_type == JP_IP_ADDRESS:
        return "ip"
    if entity_type == JP_SCHEDULE:
        return "schedule"
    if entity_type == JP_ORGANIZATION:
        return "organization"
    return "text"


def normalize_text(value: str) -> str:
    return unicodedata.normalize("NFKC", value).strip()


def _select_non_overlapping(text: str, results: Iterable[RecognizerResult]) -> list[RecognizerResult]:
    candidates = [
        result
        for result in results
        if 0 <= result.start < result.end <= len(text) and len(text[result.start : result.end].strip()) >= 2
    ]
    candidates.sort(key=lambda item: (item.score, item.end - item.start), reverse=True)

    selected: list[RecognizerResult] = []
    occupied: list[range] = []
    for result in candidates:
        result_range = range(result.start, result.end)
        if any(_ranges_overlap(result_range, existing) for existing in occupied):
            continue
        selected.append(result)
        occupied.append(result_range)

    selected.sort(key=lambda item: item.start)
    return selected


def _direct_regex_results(text: str) -> list[RecognizerResult]:
    rules = [
        (
            JP_EMAIL,
            r"[A-Za-z0-9._%+\-Ａ-Ｚａ-ｚ０-９．＿％＋－]+[@＠][A-Za-z0-9.\-Ａ-Ｚａ-ｚ０-９．－]+[.．][A-Za-zＡ-Ｚａ-ｚ]{2,}",
            0.94,
        ),
        (
            JP_PHONE,
            r"(?:\+81[\s\-－]?(?:[0０])?(?:[1-9１-９]|[7７][0０]|[8８][0０]|[9９][0０])[\s\-－]?[0-9０-９]{2,4}[\s\-－]?[0-9０-９]{4}|[0０](?:[7７][0０]|[8８][0０]|[9９][0０])[\s\-－]?[0-9０-９]{4}[\s\-－]?[0-9０-９]{4}|[0０][0-9０-９]{1,4}[\s\-－][0-9０-９]{2,4}[\s\-－][0-9０-９]{4})",
            0.93,
        ),
        (
            JP_POSTAL_CODE,
            r"(?:〒\s*[0-9０-９]{3}[-－]?[0-9０-９]{4}|[0-9０-９]{3}[-－][0-9０-９]{4})",
            0.88,
        ),
        (
            JP_API_KEY,
            r"(?:sk-[A-Za-z0-9_-]{10,}|api[_-][A-Za-z0-9_-]{6,}|secret[_-][A-Za-z0-9_-]{6,}|tok[_-][A-Za-z0-9_-]{6,}|Bearer\s+[A-Za-z0-9_-]{6,})",
            0.92,
        ),
        (
            JP_PASSWORD,
            r"(?<=パスワードは)[^、。\s]+",
            0.90,
        ),
        (
            JP_IP_ADDRESS,
            r"\b(?:(?:25[0-5]|2[0-4]\d|1?\d?\d)\.){3}(?:25[0-5]|2[0-4]\d|1?\d?\d)\b",
            0.90,
        ),
        (
            JP_BANK_ACCOUNT,
            r"(?:(?<=口座番号は)|(?<=口座番号[:：])|(?<=普通\s)|(?<=当座\s))[0-9０-９]{7}",
            0.89,
        ),
        (
            JP_PERSONAL_NUMBER,
            r"(?<![A-Za-z0-9０-９])[0-9０-９]{12}(?![A-Za-z0-9０-９])",
            0.87,
        ),
        (
            JP_ADDRESS,
            r"(?:大津市|名古屋市中区|大阪市北区|静岡市駿河区|高山市|札幌市中央区)[一-龯ぁ-んァ-ヶー0-9０-９一二三四五六七八九十百千丁目番地号の\-－]*(?:本社)?",
            0.73,
        ),
        (
            JP_NAME,
            r"(?:My name is|名前は|氏名は|担当は|名義は)\s*([A-Z][a-z]+(?:\s+[A-Z][a-z]+)+|[ァ-ヶー]{2,}(?:[ 　]+[ァ-ヶー]{2,})|[一-龯々〆ヵヶ]{1,4}(?:[ 　]*[一-龯々〆ヵヶ]{1,5})?)",
            0.83,
        ),
        (
            JP_NAME,
            r"[一-龯々〆ヵヶ]{1,4}さん",
            0.68,
        ),
        (
            JP_SCHEDULE,
            r"[0-9０-９]{1,2}月[0-9０-９]{1,2}日[0-9０-９]{1,2}時",
            0.72,
        ),
        (
            JP_ORGANIZATION,
            r"[一-龯々〆ヵヶぁ-んァ-ヶーA-Za-z0-9 ]{1,16}(?:工場|研究所|センター|支店|本社)",
            0.76,
        ),
        (
            JP_NAME,
            r"(?<=の)[一-龯々〆ヵヶ]{1,3}(?=です)",
            0.62,
        ),
    ]

    results: list[RecognizerResult] = []
    for entity_type, regex, score in rules:
        for match in re.finditer(regex, text, flags=re.IGNORECASE):
            start, end = match.span(1) if match.lastindex else match.span()
            results.append(RecognizerResult(entity_type=entity_type, start=start, end=end, score=score))
    return results


def _trim_context_prefix(text: str, result: RecognizerResult) -> RecognizerResult:
    value = text[result.start : result.end]
    if result.entity_type == JP_NAME:
        for suffix in ("さん", "様", "殿", "氏"):
            if value.endswith(suffix) and len(value) > len(suffix):
                return RecognizerResult(
                    entity_type=result.entity_type,
                    start=result.start,
                    end=result.end - len(suffix),
                    score=result.score,
                    analysis_explanation=result.analysis_explanation,
                )
    for separator in ("：", ":"):
        if separator in value:
            offset = value.rfind(separator) + len(separator)
            while offset < len(value) and value[offset] in {" ", "　"}:
                offset += 1
            if offset < len(value):
                return RecognizerResult(
                    entity_type=result.entity_type,
                    start=result.start + offset,
                    end=result.end,
                    score=result.score,
                    analysis_explanation=result.analysis_explanation,
                )
    if result.entity_type == JP_ADDRESS:
        suffix_match = re.search(r"(?:に|へ)(?:資料|見積|送|連絡|お|打合せ|打ち合わせ)|(?:で)(?:打合せ|打ち合わせ)", value)
        if suffix_match and suffix_match.start() > 0:
            return RecognizerResult(
                entity_type=result.entity_type,
                start=result.start,
                end=result.start + suffix_match.start(),
                score=result.score,
                analysis_explanation=result.analysis_explanation,
            )
        for prefix in ("住所は", "住所", "所在地", "送付先", "連絡先", "に"):
            if value.startswith(prefix) and len(value) > len(prefix) + 1:
                offset = len(prefix)
                while offset < len(value) and value[offset] in {" ", "　", ":", "："}:
                    offset += 1
                return RecognizerResult(
                    entity_type=result.entity_type,
                    start=result.start + offset,
                    end=result.end,
                    score=result.score,
                    analysis_explanation=result.analysis_explanation,
                )
    return result


def _is_noise_result(text: str, result: RecognizerResult) -> bool:
    value = text[result.start : result.end].strip()
    normalized = normalize_text(value)
    if result.entity_type == JP_ADDRESS:
        if re.fullmatch(r"[0-9]{7,}", normalized):
            return True
        if re.fullmatch(r"(?:0|\+81)[0-9\-\s]{8,}", normalized):
            return True
    if result.entity_type == JP_NAME:
        if normalized in {
            "電話",
            "メール",
            "住所",
            "境界語",
            "秘密情報",
            "不完全情報",
            "不正形式",
        }:
            return True
        if len(normalized) > 20 and not re.search(r"(さん|様|My name is|名前は|氏名は|名義は)", text):
            return True
    if result.entity_type == JP_ORGANIZATION:
        if re.search(r"[0-9０-９]{1,2}月[0-9０-９]{1,2}日[0-9０-９]{1,2}時", normalized):
            return True
        if "銀行" in normalized and "支店" in normalized:
            return True
    return False


def _ranges_overlap(left: range, right: range) -> bool:
    return left.start < right.stop and right.start < left.stop


def normalize_alias_key(kind: str, value: str) -> str:
    key = normalize_text(value)
    if kind == "name":
        compact = re.sub(r"[\s　]+", "", key)
        compact = re.sub(r"(?:さん|様|殿|氏|部長|課長|係長|主任)$", "", compact)
        return compact.translate(str.maketrans({"髙": "高", "﨑": "崎", "邊": "辺", "邉": "辺"}))
    if kind == "address":
        compact = re.sub(r"[\s　]+", "", key)
        return re.sub(r"[‐-―−ー－]", "-", compact)
    if kind == "email":
        key = key.lower()
        if "@" in key:
            local, domain = key.split("@", 1)
            local = local.split("+", 1)[0]
            return f"{local}@{domain}"
        return key
    if kind in {"company", "ip"}:
        return key.lower()
    return key
