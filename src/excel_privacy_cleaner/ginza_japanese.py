from __future__ import annotations

from dataclasses import dataclass

# GiNZA's NER uses Sekine's Extended Named Entity Hierarchy (~200 fine-grained
# labels), not this app's coarse 会社名/氏名 categories. This is a best-effort,
# hand-curated allow-list of the organization-family labels worth surfacing as
# 会社名 candidates; unlisted labels (dates, locations, products, ...) are
# dropped. Extend this list as real documents surface more label variants.
GINZA_ORGANIZATION_LABELS = {
    "Company",
    "Company_Group",
    "Corporation_Other",
    "International_Organization",
    "Show_Organization",
    "Political_Organization",
    "Political_Organization_Other",
    "Political_Party",
    "Cabinet",
    "Government",
    "Sports_Organization",
    "Pro_Sports_Organization",
    "Sports_League",
    "School",
    "Ethnic_Group_Other",
    "Nationality",
    "Military",
}
GINZA_PERSON_LABELS = {"Person"}

# GiNZA's standard NER pipeline does not expose a calibrated per-entity
# confidence score. Every candidate from this source is given this fixed,
# below-review-threshold confidence so it always lands in the "要確認"
# (review-required) bucket rather than being auto-converted -- see
# WORD_REVIEW_CONFIDENCE_THRESHOLD in word_processor.py. The detection_rule
# name is also distinct from every pattern-based rule so it is clearly
# visible in the findings list/CSV/audit log which candidates came from NLP
# judgment rather than a fixed pattern or dictionary.
WORD_NLP_DETECTION_RULE = "ginza_ner"
WORD_NLP_CONFIDENCE = 0.50

_GINZA_MODEL_NAME = "ja_ginza"


@dataclass(frozen=True)
class GinzaNlpResult:
    start: int
    end: int
    category: str
    raw_label: str


_nlp = None


def _load_nlp():
    global _nlp
    if _nlp is None:
        import spacy

        _nlp = spacy.load(
            _GINZA_MODEL_NAME,
            config={"components": {"compound_splitter": {"split_mode": "A"}}},
        )
    return _nlp


class GinzaEntityDetector:
    """Best-effort Japanese NER detector (GiNZA) for candidates that the
    pattern/dictionary rules structurally cannot see (e.g. a company's short
    form used without its legal-entity suffix elsewhere in the document)."""

    def __init__(self) -> None:
        self._nlp = _load_nlp()

    def analyze(self, text: str) -> list[GinzaNlpResult]:
        if not text or not text.strip():
            return []
        doc = self._nlp(text)
        results: list[GinzaNlpResult] = []
        for ent in doc.ents:
            if ent.label_ in GINZA_ORGANIZATION_LABELS:
                category = "会社名"
            elif ent.label_ in GINZA_PERSON_LABELS:
                category = "氏名"
            else:
                continue
            results.append(
                GinzaNlpResult(start=ent.start_char, end=ent.end_char, category=category, raw_label=ent.label_)
            )
        return results
