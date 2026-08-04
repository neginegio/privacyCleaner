from __future__ import annotations

from dataclasses import dataclass


@dataclass
class Finding:
    enabled: bool
    sheet: str
    cell: str
    entity_type: str
    detection_kind: str
    original: str
    replacement: str
    reason: str
    start: int | None = None
    end: int | None = None
    # True only when a reviewer looked at a review-required candidate and
    # explicitly decided to keep the original text (not convert it), as
    # distinct from simply never having been reviewed yet. Word tracks this
    # on WordReplacementDecision instead and does not populate this field
    # for its own bookkeeping; Excel uses this field directly.
    excluded: bool = False

    @property
    def dedupe_key(self) -> tuple[str, str, str, str, str]:
        return (
            self.sheet,
            self.cell,
            self.entity_type,
            self.detection_kind,
            self.original,
        )
