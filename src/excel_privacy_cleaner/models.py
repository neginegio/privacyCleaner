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

    @property
    def dedupe_key(self) -> tuple[str, str, str, str, str]:
        return (
            self.sheet,
            self.cell,
            self.entity_type,
            self.detection_kind,
            self.original,
        )
