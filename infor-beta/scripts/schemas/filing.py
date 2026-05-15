"""Filing schema — Phase 1.

Per Obsidian note 12, G2: fixed enum with an "other" escape hatch. When
`type == FilingType.OTHER`, the `type_other` free-text field is required.
"""

from __future__ import annotations

from datetime import date
from enum import Enum
from pathlib import Path

from pydantic import BaseModel, ConfigDict, Field, model_validator


class FilingType(str, Enum):
    """Locked enum of recognised filing types (G2)."""

    TEN_K = "10-K"
    TEN_Q = "10-Q"
    EIGHT_K = "8-K"
    PROXY = "proxy"
    S_1 = "S-1"
    TWENTY_F = "20-F"
    SIX_K = "6-K"
    AIF = "AIF"
    MDA = "MD&A"
    MANAGEMENT_CIRCULAR = "management-circular"
    ANNUAL_REPORT = "annual-report"
    ARS = "ARS"
    PROSPECTUS = "prospectus"
    PRESS_RELEASE = "press-release"
    TRANSCRIPT = "transcript"
    INVESTOR_DECK = "investor-deck"
    OTHER = "other"


class Filing(BaseModel):
    """A single filing or attachment associated with a deal.

    The path is the canonical on-disk location after the conductor saves the
    analyst-attached file into the deal directory.
    """

    model_config = ConfigDict(extra="forbid", str_strip_whitespace=True)

    type: FilingType = Field(..., description="Filing class. Use FilingType.OTHER for anything not enumerated.")
    type_other: str | None = Field(
        default=None,
        description="Free-text description, required when `type == FilingType.OTHER`.",
    )
    title: str | None = Field(default=None, description="Display title (e.g. 'OpenText FY2025 10-K').")
    filed_on: date | None = Field(default=None, description="Date the filing was originally filed.")
    period_end: date | None = Field(default=None, description="Fiscal period covered, where applicable.")
    source_url: str | None = Field(default=None, description="Source URL (SEDAR / EDGAR / IR site).")
    local_path: Path | None = Field(
        default=None,
        description="Path under the deal directory once persisted by the conductor.",
    )
    notes: str | None = Field(default=None, description="Free-form analyst notes about this filing.")

    @model_validator(mode="after")
    def _require_type_other_when_other(self) -> "Filing":
        if self.type == FilingType.OTHER:
            if not self.type_other or not self.type_other.strip():
                raise ValueError(
                    "When `type` is FilingType.OTHER, `type_other` must be a non-empty string."
                )
        return self
