"""Typed content handoff for the 12-slide INFOR slide-library POC deck."""

from __future__ import annotations

import re

from pydantic import BaseModel, ConfigDict, Field, field_validator, model_validator


_DATE_RE = re.compile(r"^[A-Z][a-z]+\s+\d{4}$")


class PitchBullet(BaseModel):
    """Flexible bullet item used by the deck-content POC."""

    model_config = ConfigDict(extra="forbid", str_strip_whitespace=True)

    text: str = Field(..., min_length=1, max_length=220)
    level: int = Field(default=0, ge=0, le=1)


class RiskMitigantRow(BaseModel):
    """One acquirer risk with exactly three concise mitigants."""

    model_config = ConfigDict(extra="forbid", str_strip_whitespace=True)

    risk: str = Field(..., min_length=1, max_length=180)
    mitigants: list[str] = Field(..., min_length=3, max_length=3)

    @field_validator("mitigants")
    @classmethod
    def mitigants_are_concise(cls, values: list[str]) -> list[str]:
        for value in values:
            if not value.strip():
                raise ValueError("mitigants cannot be blank")
            if len(value) > 90:
                raise ValueError("mitigants must stay concise to avoid PowerPoint overflow")
        return values


class PitchSourceNote(BaseModel):
    """Source note for analyst-auditable citations."""

    model_config = ConfigDict(extra="forbid", str_strip_whitespace=True)

    section: str = Field(..., min_length=1, max_length=80)
    citation: str = Field(..., min_length=1, max_length=240)


class PitchDeckContent(BaseModel):
    """Single broad typed content bundle for the slide-library POC deck."""

    model_config = ConfigDict(extra="forbid", str_strip_whitespace=True)

    client_name: str = Field(..., min_length=1, max_length=120)
    presentation_date: str = Field(
        ...,
        description="Fully spelled-out month plus four-digit year, e.g. 'April 2026'.",
    )
    executive_summary_bullets: list[PitchBullet] = Field(..., min_length=1, max_length=12)
    section_labels: list[str] = Field(..., min_length=1, max_length=8)
    current_section: str = Field(..., min_length=1, max_length=80)
    company_overview_bullets: list[PitchBullet] = Field(..., min_length=1, max_length=10)
    financial_metric_labels: list[str] = Field(..., min_length=4, max_length=4)
    risk_mitigants: list[RiskMitigantRow] = Field(..., min_length=1, max_length=5)
    risks_tagline: str = Field(..., min_length=1, max_length=180)
    comps_takeaway: str = Field(..., min_length=1, max_length=180)
    sources: list[PitchSourceNote] = Field(default_factory=list)
    manual_steps: list[str] = Field(default_factory=list)

    @field_validator("presentation_date")
    @classmethod
    def date_is_month_year(cls, value: str) -> str:
        if not _DATE_RE.match(value):
            raise ValueError("presentation_date must be fully spelled-out month plus four-digit year, e.g. 'April 2026'")
        return value

    @field_validator("section_labels", "financial_metric_labels", "manual_steps")
    @classmethod
    def strings_not_blank(cls, values: list[str]) -> list[str]:
        if any(not value.strip() for value in values):
            raise ValueError("list values cannot be blank")
        return values

    @model_validator(mode="after")
    def current_section_must_be_listed(self) -> "PitchDeckContent":
        if self.current_section not in self.section_labels:
            raise ValueError("current_section must be one of section_labels")
        return self
