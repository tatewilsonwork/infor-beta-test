"""Typed content handoff for the Phase 3 earnings-update POC.

The wireframe stage emits a generic SlidePlan. The content stage fills the
specific fields the earnings-update deck assembler needs and serialises this
model as the typed stage artefact.
"""

from __future__ import annotations

from typing import Literal

from pydantic import BaseModel, ConfigDict, Field, field_validator, model_validator

VarianceSign = Literal[-1, 0, 1]


class CompanyOverviewBullet(BaseModel):
    """One bullet for slide 2's company overview block."""

    model_config = ConfigDict(extra="forbid", str_strip_whitespace=True)

    bold_prefix: str | None = Field(default=None, description="Optional bold prefix such as a segment name.")
    text: str = Field(..., min_length=1, max_length=250, description="Bullet body text, without terminal period/semicolon.")
    level: int = Field(default=0, ge=0, le=2, description="Bullet indentation level; 0 main, 1 sub-bullet.")

    @field_validator("text")
    @classmethod
    def _no_terminal_period_or_semicolon(cls, value: str) -> str:
        if value.rstrip().endswith((".", ";")):
            raise ValueError("company overview bullets must not end with a period or semicolon")
        return value


class KpiRow(BaseModel):
    """One KPI tile row on slide 3."""

    model_config = ConfigDict(extra="forbid", str_strip_whitespace=True)

    name: str = Field(..., min_length=1)
    prior_value: str = Field(..., min_length=1)
    current_value: str = Field(..., min_length=1)
    delta_str: str = Field(..., min_length=1)
    delta_sign: VarianceSign = Field(..., description="Positive=green, negative=red, zero=neutral.")

    @field_validator("delta_str")
    @classmethod
    def _rate_deltas_are_percent_not_bps(cls, value: str) -> str:
        if "bps" in value.lower():
            raise ValueError("rate deltas must be expressed as percentages, not bps")
        return value


class BrokerRow(BaseModel):
    """One row in the Broker Estimates vs Actuals table."""

    model_config = ConfigDict(extra="forbid", str_strip_whitespace=True)

    label: str = Field(..., min_length=1)
    reported: str = Field(..., min_length=1)
    estimate: str = Field(..., min_length=1)
    variance: str = Field(..., min_length=1)
    variance_sign: VarianceSign = Field(..., description="Positive=green, negative=red, zero=neutral.")

    @field_validator("reported", "estimate", "variance")
    @classmethod
    def _no_na_cells(cls, value: str) -> str:
        if value.strip().lower() in {"n/a", "na", "-"}:
            raise ValueError("broker table cells must not be N/A / NA / dash")
        return value


class ManagementQuote(BaseModel):
    """One management quote block on slide 3."""

    model_config = ConfigDict(extra="forbid", str_strip_whitespace=True)

    quote: str = Field(..., min_length=1, max_length=200)
    speaker: str = Field(..., min_length=1)
    role: str = Field(..., min_length=1)

    @field_validator("quote")
    @classmethod
    def _quote_word_cap(cls, value: str) -> str:
        if len(value.split()) > 30:
            raise ValueError("management quote must be <= 30 words")
        return value


class SourceNote(BaseModel):
    """Simple source citation for a content block."""

    model_config = ConfigDict(extra="forbid", str_strip_whitespace=True)

    section: str = Field(..., min_length=1)
    citation: str = Field(..., min_length=1)


class EarningsUpdateContent(BaseModel):
    """Typed content bundle consumed by the POC deck assembler."""

    model_config = ConfigDict(extra="forbid", str_strip_whitespace=True)

    company_name: str = Field(..., min_length=1)
    ticker: str | None = None
    reporting_quarter: str = Field(..., min_length=1)
    comparison_quarter: str = Field(..., min_length=1)
    currency: str = Field(..., min_length=1, description="Footnote currency, e.g. C$MM or US$MM.")
    currency_short: str = Field(..., min_length=1, description="Broker-table currency header.")
    cover_date: str = Field(..., min_length=1, description="Current month/year shown on cover.")
    company_overview_bullets: list[CompanyOverviewBullet] = Field(..., min_length=6, max_length=10)
    business_updates: list[str] = Field(..., min_length=4, max_length=6)
    kpi_rows: list[KpiRow] = Field(..., min_length=4, max_length=4)
    broker_rows: list[BrokerRow] = Field(..., min_length=5, max_length=5)
    management_quotes: list[ManagementQuote] = Field(..., min_length=2, max_length=2)
    performance_summary: str = Field(..., min_length=1, max_length=150)
    sources: list[SourceNote] = Field(default_factory=list)
    manual_steps: list[str] = Field(default_factory=list)

    @field_validator("business_updates")
    @classmethod
    def _business_update_caps(cls, values: list[str]) -> list[str]:
        for value in values:
            if len(value) > 250:
                raise ValueError("each business update must be <= 250 characters")
            if value.rstrip().endswith((".", ";")):
                raise ValueError("business updates must not end with a period or semicolon")
        if sum(len(v) for v in values) > 900:
            raise ValueError("business updates must be <= 900 characters in aggregate")
        return values

    @field_validator("performance_summary")
    @classmethod
    def _summary_word_cap(cls, value: str) -> str:
        if len(value.split()) > 25:
            raise ValueError("performance summary must be <= 25 words")
        return value

    @model_validator(mode="after")
    def _company_overview_caps(self) -> "EarningsUpdateContent":
        # Tighter budget than the legacy full-height block: library slide 7 now
        # reserves the lower-left quadrant for the LTM revenue pie, so the
        # overview text occupies a shorter region and must stay concise.
        total_chars = sum(len(((b.bold_prefix or "") + b.text)) for b in self.company_overview_bullets)
        if not 650 <= total_chars <= 1050:
            raise ValueError("company overview bullets must total 650-1,050 characters")
        return self
