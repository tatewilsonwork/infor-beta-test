"""Typed content handoff for the Phase 3 earnings-update POC.

The wireframe stage emits a generic SlidePlan. The content stage fills the
specific fields the earnings-update deck assembler needs and serialises this
model as the typed stage artefact.
"""

from __future__ import annotations

import re
from typing import Literal

from pydantic import BaseModel, ConfigDict, Field, field_validator, model_validator

VarianceSign = Literal[-1, 0, 1]

# Company-overview bullet budget, calibrated against the shared library's
# overview slide geometry in the Test #5 (OTEX) live-run fix (2026-05-29): the
# text box sits directly above the "LTM Revenue Breakdown" header, so the
# budget is tighter than the legacy full-height block. The SKILL.md prose and
# these constants must quote the same numbers — a drift-lock test in
# test_earnings_update_content.py parses the SKILL.md line and asserts it.
OVERVIEW_BULLETS_MIN = 6
OVERVIEW_BULLETS_MAX = 8
OVERVIEW_BULLET_MAX_CHARS = 220
OVERVIEW_TOTAL_MIN_CHARS = 560
OVERVIEW_TOTAL_MAX_CHARS = 820

_ZERO_WORDS = {"flat", "unchanged", "unch"}


def _display_sign(value: str) -> int | None:
    """Infer the sign a delta/variance display string reads as, when unambiguous.

    Returns ``1`` for a leading ``+``, ``-1`` for a leading ``-``/``−`` or
    accounting parentheses, and ``0`` for "flat"/zero figures — after stripping
    ``$``/``%``/thousands separators and unit words (``bps``, ``pts``). An
    unsigned magnitude (e.g. ``"12.5%"``) returns ``None``: no sign is legible
    from the string, so the declared sign field is the only signal.
    """
    s = value.strip()
    if not s:
        return None
    if s.rstrip(".").lower() in _ZERO_WORDS:
        return 0
    cleaned = re.sub(r"(?i)\b(bps|bp|pts|pp)\b", "", s)
    cleaned = cleaned.replace("$", "").replace("%", "").replace(",", "").strip()
    negative = positive = False
    if cleaned.startswith("(") and cleaned.endswith(")"):
        negative = True
        cleaned = cleaned[1:-1].strip()
    if cleaned[:1] in "+-−":
        positive = cleaned[0] == "+"
        negative = negative or cleaned[0] in "-−"
        cleaned = cleaned[1:].strip()
    try:
        magnitude = float(cleaned)
    except ValueError:
        magnitude = None
    if magnitude == 0:
        return 0
    if negative:
        return -1
    if positive:
        return 1
    return None


class CompanyOverviewBullet(BaseModel):
    """One bullet for slide 2's company overview block."""

    model_config = ConfigDict(extra="forbid", str_strip_whitespace=True)

    bold_prefix: str | None = Field(default=None, description="Optional bold prefix such as a segment name.")
    text: str = Field(..., min_length=1, max_length=OVERVIEW_BULLET_MAX_CHARS, description="Bullet body text, without terminal period/semicolon.")
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

    @model_validator(mode="after")
    def _delta_sign_matches_display(self) -> "KpiRow":
        # The assembler colors the tile purely off delta_sign, so a sign that
        # contradicts what the display string reads as would paint a beat red
        # (or a miss green). Enforce agreement whenever the string's sign is
        # legible; an unsigned magnitude is accepted as-is.
        inferred = _display_sign(self.delta_str)
        if inferred is not None and inferred != self.delta_sign:
            raise ValueError(
                f"delta_sign={self.delta_sign} contradicts delta_str "
                f"{self.delta_str!r} (which reads as sign {inferred})"
            )
        return self


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

    @model_validator(mode="after")
    def _variance_sign_matches_display(self) -> "BrokerRow":
        # Same contract as KpiRow: the assembler colors the variance cell purely
        # off variance_sign, so it must agree with any sign the display string
        # carries (leading +/-, accounting parentheses, flat/zero forms).
        inferred = _display_sign(self.variance)
        if inferred is not None and inferred != self.variance_sign:
            raise ValueError(
                f"variance_sign={self.variance_sign} contradicts variance "
                f"{self.variance!r} (which reads as sign {inferred})"
            )
        return self


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
    company_overview_bullets: list[CompanyOverviewBullet] = Field(
        ..., min_length=OVERVIEW_BULLETS_MIN, max_length=OVERVIEW_BULLETS_MAX
    )
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
        # Tighter budget than the legacy full-height block: library slide 7
        # reserves the lower-left quadrant for the LTM revenue pie, so the
        # overview text occupies a shorter region and must stay concise. The
        # range is the Test #5 (OTEX) calibration — see the module constants.
        total_chars = sum(len(((b.bold_prefix or "") + b.text)) for b in self.company_overview_bullets)
        if not OVERVIEW_TOTAL_MIN_CHARS <= total_chars <= OVERVIEW_TOTAL_MAX_CHARS:
            raise ValueError(
                f"company overview bullets must total "
                f"{OVERVIEW_TOTAL_MIN_CHARS}-{OVERVIEW_TOTAL_MAX_CHARS} characters"
            )
        return self
