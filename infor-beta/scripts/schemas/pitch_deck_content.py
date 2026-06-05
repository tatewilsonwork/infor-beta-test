"""Typed content handoff for the 14-slide INFOR slide-library POC deck."""

from __future__ import annotations

import re

from pydantic import BaseModel, ConfigDict, Field, field_validator, model_validator


_DATE_RE = re.compile(r"^[A-Z][a-z]+\s+\d{4}$")

# A financial-summary tile shows the metric NAME only — the (placeholder) chart
# carries the amount. Reject anything that looks value-laden: digits, currency or
# percent tokens, or a colon (as in "FY2025 Revenue: US$589.8MM (+31% YoY)").
_METRIC_VALUE_TOKEN_RE = re.compile(r"[\d$%:]")

# The market-entry comparison table is a fixed 12-row structure: three fixed
# top rows, seven industry-relevant metric rows chosen once per deck (and so
# identical across every target slide), then two fixed bottom rows.
_MARKET_ENTRY_FIXED_TOP = ("Overview", "Headquarters", "Year Founded")
_MARKET_ENTRY_FIXED_BOTTOM = ("Scale KPIs", "Strategic Rationale")
_MARKET_ENTRY_ROW_COUNT = 12


class PitchBullet(BaseModel):
    """Flexible bullet item used by the deck-content POC."""

    model_config = ConfigDict(extra="forbid", str_strip_whitespace=True)

    text: str = Field(..., min_length=1, max_length=220)
    level: int = Field(default=0, ge=0, le=1)


class RiskMitigantRow(BaseModel):
    """One acquirer risk with exactly three mitigants, each ~1 short sentence."""

    model_config = ConfigDict(extra="forbid", str_strip_whitespace=True)

    risk: str = Field(..., min_length=1, max_length=180)
    mitigants: list[str] = Field(..., min_length=3, max_length=3)

    @field_validator("mitigants")
    @classmethod
    def mitigants_are_concise(cls, values: list[str]) -> list[str]:
        for value in values:
            if not value.strip():
                raise ValueError("mitigants cannot be blank")
            # One very short sentence per mitigant; the cap stops paragraph-length
            # entries from overflowing the slide-9 table cell, not short sentences.
            if len(value) > 160:
                raise ValueError("each mitigant must stay to ~1 short sentence (<= 160 chars) to avoid PowerPoint overflow")
        return values


class InvestmentHighlight(BaseModel):
    """One numbered quadrant on the Key Investment Highlights slide."""

    model_config = ConfigDict(extra="forbid", str_strip_whitespace=True)

    header: str = Field(..., min_length=1, max_length=90)
    bullets: list[str] = Field(..., min_length=1, max_length=3)

    @field_validator("bullets")
    @classmethod
    def bullets_are_concise(cls, values: list[str]) -> list[str]:
        for value in values:
            if not value.strip():
                raise ValueError("investment-highlight bullets cannot be blank")
            if len(value) > 240:
                raise ValueError("investment-highlight bullets must stay concise to avoid overflow")
        return values


class MarketEntryTarget(BaseModel):
    """One target column on the Potential Market Entry Targets slide.

    `name` is the target company's name, used to label its deferred logo box as
    `[<name> Logo]`. `cells` are positional values aligned 1:1 with the fixed
    12-row `market_entry_row_labels` (Overview / HQ / Year Founded → 7 industry
    metrics → Scale KPIs / Strategic Rationale).
    """

    model_config = ConfigDict(extra="forbid", str_strip_whitespace=True)

    name: str | None = Field(
        default=None,
        max_length=80,
        description=(
            "Optional target company name. Labels the deferred logo box as "
            "'[<name> Logo]' (e.g. '[Kueski Logo]'); falls back to the generic "
            "'[Company Name Logo]' when absent."
        ),
    )
    cells: list[str] = Field(..., min_length=1, max_length=_MARKET_ENTRY_ROW_COUNT)


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
    financial_metric_labels: list[str] = Field(
        ...,
        min_length=4,
        max_length=4,
        description=(
            "Exactly four financial-summary tile labels — metric NAMES ONLY "
            "(e.g. 'Revenue', 'Adjusted EBITDA', 'Combined Loan Balances', "
            "'Adjusted Return on Equity'). No amounts, currency/percent tokens, "
            "colons, periods, or YoY deltas: the (placeholder) charts show the values."
        ),
    )
    risk_mitigants: list[RiskMitigantRow] = Field(..., min_length=1, max_length=5)
    risks_tagline: str = Field(..., min_length=1, max_length=180)
    comps_takeaway: str = Field(..., min_length=1, max_length=180)
    investment_highlights: list[InvestmentHighlight] = Field(default_factory=list, max_length=4)
    investment_highlights_tagline: str | None = Field(default=None, max_length=240)
    market_entry_market: str | None = Field(default=None, max_length=60)
    market_entry_row_labels: list[str] = Field(
        default_factory=list,
        max_length=_MARKET_ENTRY_ROW_COUNT,
        description=(
            "The fixed 12-row comparison labels, in order: 'Overview', "
            "'Headquarters', 'Year Founded', then exactly seven industry-relevant "
            "metric labels (chosen once for the deck and identical across every "
            "target slide), then 'Scale KPIs', 'Strategic Rationale'. Required "
            "whenever market_entry_targets is non-empty."
        ),
    )
    market_entry_targets: list[MarketEntryTarget] = Field(
        default_factory=list,
        max_length=8,
        description=(
            "Up to eight target companies. The deck lays them out two per slide "
            "(ceil(N/2) market-entry slides), each target's cells aligning 1:1 "
            "with the 12 market_entry_row_labels."
        ),
    )
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

    @field_validator("financial_metric_labels")
    @classmethod
    def metric_labels_are_names_only(cls, values: list[str]) -> list[str]:
        for value in values:
            if _METRIC_VALUE_TOKEN_RE.search(value):
                raise ValueError(
                    f"financial_metric_labels must be metric NAMES only, not a "
                    f"value-laden string like {value!r}. Drop the amount, currency "
                    f"and YoY delta (the chart shows them) — e.g. 'Revenue', "
                    f"'Adjusted EBITDA', 'Combined Loan Balances'."
                )
            if len(value) > 40:
                raise ValueError(
                    f"financial_metric_labels must stay short enough for a tile "
                    f"(<= 40 chars); {value!r} reads like a phrase, not a metric name."
                )
        return values

    @model_validator(mode="after")
    def current_section_must_be_listed(self) -> "PitchDeckContent":
        if self.current_section not in self.section_labels:
            raise ValueError("current_section must be one of section_labels")
        return self

    @model_validator(mode="after")
    def market_entry_structure(self) -> "PitchDeckContent":
        if self.market_entry_targets and not self.market_entry_row_labels:
            raise ValueError("market_entry_row_labels are required when market_entry_targets are provided")
        labels = self.market_entry_row_labels
        if labels:
            if len(labels) != _MARKET_ENTRY_ROW_COUNT:
                raise ValueError(
                    f"market_entry_row_labels must be exactly {_MARKET_ENTRY_ROW_COUNT} rows: "
                    f"{list(_MARKET_ENTRY_FIXED_TOP)}, then 7 industry-relevant metrics "
                    f"(consistent across every target slide), then {list(_MARKET_ENTRY_FIXED_BOTTOM)} "
                    f"(got {len(labels)})"
                )
            if tuple(labels[:3]) != _MARKET_ENTRY_FIXED_TOP:
                raise ValueError(
                    f"the first three market_entry_row_labels must be {list(_MARKET_ENTRY_FIXED_TOP)}"
                )
            if tuple(labels[-2:]) != _MARKET_ENTRY_FIXED_BOTTOM:
                raise ValueError(
                    f"the last two market_entry_row_labels must be {list(_MARKET_ENTRY_FIXED_BOTTOM)}"
                )
        for target in self.market_entry_targets:
            if len(target.cells) != len(labels):
                raise ValueError("each market-entry target's cells must align 1:1 with market_entry_row_labels")
        return self
