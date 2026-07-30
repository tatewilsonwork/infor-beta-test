"""Typed content handoff for the 16-slide INFOR slide-library POC deck."""

from __future__ import annotations

import re

from pydantic import BaseModel, ConfigDict, Field, field_validator, model_validator


_DATE_RE = re.compile(r"^[A-Z][a-z]+\s+\d{4}$")

# An ISO 4217 alphabetic code. The bundle states its currency as a code, not as a
# rendered footnote token ('US$MM'), because the deck renders the token from the
# code and two slides rendering it differently is exactly the defect the field
# exists to close.
_CURRENCY_RE = re.compile(r"^[A-Z]{3}$")

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
    """One numbered quadrant on the Key Investment Highlights slide.

    At most TWO bullets per quadrant (analyst-locked; three crowded the boxes).
    """

    model_config = ConfigDict(extra="forbid", str_strip_whitespace=True)

    header: str = Field(..., min_length=1, max_length=90)
    bullets: list[str] = Field(..., min_length=1, max_length=2)

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


class PitchContact(BaseModel):
    """One banker card on the deck's Contact slide.

    The library ships a 2x2 grid of three-row cards — ``"<name>, <title>"``, then
    the phone, then the email. INFOR is single-tenant and the deal team is
    knowable, so the cards are a declared input; a card the deck has no contact
    for is **deleted**, never left as the template's ``[x]``. A deck reaching a
    client with three empty contact cards reads as unfinished.
    """

    model_config = ConfigDict(extra="forbid", str_strip_whitespace=True)

    name: str = Field(..., min_length=1, max_length=80)
    title: str = Field(..., min_length=1, max_length=80)
    phone: str = Field(..., min_length=1, max_length=40)
    email: str = Field(..., min_length=1, max_length=80)


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
    figure_currency: str = Field(
        ...,
        description=(
            "ISO 4217 code for the currency every figure in THIS bundle is stated "
            "in — the target's filing reporting currency, which is also what the "
            "financial-summary and ltm-metrics tabs are locked to. The deck labels "
            "each slide from whatever populates it, so this is the label on every "
            "slide the bundle's own copy fills. It is NOT necessarily the cap "
            "table's output currency: the cap table converts, and where the two "
            "differ the overview slide names both plus the rate between them."
        ),
    )
    executive_summary_bullets: list[PitchBullet] = Field(..., min_length=1, max_length=12)
    section_labels: list[str] = Field(..., min_length=1, max_length=8)
    current_section: str = Field(..., min_length=1, max_length=80)
    company_overview_bullets: list[PitchBullet] = Field(..., min_length=1, max_length=10)
    risk_mitigants: list[RiskMitigantRow] = Field(..., min_length=1, max_length=5)
    risks_tagline: str = Field(..., min_length=1, max_length=180)
    comps_takeaway: str = Field(..., min_length=1, max_length=180)
    precedents_takeaway: str = Field(..., min_length=1, max_length=180)
    ownership_takeaway: str = Field(
        ...,
        min_length=1,
        max_length=180,
        description=(
            "One sentence on the ownership slide's takeaway line — the sibling of "
            "comps_takeaway / precedents_takeaway. Required: the slide ships a "
            "takeaway box, and with no field to fill it the delivered deck printed "
            "a bare '[x]' under a table of 24 insiders and 118 institutions."
        ),
    )
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
    contacts: list[PitchContact] = Field(
        default_factory=list,
        max_length=4,
        description=(
            "Deal-team cards for the Contact slide, in reading order (the library "
            "ships a 2x2 grid of four). Left empty, the deck keeps whichever cards "
            "the library ships already filled and DELETES the placeholder ones — "
            "so the default is INFOR's own card rather than a name invented here."
        ),
    )
    sources: list[PitchSourceNote] = Field(default_factory=list)
    manual_steps: list[str] = Field(default_factory=list)

    @field_validator("figure_currency")
    @classmethod
    def currency_is_iso_code(cls, value: str) -> str:
        code = value.strip().upper()
        if not _CURRENCY_RE.match(code):
            raise ValueError(
                "figure_currency must be an ISO 4217 alphabetic code, e.g. 'USD' or "
                f"'CAD' — not a rendered footnote token like 'US$MM' (got {value!r})"
            )
        return code

    @field_validator("presentation_date")
    @classmethod
    def date_is_month_year(cls, value: str) -> str:
        if not _DATE_RE.match(value):
            raise ValueError("presentation_date must be fully spelled-out month plus four-digit year, e.g. 'April 2026'")
        return value

    @field_validator("section_labels", "manual_steps")
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
