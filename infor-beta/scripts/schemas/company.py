"""Company schema — Phase 1.

Designed from public-filing-derived fields only per Obsidian note 12, G1.
No CapIQ fields — those will be added additively when the connector lands.
"""

from __future__ import annotations

from pydantic import BaseModel, ConfigDict, Field


class Company(BaseModel):
    """A subject company. Sourced from 10-K cover-page-class fields + analyst input.

    Only `legal_name` is required. Everything else is optional so a profile can be
    built up progressively during deal-init.
    """

    model_config = ConfigDict(extra="forbid", str_strip_whitespace=True)

    legal_name: str = Field(
        ...,
        min_length=1,
        description="Full legal name as it appears on the cover page of the latest annual filing.",
    )
    hq: str | None = Field(
        default=None,
        description="Headquarters location, free-form (e.g. 'Waterloo, Ontario, Canada').",
    )
    jurisdiction_of_incorporation: str | None = Field(
        default=None,
        description="Where the entity is incorporated (e.g. 'Ontario', 'Delaware', 'England and Wales').",
    )
    ticker: str | None = Field(
        default=None,
        description="Primary trading symbol if public (e.g. 'OTEX'). None for private companies.",
    )
    exchange: str | None = Field(
        default=None,
        description="Primary listing exchange if public (e.g. 'NASDAQ', 'TSX').",
    )
    fy_end: str | None = Field(
        default=None,
        description="Fiscal year end as MM-DD, e.g. '06-30', or a free-form description.",
    )
    employees: int | None = Field(
        default=None,
        ge=0,
        description="Approximate full-time-equivalent headcount.",
    )
    revenue_range: str | None = Field(
        default=None,
        description="Free-form revenue indicator (e.g. '$1B–$5B', 'Not disclosed').",
    )
    sector: str | None = Field(default=None, description="High-level sector classification.")
    industry: str | None = Field(default=None, description="Narrower industry classification.")
    notes: str | None = Field(
        default=None,
        description="Free-form analyst notes that do not fit the structured fields.",
    )
