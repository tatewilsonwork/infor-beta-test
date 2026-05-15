"""SlidePlan schema — Phase 1.

Per Obsidian note 12, G5: `library_entry_id` is an opaque string in Phase 1.
The slide library does not yet exist, so we ship no validators on its content.
Validators (registry lookup) get added in Phase 3 once the library is real.

Markdown views of plans are kept as analyst-readable previews, but the
conductor and deck-assembler consume the typed object.
"""

from __future__ import annotations

from typing import Any

from pydantic import BaseModel, ConfigDict, Field


class SlideEntry(BaseModel):
    """A single slide within a SlidePlan.

    `library_entry_id` is opaque in Phase 1 — any non-empty string is accepted.
    `content_block` is a free-form dict so wireframe-style skills can stash
    placeholder data without locking the schema before the library exists.
    """

    model_config = ConfigDict(extra="forbid", str_strip_whitespace=True)

    library_entry_id: str = Field(
        ...,
        min_length=1,
        description="Opaque ID referencing an entry in the slide library (validated in Phase 3).",
    )
    title: str = Field(..., min_length=1, description="Slide title as it should appear in the deck.")
    section: str | None = Field(
        default=None,
        description="Optional section grouping (e.g. 'Executive Summary', 'Valuation').",
    )
    order: int = Field(..., ge=0, description="Zero-based position within the SlidePlan.")
    content_block: dict[str, Any] = Field(
        default_factory=dict,
        description="Free-form placeholder data passed to the assembler / library entry.",
    )
    layout_variant: str | None = Field(
        default=None,
        description="Optional variant key per C2 (e.g. 'two_col', 'three_col').",
    )


class SlidePlan(BaseModel):
    """A typed plan for a deck. Markdown view is a downstream rendering of this object."""

    model_config = ConfigDict(extra="forbid", str_strip_whitespace=True)

    deliverable_type: str = Field(
        ...,
        min_length=1,
        description="Free-form deliverable label (e.g. 'cim', 'earnings-update', 'pitch').",
    )
    deck_title: str = Field(..., min_length=1, description="Top-level deck title.")
    slides: list[SlideEntry] = Field(default_factory=list, description="Ordered slides in the plan.")
    notes: str | None = Field(default=None, description="Optional plan-level notes for the assembler / analyst.")
