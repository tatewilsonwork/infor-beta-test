"""Helpers shared by the earnings-update and pitch wireframe stages.

Both wireframes derive a display name from a `Company` (or raw dict) and write a
typed `SlidePlan` JSON artefact the same way; these used to be byte-identical
copies in each module. They are re-exported from each wireframe module so the
existing ``from <wireframe> import write_slide_plan`` import path still works.

`slide_placement` is here for the same reason one level on: a content stage has to
say WHICH SLIDE a figure lands on for `deckcheck` to join it by identity, and the
answer is in the slide plan it was already handed — not in a number a SKILL.md
could write down, because the deck's slide mix is a deck-spec option.
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from provenance import DeckPlacement
from schemas import Company, SlidePlan


def load_slide_plan(path: Path | str) -> SlidePlan:
    """Read a typed `SlidePlan` back off disk — the inverse of `write_slide_plan`."""
    return SlidePlan.model_validate(json.loads(Path(path).read_text(encoding="utf-8")))


def slide_number_for(
    plan: SlidePlan | Path | str,
    library_entry_id: str,
    *,
    occurrence: int = 0,
) -> int:
    """The 1-based deck slide number a library entry lands on, from the plan.

    Both assemblers emit one deck slide per plan entry, in plan order, so a plan
    entry's ``order`` **is** its zero-based position in the finished deck. Resolving
    through the entry id rather than counting is the same rule as everything else
    here: the market-entry section can be one slide or four, the Financial Summary
    section one or two, and Key Investment Highlights may be absent entirely — so a
    written-down slide number is wrong for most decks the plan can describe.

    ``occurrence`` picks among repeated entries (market-entry slide 2 of 4).
    """
    resolved = plan if isinstance(plan, SlidePlan) else load_slide_plan(plan)
    entries = [e for e in resolved.slides if e.library_entry_id == library_entry_id]
    if not entries:
        raise KeyError(
            f"the slide plan has no {library_entry_id!r} entry (it holds "
            f"{sorted({e.library_entry_id for e in resolved.slides})}) — the deck spec "
            f"may have dropped that slide, in which case nothing lands on it."
        )
    if occurrence >= len(entries):
        raise IndexError(
            f"the slide plan holds {len(entries)} {library_entry_id!r} slide(s); asked for "
            f"occurrence {occurrence}"
        )
    return entries[occurrence].order + 1


def slide_placement(
    plan: SlidePlan | Path | str,
    library_entry_id: str,
    field: str,
    *,
    occurrence: int = 0,
) -> DeckPlacement:
    """A `DeckPlacement` for a typed content field: the plan's slide plus the field.

    What a content stage passes to ``ledger.record(..., placement=…)``. The field is
    the bundle path the figure sits in — ``"executive_summary_bullets[1]"``,
    ``"market_entry_targets[3].cells[10]"`` — which is what the assembler writes, so
    it stays right across a library re-layout that would move any shape name.
    """
    return DeckPlacement(
        slide=slide_number_for(plan, library_entry_id, occurrence=occurrence), field=field
    )


def company_display_name(company: Company | dict[str, Any]) -> str:
    """Best-effort display name from a `Company` model or a raw dict."""
    if isinstance(company, dict):
        return str(company.get("legal_name") or company.get("name") or "Company")
    return company.legal_name


def write_slide_plan(plan: SlidePlan, path: Path | str) -> Path:
    """Write a SlidePlan JSON artefact and return its path."""
    target = Path(path)
    target.parent.mkdir(parents=True, exist_ok=True)
    target.write_text(plan.model_dump_json(indent=2) + "\n", encoding="utf-8")
    return target
