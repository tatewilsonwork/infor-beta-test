"""Unit tests for the SlidePlan / SlideEntry schemas."""

import pytest
from pydantic import ValidationError

from schemas import SlideEntry, SlidePlan


def test_minimal_slide_entry():
    e = SlideEntry(library_entry_id="cover", title="Cover", order=0)
    assert e.library_entry_id == "cover"
    assert e.content_block == {}


def test_opaque_library_entry_id_accepted():
    """Per G5: no validators on `library_entry_id` content in Phase 1.
    Any non-empty string is fine — even nonsense / future IDs."""
    for entry_id in (
        "cover",
        "football-field",
        "this/is/not/real-yet",
        "abc-123-xyz",
        "company.profile.public",
    ):
        e = SlideEntry(library_entry_id=entry_id, title="x", order=0)
        assert e.library_entry_id == entry_id


def test_empty_library_entry_id_rejected():
    with pytest.raises(ValidationError):
        SlideEntry(library_entry_id="", title="x", order=0)


def test_order_must_be_non_negative():
    with pytest.raises(ValidationError):
        SlideEntry(library_entry_id="x", title="y", order=-1)


def test_slide_plan_round_trip():
    plan = SlidePlan(
        deliverable_type="earnings-update",
        deck_title="Project OpenText — Q4 FY25 Earnings Update",
        slides=[
            SlideEntry(library_entry_id="cover", title="Cover", order=0),
            SlideEntry(
                library_entry_id="financial-summary",
                title="Financial Summary",
                section="Results",
                order=1,
                content_block={"highlight": "Beat consensus on EPS by $0.04."},
            ),
        ],
        notes="Slim 5-slide layout.",
    )
    raw = plan.model_dump_json()
    plan2 = SlidePlan.model_validate_json(raw)
    assert plan2 == plan
    assert len(plan2.slides) == 2
    assert plan2.slides[1].content_block["highlight"].startswith("Beat consensus")


def test_layout_variant_is_rejected_after_i1_decision():
    """Per I1/I6: library entry IDs are concrete; layout_variant is dead code."""
    with pytest.raises(ValidationError):
        SlideEntry(library_entry_id="cover", title="Cover", order=0, layout_variant="two_col")


def test_empty_slides_allowed():
    """A skeleton plan with no slides yet is valid — it gets populated
    progressively by the wireframe / writing skills."""
    plan = SlidePlan(deliverable_type="cim", deck_title="Project Atlas")
    assert plan.slides == []


# ─── Resolving a placement from the plan ─────────────────────────────────────
#
# A content stage must say WHICH SLIDE a figure lands on for `deckcheck` to join it
# by identity, and the number cannot be written into a SKILL.md: the deck's slide
# mix is a deck-spec option (one Financial Summary slide or two, one market-entry
# slide or four, Key Investment Highlights present or dropped). Both assemblers
# emit one deck slide per plan entry, in order, so the plan's `order` is the answer.


def _plan(*entry_ids: str) -> SlidePlan:
    return SlidePlan(
        deliverable_type="pitch",
        deck_title="Deck",
        slides=[
            SlideEntry(library_entry_id=entry_id, title=entry_id, order=i)
            for i, entry_id in enumerate(entry_ids)
        ],
    )


def test_slide_number_comes_from_the_plans_order():
    from wireframe_common import slide_number_for

    plan = _plan("pitch-cover", "executive-summary", "public-company-overview")
    assert slide_number_for(plan, "executive-summary") == 2
    assert slide_number_for(plan, "public-company-overview") == 3


def test_a_dropped_slide_shifts_every_later_one():
    # Exactly the case a hardcoded number gets wrong: `include_investment_highlights`
    # is false, so market-entry moves up a slide.
    from wireframe_common import slide_number_for

    with_kih = _plan("pitch-cover", "key-investment-highlights", "market-entry-targets")
    without = _plan("pitch-cover", "market-entry-targets")
    assert slide_number_for(with_kih, "market-entry-targets") == 3
    assert slide_number_for(without, "market-entry-targets") == 2


def test_a_repeated_entry_is_picked_by_occurrence():
    from wireframe_common import slide_number_for

    plan = _plan("pitch-cover", "market-entry-targets", "market-entry-targets")
    assert slide_number_for(plan, "market-entry-targets") == 2
    assert slide_number_for(plan, "market-entry-targets", occurrence=1) == 3
    with pytest.raises(IndexError, match="occurrence 2"):
        slide_number_for(plan, "market-entry-targets", occurrence=2)


def test_an_entry_the_plan_does_not_carry_raises_naming_what_it_has():
    from wireframe_common import slide_number_for

    with pytest.raises(KeyError, match="key-investment-highlights"):
        slide_number_for(_plan("pitch-cover", "executive-summary"),
                         "key-investment-highlights")


def test_a_placement_carries_the_slide_and_the_typed_field():
    from provenance import DeckPlacement
    from wireframe_common import slide_placement

    plan = _plan("pitch-cover", "executive-summary")
    assert slide_placement(plan, "executive-summary", "executive_summary_bullets[1]") == (
        DeckPlacement(slide=2, field="executive_summary_bullets[1]")
    )


def test_a_placement_resolves_from_a_written_plan_file(tmp_path):
    # What the content stage actually holds is `slide_plan_path`, not the object.
    from wireframe_common import slide_placement, write_slide_plan

    path = write_slide_plan(_plan("pitch-cover", "executive-summary"), tmp_path / "plan.json")
    placement = slide_placement(path, "executive-summary", "executive_summary_bullets[0]")
    assert placement.slide == 2
