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
                layout_variant="two_col",
            ),
        ],
        notes="Slim 5-slide layout.",
    )
    raw = plan.model_dump_json()
    plan2 = SlidePlan.model_validate_json(raw)
    assert plan2 == plan
    assert len(plan2.slides) == 2
    assert plan2.slides[1].content_block["highlight"].startswith("Beat consensus")


def test_empty_slides_allowed():
    """A skeleton plan with no slides yet is valid — it gets populated
    progressively by the wireframe / writing skills."""
    plan = SlidePlan(deliverable_type="cim", deck_title="Project Atlas")
    assert plan.slides == []
