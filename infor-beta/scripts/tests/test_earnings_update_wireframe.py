"""Unit tests for earnings-update SlidePlan construction."""

from pathlib import Path

from earnings_update_wireframe import build_earnings_update_slide_plan, write_slide_plan
from schemas import Company, SlidePlan


def test_build_earnings_update_slide_plan_has_fixed_five_slide_structure():
    company = Company(legal_name="SampleCo", ticker="TSX:SMPL")

    plan = build_earnings_update_slide_plan(
        company=company,
        reporting_quarter="Q4 2025",
        comparison_quarter="Q4 2024",
    )

    assert isinstance(plan, SlidePlan)
    assert plan.deliverable_type == "earnings-update"
    assert plan.deck_title == "SampleCo Earnings Update"
    assert [s.library_entry_id for s in plan.slides] == [
        "earnings-update-cover",
        "earnings-update-company-overview",
        "earnings-update-earnings-summary",
        "earnings-update-disclaimer",
        "earnings-update-contact",
    ]
    assert [s.order for s in plan.slides] == [0, 1, 2, 3, 4]
    assert plan.slides[2].content_block["requires"] == [
        "kpi_rows",
        "business_updates",
        "broker_rows",
        "management_quotes",
        "performance_summary",
    ]


def test_write_slide_plan_round_trips_json(tmp_path: Path):
    plan = build_earnings_update_slide_plan(
        company=Company(legal_name="SampleCo"),
        reporting_quarter="Q4 2025",
        comparison_quarter="Q4 2024",
    )
    path = write_slide_plan(plan, tmp_path / "slide_plan.json")

    restored = SlidePlan.model_validate_json(path.read_text())
    assert restored == plan
