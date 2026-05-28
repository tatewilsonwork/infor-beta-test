"""TDD tests for the 14-slide INFOR slide-library POC."""

from pathlib import Path

import yaml
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pydantic import ValidationError
import pytest

from schemas import Company, PitchDeckContent, Plan
from pitch_deck_wireframe import build_pitch_deck_slide_plan, write_slide_plan
from pitch_deck_assembler import assemble_pitch_deck
from slide_library_registry import load_slide_library_registry


PLUGIN_ROOT = Path(__file__).resolve().parents[2]
TEMPLATE = PLUGIN_ROOT / "templates" / "INFOR Slide Library.pptx"


def _sample_content() -> PitchDeckContent:
    return PitchDeckContent(
        client_name="SampleCo Ltd.",
        presentation_date="April 2026",
        executive_summary_bullets=[
            {"text": "SampleCo is a compelling public-company advisory candidate supported by durable market positions and recurring revenue visibility", "level": 0},
            {"text": "The Company benefits from diversified end-market exposure and a management team focused on profitable growth", "level": 0},
            {"text": "Potential acquirors are expected to focus on scale, margin durability and cash conversion", "level": 0},
            {"text": "INFOR is well positioned to help evaluate strategic alternatives and prepare the Company for a targeted process", "level": 0},
        ],
        section_labels=["Overview", "Financial Summary", "Valuation", "Process"],
        current_section="Overview",
        company_overview_bullets=[
            {"text": "Provides mission-critical software and services to enterprise customers across North America", "level": 0},
            {"text": "Recurring revenue model supported by multi-year contracts and high customer retention", "level": 0},
            {"text": "Publicly listed issuer with established reporting history and access to capital markets", "level": 0},
        ],
        financial_metric_labels=["Revenue", "Gross Margin", "Adjusted EBITDA", "Free Cash Flow"],
        risk_mitigants=[
            {
                "risk": "Acquirors may question the durability of organic growth as market conditions normalize",
                "mitigants": [
                    "Highlight recurring revenue base",
                    "Show multi-year retention trends",
                    "Frame growth by segment",
                ],
            },
            {
                "risk": "Scale relative to larger public peers may affect valuation expectations",
                "mitigants": [
                    "Benchmark to focused peers",
                    "Emphasize margin profile",
                    "Position scarcity value",
                ],
            },
            {
                "risk": "Buyers may diligence customer concentration and renewal quality",
                "mitigants": [
                    "Prepare cohort support",
                    "Show renewal history",
                    "Document top-customer exposure",
                ],
            },
        ],
        risks_tagline="INFOR will help proactively frame key diligence topics to support a constructive acquiror dialogue.",
        comps_takeaway="SampleCo trades at a discount to higher-growth peers, creating a clear valuation framing opportunity.",
        investment_highlights=[
            {"header": "Durable Recurring Revenue", "bullets": ["High retention across multi-year contracts", "Predictable cash conversion through cycles"]},
            {"header": "Fragmented, Underpenetrated Market", "bullets": ["Few scaled competitors", "Clear whitespace for consolidation"]},
            {"header": "Strategic Buyer Interest", "bullets": ["Active strategic and sponsor consolidation", "Scarcity value for scaled assets"]},
            {"header": "Multiple Expansion Paths", "bullets": ["Organic, greenfield and M&A optionality", "Growth not reliant on any single lever"]},
        ],
        investment_highlights_tagline="SampleCo offers a scarce, scaled platform in a structurally attractive market with multiple value-creation paths.",
        market_entry_market="Canada",
        market_entry_row_labels=["Overview", "Headquarters", "Year Founded", "Product / Channel", "Target Segment", "Funding Model", "Scale KPIs", "Strategic Rationale"],
        market_entry_targets=[
            {"cells": ["Digital lender", "Toronto, ON", "2014", "Online & retail", "Prime / near-prime", "Equity and debt facilities", ">1MM customers", "Scaled platform with brand recognition"]},
            {"cells": ["Mobile-first lender", "Vancouver, BC", "2017", "Direct-to-consumer app", "Thin-file", "Venture capital and debt warehouse", "~0.5MM customers", "Actionable platform subject to diligence"]},
        ],
        sources=[{"section": "Company Overview", "citation": "Company filings, company website and analyst notes"}],
        manual_steps=["LTM revenue breakdown and financial summary charts remain placeholders in this POC."],
    )


def test_pitch_deck_content_schema_validates_concise_risk_mitigants():
    content = _sample_content()

    assert content.presentation_date == "April 2026"
    assert content.risk_mitigants[0].mitigants == [
        "Highlight recurring revenue base",
        "Show multi-year retention trends",
        "Frame growth by segment",
    ]

    bad = content.model_dump()
    bad["risk_mitigants"][0]["mitigants"] = ["one", "two"]
    with pytest.raises(ValidationError):
        PitchDeckContent.model_validate(bad)


def test_registry_loads_14_blank_slide_library_entries():
    entries = load_slide_library_registry()

    assert len(entries) == 14
    assert entries[0].library_entry_id == "pitch-cover"
    assert entries[6].library_entry_id == "public-company-overview"
    assert entries[10].library_entry_id == "key-investment-highlights"
    assert entries[11].library_entry_id == "market-entry-targets"
    assert entries[10].static is False
    assert entries[11].static is False
    assert entries[12].library_entry_id == "disclaimer"
    assert entries[12].static is True
    assert entries[13].library_entry_id == "contact"
    assert entries[13].static is True


def test_pitch_deck_wireframe_uses_blank_library_order():
    plan = build_pitch_deck_slide_plan(
        company=Company(legal_name="SampleCo Ltd.", ticker="TSX:SMP"),
        section_labels=["Overview", "Financial Summary", "Valuation", "Process"],
    )

    assert plan.deliverable_type == "pitch"
    assert len(plan.slides) == 14
    assert [slide.library_entry_id for slide in plan.slides[10:12]] == [
        "key-investment-highlights",
        "market-entry-targets",
    ]
    assert plan.slides[10].content_block["requires"] == ["investment_highlights"]
    assert plan.slides[11].content_block["requires"] == ["market_entry_row_labels", "market_entry_targets"]
    assert [slide.library_entry_id for slide in plan.slides[:8]] == [
        "pitch-cover",
        "executive-summary",
        "infor-overview",
        "infor-ma-advisor",
        "infor-key-highlights",
        "section-divider",
        "public-company-overview",
        "financial-summary",
    ]
    assert plan.slides[2].content_block["template_static"] is True
    assert plan.slides[6].content_block["requires"] == ["company_overview_bullets", "cap_table"]


def test_assemble_pitch_deck_preserves_static_slides_and_fills_allowed_fields(tmp_path: Path):
    plan = build_pitch_deck_slide_plan(
        company=Company(legal_name="SampleCo Ltd.", ticker="TSX:SMP"),
        section_labels=["Overview", "Financial Summary", "Valuation", "Process"],
    )
    plan_path = write_slide_plan(plan, tmp_path / "slide_plan.json")
    content = _sample_content()
    content_path = tmp_path / "content.json"
    content_path.write_text(content.model_dump_json(indent=2), encoding="utf-8")

    deck_path = assemble_pitch_deck(
        slide_plan_path=plan_path,
        content_path=content_path,
        template_path=TEMPLATE,
        output_dir=tmp_path,
    )

    assert deck_path.exists()
    prs = Presentation(deck_path)
    assert len(prs.slides) == 14
    text_parts = []

    def _collect(shapes):
        for shape in shapes:
            if getattr(shape, "has_text_frame", False):
                text_parts.append(shape.text)
            if getattr(shape, "has_table", False):
                for row in shape.table.rows:
                    for cell in row.cells:
                        text_parts.append(cell.text)
            if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
                _collect(shape.shapes)

    for slide in prs.slides:
        _collect(slide.shapes)
    all_text = "\n".join(text_parts)
    assert "[CLIENT NAME]" not in all_text
    assert "[Date]" not in all_text
    assert "SampleCo Ltd." in all_text
    assert "April 2026" in all_text
    assert "[Cap Table Placeholder]" in all_text  # insertion skill fills this later
    assert "[Pie Chart Placeholder]" in all_text  # intentionally deferred
    assert "[Placeholder for Metric #1 Chart]" in all_text  # intentionally deferred
    assert "Neil Selfe, Managing Principal" in all_text

    # New content slides are filled and their placeholders replaced.
    assert "Durable Recurring Revenue" in all_text
    assert "SampleCo offers a scarce, scaled platform" in all_text
    assert "Potential Canada Market Entry Targets" in all_text
    assert "[Investment Highlight #1]" not in all_text
    assert "Potential [x] Market Entry Targets" not in all_text
    # Logo placeholders on the market-entry slide remain deferred.
    assert "[Placeholder for Logo]" in all_text

    # Static credential slide text remains present.
    assert "INFOR is Canada’s leading provider of innovative, independent, forward thinking financial & strategic advice" in all_text
    assert "These materials are confidential and proprietary" in all_text


def test_earnings_update_plan_runs_captable_before_deck_for_insertion():
    plan_path = PLUGIN_ROOT / "plans" / "earnings-update.yaml"
    plan = Plan.model_validate(yaml.safe_load(plan_path.read_text(encoding="utf-8")))

    assert [stage.id for stage in plan.stages] == [
        "wireframe",
        "content",
        "captable",
        "ltm-revenue",
        "deck",
    ]
    deck_stage = next(s for s in plan.stages if s.id == "deck")
    assert deck_stage.inputs["captable_workbook_path"] == "$stages.captable.workbook_path"
    assert deck_stage.inputs["template_name"] == "INFOR Slide Library.pptx"


def test_pitch_library_poc_plan_stage_order():
    plan_path = PLUGIN_ROOT / "plans" / "pitch-library-poc.yaml"
    plan = Plan.model_validate(yaml.safe_load(plan_path.read_text(encoding="utf-8")))

    assert [stage.id for stage in plan.stages] == ["wireframe", "content", "captable", "deck"]
    assert plan.stages[0].skill == "pitch-wireframe-infor"
    assert plan.stages[1].skill == "pitch-content-infor"
    assert plan.stages[2].skill == "captable-infor"
    assert plan.stages[3].skill == "deck-assembler"
    assert plan.stages[3].inputs["slide_plan_path"] == "$stages.wireframe.slide_plan_path"
    assert plan.stages[3].inputs["content_bundle_path"] == "$stages.content.content_bundle_path"
    assert plan.stages[3].inputs["captable_workbook_path"] == "$stages.captable.workbook_path"
