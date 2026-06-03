"""TDD tests for the 14-slide INFOR slide-library POC."""

from pathlib import Path

import yaml
from openpyxl import Workbook
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.util import Pt
from pydantic import ValidationError
import pytest

from schemas import Company, PitchDeckContent, Plan
from pptx_helpers import find_shape
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
        market_entry_row_labels=[
            "Overview",
            "Headquarters",
            "Year Founded",
            "Lending Product",
            "Customer Segment",
            "Funding Model",
            "Credit Approach",
            "Geographic Footprint",
            "Regulatory Status",
            "Technology Platform",
            "Scale KPIs",
            "Strategic Rationale",
        ],
        market_entry_targets=[
            {"cells": ["Digital lender", "Toronto, ON", "2014", "Unsecured instalment", "Prime / near-prime", "Equity and debt facilities", "Proprietary scoring", "Canada-wide", "Provincially licensed", "Cloud-native", ">1MM customers", "Scaled platform with brand recognition"]},
            {"cells": ["Mobile-first lender", "Vancouver, BC", "2017", "Line of credit", "Thin-file", "Venture capital and debt warehouse", "Alternative data", "Western Canada", "Provincially licensed", "Mobile-first", "~0.5MM customers", "Actionable platform subject to diligence"]},
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


def test_registry_loads_15_blank_slide_library_entries():
    entries = load_slide_library_registry()

    assert len(entries) == 15
    assert entries[0].library_entry_id == "pitch-cover"
    assert entries[6].library_entry_id == "public-company-overview"
    assert entries[10].library_entry_id == "key-investment-highlights"
    assert entries[11].library_entry_id == "market-entry-targets"
    assert entries[10].static is False
    assert entries[11].static is False
    # Insider-ownership slide inserted before the static disclaimer/contact closers.
    assert entries[12].library_entry_id == "insider-ownership"
    assert entries[12].static is False
    assert entries[13].library_entry_id == "disclaimer"
    assert entries[13].static is True
    assert entries[14].library_entry_id == "contact"
    assert entries[14].static is True


def test_pitch_deck_wireframe_uses_blank_library_order():
    plan = build_pitch_deck_slide_plan(
        company=Company(legal_name="SampleCo Ltd.", ticker="TSX:SMP"),
        section_labels=["Overview", "Financial Summary", "Valuation", "Process"],
    )

    assert plan.deliverable_type == "pitch"
    assert len(plan.slides) == 15
    assert [slide.library_entry_id for slide in plan.slides[10:12]] == [
        "key-investment-highlights",
        "market-entry-targets",
    ]
    assert plan.slides[12].library_entry_id == "insider-ownership"
    assert plan.slides[12].content_block["requires"] == ["ownership_table"]
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
    assert len(prs.slides) == 15
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
    # Ownership slide present; both sides are placeholders when no ownership workbook is supplied.
    assert "[Placeholder for Insider Ownership]" in all_text
    assert "[Placeholder for Institutional Ownership]" in all_text

    # New content slides are filled and their placeholders replaced.
    assert "Durable Recurring Revenue" in all_text
    assert "SampleCo offers a scarce, scaled platform" in all_text
    assert "Potential Canada Market Entry Targets" in all_text
    assert "[Investment Highlight #1]" not in all_text
    assert "Potential [x] Market Entry Targets" not in all_text
    # Market-entry logo boxes are relabelled to name the company; with no target
    # name supplied they fall back to the generic label and the template's
    # '[Placeholder for Logo]' string is gone.
    assert "[Company Name Logo]" in all_text
    assert "[Placeholder for Logo]" not in all_text

    # Static credential slide text remains present.
    assert "INFOR is Canada’s leading provider of innovative, independent, forward thinking financial & strategic advice" in all_text
    assert "These materials are confidential and proprietary" in all_text


def test_earnings_update_plan_runs_captable_before_deck_for_insertion():
    plan_path = PLUGIN_ROOT / "plans" / "earnings-update.yaml"
    plan = Plan.model_validate(yaml.safe_load(plan_path.read_text(encoding="utf-8")))

    assert [stage.id for stage in plan.stages] == [
        "wireframe",
        "content",
        "ltm-metrics",
        "captable",
        "deck",
        "workbook-aggregation",
    ]
    deck_stage = next(s for s in plan.stages if s.id == "deck")
    assert deck_stage.inputs["captable_workbook_path"] == "$stages.captable.workbook_path"
    assert deck_stage.inputs["template_name"] == "INFOR Slide Library.pptx"
    # Aggregation runs last so the deck stage can still read the standalone cap table.
    assert plan.stages[-1].id == "workbook-aggregation"


def test_pitch_library_poc_plan_stage_order():
    plan_path = PLUGIN_ROOT / "plans" / "pitch.yaml"
    plan = Plan.model_validate(yaml.safe_load(plan_path.read_text(encoding="utf-8")))

    assert [stage.id for stage in plan.stages] == [
        "wireframe",
        "content",
        "ltm-metrics",
        "captable",
        "ownership",
        "deck",
        "workbook-aggregation",
    ]
    assert plan.stages[0].skill == "pitch-wireframe"
    assert plan.stages[1].skill == "pitch-content"
    assert plan.stages[2].skill == "ltm-metrics"
    assert plan.stages[3].skill == "captable"
    assert plan.stages[4].skill == "ownership"
    assert plan.stages[5].skill == "deck-assembler"
    assert plan.stages[6].skill == "workbook-aggregator"
    deck_stage = next(s for s in plan.stages if s.id == "deck")
    assert deck_stage.inputs["slide_plan_path"] == "$stages.wireframe.slide_plan_path"
    assert deck_stage.inputs["content_bundle_path"] == "$stages.content.content_bundle_path"
    assert deck_stage.inputs["captable_workbook_path"] == "$stages.captable.workbook_path"
    assert deck_stage.inputs["ownership_workbook_path"] == "$stages.ownership.workbook_path"
    # Ownership runs after captable so F35 can be sourced from the cap table's basic shares.
    ownership_stage = next(s for s in plan.stages if s.id == "ownership")
    assert ownership_stage.inputs["captable_workbook_path"] == "$stages.captable.workbook_path"
    # LTM metrics feed the cap table's LTM valuation column (mirrors earnings update).
    captable_stage = next(s for s in plan.stages if s.id == "captable")
    assert captable_stage.inputs["ltm_revenue"] == "$stages.ltm-metrics.ltm_revenue"
    assert captable_stage.inputs["ltm_adj_ebitda"] == "$stages.ltm-metrics.ltm_adj_ebitda"
    # The LTM and ownership workbooks are folded into the combined pitch workbook.
    agg_stage = next(s for s in plan.stages if s.id == "workbook-aggregation")
    assert agg_stage.inputs["workbooks"]["ltm-metrics"] == "$stages.ltm-metrics.workbook_path"
    assert agg_stage.inputs["workbooks"]["ownership"] == "$stages.ownership.workbook_path"


# ─── Helpers for the post-review fixes ───────────────────────────────────────

def _assemble(tmp_path: Path, content: PitchDeckContent, *, market_entry_target_count=None, **kwargs) -> Path:
    plan = build_pitch_deck_slide_plan(
        company=Company(legal_name="SampleCo Ltd.", ticker="TSX:SMP"),
        section_labels=["Overview", "Financial Summary", "Valuation", "Process"],
        market_entry_target_count=market_entry_target_count,
    )
    plan_path = write_slide_plan(plan, tmp_path / "slide_plan.json")
    content_path = tmp_path / "content.json"
    content_path.write_text(content.model_dump_json(indent=2), encoding="utf-8")
    return assemble_pitch_deck(
        slide_plan_path=plan_path,
        content_path=content_path,
        template_path=TEMPLATE,
        output_dir=tmp_path,
        **kwargs,
    )


def _content_with_targets(n: int) -> PitchDeckContent:
    """Clone the sample content but with `n` market-entry targets (12 cells each)."""
    base = _sample_content().model_dump()
    template_cells = base["market_entry_targets"][0]["cells"]
    base["market_entry_targets"] = [
        {"cells": [f"Target {i + 1}"] + template_cells[1:]} for i in range(n)
    ]
    return PitchDeckContent.model_validate(base)


def _write_sample_cap_table(path: Path, currency: str = "CAD") -> Path:
    wb = Workbook()
    ws = wb.active
    ws.title = "Cap with Links"
    ws["F5"] = currency  # output currency drives the footnote letter
    rows = {
        15: ("Capitalization Summary", None),
        16: ("Share Price", "12.34"),
        17: ("Basic Shares Outstanding", "100.0"),
        18: ("Basic Market Cap", "1,234.0"),
        31: ("Enterprise Value", "1,483.4"),
        40: ("EV / Adj. EBITDA", "10.0x"),
    }
    for row, (label, value) in rows.items():
        ws.cell(row=row, column=2).value = label
        if value is not None:
            ws.cell(row=row, column=6).value = value
    wb.save(path)
    return path


def _all_slides_text(prs: Presentation) -> str:
    parts: list[str] = []

    def _collect(shapes):
        for shape in shapes:
            if getattr(shape, "has_text_frame", False):
                parts.append(shape.text)
            if getattr(shape, "has_table", False):
                for row in shape.table.rows:
                    for cell in row.cells:
                        parts.append(cell.text)
            if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
                _collect(shape.shapes)

    for slide in prs.slides:
        _collect(slide.shapes)
    return "\n".join(parts)


# The two market-entry logo boxes in the library, ordered left→right by .left
# (matches the assembler's own sort): Rectangle 7 sits above the first target
# column, Rectangle 5 above the second.
_LOGO_SHAPE_NAMES = {"Rectangle 5", "Rectangle 7"}


def _logo_boxes(slide):
    return sorted(
        (s for s in slide.shapes if s.name in _LOGO_SHAPE_NAMES),
        key=lambda s: s.left,
    )


# ─── Fix 1: slide 2 exec summary keeps template colour + bullets ─────────────

def test_slide2_exec_summary_keeps_template_colour_and_bullets(tmp_path: Path):
    deck_path = _assemble(tmp_path, _sample_content())
    shape = find_shape(Presentation(deck_path).slides[1], "Content Placeholder 7")
    paras = [p for p in shape.text_frame.paragraphs if p.text.strip()]
    assert len(paras) == 4  # the four exec-summary bullets
    for p in paras:
        buchars = [e for e in p._p.iter() if e.tag.endswith("}buChar")]
        assert buchars, "every exec-summary bullet must keep its template glyph"
        for run in p.runs:
            xml = run._r.xml
            assert "<a:solidFill>" in xml, "run must carry an explicit template colour"
            assert "1B2759" not in xml, "body text must not fall back to the navy list colour"


# ─── Fix 3: financial-summary tiles are metric NAMES only ────────────────────

def test_financial_metric_labels_reject_value_laden_strings():
    base = _sample_content().model_dump()
    base["financial_metric_labels"] = [
        "FY2025 Revenue: US$589.8MM (+31% YoY)", "Gross Margin", "Adjusted EBITDA", "Free Cash Flow",
    ]
    with pytest.raises(ValidationError):
        PitchDeckContent.model_validate(base)
    base["financial_metric_labels"] = [
        "Revenue", "Adjusted EBITDA", "Combined Loan Balances", "Adjusted Return on Equity",
    ]
    assert PitchDeckContent.model_validate(base).financial_metric_labels[2] == "Combined Loan Balances"


# ─── Fix 4a: fixed 12-row market-entry structure ─────────────────────────────

def test_market_entry_row_labels_enforce_fixed_12_row_structure():
    # Wrong count (8 rows) rejected even when targets align 1:1.
    bad = _sample_content().model_dump()
    bad["market_entry_row_labels"] = [
        "Overview", "Headquarters", "Year Founded", "A", "B", "C", "Scale KPIs", "Strategic Rationale",
    ]
    bad["market_entry_targets"] = [{"cells": ["x"] * 8}, {"cells": ["y"] * 8}]
    with pytest.raises(ValidationError):
        PitchDeckContent.model_validate(bad)

    # Wrong fixed top label rejected.
    bad2 = _sample_content().model_dump()
    labels = list(bad2["market_entry_row_labels"])
    labels[0] = "Summary"  # must be 'Overview'
    bad2["market_entry_row_labels"] = labels
    with pytest.raises(ValidationError):
        PitchDeckContent.model_validate(bad2)

    # Wrong fixed bottom label rejected.
    bad3 = _sample_content().model_dump()
    labels = list(bad3["market_entry_row_labels"])
    labels[-1] = "Why Acquire"  # must be 'Strategic Rationale'
    bad3["market_entry_row_labels"] = labels
    with pytest.raises(ValidationError):
        PitchDeckContent.model_validate(bad3)


# ─── Fix 4b / 4c: N targets across N/2 slides, formatting ────────────────────

def test_market_entry_expands_two_targets_per_slide(tmp_path: Path):
    deck_path = _assemble(tmp_path, _content_with_targets(8))
    prs = Presentation(deck_path)
    # 15 base - 1 market-entry + 4 market-entry = 18 slides.
    assert len(prs.slides) == 18
    titles = [find_shape(prs.slides[11 + j], "Title 1").text for j in range(4)]
    assert titles == [
        "Potential Canada Market Entry Targets (1 of 4)",
        "Potential Canada Market Entry Targets (2 of 4)",
        "Potential Canada Market Entry Targets (3 of 4)",
        "Potential Canada Market Entry Targets (4 of 4)",
    ]
    # Static disclaimer + contact preserved at the tail.
    tail = _all_slides_text(prs)
    assert "These materials are confidential and proprietary" in tail
    assert "Neil Selfe, Managing Principal" in tail


def test_market_entry_table_formatting_no_blank_rows(tmp_path: Path):
    deck_path = _assemble(tmp_path, _sample_content())  # 2 targets -> 1 slide
    table = next(
        s for s in Presentation(deck_path).slides[11].shapes if getattr(s, "has_table", False)
    ).table
    # 12 data rows (rows 1-12); every one populated, label white@11, values @10.
    for row in range(1, 13):
        label_run = table.cell(row, 0).text_frame.paragraphs[0].runs[0]
        assert label_run.font.size == Pt(11), "label column must be 11 pt"
        assert str(label_run.font.color.rgb) == "FFFFFF", "label column must be white"
        for col in (1, 2):
            cell = table.cell(row, col)
            assert cell.text.strip(), f"data cell ({row},{col}) must be populated"
            assert cell.text_frame.paragraphs[0].runs[0].font.size == Pt(10), "values must be 10 pt"


def test_market_entry_odd_count_blanks_unused_column_and_logo(tmp_path: Path):
    deck_path = _assemble(tmp_path, _content_with_targets(3))  # -> 2 slides
    prs = Presentation(deck_path)
    assert len(prs.slides) == 16  # 14 base + 2 market-entry
    last_me = prs.slides[12]  # the second (final) market-entry slide
    table = next(s for s in last_me.shapes if getattr(s, "has_table", False)).table
    assert table.cell(1, 1).text.strip(), "the single target fills the first target column"
    assert table.cell(1, 2).text.strip() == "", "the unused target column is blanked"
    boxes = _logo_boxes(last_me)
    assert len(boxes) == 2
    assert boxes[0].text.strip() == "[Company Name Logo]", "the single target's logo box is labelled"
    assert boxes[1].text.strip() == "", "the unused (rightmost) logo box is blanked"


def test_market_entry_logo_box_names_each_target(tmp_path: Path):
    base = _sample_content().model_dump()
    base["market_entry_targets"] = [
        {"name": "Kueski", "cells": base["market_entry_targets"][0]["cells"]},
        {"name": "Nubank", "cells": base["market_entry_targets"][1]["cells"]},
    ]
    deck_path = _assemble(tmp_path, PitchDeckContent.model_validate(base))  # 2 targets -> 1 slide
    boxes = _logo_boxes(Presentation(deck_path).slides[11])
    assert [b.text.strip() for b in boxes] == ["[Kueski Logo]", "[Nubank Logo]"]


# ─── Slide 9: Considerations/Mitigants font sizes match the library ──────────

def test_slide9_risk_table_uses_library_font_sizes(tmp_path: Path):
    deck_path = _assemble(tmp_path, _sample_content())  # 3 risk/mitigant rows
    table = next(
        s for s in Presentation(deck_path).slides[8].shapes if getattr(s, "has_table", False)
    ).table
    # Header row at 12 pt (was 9 pt), matching the library.
    for col, label in ((0, "Considerations"), (1, "Mitigants")):
        run = table.cell(0, col).text_frame.paragraphs[0].runs[0]
        assert run.text == label
        assert run.font.size == Pt(12), "risk/mitigant header must be 12 pt"
    # Body rows at 10 pt (was 8 pt) — three populated rows in the sample content.
    for row in range(1, 4):
        for col in (0, 1):
            run = table.cell(row, col).text_frame.paragraphs[0].runs[0]
            assert run.font.size == Pt(10), "risk/mitigant body must be 10 pt"


def test_pitch_wireframe_expands_market_entry_slides():
    plan = build_pitch_deck_slide_plan(
        company=Company(legal_name="SampleCo Ltd.", ticker="TSX:SMP"),
        market_entry_target_count=8,
    )
    me = [s for s in plan.slides if s.library_entry_id == "market-entry-targets"]
    assert len(me) == 4
    assert len(plan.slides) == 18
    assert me[0].title.endswith("(1 of 4)")
    assert me[3].title.endswith("(4 of 4)")
    # Count unknown (live pipeline) -> single market-entry slide, 15-slide plan.
    default_plan = build_pitch_deck_slide_plan(company=Company(legal_name="X Co", ticker="T:X"))
    assert sum(1 for s in default_plan.slides if s.library_entry_id == "market-entry-targets") == 1
    assert len(default_plan.slides) == 15


# ─── Fix 2: cap table pasted onto slide 7 + footnote currency ────────────────

def test_pitch_deck_inserts_cap_table_into_slide7(tmp_path: Path):
    pytest.importorskip("win32com.client", reason="picture-based insertion requires pywin32 + Excel")
    workbook = _write_sample_cap_table(tmp_path / "cap-table.xlsx", currency="USD")
    deck_path = _assemble(tmp_path, _sample_content(), captable_workbook_path=workbook)
    prs = Presentation(deck_path)
    slide7 = prs.slides[6]
    assert next((s for s in slide7.shapes if s.name == "Rectangle 3"), None) is None, (
        "slide 7 cap-table placeholder should be replaced by the picture"
    )
    pictures = [s for s in slide7.shapes if s.shape_type == MSO_SHAPE_TYPE.PICTURE]
    assert pictures, "expected a cap-table picture on slide 7"
    note = find_shape(slide7, "Text Placeholder 1").text
    assert "US$MM" in note and "[x]" not in note, "footnote currency derived from the cap table (USD)"
    # The figure-footnote currency token is resolved on slide 11 too (no stray [x]).
    note11 = find_shape(prs.slides[10], "Text Placeholder 13").text
    assert "US$MM" in note11 and "[x]" not in note11, "slide 11 footnote currency derived from the cap table"


# ─── Ownership: insider table pasted onto the ownership slide ─────────────────

def test_pitch_deck_inserts_ownership_into_slide(tmp_path: Path):
    pytest.importorskip("win32com.client", reason="picture-based insertion requires pywin32 + Excel")
    import pywintypes

    from ownership_workbook import build_ownership_workbook, InsiderHolding

    own_wb = build_ownership_workbook(
        template_path=PLUGIN_ROOT / "templates" / "INFOR Ownership Template.xlsx",
        insiders=[
            InsiderHolding("Barrenechea, Mark James", "Mark Barrenechea (CEO & Director)", 1219092, "2025-03-31"),
            InsiderHolding("Fowlie, Randy", "Randy Fowlie (Director)", [193000, 0], "2025-12-01"),
        ],
        total_shares_outstanding=261_000_000,
        output_path=tmp_path / "Ownership.xlsx",
    )
    try:
        deck_path = _assemble(tmp_path, _sample_content(), ownership_workbook_path=own_wb)
    except (pywintypes.com_error, RuntimeError) as exc:  # Excel/LibreOffice unavailable here
        pytest.skip(f"range render backend unavailable in this environment: {exc}")

    prs = Presentation(deck_path)
    own_slide = prs.slides[12]  # ownership follows the single market-entry slide (index 11)
    assert next((s for s in own_slide.shapes if s.name == "Rectangle 1"), None) is None, (
        "ownership insider placeholder (Rectangle 1) should be replaced by the picture"
    )
    assert [s for s in own_slide.shapes if s.shape_type == MSO_SHAPE_TYPE.PICTURE], (
        "expected an insider-ownership picture on the ownership slide"
    )
    text = _all_slides_text(prs)
    assert "[Placeholder for Insider Ownership]" not in text, "insider placeholder replaced"
    # The institutional / Bloomberg side stays a placeholder (SEDI fills only insiders).
    assert "[Placeholder for Institutional Ownership]" in text
