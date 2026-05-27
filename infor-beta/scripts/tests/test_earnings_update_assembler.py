"""Unit tests for the POC earnings-update deck assembler."""

from pathlib import Path

import pytest
from openpyxl import Workbook
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from earnings_update_assembler import assemble_earnings_update_deck
from earnings_update_wireframe import build_earnings_update_slide_plan, write_slide_plan
from schemas import Company, EarningsUpdateContent


def _sample_content() -> EarningsUpdateContent:
    return EarningsUpdateContent(
        company_name="SampleCo",
        ticker="TSX:SMPL",
        reporting_quarter="Q4 2025",
        comparison_quarter="Q4 2024",
        currency="C$MM",
        currency_short="C$MM",
        cover_date="May 2026",
        company_overview_bullets=[
                {"text": "Leading provider of mission-critical software serving blue-chip enterprise customers across Canada with a platform spanning compliance, workflow automation and analytics", "level": 0},
                {"text": "Recurring revenue model supported by multi-year contracts, high customer retention and a growing installed base across regulated end markets", "level": 0},
                {"text": "Diversified product suite addressing daily operational pain points for finance, legal and compliance teams that require reliable data and auditability", "level": 0},
                {"text": "Established go-to-market platform with direct sales coverage, partner channels and a repeatable land-and-expand motion across enterprise accounts", "level": 0},
                {"text": "Meaningful operating leverage as the Company scales across existing infrastructure while maintaining disciplined product investment and customer support", "level": 0},
                {"text": "Experienced Management team with a demonstrated record of disciplined execution, prudent capital allocation and successful integration of tuck-in acquisitions", "level": 0},
                {"text": "Strong balance sheet and flexible capital structure supporting organic growth initiatives, selective acquisitions and continued investment in the platform", "level": 0},
                {"text": "Well-positioned to benefit from continued digitization of compliance workflows as customers prioritize efficiency, accuracy and defensible reporting", "level": 0},
        ],
        business_updates=[
            "Revenue growth reflected continued enterprise demand and disciplined customer expansion",
            "Management highlighted improving sales productivity and constructive renewal activity",
            "The Company continued to invest in product development while maintaining margin discipline",
            "Near-term priorities remain focused on enterprise execution and cash conversion",
        ],
        kpi_rows=[
            {"name": "Revenue", "prior_value": "$100.0", "current_value": "$125.0", "delta_str": "+$25.0", "delta_sign": 1},
            {"name": "Adjusted EBITDA", "prior_value": "$20.0", "current_value": "$22.0", "delta_str": "+$2.0", "delta_sign": 1},
            {"name": "Net Income", "prior_value": "$8.0", "current_value": "$7.0", "delta_str": "-$1.0", "delta_sign": -1},
            {"name": "Gross Margin", "prior_value": "60.0%", "current_value": "62.0%", "delta_str": "+2.0%", "delta_sign": 1},
        ],
        broker_rows=[
            {"label": "Revenue", "reported": "$125.0", "estimate": "$120.0", "variance": "+$5.0", "variance_sign": 1},
            {"label": "Adjusted EBITDA", "reported": "$22.0", "estimate": "$21.0", "variance": "+$1.0", "variance_sign": 1},
            {"label": "EPS", "reported": "$0.10", "estimate": "$0.12", "variance": "-$0.02", "variance_sign": -1},
            {"label": "Gross Margin", "reported": "62.0%", "estimate": "61.0%", "variance": "+1.0%", "variance_sign": 1},
            {"label": "Free Cash Flow", "reported": "$15.0", "estimate": "$14.0", "variance": "+$1.0", "variance_sign": 1},
        ],
        management_quotes=[
            {"quote": "We delivered a strong quarter driven by enterprise execution and continued customer expansion", "speaker": "Jane Doe", "role": "CEO"},
            {"quote": "Our focus remains on profitable growth, cash conversion and disciplined capital allocation", "speaker": "John Smith", "role": "CFO"},
        ],
        performance_summary="SampleCo beat consensus revenue while maintaining disciplined operating execution",
        sources=[{"section": "All", "citation": "Company filings, S&P CapIQ, equity research"}],
        manual_steps=["Refresh CapIQ in the companion cap table before pasting into slide 2"],
    )


def _write_sample_cap_table(path: Path) -> Path:
    wb = Workbook()
    ws = wb.active
    ws.title = "Cap with Links"
    ws["B13"] = "SampleCo Cap Table"
    rows = {
        15: ("Company Ticker:", "TSX:SMPL"),
        16: ("Share Price (21-Apr-26)", "C$12.34"),
        17: ("Basic Shares Outstanding", "100.0"),
        18: ("Basic Market Cap", "C$1,234.0"),
        21: ("Fully-Diluted Shares Outstanding", "104.0"),
        22: ("Fully-Diluted Market Cap", "C$1,283.4"),
        28: ("Net Debt", "C$200.0"),
        31: ("Enterprise Value", "C$1,483.4"),
    }
    for row, (label, value) in rows.items():
        ws.cell(row=row, column=2).value = label
        ws.cell(row=row, column=6).value = value
    wb.save(path)
    return path


def _assemble_sample_deck(tmp_path: Path, content: EarningsUpdateContent | None = None, **kwargs) -> Path:
    template = Path("infor-beta/templates/INFOR Earnings Update Template.pptx")
    slide_plan = build_earnings_update_slide_plan(
        company=Company(legal_name="SampleCo", ticker="TSX:SMPL"),
        reporting_quarter="Q4 2025",
        comparison_quarter="Q4 2024",
    )
    slide_plan_path = write_slide_plan(slide_plan, tmp_path / "slide_plan.json")
    content = content or _sample_content()
    content_path = tmp_path / "content.json"
    content_path.write_text(content.model_dump_json(indent=2), encoding="utf-8")

    return assemble_earnings_update_deck(
        slide_plan_path=slide_plan_path,
        content_path=content_path,
        template_path=template,
        output_dir=tmp_path,
        **kwargs,
    )


def test_assemble_earnings_update_deck_populates_template(tmp_path: Path):
    deck_path = _assemble_sample_deck(tmp_path)

    assert deck_path.exists()
    prs = Presentation(deck_path)
    assert len(prs.slides) == 5
    all_text = "\n".join(shape.text for slide in prs.slides for shape in slide.shapes if getattr(shape, "has_text_frame", False))
    assert "[Current Month]" not in all_text
    assert "[Client Name]" not in all_text
    assert "Q[x]" not in all_text
    assert "SampleCo Overview" in all_text
    assert "SampleCo Q4 2025 Earnings Summary" in all_text
    assert "[Macabacus Placeholder]" in all_text


def test_assemble_earnings_update_deck_strips_currency_units_from_kpi_labels(tmp_path: Path):
    content = _sample_content()
    content.kpi_rows[0].name = "Cloud Revenue (C$MM)"
    content.kpi_rows[1].name = "Recurring Revenue (C$MM)"
    content.kpi_rows[2].name = "Adjusted EBITDA (C$MM)"

    deck_path = _assemble_sample_deck(tmp_path, content)

    slide3_text = "\n".join(
        shape.text
        for shape in Presentation(deck_path).slides[2].shapes
        if getattr(shape, "has_text_frame", False)
    )
    assert "Q4 2025 Cloud Revenue" in slide3_text
    assert "Q4 2025 Cloud Revenue (C$MM)" not in slide3_text
    assert "Q4 2026 Recurring Revenue (C$MM)" not in slide3_text
    assert "Q4 2025 Adjusted EBITDA (C$MM)" not in slide3_text
    assert "Note: All figures in C$MM" in slide3_text


def test_assemble_earnings_update_deck_does_not_bold_overview_headers(tmp_path: Path):
    content = _sample_content()
    content.company_overview_bullets[0].bold_prefix = "Header:"
    content.company_overview_bullets[0].text = " Enterprise software provider with recurring revenue visibility across regulated markets and a diversified global customer base supporting resilient growth"

    deck_path = _assemble_sample_deck(tmp_path, content)

    overview_shape = next(
        shape for shape in Presentation(deck_path).slides[1].shapes if shape.name == "TextBox 16"
    )
    header_runs = [
        run
        for para in overview_shape.text_frame.paragraphs
        for run in para.runs
        if "Header:" in run.text
    ]
    assert header_runs, "expected the overview bullet prefix text to be present"
    assert all(run.font.bold is not True for run in header_runs)


def test_assemble_earnings_update_deck_inserts_cap_table_from_workbook(tmp_path: Path):
    pytest.importorskip("win32com.client", reason="picture-based insertion requires pywin32 + Excel")

    workbook_path = _write_sample_cap_table(tmp_path / "cap-table.xlsx")

    deck_path = _assemble_sample_deck(tmp_path, captable_workbook_path=workbook_path)

    prs = Presentation(deck_path)
    slide2 = prs.slides[1]

    assert next((s for s in slide2.shapes if s.name == "Rectangle 4"), None) is None, (
        "slide 2 Macabacus placeholder should be removed after picture insertion"
    )

    pictures = [s for s in slide2.shapes if s.shape_type == MSO_SHAPE_TYPE.PICTURE]
    assert pictures, "expected a picture shape on slide 2 after cap-table insertion"

    pic = pictures[-1]
    assert pic.width == 4140000, f"picture width should match placeholder (4140000 EMU), got {pic.width}"
    assert pic.height == 4947508, f"picture height should match placeholder (4947508 EMU), got {pic.height}"
