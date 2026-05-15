"""Unit tests for the POC earnings-update deck assembler."""

from pathlib import Path

from pptx import Presentation

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


def test_assemble_earnings_update_deck_populates_template(tmp_path: Path):
    template = Path("infor-beta/templates/INFOR Earnings Update Template.pptx")
    slide_plan = build_earnings_update_slide_plan(
        company=Company(legal_name="SampleCo", ticker="TSX:SMPL"),
        reporting_quarter="Q4 2025",
        comparison_quarter="Q4 2024",
    )
    slide_plan_path = write_slide_plan(slide_plan, tmp_path / "slide_plan.json")
    content = _sample_content()
    content_path = tmp_path / "content.json"
    content_path.write_text(content.model_dump_json(indent=2), encoding="utf-8")

    deck_path = assemble_earnings_update_deck(
        slide_plan_path=slide_plan_path,
        content_path=content_path,
        template_path=template,
        output_dir=tmp_path,
    )

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
