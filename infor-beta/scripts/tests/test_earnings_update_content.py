"""Unit tests for the EarningsUpdateContent typed handoff."""

import pytest
from pydantic import ValidationError

from schemas import (
    BrokerRow,
    CompanyOverviewBullet,
    EarningsUpdateContent,
    KpiRow,
    ManagementQuote,
    SourceNote,
)


def _valid_content(**overrides):
    data = {
        "company_name": "SampleCo",
        "ticker": "TSX:SMPL",
        "reporting_quarter": "Q4 2025",
        "comparison_quarter": "Q4 2024",
        "currency": "C$MM",
        "currency_short": "C$MM",
        "cover_date": "May 2026",
        "company_overview_bullets": [
                {"bold_prefix": "", "text": "Leading provider of mission-critical compliance and workflow software serving blue-chip enterprises across Canada", "level": 0},
                {"bold_prefix": "", "text": "Recurring revenue model supported by multi-year contracts and a high-retention installed base in regulated markets", "level": 0},
                {"bold_prefix": "", "text": "Diversified product suite addressing daily operational needs for finance, legal and compliance teams", "level": 0},
                {"bold_prefix": "", "text": "Experienced management team with a record of disciplined execution and successful tuck-in acquisitions", "level": 0},
                {"bold_prefix": "", "text": "Strong balance sheet supporting organic growth, selective acquisitions and continued platform investment", "level": 0},
                {"bold_prefix": "", "text": "Established go-to-market platform with direct sales coverage and a repeatable land-and-expand motion", "level": 0},
                {"bold_prefix": "", "text": "Well-positioned to benefit from the ongoing digitization of enterprise compliance workflows", "level": 0},
        ],
        "business_updates": [
            "Revenue growth reflected continued enterprise demand and disciplined customer expansion",
            "Management highlighted improving sales productivity and constructive renewal activity",
            "The Company continued to invest in product development while maintaining margin discipline",
            "Near-term priorities remain focused on enterprise execution and cash conversion",
        ],
        "kpi_rows": [
            {"name": "Revenue", "prior_value": "$100.0", "current_value": "$125.0", "delta_str": "+$25.0", "delta_sign": 1},
            {"name": "Adjusted EBITDA", "prior_value": "$20.0", "current_value": "$22.0", "delta_str": "+$2.0", "delta_sign": 1},
            {"name": "Net Income", "prior_value": "$8.0", "current_value": "$7.0", "delta_str": "-$1.0", "delta_sign": -1},
            {"name": "Gross Margin", "prior_value": "60.0%", "current_value": "62.0%", "delta_str": "+2.0%", "delta_sign": 1},
        ],
        "broker_rows": [
            {"label": "Revenue", "reported": "$125.0", "estimate": "$120.0", "variance": "+$5.0", "variance_sign": 1},
            {"label": "Adjusted EBITDA", "reported": "$22.0", "estimate": "$21.0", "variance": "+$1.0", "variance_sign": 1},
            {"label": "EPS", "reported": "$0.10", "estimate": "$0.12", "variance": "-$0.02", "variance_sign": -1},
            {"label": "Gross Margin", "reported": "62.0%", "estimate": "61.0%", "variance": "+1.0%", "variance_sign": 1},
            {"label": "Free Cash Flow", "reported": "$15.0", "estimate": "$14.0", "variance": "+$1.0", "variance_sign": 1},
        ],
        "management_quotes": [
            {"quote": "We delivered a strong quarter driven by enterprise execution and continued customer expansion", "speaker": "Jane Doe", "role": "CEO"},
            {"quote": "Our focus remains on profitable growth, cash conversion and disciplined capital allocation", "speaker": "John Smith", "role": "CFO"},
        ],
        "performance_summary": "SampleCo beat consensus revenue while maintaining disciplined operating execution",
        "sources": [{"section": "Business Updates", "citation": "Company filings"}],
        "manual_steps": ["Refresh CapIQ in the companion cap table before pasting into slide 2"],
    }
    data.update(overrides)
    return EarningsUpdateContent(**data)


def test_valid_earnings_update_content_round_trips():
    content = _valid_content()
    raw = content.model_dump_json()
    restored = EarningsUpdateContent.model_validate_json(raw)
    assert restored == content
    assert restored.broker_rows[0].variance_sign == 1


def test_company_overview_requires_6_to_10_bullets():
    with pytest.raises(ValidationError):
        _valid_content(company_overview_bullets=[{"text": "Too few", "level": 0}] * 5)
    with pytest.raises(ValidationError):
        _valid_content(company_overview_bullets=[{"text": "Too many", "level": 0}] * 11)


def test_company_overview_rejects_text_over_char_budget():
    # 10 bullets of ~120 chars each blows past the 1,050-char ceiling.
    long_bullet = {"text": "x" * 120, "level": 0}
    with pytest.raises(ValidationError):
        _valid_content(company_overview_bullets=[long_bullet] * 10)


def test_company_overview_rejects_trailing_period_or_semicolon():
    with pytest.raises(ValidationError):
        CompanyOverviewBullet(text="This should not end with a period.", level=0)
    with pytest.raises(ValidationError):
        CompanyOverviewBullet(text="This should not end with a semicolon;", level=0)


def test_business_updates_require_4_to_6_items_and_900_char_budget():
    with pytest.raises(ValidationError):
        _valid_content(business_updates=["only three"] * 3)
    with pytest.raises(ValidationError):
        _valid_content(business_updates=["x" * 226] * 4)


def test_broker_rows_must_have_exactly_five_non_na_rows():
    with pytest.raises(ValidationError):
        _valid_content(broker_rows=[{"label": "Revenue", "reported": "$1", "estimate": "$1", "variance": "$0", "variance_sign": 0}])
    with pytest.raises(ValidationError):
        BrokerRow(label="Revenue", reported="N/A", estimate="$1", variance="$0", variance_sign=0)


def test_quote_and_summary_caps_are_enforced():
    with pytest.raises(ValidationError):
        ManagementQuote(quote="word " * 31, speaker="Jane Doe", role="CEO")
    with pytest.raises(ValidationError):
        _valid_content(performance_summary="word " * 26)


def test_kpi_rows_reject_bps_rate_deltas():
    with pytest.raises(ValidationError):
        KpiRow(name="Gross Margin", prior_value="60%", current_value="62%", delta_str="+200 bps", delta_sign=1)


def test_source_note_requires_non_empty_fields():
    with pytest.raises(ValidationError):
        SourceNote(section="", citation="Company filings")
