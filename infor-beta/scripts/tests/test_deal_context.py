"""Unit tests for DealContext."""

from pathlib import Path

import pytest
from pydantic import ValidationError

from schemas import Company, DealContext, Filing, FilingType


def test_minimum_deal_context():
    ctx = DealContext(
        codename="Project OpenText",
        deal_dir=Path("~/Documents/INFOR Deals/Project OpenText"),
        deliverable_type="earnings-update",
    )
    assert ctx.codename == "Project OpenText"
    assert ctx.subject_company is None
    assert ctx.filings == []


def test_deal_dir_accepts_relative_path():
    ctx = DealContext(
        codename="Project OpenText",
        deal_dir=Path("./deals/Project OpenText"),
        deliverable_type="pitch",
    )
    assert ctx.deal_dir == Path("./deals/Project OpenText")


def test_deal_dir_accepts_absolute_path(tmp_path: Path):
    abs_dir = tmp_path / "Project OpenText"  # tmp_path is absolute on every OS
    ctx = DealContext(
        codename="Project OpenText",
        deal_dir=abs_dir,
        deliverable_type="pitch",
    )
    assert ctx.deal_dir.is_absolute()
    assert ctx.deal_dir == abs_dir


def test_invalid_deliverable_type_rejected():
    with pytest.raises(ValidationError):
        DealContext(
            codename="Project OpenText",
            deal_dir=Path("/tmp/x"),
            deliverable_type="management-presentation",  # removed from scope, B2
        )


def test_round_trip_with_company_and_filings():
    ctx = DealContext(
        codename="Project OpenText",
        deal_dir=Path("/tmp/Project OpenText"),
        deliverable_type="earnings-update",
        subject_company=Company(legal_name="OpenText Corporation", ticker="OTEX"),
        filings=[Filing(type=FilingType.TEN_K, title="OpenText FY2025 10-K")],
        notes="Pitching as part of broader software thesis.",
    )
    raw = ctx.model_dump_json()
    ctx2 = DealContext.model_validate_json(raw)
    assert ctx2 == ctx


def test_empty_codename_rejected():
    with pytest.raises(ValidationError):
        DealContext(codename="", deal_dir=Path("/tmp/x"), deliverable_type="earnings-update")
