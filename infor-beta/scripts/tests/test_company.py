"""Unit tests for the Company schema."""

import pytest
from pydantic import ValidationError

from schemas import Company


def test_minimal_company_only_legal_name():
    c = Company(legal_name="OpenText Corporation")
    assert c.legal_name == "OpenText Corporation"
    assert c.ticker is None
    assert c.hq is None


def test_legal_name_required():
    with pytest.raises(ValidationError):
        Company()  # type: ignore[call-arg]


def test_empty_legal_name_rejected():
    with pytest.raises(ValidationError):
        Company(legal_name="")


def test_all_optional_fields_round_trip():
    c = Company(
        legal_name="OpenText Corporation",
        hq="Waterloo, Ontario, Canada",
        jurisdiction_of_incorporation="Ontario",
        ticker="OTEX",
        exchange="NASDAQ",
        fy_end="06-30",
        employees=23000,
        revenue_range="$5B–$10B",
        sector="Information Technology",
        industry="Application Software",
        notes="One of the largest Canadian software companies.",
    )
    raw = c.model_dump_json()
    c2 = Company.model_validate_json(raw)
    assert c2 == c


def test_employees_must_be_non_negative():
    with pytest.raises(ValidationError):
        Company(legal_name="X Corp", employees=-1)


def test_extra_fields_rejected():
    with pytest.raises(ValidationError):
        Company(legal_name="X Corp", capiq_id="IQ123")  # type: ignore[call-arg]


def test_strip_whitespace():
    c = Company(legal_name="  Acme Corp  ")
    assert c.legal_name == "Acme Corp"
