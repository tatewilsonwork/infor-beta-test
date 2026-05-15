"""Unit tests for the Filing schema and its FilingType enum."""

from datetime import date
from pathlib import Path

import pytest
from pydantic import ValidationError

from schemas import Filing, FilingType


def test_minimal_filing():
    f = Filing(type=FilingType.TEN_K)
    assert f.type == FilingType.TEN_K
    assert f.type_other is None


def test_all_enum_values_construct():
    """Every locked enum value (G2) must round-trip cleanly."""
    expected = {
        "10-K",
        "10-Q",
        "8-K",
        "proxy",
        "S-1",
        "20-F",
        "6-K",
        "AIF",
        "MD&A",
        "management-circular",
        "annual-report",
        "ARS",
        "prospectus",
        "press-release",
        "transcript",
        "investor-deck",
        "other",
    }
    actual = {member.value for member in FilingType}
    assert actual == expected, f"missing/extra enum values: {actual ^ expected}"

    # type=other needs type_other; use a stub for that one
    for value in expected:
        if value == "other":
            f = Filing(type=FilingType(value), type_other="stub")
        else:
            f = Filing(type=FilingType(value))
        assert f.type.value == value


def test_other_requires_type_other():
    with pytest.raises(ValidationError) as exc:
        Filing(type=FilingType.OTHER)
    assert "type_other" in str(exc.value)


def test_other_with_blank_type_other_rejected():
    with pytest.raises(ValidationError):
        Filing(type=FilingType.OTHER, type_other="   ")


def test_other_with_real_type_other_accepted():
    f = Filing(type=FilingType.OTHER, type_other="court-filing")
    assert f.type == FilingType.OTHER
    assert f.type_other == "court-filing"


def test_non_other_with_type_other_allowed_but_unused():
    """Setting type_other on a non-OTHER filing is permitted (no harm). The
    field is only *required* when type is OTHER."""
    f = Filing(type=FilingType.TEN_K, type_other="ignored")
    assert f.type == FilingType.TEN_K
    assert f.type_other == "ignored"


def test_optional_metadata_fields():
    f = Filing(
        type=FilingType.MDA,
        title="OpenText FY2025 MD&A",
        filed_on=date(2025, 8, 1),
        period_end=date(2025, 6, 30),
        source_url="https://sedarplus.ca/example",
        local_path=Path("/tmp/example.pdf"),
        notes="Snapped from the IR site.",
    )
    raw = f.model_dump_json()
    assert "FY2025" in raw
    f2 = Filing.model_validate_json(raw)
    assert f2 == f
