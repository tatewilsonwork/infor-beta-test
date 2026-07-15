"""Unit tests for the codename resolver (Phase 1, locked decision G3)."""

from pathlib import Path

import pytest

from codename import codename_from_company, disambiguate, find_existing, resolve


def test_strip_unsafe_chars():
    display, _ = resolve("Project Open/Text*?", deals_root="/tmp/deals")
    assert display == "Project OpenText"


def test_preserves_case():
    display, _ = resolve("Project OpenText", deals_root="/tmp/deals")
    assert display == "Project OpenText"
    display2, _ = resolve("project opentext", deals_root="/tmp/deals")
    assert display2 == "project opentext"


def test_preserves_common_chars():
    """& , . spaces and apostrophes are all macOS-safe and must survive."""
    display, _ = resolve("Project Smith & Co., Inc.", deals_root="/tmp/deals")
    assert display == "Project Smith & Co., Inc."


def test_dir_path_under_deals_root():
    _, path = resolve("Project OpenText", deals_root="/tmp/deals")
    assert path == Path("/tmp/deals/Project OpenText")


def test_empty_after_strip_raises():
    with pytest.raises(ValueError):
        resolve("///", deals_root="/tmp/deals")


def test_none_raises():
    with pytest.raises(ValueError):
        resolve(None)  # type: ignore[arg-type]


@pytest.mark.parametrize(
    ("company", "expected"),
    [
        ("OpenText Corporation", "Project OpenText"),
        ("ACME Corp", "Project ACME"),
        ("ACME Corp.", "Project ACME"),
        ("ACME, Inc.", "Project ACME"),
        ("acme incorporated", "Project acme"),
        ("ACME Holdings Inc.", "Project ACME"),  # suffixes strip repeatedly
        ("Northern Rail Ltd", "Project Northern Rail"),
        ("Northern Rail Limited", "Project Northern Rail"),
        ("Maple Leaf Foods LLC", "Project Maple Leaf Foods"),
        ("Brookfield LP", "Project Brookfield"),
        ("Vodafone PLC", "Project Vodafone"),
        ("Boston Consulting Group", "Project Boston Consulting"),
        ("The Coca-Cola Company", "Project The Coca-Cola"),
        ("Smith & Co", "Project Smith & Co"),  # "& Co" brands kept whole
        ("Kelso & Co LP", "Project Kelso & Co"),
        ("NoSuffix", "Project NoSuffix"),
        ("Cobalt", "Project Cobalt"),  # "Co" only strips as a whole word
    ],
)
def test_codename_from_company_strips_corporate_suffixes(company, expected):
    assert codename_from_company(company) == expected


def test_codename_from_company_strips_unsafe_and_collapses_whitespace():
    assert codename_from_company("Open/Text*  Corp?") == "Project OpenText"


def test_codename_from_company_all_suffix_name_keeps_sanitised_form():
    # A name that is nothing but suffix words must not empty out.
    assert codename_from_company("Limited") == "Project Limited"


def test_codename_from_company_rejects_empty():
    with pytest.raises(ValueError):
        codename_from_company(None)  # type: ignore[arg-type]
    with pytest.raises(ValueError):
        codename_from_company("///")


def test_find_existing_case_insensitive(tmp_path: Path):
    (tmp_path / "Project OpenText").mkdir()
    (tmp_path / "Project OpenText" / "comps.xlsx").write_bytes(b"")
    found = find_existing(tmp_path, "project opentext")
    assert found is not None
    assert found.name == "Project OpenText"


def test_find_existing_returns_none_when_missing(tmp_path: Path):
    (tmp_path / "Project OpenText").mkdir()
    assert find_existing(tmp_path, "Project Atlas") is None


def test_find_existing_handles_missing_root(tmp_path: Path):
    missing = tmp_path / "does-not-exist"
    assert find_existing(missing, "Project OpenText") is None


def test_find_existing_ignores_files(tmp_path: Path):
    """A *file* named like the codename must not be returned — we only look at directories."""
    (tmp_path / "Project OpenText").write_bytes(b"")
    assert find_existing(tmp_path, "Project OpenText") is None


def test_disambiguate_suggestions_avoid_existing(tmp_path: Path):
    (tmp_path / "Project OpenText").mkdir()
    (tmp_path / "Project OpenText II").mkdir()  # already taken
    suggestions = disambiguate(tmp_path, "Project OpenText")
    assert "Project OpenText II" not in suggestions
    assert "Project OpenText 2026" in suggestions
    assert len(suggestions) <= 4


def test_disambiguate_returns_at_least_one(tmp_path: Path):
    (tmp_path / "Project OpenText").mkdir()
    suggestions = disambiguate(tmp_path, "Project OpenText")
    assert len(suggestions) >= 1
