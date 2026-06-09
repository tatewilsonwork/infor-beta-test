"""Tests for the ownership workbook builder (deterministic; no Excel needed).

The picture render (Excel COM / LibreOffice) is covered separately and skips
where unavailable; these tests pin the openpyxl write that fills the template's
Select-Insiders block and the vestigial-cruft strip that keeps the output
openable by the render step.
"""

from datetime import date, datetime
from pathlib import Path

import pytest
from openpyxl import Workbook, load_workbook

from ownership_workbook import (
    InsiderHolding,
    build_ownership_workbook,
    read_basic_shares_from_cap_table,
)

PLUGIN_ROOT = Path(__file__).resolve().parents[2]
TEMPLATE = PLUGIN_ROOT / "templates" / "INFOR Ownership Template.xlsx"


def _build(tmp_path, insiders, total=261_000_000):
    return build_ownership_workbook(
        template_path=TEMPLATE,
        insiders=insiders,
        total_shares_outstanding=total,
        output_path=tmp_path / "Ownership.xlsx",
    )


def test_writes_insider_block_with_names_shares_dates_and_roles(tmp_path: Path):
    out = _build(
        tmp_path,
        [
            InsiderHolding("Barrenechea, Mark James", "Mark Barrenechea (CEO & Director)", 1219092, "2025-03-31"),
            InsiderHolding("Fowlie, Randy", "Randy Fowlie (Director)", [193000, 0, 0], "2025-12-01"),
        ],
    )
    ws = load_workbook(out)["Ownership"]  # formulas preserved (no data_only)

    # Row 39 — single tranche -> plain int.
    assert ws["B39"].value == "Barrenechea, Mark James"
    assert ws["F39"].value == 1219092
    # openpyxl reads a date-formatted cell back as datetime; compare the date part.
    g39 = ws["G39"].value
    assert (g39.date() if isinstance(g39, datetime) else g39) == date(2025, 3, 31)
    assert ws["G39"].number_format == "yyyy-mm-dd"
    assert ws["J39"].value == "Mark Barrenechea (CEO & Director)"
    # Hardcoded values are blue; H/I template formulas are preserved.
    assert str(ws["F39"].font.color.rgb).endswith("0000FF")
    assert ws["H39"].value == 1
    assert ws["I39"].value == "=H39*F39"

    # Row 40 — multiple tranches -> in-cell sum formula (not pre-computed).
    assert ws["F40"].value == "=193000+0+0"
    assert ws["J40"].value == "Randy Fowlie (Director)"

    # F35 (% denominator) is written with a source comment.
    assert ws["F35"].value == 261_000_000
    assert ws["F35"].comment is not None and "cap table" in ws["F35"].comment.text


def test_accepts_plain_dicts_and_handles_missing_date(tmp_path: Path):
    out = _build(
        tmp_path,
        [{"sedi_name": "Doe, Jane", "adjusted_name": "Jane Doe (CFO)", "common_shares": 500}],
    )
    ws = load_workbook(out)["Ownership"]
    assert ws["B39"].value == "Doe, Jane"
    assert ws["F39"].value == 500
    assert ws["G39"].value is None  # no date supplied


def test_too_many_insiders_raises(tmp_path: Path):
    too_many = [
        InsiderHolding(f"Name {i}", f"Name {i} (Director)", i + 1, "2025-01-01") for i in range(28)
    ]
    with pytest.raises(ValueError, match="insider rows"):
        _build(tmp_path, too_many)


def test_template_and_output_carry_no_external_links_or_defined_names(tmp_path: Path):
    """The shipped template is pre-cleaned of dead external links + legacy
    defined names so the openpyxl output opens in Excel for the render step.
    Guards against a re-dirtied template and confirms the build stays clean."""
    src = load_workbook(TEMPLATE)
    assert list(getattr(src, "_external_links", [])) == [], "shipped template must carry no external links"
    assert len(src.defined_names) == 0, "shipped template must carry no defined names"
    out = _build(tmp_path, [InsiderHolding("Doe, Jane", "Jane Doe (CFO)", 500, "2025-01-01")])
    cleaned = load_workbook(out)
    assert list(getattr(cleaned, "_external_links", [])) == []
    assert len(cleaned.defined_names) == 0


def test_f35_keeps_template_palatino_font(tmp_path: Path):
    """F35 keeps the template's Palatino typeface — a bare Font() would reset it to
    Calibri 11. The aggregator later relinks F35 to the cap table, preserving font."""
    out = _build(tmp_path, [InsiderHolding("Doe, Jane", "Jane Doe (CFO)", 500, "2025-01-01")])
    f35 = load_workbook(out)["Ownership"]["F35"].font
    assert f35.name == "Palatino Linotype", "F35 must keep the template's Palatino font"
    # The hardcoded insider data cells stay blue but also keep Palatino (not Calibri 11).
    f39 = load_workbook(out)["Ownership"]["F39"].font
    assert f39.name == "Palatino Linotype", "insider data cells must keep Palatino"


def test_total_shares_none_leaves_f35_blank(tmp_path: Path):
    out = build_ownership_workbook(
        template_path=TEMPLATE,
        insiders=[InsiderHolding("Doe, Jane", "Jane Doe (CFO)", 500, "2025-01-01")],
        total_shares_outstanding=None,
        output_path=tmp_path / "Ownership.xlsx",
    )
    assert load_workbook(out)["Ownership"]["F35"].value is None


def _cap_table(path: Path, rows: dict[int, float]) -> Path:
    wb = Workbook()
    ws = wb.active
    ws.title = "Cap with Links"
    for row, value in rows.items():
        ws.cell(row=row, column=6, value=value)  # column F
    wb.save(path)
    return path


def test_read_basic_shares_sums_section_vii_to_full_units(tmp_path: Path):
    # Section VII basic-share inputs are in millions; reader returns full units.
    cap = _cap_table(tmp_path / "cap.xlsx", {168: 100.0, 169: 72.34})
    assert read_basic_shares_from_cap_table(cap) == 172_340_000


def test_read_basic_shares_missing_or_empty_returns_none(tmp_path: Path):
    assert read_basic_shares_from_cap_table(tmp_path / "nope.xlsx") is None
    empty = _cap_table(tmp_path / "empty.xlsx", {})
    assert read_basic_shares_from_cap_table(empty) is None
