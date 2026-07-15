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
    bloomberg_matches_sedi,
    build_ownership_workbook,
    match_bloomberg_to_sedi,
    read_basic_shares_from_cap_table,
    read_bloomberg_export,
    strip_legal_suffixes,
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
    # Section VII sentinel labels — read_basic_shares_from_cap_table verifies
    # these anchors (template_layout) before summing the F168:F185 window.
    ws["B166"] = "VII. BASIC SHARES OUTSTANDING"
    ws["B167"] = "Description"
    ws["F167"] = "Amount"
    ws["B186"] = "Total Basic Shares Outstanding"
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


# ─── Bloomberg institutional side ─────────────────────────────────────────────

def _bbg_export(path: Path, holders, sheet_title: str = "Summary View") -> Path:
    """A minimal BBG Excel add-in ownership export (Summary View layout).

    ``holders`` is a list of ``(name, position, insider_status)`` tuples.
    """
    wb = Workbook()
    ws = wb.active
    ws.title = sheet_title
    ws["C7"] = "Name"
    ws["E7"] = "Test Co Inc."
    ws["C9"] = "Ticker"
    ws["E9"] = "TST CN EQUITY"
    ws["G9"] = "View"
    ws["I9"] = "Summary"
    ws["C11"] = "Sort By"
    ws["E11"] = "Position"
    ws["G11"] = "Order"
    ws["I11"] = "Descending"
    for col, header in {
        "C": "Holder Name", "G": "Portfolio Name", "L": "Position", "M": "Latest Chg",
        "N": "Filing Date", "O": "Source", "P": "% Out", "R": "Insider Status",
        "S": "Institution Type", "U": "Country", "V": "Market Value",
    }.items():
        ws[f"{col}13"] = header
    for i, (name, position, status) in enumerate(holders):
        row = 14 + i
        ws[f"B{row}"] = i + 1
        ws[f"C{row}"] = name
        ws[f"L{row}"] = position
        ws[f"N{row}"] = datetime(2026, 5, 21)
        ws[f"N{row}"].number_format = "m/d/yyyy"
        ws[f"O{row}"] = "Form 4" if status == "Y" else "ULT-AGG"
        ws[f"R{row}"] = status
        ws[f"S{row}"] = "Unclassified" if status == "Y" else "Investment Advisor"
    wb.save(path)
    return path


_SAMPLE_HOLDERS = [
    ("Malchow Eric D", 4_972_500, "Y"),
    ("T Rowe Price Group Inc", 150_311, "N-P"),
    ("Kelso & Co LP", 14_983, "N-P"),
]


@pytest.mark.parametrize(
    ("raw", "adjusted"),
    [
        ("T Rowe Price Group Inc", "T Rowe Price"),
        ("Kelso & Co LP", "Kelso & Co"),  # '& Co' is the brand, not a suffix
        ("Vanguard Group Inc", "Vanguard"),
        ("BlackRock Inc.", "BlackRock"),
        ("Malchow Eric D", "Malchow Eric D"),  # person names pass through
        ("Royal Bank of Canada", "Royal Bank of Canada"),
        ("Inc", "Inc"),  # never stripped to nothing
    ],
)
def test_strip_legal_suffixes(raw: str, adjusted: str):
    assert strip_legal_suffixes(raw) == adjusted


@pytest.mark.parametrize(
    ("bbg", "sedi", "expected"),
    [
        ("Malchow Eric D", "Malchow, Eric D", True),
        ("Malchow Eric", "Malchow, Eric D", True),  # middle initial missing on one side
        ("Weber Mary R.", "Weber, Mary R", True),  # punctuation-insensitive
        ("Smith M Christine", "Smith, M Christine", True),
        ("Smith J", "Smith, Jane", True),  # initial-only given name
        ("Van Der Berg Jan", "Van Der Berg, Jan", True),  # multi-token surname
        ("Fairfax Financial Holdings Ltd", "Fairfax Financial Holdings Limited", True),  # corporate insider
        ("Smith M Christine", "Smith, Mary", False),  # full given names disagree
        ("Brown Robert T", "Brown, Richard", False),
        ("T Rowe Price Group Inc", "Barrenechea, Mark", False),
        ("Malchow Eric D", "Malchowski, Eric", False),  # surname must match exactly
    ],
)
def test_bloomberg_matches_sedi(bbg: str, sedi: str, expected: bool):
    assert bloomberg_matches_sedi(bbg, sedi) is expected


def test_match_bloomberg_to_sedi_maps_first_match_only():
    matches = match_bloomberg_to_sedi(
        [h[0] for h in _SAMPLE_HOLDERS], ["Malchow, Eric D", "Brown, Robert T"]
    )
    assert matches == {"Malchow Eric D": "Malchow, Eric D"}


def test_read_bloomberg_export_parses_summary_view(tmp_path: Path):
    src = _bbg_export(tmp_path / "bbg.xlsx", _SAMPLE_HOLDERS)
    export = read_bloomberg_export(src)
    assert export.company_name == "Test Co Inc."
    assert export.ticker == "TST CN EQUITY"
    assert [h.name for h in export.holders] == [h[0] for h in _SAMPLE_HOLDERS]
    assert export.holders[0].position == 4_972_500
    assert export.holders[0].insider_status == "Y"
    assert export.holders[1].institution_type == "Investment Advisor"
    assert export.holders[0].number_formats["N"] == "m/d/yyyy"
    assert export.info["E7"] == "Test Co Inc."


def test_read_bloomberg_export_finds_sheet_by_header_row(tmp_path: Path):
    src = _bbg_export(tmp_path / "bbg.xlsx", _SAMPLE_HOLDERS, sheet_title="My Export")
    assert [h.name for h in read_bloomberg_export(src).holders] == [h[0] for h in _SAMPLE_HOLDERS]


def test_read_bloomberg_export_rejects_non_bbg_workbook(tmp_path: Path):
    wb = Workbook()
    wb.active["A1"] = "not a bloomberg export"
    wb.save(tmp_path / "other.xlsx")
    with pytest.raises(ValueError, match="Summary View"):
        read_bloomberg_export(tmp_path / "other.xlsx")
    with pytest.raises(FileNotFoundError):
        read_bloomberg_export(tmp_path / "missing.xlsx")


def _build_with_bbg(tmp_path: Path, holders=None, **kwargs):
    src = _bbg_export(tmp_path / "bbg.xlsx", holders if holders is not None else _SAMPLE_HOLDERS)
    return build_ownership_workbook(
        template_path=TEMPLATE,
        insiders=[
            InsiderHolding("Malchow, Eric D", "Eric Malchow (CEO & Director)", 4_972_500, "2026-05-21"),
            InsiderHolding("Brown, Robert T", "Robert Brown (Director)", 1_665_600, "2026-05-21"),
        ],
        total_shares_outstanding=31_650_000,
        output_path=tmp_path / "Ownership.xlsx",
        bloomberg_export_path=src,
        **kwargs,
    )


def test_build_with_bloomberg_fills_output_tab_and_links(tmp_path: Path):
    out = _build_with_bbg(tmp_path)
    wb = load_workbook(out)
    bbg = wb["Bloomberg Output"]
    own = wb["Ownership"]

    # Export copied into the Bloomberg Output tab (values + formats + info cells).
    assert bbg["E7"].value == "Test Co Inc."
    assert bbg["E9"].value == "TST CN EQUITY"
    assert bbg["C14"].value == "Malchow Eric D"
    assert bbg["L14"].value == 4_972_500
    assert bbg["R14"].value == "Y"
    assert bbg["N14"].number_format == "m/d/yyyy"
    assert bbg["C16"].value == "Kelso & Co LP"

    # The Ownership tab's pre-wired link formulas are untouched.
    assert own["B68"].value == "='Bloomberg Output'!C14"
    assert "XLOOKUP" in own["F68"].value

    # SEDI duplicate excluded (H=0 + audit comment); institutions stay included.
    assert own["H68"].value == 0
    assert str(own["H68"].font.color.rgb).endswith("0000FF")
    assert "Malchow, Eric D" in own["H68"].comment.text
    assert own["H69"].value == 1
    assert own["H70"].value == 1

    # Adjusted display names (col J) for every populated Bloomberg row.
    assert own["J68"].value == "Malchow Eric D"
    assert own["J69"].value == "T Rowe Price"
    assert own["J70"].value == "Kelso & Co"

    # Unused link rows are neutralised so LARGE($I$68:$I$185) stays numeric.
    for row in (71, 120, 185):
        assert own[f"B{row}"].value is None
        assert own[f"F{row}"].value is None
        assert own[f"G{row}"].value is None
        assert own[f"H{row}"].value == 0
        assert own[f"I{row}"].value == f"=H{row}*F{row}"

    # Insider block untouched by the Bloomberg write.
    assert own["B39"].value == "Malchow, Eric D"
    assert own["F35"].value == 31_650_000


def test_build_with_bloomberg_honours_overrides(tmp_path: Path):
    out = _build_with_bbg(
        tmp_path,
        bloomberg_adjusted_names={"T Rowe Price Group Inc": "T. Rowe Price"},
        bloomberg_include_overrides={"Kelso & Co LP": 0, "Malchow Eric D": 1},
    )
    own = load_workbook(out)["Ownership"]
    assert own["J69"].value == "T. Rowe Price"
    assert own["H70"].value == 0
    assert "override" in own["H70"].comment.text
    # Forced include on a matched duplicate leaves the template's 1 in place.
    assert own["H68"].value == 1
    assert own["H68"].comment is None


def test_build_without_bloomberg_leaves_institutional_side_untouched(tmp_path: Path):
    out = _build(tmp_path, [InsiderHolding("Doe, Jane", "Jane Doe (CFO)", 500, "2025-01-01")])
    wb = load_workbook(out)
    own = wb["Ownership"]
    assert wb["Bloomberg Output"]["C14"].value is None
    assert own["B68"].value == "='Bloomberg Output'!C14"
    assert own["H68"].value == 1
    assert own["H185"].value == 1
    assert own["J68"].value is None


def test_build_with_bloomberg_truncates_at_118_holders(tmp_path: Path):
    holders = [(f"Holder {i}", 1_000 - i, "N-P") for i in range(1, 121)]  # 120 holders
    out = _build_with_bbg(tmp_path, holders=holders)
    wb = load_workbook(out)
    bbg = wb["Bloomberg Output"]
    assert bbg["C131"].value == "Holder 118"
    assert bbg["C132"].value is None
    own = wb["Ownership"]
    assert own["H185"].value == 1  # all 118 slots used, none neutralised
    assert own["J185"].value == "Holder 118"


def test_build_with_bloomberg_stays_excel_clean(tmp_path: Path):
    out = _build_with_bbg(tmp_path)
    cleaned = load_workbook(out)
    assert list(getattr(cleaned, "_external_links", [])) == []
    assert len(cleaned.defined_names) == 0
