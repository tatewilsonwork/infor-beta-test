"""Unit tests for the workbook aggregator helper.

These exercise the platform-independent pieces — filename normalization,
Excel-safe sheet naming, and the openpyxl merge backend — so they run on
any platform without Microsoft Excel. The COM backend is covered only on
Windows-with-Excel runtimes and is not unit-tested here.
"""

from pathlib import Path

import pytest
from openpyxl import Workbook, load_workbook

from workbook_aggregator import (
    _combine_via_openpyxl,
    _excel_safe_sheet_name,
    _unique_sheet_name,
    combine_workbooks,
    combined_filename,
)


def _make_workbook(path: Path, sheets: dict[str, list[list]]) -> Path:
    wb = Workbook()
    wb.remove(wb.active)
    for title, rows in sheets.items():
        ws = wb.create_sheet(title=title)
        for r, row in enumerate(rows, start=1):
            for c, value in enumerate(row, start=1):
                ws.cell(row=r, column=c, value=value)
    wb.save(path)
    return path


# --- combined_filename -------------------------------------------------------


def test_combined_filename_drops_hyphens_from_deliverable():
    assert combined_filename("earnings-update", "Project Atlas") == "earningsupdate-Project Atlas.xlsx"


def test_combined_filename_keeps_plain_deliverable():
    assert combined_filename("pitch", "Project Atlas") == "pitch-Project Atlas.xlsx"


def test_combined_filename_sanitizes_deal_name():
    name = combined_filename("pitch", "Project: A/B?")
    assert name == "pitch-Project- A-B-.xlsx"


# --- sheet-name rules --------------------------------------------------------


def test_excel_safe_sheet_name_strips_forbidden_and_truncates():
    assert _excel_safe_sheet_name("Cap:With/Links") == "Cap-With-Links"
    assert len(_excel_safe_sheet_name("x" * 50)) == 31


def test_unique_sheet_name_disambiguates_collisions():
    used: set[str] = set()
    assert _unique_sheet_name("captable", used) == "captable"
    assert _unique_sheet_name("captable", used) == "captable (2)"
    assert _unique_sheet_name("CAPTABLE", used) == "CAPTABLE (3)"


# --- openpyxl merge backend --------------------------------------------------


def test_single_sheet_source_tab_named_after_skill(tmp_path: Path):
    a = _make_workbook(tmp_path / "a.xlsx", {"LTM Revenue": [["x", 1]]})
    out = tmp_path / "out.xlsx"
    _combine_via_openpyxl([("ltm-revenue", a)], out)
    wb = load_workbook(out)
    assert wb.sheetnames == ["ltm-revenue"]


def test_multi_sheet_source_tabs_prefixed_with_skill(tmp_path: Path):
    a = _make_workbook(tmp_path / "a.xlsx", {"Cap with Links": [["a"]], "Inputs": [["b"]]})
    out = tmp_path / "out.xlsx"
    _combine_via_openpyxl([("captable", a)], out)
    wb = load_workbook(out)
    assert wb.sheetnames == ["captable-Cap with Links", "captable-Inputs"]


def test_capiq_helper_sheet_is_dropped(tmp_path: Path):
    # The CapIQ add-in's "__snloffice" metadata sheet must not surface as a tab;
    # with it filtered out the lone content sheet collapses to the skill name.
    a = _make_workbook(
        tmp_path / "a.xlsx",
        {"__snloffice": [["garbage"]], "Cap with Links": [["a"]]},
    )
    out = tmp_path / "out.xlsx"
    _combine_via_openpyxl([("captable", a)], out)
    wb = load_workbook(out)
    assert wb.sheetnames == ["captable"]


def test_merge_preserves_values_and_formulas(tmp_path: Path):
    a = _make_workbook(tmp_path / "a.xlsx", {"S": [["Total", 10], [None, "=B1*2"]]})
    out = tmp_path / "out.xlsx"
    _combine_via_openpyxl([("comps", a)], out)
    ws = load_workbook(out)["comps"]
    assert ws["A1"].value == "Total"
    assert ws["B1"].value == 10
    assert ws["B2"].value == "=B1*2"


# --- combine_workbooks end-to-end (openpyxl path off-Windows) ---------------


def test_combine_writes_named_file_and_deletes_sources(tmp_path: Path, monkeypatch):
    # Force the openpyxl backend so the test never shells out to Excel.
    monkeypatch.setattr("workbook_aggregator.sys.platform", "linux")
    a = _make_workbook(tmp_path / "cap.xlsx", {"Cap": [[1]]})
    b = _make_workbook(tmp_path / "ltm.xlsx", {"LTM": [[2]]})

    out = combine_workbooks(
        sources={"captable": a, "ltm-revenue": b},
        output_dir=tmp_path,
        deliverable_type="earnings-update",
        deal_name="Project Atlas",
    )

    assert out.name == "earningsupdate-Project Atlas.xlsx"
    assert out.exists()
    assert load_workbook(out).sheetnames == ["captable", "ltm-revenue"]
    # Sources replaced by the combined workbook.
    assert not a.exists()
    assert not b.exists()


def test_combine_skips_none_and_missing_sources(tmp_path: Path, monkeypatch):
    monkeypatch.setattr("workbook_aggregator.sys.platform", "linux")
    a = _make_workbook(tmp_path / "cap.xlsx", {"Cap": [[1]]})

    out = combine_workbooks(
        sources={
            "captable": a,
            "comps": None,
            "precedents": tmp_path / "does-not-exist.xlsx",
        },
        output_dir=tmp_path,
        deliverable_type="pitch",
        deal_name="Project Atlas",
    )
    assert load_workbook(out).sheetnames == ["captable"]


def test_combine_keeps_sources_when_requested(tmp_path: Path, monkeypatch):
    monkeypatch.setattr("workbook_aggregator.sys.platform", "linux")
    a = _make_workbook(tmp_path / "cap.xlsx", {"Cap": [[1]]})
    out = combine_workbooks(
        sources={"captable": a},
        output_dir=tmp_path,
        deliverable_type="pitch",
        deal_name="Atlas",
        delete_sources=False,
    )
    assert out.exists()
    assert a.exists()


def test_combine_raises_when_no_valid_sources(tmp_path: Path, monkeypatch):
    monkeypatch.setattr("workbook_aggregator.sys.platform", "linux")
    with pytest.raises(ValueError):
        combine_workbooks(
            sources={"comps": None},
            output_dir=tmp_path,
            deliverable_type="pitch",
            deal_name="Atlas",
        )
