"""Unit tests for the workbook aggregator helper.

These exercise the platform-independent pieces — filename normalization,
Excel-safe sheet naming, theme resolution, and the openpyxl merge backend — so
they run on any platform without Microsoft Excel. The COM backend is covered
only on Windows-with-Excel runtimes and is not unit-tested here.
"""

import re
import zipfile
from pathlib import Path

import pytest
from openpyxl import Workbook, load_workbook

from workbook_aggregator import (
    _combine_via_openpyxl,
    _default_theme_path,
    _excel_safe_sheet_name,
    _extract_theme_xml,
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
    a = _make_workbook(tmp_path / "a.xlsx", {"LTM Metrics": [["x", 1]]})
    out = tmp_path / "out.xlsx"
    _combine_via_openpyxl([("ltm-metrics", a)], out)
    wb = load_workbook(out)
    assert wb.sheetnames == ["ltm-metrics"]


def test_multi_sheet_source_keeps_original_sheet_names(tmp_path: Path):
    # A multi-sheet source keeps its (self-describing) sheet names unprefixed —
    # the skill prefix is only used to label a single-sheet source.
    a = _make_workbook(tmp_path / "a.xlsx", {"Ownership": [["a"]], "Bloomberg Output": [["b"]]})
    out = tmp_path / "out.xlsx"
    _combine_via_openpyxl([("ownership", a)], out)
    wb = load_workbook(out)
    assert wb.sheetnames == ["Ownership", "Bloomberg Output"]


def test_multi_sheet_source_preserves_cross_sheet_refs(tmp_path: Path):
    # Because multi-sheet tabs keep their original names, a sheet's formula that
    # references a sibling sheet by name stays valid (no rename -> no #REF). This
    # is the ownership `Ownership` -> `Bloomberg Output` case in miniature.
    a = _make_workbook(
        tmp_path / "a.xlsx",
        {"Ownership": [["=+'Bloomberg Output'!C14"]], "Bloomberg Output": [["x"]]},
    )
    out = tmp_path / "out.xlsx"
    _combine_via_openpyxl([("ownership", a)], out)
    wb = load_workbook(out)
    assert "Bloomberg Output" in wb.sheetnames
    assert wb["Ownership"]["A1"].value == "=+'Bloomberg Output'!C14"


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
        sources={"captable": a, "ltm-metrics": b},
        output_dir=tmp_path,
        deliverable_type="earnings-update",
        deal_name="Project Atlas",
    )

    assert out.name == "earningsupdate-Project Atlas.xlsx"
    assert out.exists()
    assert load_workbook(out).sheetnames == ["captable", "ltm-metrics"]
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


# --- cross-tab relink --------------------------------------------------------


def test_relink_wires_cross_tab_formulas(tmp_path: Path, monkeypatch):
    """In the combined workbook, the cap table's LTM cells link to the ltm-metrics
    bridge totals (found by label, since their rows are dynamic) and the ownership
    denominator links to the cap table's basic shares (millions -> full units)."""
    monkeypatch.setattr("workbook_aggregator.sys.platform", "linux")
    cap = _make_workbook(tmp_path / "cap.xlsx", {"Cap with Links": [["cap"]]})
    ltm = _make_workbook(
        tmp_path / "ltm.xlsx",
        {
            "LTM Metrics": [
                ["LTM Revenue Overview", None],
                ["Segment A", 100],
                ["Total", 100],
                [None, None],
                ["LTM Revenue Bridge", None],
                ["Component", "Amount"],
                ["(+) FY", 90],
                ["(+) YTD", 30],
                ["(−) Prior YTD", 20],
                ["(=) LTM Revenue", "=B7+B8-B9"],      # row 10
                [None, None],
                ["LTM Adj. EBITDA Bridge", None],
                ["Component", "Amount"],
                ["(+) FY EBITDA", 40],
                ["(=) LTM Adj. EBITDA", "=B14"],        # row 15
            ]
        },
    )
    own = _make_workbook(tmp_path / "own.xlsx", {"Ownership": [["own"]]})

    out = combine_workbooks(
        sources={"captable": cap, "ltm-metrics": ltm, "ownership": own},
        output_dir=tmp_path,
        deliverable_type="pitch",
        deal_name="Atlas",
    )
    wb = load_workbook(out)
    assert wb["captable"]["D47"].value == "='ltm-metrics'!B10*F7"
    assert wb["captable"]["D48"].value == "='ltm-metrics'!B15*F7"
    assert wb["ownership"]["F35"].value == "='captable'!F17*1000000"


def test_relink_is_noop_without_ltm_bridge_labels(tmp_path: Path, monkeypatch):
    # No "(=) LTM ..." labels in the ltm tab -> the cap table's D47/D48 are left
    # alone (the captable skill's own scalar/CapIQ values stand).
    monkeypatch.setattr("workbook_aggregator.sys.platform", "linux")
    cap = _make_workbook(tmp_path / "cap.xlsx", {"Cap with Links": [["cap"]]})
    ltm = _make_workbook(tmp_path / "ltm.xlsx", {"LTM Metrics": [["Total", 100]]})
    out = combine_workbooks(
        sources={"captable": cap, "ltm-metrics": ltm},
        output_dir=tmp_path,
        deliverable_type="pitch",
        deal_name="Atlas",
    )
    wb = load_workbook(out)
    assert wb["captable"]["D47"].value is None
    assert wb["captable"]["D48"].value is None


def test_financial_summary_ltm_link_resolves_after_merge(tmp_path: Path, monkeypatch):
    """The financial-summary tab's label-keyed LTM links (written before the
    ltm-metrics tab exists) survive the merge and target the renamed 'ltm-metrics'
    tab — so they resolve in the combined workbook, like the cap table's CapIQ
    formulas."""
    monkeypatch.setattr("workbook_aggregator.sys.platform", "linux")
    from financial_summary_workbook import MetricSeries, build_financial_summary_workbook

    fs = build_financial_summary_workbook(
        company_name="SampleCo",
        currency_note="Figures in US$MM",
        period_note="FY = fiscal year; LTM = trailing twelve months",
        fiscal_labels=["FY2021", "FY2022", "FY2023", "FY2024", "FY2025"],
        metrics=[
            MetricSeries("Revenue", "US$MM", [1, 2, 3, 4, 5], result_label="LTM Revenue"),
            MetricSeries("Gross Profit", "US$MM", [1, 2, 3, 4, 5], result_label="LTM Gross Profit"),
            MetricSeries("Adjusted EBITDA", "US$MM", [1, 2, 3, 4, 5], result_label="LTM Adj. EBITDA"),
            MetricSeries("Net Income", "US$MM", [1, 2, 3, 4, 5], result_label="LTM Net Income"),
        ],
        output_dir=tmp_path,
        file_stem="fs",
    )
    ltm = _make_workbook(tmp_path / "ltm.xlsx", {"LTM Metrics": [["(=) LTM Revenue", 4520.0]]})

    out = combine_workbooks(
        sources={"financial-summary": fs, "ltm-metrics": ltm},
        output_dir=tmp_path,
        deliverable_type="pitch",
        deal_name="Atlas",
    )
    wb = load_workbook(out)
    assert "financial-summary" in wb.sheetnames
    assert "ltm-metrics" in wb.sheetnames  # the link's target tab is present post-merge
    assert wb["financial-summary"]["G6"].value == (
        "=INDEX('ltm-metrics'!$B:$B, MATCH(\"(=) LTM Revenue\", 'ltm-metrics'!$A:$A, 0))"
    )


def test_internalize_external_sheet_ref_rewrites_workbook_index_and_path_forms():
    from workbook_aggregator import _internalize_external_sheet_ref

    idx_form = "=INDEX('[1]ltm-metrics'!$B:$B, MATCH(\"(=) LTM Revenue\", '[1]ltm-metrics'!$A:$A, 0))"
    assert _internalize_external_sheet_ref(idx_form, "ltm-metrics", "ltm-metrics") == (
        "=INDEX('ltm-metrics'!$B:$B, MATCH(\"(=) LTM Revenue\", 'ltm-metrics'!$A:$A, 0))"
    )
    # Full-path external form (Excel emits this once the source workbook is closed).
    path_form = "=INDEX('C:\\d\\[SampleCo - LTM Metrics.xlsx]ltm-metrics'!$B:$B, MATCH(\"k\", 'C:\\d\\[SampleCo - LTM Metrics.xlsx]ltm-metrics'!$A:$A, 0))"
    assert _internalize_external_sheet_ref(path_form, "ltm-metrics", "ltm-metrics") == (
        "=INDEX('ltm-metrics'!$B:$B, MATCH(\"k\", 'ltm-metrics'!$A:$A, 0))"
    )
    # A purely-internal ref (the openpyxl path) is left untouched.
    internal = "=INDEX('ltm-metrics'!$B:$B, MATCH(\"k\", 'ltm-metrics'!$A:$A, 0))"
    assert _internalize_external_sheet_ref(internal, "ltm-metrics", "ltm-metrics") == internal


def test_relink_financial_summary_openpyxl_internalizes_external_links(tmp_path: Path):
    from workbook_aggregator import _relink_financial_summary_openpyxl

    wb = Workbook()
    wb.remove(wb.active)
    fs = wb.create_sheet("financial-summary")
    wb.create_sheet("ltm-metrics")
    # Simulate the external link a COM copy would leave behind.
    fs["G6"] = "=INDEX('[1]ltm-metrics'!$B:$B, MATCH(\"(=) LTM Revenue\", '[1]ltm-metrics'!$A:$A, 0))"
    fs["B6"] = 5.0  # a non-LTM numeric cell is left alone
    _relink_financial_summary_openpyxl(wb, {"financial-summary": "financial-summary", "ltm-metrics": "ltm-metrics"})
    assert fs["G6"].value == (
        "=INDEX('ltm-metrics'!$B:$B, MATCH(\"(=) LTM Revenue\", 'ltm-metrics'!$A:$A, 0))"
    )
    assert fs["B6"].value == 5.0


# --- brand theme -------------------------------------------------------------

_INFOR_ACCENT1 = "0E213F"  # INFOR (New) accent1 — distinguishes it from Office's 4F81BD


def test_infor_theme_is_shipped():
    # The aggregator stamps templates/INFORFG.thmx on the combined workbook, so
    # the file must ship and parse to the INFOR colour scheme.
    theme = _default_theme_path()
    assert theme.exists(), f"shipped theme missing: {theme}"
    xml = _extract_theme_xml(theme)
    assert xml is not None
    assert _INFOR_ACCENT1.encode() in xml


def test_extract_theme_xml_returns_none_for_bad_file(tmp_path: Path):
    bad = tmp_path / "not-a-theme.thmx"
    bad.write_text("not a zip")
    assert _extract_theme_xml(bad) is None
    assert _extract_theme_xml(tmp_path / "missing.thmx") is None


def test_combined_workbook_carries_infor_theme(tmp_path: Path, monkeypatch):
    # End-to-end (openpyxl backend): a fresh openpyxl workbook would carry the
    # default Office theme; the aggregator must inject the INFOR theme instead.
    monkeypatch.setattr("workbook_aggregator.sys.platform", "linux")
    a = _make_workbook(tmp_path / "cap.xlsx", {"Cap with Links": [["a"]]})
    out = combine_workbooks(
        sources={"captable": a},
        output_dir=tmp_path,
        deliverable_type="pitch",
        deal_name="Atlas",
    )
    with zipfile.ZipFile(out) as z:
        theme_xml = z.read("xl/theme/theme1.xml").decode("utf-8", "replace")
    accent1 = re.search(r"<a:accent1>\s*<a:srgbClr val=\"([0-9A-Fa-f]{6})\"", theme_xml)
    assert accent1 is not None and accent1.group(1).upper() == _INFOR_ACCENT1


def test_theme_override_can_be_disabled(tmp_path: Path, monkeypatch):
    # Passing a non-existent theme path leaves the default Office theme in place
    # (no INFOR injection) — proving the stamp is driven by theme_path.
    monkeypatch.setattr("workbook_aggregator.sys.platform", "linux")
    a = _make_workbook(tmp_path / "cap.xlsx", {"Cap with Links": [["a"]]})
    out = combine_workbooks(
        sources={"captable": a},
        output_dir=tmp_path,
        deliverable_type="pitch",
        deal_name="Atlas",
        theme_path=tmp_path / "does-not-exist.thmx",
    )
    with zipfile.ZipFile(out) as z:
        theme_xml = z.read("xl/theme/theme1.xml").decode("utf-8", "replace")
    accent1 = re.search(r"<a:accent1>\s*<a:srgbClr val=\"([0-9A-Fa-f]{6})\"", theme_xml)
    assert accent1 is not None and accent1.group(1).upper() != _INFOR_ACCENT1
