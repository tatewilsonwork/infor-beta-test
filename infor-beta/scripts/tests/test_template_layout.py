"""Tests for the template_layout runtime-verification map.

Two halves, per the design contract:

(a) every sentinel anchor / slide marker passes against the SHIPPED templates
    (verification is a no-op change for an unmodified plugin), and
(b) mutating a temp COPY of a template (openpyxl ``insert_rows`` / a
    python-pptx slide re-order) makes the wired-in writers raise
    ``TemplateLayoutError`` with a message naming the template, the protected
    address, what was expected, and what was found.

The shipped templates are never modified — every mutation happens on a copy
under ``tmp_path``.
"""

from __future__ import annotations

import shutil
from pathlib import Path

import pytest
from openpyxl import Workbook, load_workbook
from pptx import Presentation

import template_layout as tl
from template_layout import CellAnchor, TemplateLayoutError

PLUGIN_ROOT = Path(__file__).resolve().parents[2]
TEMPLATES = PLUGIN_ROOT / "templates"
CAP_TABLE = TEMPLATES / "INFOR Cap Table Template.xlsx"
OWNERSHIP = TEMPLATES / "INFOR Ownership Template.xlsx"
COMPS = TEMPLATES / "INFOR Comps Template.xlsx"
PRECEDENTS = TEMPLATES / "INFOR Precedents Template.xlsx"
LIBRARY = TEMPLATES / "INFOR Slide Library.pptx"


# ─── (a) every verification passes against the shipped templates ─────────────


def test_cap_table_template_passes_every_anchor_group():
    ws = load_workbook(CAP_TABLE)[tl.CAP_TABLE_SHEET]
    tl.verify_cap_table_before_write(ws)  # header + LTM + Section VII
    tl.verify_anchors(ws, tl.CAP_TABLE_PICTURE_ANCHORS, template=tl.CAP_TABLE_TEMPLATE)


def test_ownership_template_passes_every_anchor_group():
    wb = load_workbook(OWNERSHIP)
    ws = wb[tl.OWNERSHIP_SHEET]
    tl.verify_anchors(
        ws,
        tl.OWNERSHIP_INSIDER_BLOCK_ANCHORS
        + tl.OWNERSHIP_TOTAL_SHARES_ANCHORS
        + tl.OWNERSHIP_BBG_LINK_ANCHORS
        + tl.OWNERSHIP_INSIDERS_PICTURE_ANCHORS
        + tl.OWNERSHIP_INSTITUTIONS_PICTURE_ANCHORS,
        template=tl.OWNERSHIP_TEMPLATE,
    )
    tl.verify_anchors(
        wb[tl.OWNERSHIP_BBG_SHEET],
        tl.OWNERSHIP_BBG_TEMPLATE_ANCHORS,
        template=tl.OWNERSHIP_TEMPLATE,
    )


def test_comps_template_passes_every_anchor_group():
    ws = load_workbook(COMPS)[tl.COMPS_SHEET]
    tl.verify_anchors(
        ws, tl.COMPS_BLOCK_ANCHORS + tl.COMPS_OUTPUT_CCY_ANCHORS, template=tl.COMPS_TEMPLATE
    )


def test_precedents_template_passes_every_anchor_group():
    ws = load_workbook(PRECEDENTS)[tl.PRECEDENTS_SHEET]
    tl.verify_anchors(
        ws,
        tl.PRECEDENTS_OUTPUT_CCY_ANCHORS + tl.PRECEDENTS_BLOCK_ANCHORS,
        template=tl.PRECEDENTS_TEMPLATE,
    )


def test_slide_library_passes_every_marker():
    prs = Presentation(str(LIBRARY))
    assert len(prs.slides) == tl.SLIDE_LIBRARY_SLIDE_COUNT
    for index in sorted(tl.LIBRARY_SLIDE_MARKERS):
        tl.verify_library_slide(prs, index)


def test_assembler_slide_indices_are_all_covered_by_markers():
    # Every raw library index the assemblers keep, clone, or delete must have a
    # marker — otherwise its verification silently cannot run.
    from earnings_update_assembler import _KEEP_LIBRARY_INDICES
    from pitch_deck_assembler import (
        _EARNINGS_LIBRARY_SLIDE_INDEX,
        _LIBRARY_FINANCIAL_SUMMARY_INDEX,
        _LIBRARY_MARKET_ENTRY_INDEX,
    )

    for index in _KEEP_LIBRARY_INDICES:
        assert index in tl.LIBRARY_SLIDE_MARKERS
    for index in (
        _EARNINGS_LIBRARY_SLIDE_INDEX,
        _LIBRARY_FINANCIAL_SUMMARY_INDEX,
        _LIBRARY_MARKET_ENTRY_INDEX,
    ):
        assert index in tl.LIBRARY_SLIDE_MARKERS


def test_overview_slide_index_is_shared_not_duplicated():
    # v0.5.32 dedupe: pitch_deck_assembler and financial_charts must both read
    # the overview index from template_layout, not carry their own copies.
    import financial_charts
    import pitch_deck_assembler

    assert financial_charts._OVERVIEW_SLIDE_INDEX == tl.OVERVIEW_SLIDE_INDEX
    assert pitch_deck_assembler.OVERVIEW_SLIDE_INDEX is tl.OVERVIEW_SLIDE_INDEX


def test_fs_marker_matches_financial_charts_self_discovery():
    # The FS slide marker must stay in lock-step with the Rectangle 17 pattern
    # financial_charts uses to self-discover FS slides in the assembled deck.
    import financial_charts

    assert tl.MARKER_FINANCIAL_SUMMARY.shape_name == financial_charts._FS_MARKER_PLACEHOLDER


# ─── verify helper semantics ──────────────────────────────────────────────────


def test_verify_anchors_reports_every_mismatch_at_once(tmp_path: Path):
    wb = Workbook()
    ws = wb.active
    ws.title = "S"
    ws["B1"] = "Right Label"
    anchors = (
        CellAnchor("F1", "B1", "Right Label"),
        CellAnchor("F2", "B2", "Missing Label"),
        CellAnchor("F3", "B3", "Also Missing"),
    )
    with pytest.raises(TemplateLayoutError) as exc:
        tl.verify_anchors(ws, anchors, template="Some Template.xlsx")
    msg = str(exc.value)
    assert "Some Template.xlsx" in msg
    assert "F2" in msg and "'Missing Label'" in msg
    assert "F3" in msg and "'Also Missing'" in msg
    assert "F1" not in msg  # the matching anchor is not reported


def test_verify_workbook_anchors_names_a_missing_sheet(tmp_path: Path):
    path = tmp_path / "wrong.xlsx"
    Workbook().save(path)
    with pytest.raises(TemplateLayoutError, match="expected sheet 'Cap with Links'"):
        tl.verify_workbook_anchors(
            path, sheet=tl.CAP_TABLE_SHEET, anchors=tl.CAP_TABLE_PICTURE_ANCHORS
        )


# ─── (b) a shifted Excel template raises through the wired-in writers ─────────


def _shifted_copy(template: Path, dest: Path, sheet: str, at_row: int) -> Path:
    """Copy a template and insert one row at ``at_row`` (shifting rows down)."""
    shutil.copyfile(template, dest)
    wb = load_workbook(dest)
    wb[sheet].insert_rows(at_row)
    wb.save(dest)
    return dest


def test_shifted_cap_table_fails_ltm_anchor_verification(tmp_path: Path):
    # A row inserted above the LTM block shifts D47/D48 — the anchors must
    # catch it and the message must name template, address, expected and found.
    shifted = _shifted_copy(CAP_TABLE, tmp_path / "cap.xlsx", tl.CAP_TABLE_SHEET, 45)
    with pytest.raises(TemplateLayoutError) as exc:
        tl.verify_workbook_anchors(
            shifted, sheet=tl.CAP_TABLE_SHEET, anchors=tl.CAP_TABLE_LTM_ANCHORS
        )
    msg = str(exc.value)
    assert "cap.xlsx" in msg
    assert "D47" in msg and "'Revenue'" in msg


def test_shifted_cap_table_fails_section_vii_read(tmp_path: Path):
    # read_basic_shares_from_cap_table must raise on a shifted Section VII
    # instead of silently summing the wrong F168:F185 window.
    from ownership_workbook import read_basic_shares_from_cap_table

    shifted = _shifted_copy(CAP_TABLE, tmp_path / "cap.xlsx", tl.CAP_TABLE_SHEET, 150)
    with pytest.raises(TemplateLayoutError, match="VII. BASIC SHARES OUTSTANDING"):
        read_basic_shares_from_cap_table(shifted)


def test_shifted_ownership_template_fails_the_builder(tmp_path: Path):
    from ownership_workbook import InsiderHolding, build_ownership_workbook

    shifted = _shifted_copy(OWNERSHIP, tmp_path / "own-template.xlsx", tl.OWNERSHIP_SHEET, 30)
    with pytest.raises(TemplateLayoutError) as exc:
        build_ownership_workbook(
            template_path=shifted,
            insiders=[InsiderHolding("Doe, Jane", "Jane Doe (CFO)", 500)],
            total_shares_outstanding=1_000_000,
            output_path=tmp_path / "Ownership.xlsx",
        )
    msg = str(exc.value)
    assert tl.OWNERSHIP_TEMPLATE in msg
    assert "B38" in msg and "'SEDI Name'" in msg


def test_shifted_comps_template_fails_the_builder(tmp_path: Path):
    from comps_workbook import build_comps_workbook

    shifted = _shifted_copy(COMPS, tmp_path / "comps-template.xlsx", tl.COMPS_SHEET, 8)
    with pytest.raises(TemplateLayoutError, match="Group Average"):
        build_comps_workbook(
            template_path=shifted,
            verticals=[{"name": "Vertical A", "companies": [{"ticker": "TSX:RY"}]}],
            output_path=tmp_path / "Comps.xlsx",
        )


def test_shifted_precedents_template_fails_the_builder(tmp_path: Path):
    from precedents_workbook import build_precedents_workbook

    shifted = _shifted_copy(PRECEDENTS, tmp_path / "prec-template.xlsx", tl.PRECEDENTS_SHEET, 3)
    with pytest.raises(TemplateLayoutError) as exc:
        build_precedents_workbook(
            template_path=shifted,
            groups=[
                {
                    "name": "Group A",
                    "transactions": [
                        {
                            "input_currency": "USD",
                            "announce_date": "2025-01-01",
                            "target": "T",
                            "acquiror": "A",
                            "tev": 100.0,
                            "hq_country": "USA",
                            "revenue_ltm": 50.0,
                        }
                    ],
                }
            ],
            output_path=tmp_path / "Precedents.xlsx",
        )
    assert "'Currency'" in str(exc.value)


def test_combine_workbooks_pre_flight_halts_before_merging(tmp_path: Path, monkeypatch):
    # The aggregator's relink pre-flight runs on the shared layer BEFORE either
    # backend: a captable whose LTM anchors shifted must abort the merge with
    # nothing written and no source deleted.
    from workbook_aggregator import combine_workbooks

    monkeypatch.setattr("workbook_aggregator.sys.platform", "linux")

    cap = tmp_path / "cap.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = tl.CAP_TABLE_SHEET
    ws["B5"] = "Output Currency:"
    ws["B7"] = "FX Rate:"
    ws["D33"] = "LTM"
    ws["B48"] = "Revenue"  # shifted: the labels sit one row below D47/D48
    ws["B49"] = "Adj. EBITDA"
    wb.save(cap)
    ltm = tmp_path / "ltm.xlsx"
    wb2 = Workbook()
    wb2.active.title = "LTM Metrics"
    wb2.active["A1"] = "(=) LTM Revenue"
    wb2.active["B1"] = 100
    wb2.save(ltm)

    with pytest.raises(TemplateLayoutError, match="D47"):
        combine_workbooks(
            sources={"captable": cap, "ltm-metrics": ltm},
            output_dir=tmp_path,
            deliverable_type="pitch",
            deal_name="Atlas",
        )
    assert not (tmp_path / "pitch-Atlas.xlsx").exists()
    assert cap.exists() and ltm.exists()  # nothing merged, nothing deleted


# ─── (b) a re-ordered slide library raises ────────────────────────────────────


def _reordered_library_copy(dest: Path, from_index: int) -> Path:
    """Save a copy of the shipped library with one slide moved to the end."""
    prs = Presentation(str(LIBRARY))
    sld_id_lst = prs.slides._sldIdLst
    element = list(sld_id_lst)[from_index]
    sld_id_lst.remove(element)
    sld_id_lst.append(element)
    prs.save(str(dest))
    return dest


def test_reordered_library_fails_marker_verification(tmp_path: Path):
    # Move the earnings-summary slide (raw index 7) to the end: index 7 is now
    # the Financial Summary slide and every later index shifts by one.
    reordered = _reordered_library_copy(tmp_path / "library.pptx", 7)
    prs = Presentation(str(reordered))
    with pytest.raises(TemplateLayoutError) as exc:
        tl.verify_library_slide(prs, 7)
    msg = str(exc.value)
    assert "slide index 7" in msg
    assert "earnings summary" in msg


def test_reordered_library_fails_the_earnings_assembler(tmp_path: Path):
    # End-to-end: the earnings assembler verifies every kept index before the
    # clone/delete pass, so a re-ordered library halts instead of shipping the
    # wrong slides. Content/plan validation happens before the library opens,
    # so real (minimal) artefacts are built via the wireframe helpers.
    from earnings_update_assembler import _KEEP_LIBRARY_INDICES, assemble_earnings_update_deck
    from earnings_update_wireframe import build_earnings_update_slide_plan, write_slide_plan
    from schemas import Company

    reordered = _reordered_library_copy(tmp_path / "library.pptx", _KEEP_LIBRARY_INDICES[1])

    slide_plan = build_earnings_update_slide_plan(
        company=Company(legal_name="SampleCo", ticker="TSX:SMPL"),
        reporting_quarter="Q4 2025",
        comparison_quarter="Q4 2024",
    )
    slide_plan_path = write_slide_plan(slide_plan, tmp_path / "slide_plan.json")
    from tests.test_earnings_update_assembler import _sample_content

    content_path = tmp_path / "content.json"
    content_path.write_text(_sample_content().model_dump_json(indent=2), encoding="utf-8")

    with pytest.raises(TemplateLayoutError):
        assemble_earnings_update_deck(
            slide_plan_path=slide_plan_path,
            content_path=content_path,
            template_path=reordered,
            output_dir=tmp_path,
        )
