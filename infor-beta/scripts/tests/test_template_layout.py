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


# ─── (a0) the shipped templates carry every declared defined name ────────────
# The Phase C contract in both directions: each `infor_` name exists on the
# sheet the registry says, resolves to the address the registry says, and agrees
# with the sentinel that independently pins the same region.


@pytest.mark.parametrize("template", sorted(tl.TEMPLATE_NAMED_RANGES))
def test_shipped_template_carries_every_declared_name(template: str):
    wb = load_workbook(TEMPLATES / template)
    try:
        for sheet, targets in tl.TEMPLATE_NAMED_RANGES[template].items():
            assert sheet in wb.sheetnames, f"{template} has no sheet {sheet!r}"
            ws = wb[sheet]
            for name, target in sorted(targets.items()):
                assert name.startswith("infor_"), f"{name} is missing the collision prefix"
                assert tl.defined_name_ref(ws, name) == tl.normalize_ref(target), (
                    f"{template} {sheet}!{name}"
                )
    finally:
        wb.close()


def test_the_registry_covers_the_names_phase_c_promised():
    # The plan named eight cells explicitly; this pins that each is reachable.
    # (`precedents_input_ccy` ships as `infor_prec_output_ccy` — the cell is
    # labelled "Output:" and the aggregator relinks it to the cap table's
    # OUTPUT currency, so the name follows the artefact.)
    cap = tl.TEMPLATE_NAMED_RANGES[tl.CAP_TABLE_TEMPLATE][tl.CAP_TABLE_SHEET]
    assert cap[tl.NAME_FX_RATE] == "F7"
    assert cap[tl.NAME_SHARE_PRICE] == "F16"
    assert cap[tl.NAME_LTM_REVENUE_VALUATION] == "D47"
    assert cap[tl.NAME_LTM_EBITDA_VALUATION] == "D48"
    assert cap[tl.NAME_CAP_PICTURE_RANGE] == "B15:F40"
    assert cap[tl.NAME_BASIC_SHARES] == "F17"
    assert tl.TEMPLATE_NAMED_RANGES[tl.COMPS_TEMPLATE][tl.COMPS_SHEET][
        tl.NAME_COMPS_OUTPUT_CCY
    ] == "F3"
    assert tl.TEMPLATE_NAMED_RANGES[tl.PRECEDENTS_TEMPLATE][tl.PRECEDENTS_SHEET][
        tl.NAME_PREC_OUTPUT_CCY
    ] == "C2"


def test_new_names_did_not_disturb_the_capiq_and_legacy_namespaces():
    # The reconnaissance rule: the cap table's Capital IQ names identify the
    # workbook to the add-in, and the comps template carries 1,245 legacy
    # artefacts. Phase C adds alongside them; it must never tidy them away.
    cap = load_workbook(CAP_TABLE)
    assert {"CIQWBGuid", "CIQWBInfo", "IQ_LTM", "CAD_USD"} <= set(cap.defined_names)
    assert len(cap.defined_names) == 33  # unchanged; the infor_ names are sheet-scoped
    comps = load_workbook(COMPS)
    assert len(comps.defined_names) == 1245
    assert not [n for n in comps.defined_names if n.startswith("infor_")]


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


def test_every_library_marker_identifies_exactly_one_slide():
    # The Phase C precondition. `find_slide_by_marker` replaced the assemblers'
    # hardcoded indices, and it is only sound if each marker is unique in the
    # library: two matches and the assembler's choice would be arbitrary.
    prs = Presentation(str(LIBRARY))
    for index, marker in sorted(tl.LIBRARY_SLIDE_MARKERS.items()):
        assert tl.find_slides_by_marker(prs, marker) == [index], marker.description


def test_earnings_assembler_finds_its_five_slides_by_marker():
    # `_KEEP_LIBRARY_INDICES = (0, 6, 7, 15, 16)` is gone; the five entries are
    # discovered. This pins that they still resolve to the historical indices —
    # so the migration is faithful — without the assembler depending on them.
    from earnings_update_assembler import _KEEP_MARKERS

    prs = Presentation(str(LIBRARY))
    found = [tl.find_slide_by_marker(prs, marker) for marker in _KEEP_MARKERS]
    assert found == [0, 6, 7, 15, 16]


def test_pitch_assembler_clone_and_delete_targets_resolve_by_marker():
    prs = Presentation(str(LIBRARY))
    assert tl.find_slide_by_marker(prs, tl.MARKER_EARNINGS_SUMMARY) == 7
    assert tl.find_slide_by_marker(prs, tl.MARKER_FINANCIAL_SUMMARY) == 8
    assert tl.find_slide_by_marker(prs, tl.MARKER_MARKET_ENTRY) == 14


def test_built_overview_marker_finds_the_overview_slide_in_an_assembled_deck():
    # `OVERVIEW_SLIDE_INDEX = 6` is gone. In a BUILT deck the library's overview
    # marker no longer matches — its title has been filled with the client name
    # — so the pie placeholder is the marker, which is also the shape
    # `financial_charts` is about to replace.
    #
    # Measured on the earnings-update fixture, whose overview slide sits at
    # index 1: the point is that the marker finds it wherever it is. The pitch
    # fixture is deliberately NOT used — it is a finished deck, so
    # `financial-charts` has already swapped the placeholder for a picture,
    # which is the same "present until this stage replaces it" contract the FS
    # self-discovery has always had.
    deck = Path(__file__).parent / "fixtures" / "earnings-update-deck.pptx"
    prs = Presentation(str(deck))
    assert tl.find_slides_by_marker(prs, tl.MARKER_OVERVIEW) == []  # title was filled
    assert tl.find_slide_by_marker(prs, tl.MARKER_BUILT_OVERVIEW) == 1


def test_the_pitch_assembler_always_leaves_the_pie_placeholder_for_the_marker():
    # `financial_charts` locates the overview slide by the pie placeholder, so
    # the assembler must always leave one. It does: `_verify_pitch_output`
    # fails the stage if the placeholder went missing.
    import pitch_deck_assembler

    source = Path(pitch_deck_assembler.__file__).read_text(encoding="utf-8")
    assert tl.MARKER_BUILT_OVERVIEW.contains in source


def test_find_slide_by_marker_rejects_zero_and_multiple_matches():
    prs = Presentation(str(LIBRARY))
    absent = tl.SlideMarker("Title 1", "no such text anywhere", "nonexistent")
    with pytest.raises(tl.TemplateLayoutError, match="no slide carries"):
        tl.find_slide_by_marker(prs, absent)
    assert tl.find_optional_slide_by_marker(prs, absent) is None

    # 'Title 1' with an empty fragment matches every slide that has the shape.
    ambiguous = tl.SlideMarker("Title 1", "", "ambiguous")
    with pytest.raises(tl.TemplateLayoutError, match="must identify exactly one"):
        tl.find_slide_by_marker(prs, ambiguous)
    with pytest.raises(tl.TemplateLayoutError, match="must identify exactly one"):
        tl.find_optional_slide_by_marker(prs, ambiguous)


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


# ─── the name ↔ sentinel cross-check ──────────────────────────────────────────
# The check that earns the sentinel tables' deletion in a later release: a name
# and a sentinel that disagree about where a cell is means the Phase C migration
# mis-mapped it, and the writers must not proceed on either answer.


def _named_sheet(name: str, ref: str):
    """A one-sheet workbook with `B1 = 'Label'` and a sheet-scoped defined name."""
    from openpyxl.workbook.defined_name import DefinedName

    wb = Workbook()
    ws = wb.active
    ws.title = "S"
    ws["B1"] = "Label"
    ws.defined_names.add(DefinedName(name, attr_text=f"'S'!{ref}"))
    return ws


def test_cross_check_passes_when_the_name_agrees_with_its_sentinel():
    ws = _named_sheet("infor_thing", "$F$1")
    anchor = CellAnchor("F1", "B1", "Label", name="infor_thing")
    tl.verify_anchors(ws, (anchor,), template="T.xlsx", require_names=True)
    assert tl.resolve_name_cell(ws, "infor_thing") == "F1"


def test_cross_check_raises_when_the_name_and_sentinel_disagree():
    # The name points one column right of where the sentinel pins the cell.
    ws = _named_sheet("infor_thing", "$G$1")
    anchor = CellAnchor("F1", "B1", "Label", name="infor_thing")
    with pytest.raises(TemplateLayoutError) as exc:
        tl.verify_anchors(ws, (anchor,), template="T.xlsx")
    msg = str(exc.value)
    assert "infor_thing" in msg and "resolves to G1" in msg and "pins F1" in msg


def test_a_missing_name_is_an_error_only_when_required():
    # A built artefact from before Phase C carries no names; it must still pass
    # on its sentinels. A shipped template must not.
    wb = Workbook()
    ws = wb.active
    ws.title = "S"
    ws["B1"] = "Label"
    anchor = CellAnchor("F1", "B1", "Label", name="infor_thing")
    tl.verify_anchors(ws, (anchor,), template="T.xlsx")  # tolerated
    with pytest.raises(TemplateLayoutError, match="expects defined name 'infor_thing'"):
        tl.verify_anchors(ws, (anchor,), template="T.xlsx", require_names=True)


def test_resolve_helpers_fail_loudly_and_distinguish_cell_from_range():
    ws = _named_sheet("infor_block", "$B$2:$D$4")
    assert tl.resolve_name_range(ws, "infor_block") == "B2:D4"
    with pytest.raises(TemplateLayoutError, match="resolves to the range 'B2:D4'"):
        tl.resolve_name_cell(ws, "infor_block")
    with pytest.raises(TemplateLayoutError, match="defined name 'infor_absent' is missing"):
        tl.resolve_name_range(ws, "infor_absent")


def test_resolve_workbook_range_falls_back_for_a_pre_phase_c_artefact(tmp_path: Path):
    # An artefact an earlier release produced has no names; the assemblers must
    # still be able to paste its picture range rather than failing for a reason
    # that is not a layout shift.
    path = tmp_path / "old.xlsx"
    wb = Workbook()
    wb.active.title = tl.CAP_TABLE_SHEET
    wb.save(path)
    assert tl.resolve_workbook_range(
        path, sheet=tl.CAP_TABLE_SHEET, name=tl.NAME_CAP_PICTURE_RANGE, fallback="B15:F40"
    ) == "B15:F40"
    # ...and the shipped template resolves through the name instead.
    assert tl.resolve_workbook_range(
        CAP_TABLE, sheet=tl.CAP_TABLE_SHEET, name=tl.NAME_CAP_PICTURE_RANGE, fallback="ZZ1"
    ) == "B15:F40"


# ─── (b) a shifted Excel template raises through the wired-in writers ─────────


def _shifted_copy(template: Path, dest: Path, sheet: str, at_row: int) -> Path:
    """Copy a template and insert one row at ``at_row`` (shifting rows down)."""
    shutil.copyfile(template, dest)
    wb = load_workbook(dest)
    wb[sheet].insert_rows(at_row)
    wb.save(dest)
    return dest


def _shifted_deal_workbook(tmp_path: Path, tab: str, at_row: int) -> Path:
    """A deal workbook with one row inserted into `tab`, shifting its layout.

    The Phase D equivalent of `_shifted_copy`: the builders write tabs of the
    deal workbook now, so a layout shift has to be introduced there for the
    sentinel/name cross-check to catch it.
    """
    from deal_workbook import init_deal_workbook

    deal = init_deal_workbook(
        deal_dir=tmp_path, deliverable_type="pitch", deal_name="Project Test"
    )
    wb = load_workbook(deal)
    wb[tab].insert_rows(at_row)
    wb.save(deal)
    return deal


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

    from deal_workbook import TAB_OWNERSHIP

    shifted = _shifted_deal_workbook(tmp_path, TAB_OWNERSHIP, 30)
    with pytest.raises(TemplateLayoutError) as exc:
        build_ownership_workbook(
            insiders=[InsiderHolding("Doe, Jane", "Jane Doe (CFO)", 500)],
            total_shares_outstanding=1_000_000,
            deal_workbook=shifted,
        )
    msg = str(exc.value)
    assert tl.OWNERSHIP_TEMPLATE in msg
    assert "B38" in msg and "'SEDI Name'" in msg


def test_shifted_comps_template_fails_the_builder(tmp_path: Path):
    from comps_workbook import build_comps_workbook

    from deal_workbook import TAB_COMPS

    shifted = _shifted_deal_workbook(tmp_path, TAB_COMPS, 8)
    with pytest.raises(TemplateLayoutError, match="Group Average"):
        build_comps_workbook(
            verticals=[{"name": "Vertical A", "companies": [{"ticker": "TSX:RY"}]}],
            deal_workbook=shifted,
        )


def test_shifted_precedents_template_fails_the_builder(tmp_path: Path):
    from precedents_workbook import build_precedents_workbook

    from deal_workbook import TAB_PRECEDENTS

    shifted = _shifted_deal_workbook(tmp_path, TAB_PRECEDENTS, 3)
    with pytest.raises(TemplateLayoutError) as exc:
        build_precedents_workbook(
            deal_workbook=shifted,
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
        )
    assert "'Currency'" in str(exc.value)


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


def _earnings_inputs(tmp_path: Path):
    """Minimal real SlidePlan + content artefacts for an end-to-end assembly."""
    from earnings_update_wireframe import build_earnings_update_slide_plan, write_slide_plan
    from schemas import Company
    from tests.test_earnings_update_assembler import _sample_content

    slide_plan = build_earnings_update_slide_plan(
        company=Company(legal_name="SampleCo", ticker="TSX:SMPL"),
        reporting_quarter="Q4 2025",
        comparison_quarter="Q4 2024",
    )
    slide_plan_path = write_slide_plan(slide_plan, tmp_path / "slide_plan.json")
    content_path = tmp_path / "content.json"
    content_path.write_text(_sample_content().model_dump_json(indent=2), encoding="utf-8")
    return slide_plan_path, content_path


def test_reordered_library_still_assembles_the_earnings_deck(tmp_path: Path):
    # The Phase C payoff, end to end. Through v0.5.39 the earnings assembler
    # kept library indices (0, 6, 7, 15, 16) and this same re-order raised
    # TemplateLayoutError — correct, but it meant every library edit was a code
    # migration. The five entries are now located by marker, so moving one is a
    # non-event: the deck still assembles, with the right five slides.
    from earnings_update_assembler import assemble_earnings_update_deck

    reordered = _reordered_library_copy(tmp_path / "library.pptx", 6)  # overview -> end
    slide_plan_path, content_path = _earnings_inputs(tmp_path)

    out = assemble_earnings_update_deck(
        slide_plan_path=slide_plan_path,
        content_path=content_path,
        template_path=reordered,
        output_dir=tmp_path,
        converge=False,  # geometry is deck_repair's business, not this test's
    )
    prs = Presentation(str(out))
    assert len(prs.slides) == 5
    # The overview slide moved to the end of the library, so it is now the last
    # of the five kept entries rather than the second — and it is still filled.
    texts = ["\n".join(s.text for s in slide.shapes if s.has_text_frame) for slide in prs.slides]
    assert any("Introduction to SampleCo" in t for t in texts)
    assert any("Earnings Summary" in t for t in texts)
    assert any("Disclaimer" in t for t in texts)


def test_inserting_a_library_slide_needs_no_code_change(tmp_path: Path):
    """The Phase C payoff, on the edit that historically forced a migration.

    Inserting a slide ahead of the disclaimer/contact closers is precisely what
    v0.5.8 and v0.5.14 did, and each time the earnings assembler shipped the
    wrong slides until someone hand-bumped `_KEEP_LIBRARY_INDICES` — once
    dropping the Contact slide entirely from client decks. Here the closers move
    15/16 -> 16/17 and both assemblers simply follow.
    """
    from earnings_update_assembler import _KEEP_MARKERS, assemble_earnings_update_deck
    from pitch_deck_assembler import _PitchLayout
    from pptx_helpers import clone_slide_after, delete_slide

    shipped = Presentation(str(LIBRARY))
    before = {
        "disclaimer": tl.find_slide_by_marker(shipped, tl.MARKER_DISCLAIMER),
        "contact": tl.find_slide_by_marker(shipped, tl.MARKER_CONTACT),
        "precedents": tl.find_slide_by_marker(shipped, tl.MARKER_PRECEDENTS),
    }

    # Insert one slide just after Comparable Companies, and retitle it so it is a
    # distinct concept rather than a second comps slide (which would make the
    # comps marker ambiguous — itself a caught error, see the finder tests).
    grown = tmp_path / "library.pptx"
    prs = Presentation(str(LIBRARY))
    comps_at = tl.find_slide_by_marker(prs, tl.MARKER_COMPS)
    clone_slide_after(prs, comps_at)
    inserted = comps_at + 1
    for shape in prs.slides[inserted].shapes:
        if shape.name == tl.MARKER_COMPS.shape_name and shape.has_text_frame:
            shape.text_frame.paragraphs[0].runs[0].text = "Trading History Analysis"
    prs.save(str(grown))

    prs = Presentation(str(grown))
    assert len(prs.slides) == len(shipped.slides) + 1
    # Everything after the insertion shifted by one.
    assert tl.find_slide_by_marker(prs, tl.MARKER_DISCLAIMER) == before["disclaimer"] + 1
    assert tl.find_slide_by_marker(prs, tl.MARKER_CONTACT) == before["contact"] + 1
    assert tl.find_slide_by_marker(prs, tl.MARKER_PRECEDENTS) == before["precedents"] + 1

    # The earnings assembler follows the shift and still builds the right deck.
    found = [tl.find_slide_by_marker(prs, marker) for marker in _KEEP_MARKERS]
    assert found == [0, 6, 7, before["disclaimer"] + 1, before["contact"] + 1]

    slide_plan_path, content_path = _earnings_inputs(tmp_path)
    out = assemble_earnings_update_deck(
        slide_plan_path=slide_plan_path,
        content_path=content_path,
        template_path=grown,
        output_dir=tmp_path,
        converge=False,
    )
    built = Presentation(str(out))
    assert len(built.slides) == 5
    texts = ["\n".join(s.text for s in slide.shapes if s.has_text_frame) for slide in built.slides]
    assert any("Introduction to SampleCo" in t for t in texts)
    assert any("Disclaimer" in t for t in texts), "the closers must still be the last two"
    assert any("Contact" in t for t in texts)

    # The pitch layout discovers its indices over the grown library too.
    pitch = Presentation(str(grown))
    delete_slide(pitch, tl.find_slide_by_marker(pitch, tl.MARKER_EARNINGS_SUMMARY))
    layout = _PitchLayout(pitch, template_name=grown.name, expected_total=len(pitch.slides))
    assert layout.overview == 6
    assert layout.precedents == before["precedents"]  # 11 -> 12 raw, -1 for the earnings delete
    assert layout.investment_highlights == layout.precedents + 1
    assert layout.market_entry == [layout.precedents + 2]


def test_library_missing_a_kept_entry_fails_the_earnings_assembler(tmp_path: Path):
    # The other direction: discovery must still be a check. A library that lost
    # one of the five entries halts the run rather than shipping a short deck.
    from earnings_update_assembler import assemble_earnings_update_deck
    from pptx_helpers import delete_slide

    prs = Presentation(str(LIBRARY))
    delete_slide(prs, tl.find_slide_by_marker(prs, tl.MARKER_CONTACT))
    broken = tmp_path / "library.pptx"
    prs.save(str(broken))

    slide_plan_path, content_path = _earnings_inputs(tmp_path)
    with pytest.raises(TemplateLayoutError, match="contact"):
        assemble_earnings_update_deck(
            slide_plan_path=slide_plan_path,
            content_path=content_path,
            template_path=broken,
            output_dir=tmp_path,
            converge=False,
        )
