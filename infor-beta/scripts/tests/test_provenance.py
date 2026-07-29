"""The provenance record — validation, rendering, and the per-run merge (Phase G).

Two properties carry the phase, and each has its own layer here:

1. **The record renders the citation, not the other way round.** Both established
   citation forms — a filing/statement/page chain and a URL + retrieval date — come
   out of one `render()`, byte-identical to what v0.5.31 / v0.5.34 wrote. That is
   what makes the promotion safe: the artefact an analyst opens does not change.
2. **A half-citation cannot be constructed.** A statement with no filing, a URL
   with no retrieval date, a figure with neither a source nor a derivation — each
   reads as provenance and cannot be followed, which is worse than none. They raise.

The merge tests cover the reason fragments exist at all: two stages in one wave run
concurrently, so each writes its own file and the run record is their merge.
"""

from __future__ import annotations

import datetime
import json
from pathlib import Path

import pytest

from provenance import (
    PROVENANCE_FILENAME,
    DeckPlacement,
    FigureProvenance,
    FigureRef,
    FigureSource,
    ProvenanceError,
    ProvenanceLedger,
    read_run_provenance,
    stage_provenance_path,
    write_run_provenance,
)


# ─── Rendering: the one place a citation's wording is decided ────────────────


def test_filing_chain_renders_as_the_v0_5_34_string():
    assert FigureSource(
        filing="FY2025 10-K", statement="Consolidated Statements of Operations"
    ).render() == "FY2025 10-K, Consolidated Statements of Operations"


def test_url_form_renders_as_the_v0_5_31_string():
    assert FigureSource(
        url="https://example.com/fx", retrieved="2026-07-15"
    ).render() == "https://example.com/fx — retrieved 2026-07-15"


def test_page_joins_the_document_chain():
    assert FigureSource(
        filing="FY2025 10-K", statement="Note 12: Long-Term Debt", page=87
    ).render() == "FY2025 10-K, Note 12: Long-Term Debt, p. 87"


def test_a_non_numeric_page_is_kept_verbatim():
    assert FigureSource(filing="FY2025 10-K", page="F-12").render() == "FY2025 10-K, p. F-12"


def test_a_document_and_a_web_citation_render_as_ordered_segments():
    assert FigureSource(
        filing="FY2025 10-K",
        statement="Note 12",
        url="https://example.com/ir",
        retrieved="2026-07-15",
    ).render() == "FY2025 10-K, Note 12 — https://example.com/ir — retrieved 2026-07-15"


def test_a_date_object_is_normalised_to_iso():
    source = FigureSource(url="https://example.com/fx", retrieved=datetime.date(2026, 7, 15))
    assert source.retrieved == "2026-07-15"
    assert source.render().endswith("retrieved 2026-07-15")


# ─── Validation: a half-citation cannot exist ────────────────────────────────


@pytest.mark.parametrize(
    "kwargs, because",
    [
        ({}, "names neither a filing nor a url"),
        ({"statement": "Consolidated Statements of Operations"}, "statement with no filing"),
        ({"page": 42}, "page with no filing"),
        ({"url": "https://example.com/fx"}, "url with no retrieval date"),
        ({"filing": "   "}, "whitespace is not a filing"),
    ],
)
def test_incomplete_sources_raise(kwargs, because):
    with pytest.raises(ProvenanceError):
        FigureSource(**kwargs)


def test_a_figure_needs_a_source_or_a_derivation():
    # "sources are REQUIRED" was a SKILL.md sentence; it is structural now.
    with pytest.raises(ProvenanceError, match="no source and no derivation"):
        FigureProvenance(figure="Revenue FY2025", value=4520.0)


def test_a_derived_figure_needs_no_source():
    entry = FigureProvenance(
        figure="LTM Revenue", value=6062.0, derivation="FY + YTD − prior YTD"
    )
    assert entry.sources == ()
    assert entry.citation_lines == ()


def test_a_blank_figure_name_raises():
    with pytest.raises(ProvenanceError, match="figure name"):
        FigureProvenance(figure="  ", derivation="x")


def test_recording_a_citation_string_as_a_source_raises():
    ledger = ProvenanceLedger(stage="ltm-metrics")
    with pytest.raises(ProvenanceError, match="FigureSource"):
        ledger.record("Revenue FY2025", sources=["FY2025 10-K, income statement"])


# ─── Serialisation ───────────────────────────────────────────────────────────


def test_to_dict_omits_the_fields_a_source_does_not_carry():
    assert FigureSource(filing="FY2025 10-K").to_dict() == {"filing": "FY2025 10-K"}


def test_a_ledger_round_trips_through_json():
    ledger = ProvenanceLedger(stage="financial-summary")
    ledger.record(
        "Revenue FY2025",
        sources=FigureSource(filing="FY2025 10-K", statement="Income statement", page=61),
        value=4520.0,
        units="US$MM",
        location="financial-summary!F6",
    )
    ledger.record("Revenue LTM", value=6062.0, location="financial-summary!G6",
                  derivation="link to the ltm-metrics tab")

    reloaded = ProvenanceLedger.from_dict(json.loads(ledger.to_json()))
    assert reloaded.stage == "financial-summary"
    assert [f.to_dict() for f in reloaded.figures] == [f.to_dict() for f in ledger.figures]
    assert reloaded.figures[0].sources[0].page == "61"


def test_write_puts_the_fragment_in_the_stage_directory(tmp_path: Path):
    ledger = ProvenanceLedger(stage="ltm-metrics")
    ledger.record("LTM Revenue", derivation="FY + YTD − prior YTD", value=6062.0)
    path = ledger.write(tmp_path)
    assert path == stage_provenance_path(tmp_path) == tmp_path / PROVENANCE_FILENAME
    assert ProvenanceLedger.read(tmp_path).stage == "ltm-metrics"


# ─── The per-run merge ───────────────────────────────────────────────────────


def _fragment(run_dir: Path, stage: str, *figures: str, label_stage: bool = True) -> None:
    ledger = ProvenanceLedger(stage=stage if label_stage else None)
    for name in figures:
        ledger.record(name, sources=FigureSource(filing=f"{name} filing"), value=1.0)
    ledger.write(run_dir / "stages" / stage)


def test_read_run_provenance_merges_every_stage_fragment(tmp_path: Path):
    _fragment(tmp_path, "financial-summary", "Revenue FY2025")
    _fragment(tmp_path, "ltm-metrics", "LTM Revenue", "LTM Adj. EBITDA")
    _fragment(tmp_path, "captable", "Share price")

    merged = read_run_provenance(tmp_path)
    assert len(merged) == 4
    # Stage-id order, so the merged record reads like the run did.
    assert merged.stages == ("captable", "financial-summary", "ltm-metrics")
    assert {f.figure for f in merged.figures} == {
        "Revenue FY2025", "LTM Revenue", "LTM Adj. EBITDA", "Share price"
    }


def test_a_fragment_that_forgot_to_label_itself_is_stamped_from_its_directory(tmp_path: Path):
    _fragment(tmp_path, "ltm-metrics", "LTM Revenue", label_stage=False)
    merged = read_run_provenance(tmp_path)
    assert [f.stage for f in merged.figures] == ["ltm-metrics"]


def test_a_stage_with_no_fragment_is_not_an_error(tmp_path: Path):
    # The wireframe and the assembler extract no figures — there is nothing to
    # record, and that must not fail the merge.
    (tmp_path / "stages" / "wireframe").mkdir(parents=True)
    _fragment(tmp_path, "ltm-metrics", "LTM Revenue")
    assert len(read_run_provenance(tmp_path)) == 1


def test_a_run_with_no_stages_directory_merges_to_nothing(tmp_path: Path):
    assert len(read_run_provenance(tmp_path)) == 0


def test_write_run_provenance_is_idempotent(tmp_path: Path):
    _fragment(tmp_path, "ltm-metrics", "LTM Revenue")
    first = write_run_provenance(tmp_path).read_text(encoding="utf-8")
    second = write_run_provenance(tmp_path).read_text(encoding="utf-8")
    assert first == second
    assert (tmp_path / PROVENANCE_FILENAME).is_file()


def test_a_fragment_lands_in_its_own_stage_directory_and_nowhere_else(tmp_path: Path):
    # The shape of the whole scheme: `ledger.write(io.stage_dir)` writes one file,
    # inside that stage's directory, and touches nothing shared.
    ledger = ProvenanceLedger(stage="comps")
    ledger.record("Comps peer — NYSE:AAAA", sources=FigureSource(url="https://example.com/ir",
                                                                retrieved="2026-07-29"))
    stage_dir = tmp_path / "stages" / "comps"
    path = ledger.write(stage_dir)

    assert path == stage_dir / PROVENANCE_FILENAME
    assert [p.relative_to(tmp_path).as_posix() for p in tmp_path.rglob("*.json")] == [
        "stages/comps/provenance.json"
    ], "a stage wrote something outside its own directory"


def test_two_stages_writing_at_the_same_time_both_survive_the_merge(tmp_path: Path):
    # Why fragments exist at all: wave-mates run CONCURRENTLY as separate
    # sub-agents. A shared provenance.json would be a read-modify-write race — the
    # second writer would land on a file the first had already replaced, and one
    # stage's figures would simply be gone with nothing failing.
    import threading

    stages = {
        "comps": [f"Comps peer — NYSE:AAA{i}" for i in range(40)],
        "precedents": [f"Example Target {i} Inc. / Example Buyer Inc. — tev" for i in range(40)],
        "ownership": [f"Insider holding — Insider {i}" for i in range(40)],
    }
    start = threading.Barrier(len(stages))

    def write(stage: str, figures: list[str]) -> None:
        ledger = ProvenanceLedger(stage=stage)
        for name in figures:
            ledger.record(name, sources=FigureSource(filing=f"{stage} source"), value=1.0)
        start.wait(timeout=10)  # every thread writes in the same instant
        ledger.write(tmp_path / "stages" / stage)

    threads = [threading.Thread(target=write, args=(s, f)) for s, f in stages.items()]
    for thread in threads:
        thread.start()
    for thread in threads:
        thread.join(timeout=30)
    assert not any(thread.is_alive() for thread in threads)

    merged = read_run_provenance(tmp_path)
    assert len(merged) == sum(len(f) for f in stages.values())
    assert merged.stages == ("comps", "ownership", "precedents")
    for stage, figures in stages.items():
        assert {f.figure for f in merged.figures if f.stage == stage} == set(figures)

    # And the consolidated file is that same merge, written once by `deckcheck`.
    consolidated = ProvenanceLedger.read(write_run_provenance(tmp_path))
    assert len(consolidated) == len(merged)


# ─── A derivation a machine can follow ───────────────────────────────────────


def _bridge_ledger() -> ProvenanceLedger:
    """An LTM bridge: three sourced components and a total derived from them."""
    ledger = ProvenanceLedger(stage="ltm-metrics")
    for name, value, filing, page in (
        ("LTM Revenue — FY2025 revenue", 5168.4, "FY2025 10-K", 142),
        ("LTM Revenue — Q1 2026 YTD", 166.1, "Q1 2026 10-Q", 4),
        ("LTM Revenue — Q1 2025 YTD", 138.4, "Q1 2026 10-Q", 4),
    ):
        ledger.record(
            name,
            sources=FigureSource(filing=filing, statement="Consolidated Statements of Operations",
                                 page=page),
            value=value,
            units="US$MM",
            location=f"ltm-metrics!B{20 + len(ledger)}",
        )
    ledger.record(
        "LTM Revenue",
        value=5196.1,
        units="US$MM",
        location="ltm-metrics!B23",
        derivation="FY2025 revenue + Q1 2026 YTD − Q1 2025 YTD",
        derived_from=[FigureRef(location=f.location, figure=f.figure) for f in ledger.figures],
    )
    return ledger


def test_a_structured_derivation_resolves_to_its_upstream_records():
    ledger = _bridge_ledger()
    trace = ledger.trace(ledger.figures[-1])

    assert trace.structured and trace.resolved
    assert [c.figure for c in trace.components] == [f.figure for f in ledger.figures[:3]]
    # The point of resolving it: the total reaches the filing PAGES underneath.
    assert {s.render() for s in trace.root_sources} == {
        "FY2025 10-K, Consolidated Statements of Operations, p. 142",
        "Q1 2026 10-Q, Consolidated Statements of Operations, p. 4",
    }


def test_an_unresolvable_component_is_reported_as_unresolvable():
    # A stage claiming a figure was built from something with no record is exactly
    # the gap prose could not express: "derived from an upstream record" and
    # "unsourced" used to read identically.
    ledger = ProvenanceLedger(stage="ltm-metrics")
    entry = ledger.record(
        "LTM Adj. EBITDA",
        value=1840.4,
        derivation="FY2025 + Q1 2026 YTD − Q1 2025 YTD",
        derived_from=["FY2025 Adj. EBITDA", FigureRef(location="ltm-metrics!B99")],
    )
    trace = ledger.trace(entry)

    assert trace.structured and not trace.resolved
    assert trace.components == () and trace.root_sources == ()
    assert [r.render() for r in trace.unresolved] == ["FY2025 Adj. EBITDA", "ltm-metrics!B99"]
    assert "UNRESOLVABLE: FY2025 Adj. EBITDA, ltm-metrics!B99" in trace.render()


def test_a_prose_only_derivation_is_reported_as_unstructured():
    # The eleven records a real run left like this: legal, and honestly labelled as
    # not followable rather than printed beside the resolved ones.
    ledger = ProvenanceLedger(stage="financial-summary")
    entry = ledger.record("Revenue LTM", value=6062.0,
                          derivation="ltm-metrics stage handoff, converted to CAD by *F7")
    trace = ledger.trace(entry)
    assert not trace.structured and not trace.resolved
    assert trace.render() == "ltm-metrics stage handoff, converted to CAD by *F7"


def test_a_derivation_ref_resolves_across_stage_fragments(tmp_path: Path):
    # The `financial-summary` LTM cell links a bridge total in ANOTHER stage's
    # fragment, so the ref names no stage and resolves only in the merge — which is
    # what the merge is for.
    _bridge_ledger().write(tmp_path / "stages" / "ltm-metrics")
    fs = ProvenanceLedger(stage="financial-summary")
    fs.record("Revenue LTM", value=5196.1, units="US$MM", location="financial-summary!G6",
              derivation="link to the ltm-metrics tab's '(=) LTM Revenue' bridge total",
              derived_from=[FigureRef(figure="LTM Revenue")])
    fs.write(tmp_path / "stages" / "financial-summary")

    fragment = ProvenanceLedger.read(tmp_path / "stages" / "financial-summary")
    alone = fragment.trace(fragment.figures[0])
    assert not alone.resolved, "on its own a fragment cannot see the other stage's record"

    merged = read_run_provenance(tmp_path)
    linked = next(f for f in merged.figures if f.figure == "Revenue LTM")
    trace = merged.trace(linked)
    assert trace.resolved
    assert [c.figure for c in trace.components] == ["LTM Revenue"]
    assert any("FY2025 10-K" in s.render() for s in trace.root_sources)


def test_a_circular_derivation_does_not_hang():
    # Nothing generates one deliberately; an analyst editing a workbook can.
    ledger = ProvenanceLedger(stage="captable")
    ledger.record("A", value="=B1", location="captable!A1",
                  derived_from=[FigureRef(location="captable!B1")])
    ledger.record("B", value="=A1", location="captable!B1",
                  derived_from=[FigureRef(location="captable!A1")])
    trace = ledger.trace(ledger.figures[0])
    assert [c.figure for c in trace.components] == ["B"]
    assert trace.unresolved == ()


def test_a_record_may_be_derived_with_refs_and_no_prose():
    entry = FigureProvenance(figure="Enterprise Value", derived_from=[FigureRef(location="captable!F28")])
    assert entry.derived
    assert entry.derivation_line == "derived from captable!F28"


def test_a_figure_with_neither_a_source_nor_any_derivation_still_raises():
    with pytest.raises(ProvenanceError, match="no source and no derivation"):
        FigureProvenance(figure="Enterprise Value", value=15796.0)


# ─── Placement: where a figure lands ─────────────────────────────────────────


def test_a_placement_needs_to_name_somewhere():
    with pytest.raises(ProvenanceError, match="must name a slide"):
        DeckPlacement()


def test_a_placement_slide_is_one_based():
    with pytest.raises(ProvenanceError, match="1-based"):
        DeckPlacement(slide=0, field="executive_summary_bullets[0]")


def test_a_ref_needs_a_figure_or_a_location():
    with pytest.raises(ProvenanceError, match="figure name or a location"):
        FigureRef()


def test_a_location_ref_ignores_absolute_markers_and_sheet_quoting():
    ledger = ProvenanceLedger(stage="captable")
    ledger.record("Basic Shares Outstanding", value="=F186", location="captable!F17",
                  derivation="cap-table formula =F186")
    assert ledger.find(FigureRef(location="captable!$F$17")).figure == "Basic Shares Outstanding"
    assert ledger.find(FigureRef(location="'captable'!f17")) is not None
    assert ledger.find(FigureRef(location="captable!F18")) is None


def test_placement_and_refs_round_trip_through_json(tmp_path: Path):
    ledger = ProvenanceLedger(stage="content")
    ledger.record(
        "ARR",
        sources=FigureSource(filing="Q1 2026 10-Q", statement="MD&A — key metrics", page=7),
        value=4190.5,
        units="US$MM",
        placement=DeckPlacement(slide=2, field="executive_summary_bullets[1]"),
    )
    ledger.record("ARR as % of revenue", value=81.0, units="%",
                  derivation="ARR ÷ FY2025 revenue",
                  derived_from=[FigureRef(figure="ARR", stage="content")],
                  placement=DeckPlacement(slide=2, field="executive_summary_bullets[1]"))

    reloaded = ProvenanceLedger.read(ledger.write(tmp_path))
    assert [f.to_dict() for f in reloaded.figures] == [f.to_dict() for f in ledger.figures]
    assert reloaded.figures[0].placement == DeckPlacement(slide=2,
                                                         field="executive_summary_bullets[1]")
    assert reloaded.trace(reloaded.figures[1]).resolved
