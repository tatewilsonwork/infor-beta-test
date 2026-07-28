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
    FigureProvenance,
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
