"""The falsification pass — figure extraction, the ledger join, and the report.

Asserted here is the **machine** half only, and deliberately: what a reviewer
concluded after reading a filing is judgement, exactly as the vision tier's
verdicts are (`test_deck_contract`'s "the vision tier is deliberately not asserted
for verdicts"). What can be pinned is that the agenda points at the right figures
and that nothing in it can ever block a run.

Both frozen fixture decks anchor the extraction, because the two failure modes are
opposite: the pitch deck's figures live in bullet text and market-entry tables, the
earnings-update deck's in a broker table, and the pitch deck carries three static
credential slides whose tombstone values are INFOR's own and must not be reported.
"""

from __future__ import annotations

from pathlib import Path

import pytest
from pptx import Presentation
from pptx.util import Inches

from deckcheck import (
    EXPECTED_ERROR_CONTEXTS,
    SEVERITY_ADVISORY,
    VERDICTS,
    CheckFinding,
    audit_deck,
    deck_pictures,
    extract_deck_figures,
    render_agenda,
    render_report,
    write_report,
)
from provenance import FigureSource, ProvenanceLedger

_FIXTURES = Path(__file__).parent / "fixtures"
_PITCH = _FIXTURES / "pitch-deck.pptx"
_EARNINGS = _FIXTURES / "earnings-update-deck.pptx"

_SRC = FigureSource(filing="FY2025 10-K", statement="Consolidated Statements of Operations")


def _deck_with(tmp_path: Path, text: str, *, name: str = "check.pptx") -> Path:
    """A one-slide deck whose single textbox holds `text`."""
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank
    box = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(8), Inches(2))
    box.name = "Figures"
    box.text_frame.text = text
    path = tmp_path / name
    prs.save(path)
    return path


# ─── Extraction ──────────────────────────────────────────────────────────────


def test_currency_figures_are_normalised_to_millions(tmp_path: Path):
    deck = _deck_with(tmp_path, "Revenue of US$589.8MM, EBITDA of $1.2B, cash of $450K")
    by_raw = {f.raw: f for f in extract_deck_figures(deck, library=None)}
    assert by_raw["US$589.8MM"].millions == pytest.approx(589.8)
    assert by_raw["$1.2B"].millions == pytest.approx(1200.0)
    assert by_raw["$450K"].millions == pytest.approx(0.45)


def test_percent_and_multiple_figures_keep_their_own_scale(tmp_path: Path):
    deck = _deck_with(tmp_path, "Margins of 28.2% at 11.5x EBITDA")
    kinds = {f.raw: (f.kind, f.value, f.millions) for f in extract_deck_figures(deck, library=None)}
    assert kinds["28.2%"] == ("percent", pytest.approx(28.2), None)
    assert kinds["11.5x"] == ("multiple", pytest.approx(11.5), None)


def test_an_accounting_negative_is_read_as_negative(tmp_path: Path):
    deck = _deck_with(tmp_path, "Net loss of ($12.3MM) for the period")
    figure = next(f for f in extract_deck_figures(deck, library=None) if "12.3" in f.raw)
    assert figure.value == pytest.approx(-12.3)
    assert figure.millions == pytest.approx(-12.3)


@pytest.mark.parametrize(
    "text",
    [
        "Q3 2026 results per the 2024 Annual Report",  # quarters, years, doc names
        "Founded 1987, page 3 of 16",                  # bare counts
        "See the 10-K and the 10-Q",                   # filing names
        "Presented April 2026",                        # a date
    ],
)
def test_bare_integers_are_not_figures(tmp_path: Path, text: str):
    # This exclusion is what makes the agenda readable — a deck is full of bare
    # integers and not one of them is a claim about the target's financials.
    assert extract_deck_figures(_deck_with(tmp_path, text), library=None) == []


def test_a_figure_inside_a_table_cell_is_found():
    # The earnings-update deck's broker table is the case: every figure it carries
    # is in a table, which a text-frame-only scan would miss entirely.
    figures = extract_deck_figures(_EARNINGS)
    tabled = [f for f in figures if f.shape.startswith("Table")]
    assert tabled, "no table-cell figures found in the earnings-update fixture"
    assert all(f.slide == 2 for f in tabled), "the broker table is on slide 3 (zero-based 2)"


def test_the_blank_library_excludes_the_static_credential_slides():
    # The pitch deck's slides 3-5 are copied from the library verbatim: their
    # tombstone values are INFOR's own and have no provenance in this run. They
    # drop out because the same shape in the blank library carries the same figure
    # — no slide list to migrate when the library gains an entry.
    with_baseline = extract_deck_figures(_PITCH)
    without = extract_deck_figures(_PITCH, library=_FIXTURES / "no-such-library.pptx")

    dropped = {(f.slide, f.shape, f.raw) for f in without} - {
        (f.slide, f.shape, f.raw) for f in with_baseline
    }
    assert dropped, "the library baseline excluded nothing"
    assert {slide for slide, _, _ in dropped} == {2, 3, 4}, "only the credential slides"
    assert {f.slide for f in with_baseline}.isdisjoint({2, 3, 4})


def test_filled_slides_survive_the_library_baseline():
    # The library holds `[x]` tokens where the fill puts figures, so a filled
    # figure can never match the baseline.
    figures = extract_deck_figures(_PITCH)
    assert any(f.slide == 1 for f in figures), "the executive summary lost its figures"
    assert len(figures) > 20


def test_pictures_are_listed_rather_than_scanned():
    # A rasterised range is out of scope for a string scan on purpose (an error
    # value inside one is usually correct), so it becomes a reader's job instead.
    pictures = deck_pictures(_EARNINGS)
    assert pictures
    assert all(isinstance(index, int) and name for index, name in pictures)


def test_the_library_baseline_also_drops_decorative_pictures():
    # 49 pictures on the pitch deck, 44 of them the library's own logos and
    # graphics. What is left is what this run pasted: the cap-table range on the
    # overview slide, the Financial Summary charts, the ownership blocks. A list of
    # 5 gets read; a list of 49 does not.
    kept = deck_pictures(_PITCH)
    everything = deck_pictures(_PITCH, library=_FIXTURES / "no-such-library.pptx")
    assert len(everything) > 40
    assert 0 < len(kept) <= 8, kept
    # Zero-based: the overview (6), the Financial Summary slide (7), ownership (8).
    assert {index for index, _ in kept} == {6, 7, 8}


def test_a_figures_context_is_a_window_not_the_whole_shape(tmp_path: Path):
    # A bullet block holds a dozen figures; repeating its full text once per figure
    # made the agenda table unreadable, which is the same failure as not reporting.
    long_text = "Lead-in. " * 40 + "Revenue of US$589.8MM. " + "Trailing. " * 40
    figure = next(f for f in extract_deck_figures(_deck_with(tmp_path, long_text), library=None)
                  if "589.8" in f.raw)
    assert "US$589.8MM" in figure.context
    assert len(figure.context) < len(long_text) / 3
    assert figure.context.startswith("…") and figure.context.endswith("…")


# ─── The ledger join ─────────────────────────────────────────────────────────


def _ledger(*figures) -> ProvenanceLedger:
    ledger = ProvenanceLedger(stage="financial-summary")
    for name, value, units in figures:
        ledger.record(name, sources=_SRC, value=value, units=units, location="financial-summary!B6")
    return ledger


def test_a_figure_matching_a_record_is_traced(tmp_path: Path):
    deck = _deck_with(tmp_path, "Revenue of US$589.8MM")
    audit = audit_deck(deck, _ledger(("Revenue FY2025", 589.8, "US$MM")), library=None)
    assert len(audit.traced) == 1
    assert audit.traced[0].record.figure == "Revenue FY2025"
    assert audit.untraced == []


def test_a_rounded_rendering_still_matches_its_record(tmp_path: Path):
    # The deck rounds; the workbook does not. Tolerance is half the last decimal
    # place the deck actually shows.
    deck = _deck_with(tmp_path, "Revenue of US$589.8MM")
    assert audit_deck(deck, _ledger(("Revenue", 589.83, "US$MM")), library=None).traced
    assert not audit_deck(deck, _ledger(("Revenue", 591.2, "US$MM")), library=None).traced


def test_a_scale_change_between_workbook_and_deck_still_matches(tmp_path: Path):
    # The workbook is locked to millions; a deck may render the same figure in
    # billions.
    deck = _deck_with(tmp_path, "Revenue of $1.2B")
    assert audit_deck(deck, _ledger(("Revenue", 1200.0, "US$MM")), library=None).traced


def test_a_figure_with_no_record_is_untraced(tmp_path: Path):
    deck = _deck_with(tmp_path, "Total addressable market of $4.2B")
    audit = audit_deck(deck, _ledger(("Revenue FY2025", 589.8, "US$MM")), library=None)
    assert len(audit.untraced) == 1
    assert audit.untraced[0].record is None


def test_a_percent_only_matches_a_percent_record(tmp_path: Path):
    deck = _deck_with(tmp_path, "Margin of 28.2%")
    assert audit_deck(deck, _ledger(("EBITDA margin", 28.2, "%")), library=None).traced
    # Same number, but the record is a dollar figure — not the same claim.
    assert not audit_deck(deck, _ledger(("Revenue", 28.2, "US$MM")), library=None).traced


def test_a_percent_record_stored_as_a_fraction_still_matches(tmp_path: Path):
    deck = _deck_with(tmp_path, "Margin of 28.2%")
    assert audit_deck(deck, _ledger(("EBITDA margin", 0.282, "%")), library=None).traced


def test_a_record_holding_an_excel_formula_cannot_be_auto_matched(tmp_path: Path):
    # A combined figure is written as a formula so the arithmetic stays in the
    # cell; openpyxl never evaluates it, so there is no number to compare. The
    # figure lands in the untraced list for a reviewer, not silently in traced.
    deck = _deck_with(tmp_path, "Combined balances of US$9,800.0MM")
    audit = audit_deck(deck, _ledger(("Combined balances", "=9000+800", "US$MM")), library=None)
    assert audit.untraced and not audit.traced


def test_audit_needs_no_renderer():
    # Deliberate: the join is deterministic and must run on a box with no
    # LibreOffice, so a missing renderer degrades the evidence, not the audit.
    audit = audit_deck(_PITCH, ProvenanceLedger())
    assert audit.matches and audit.untraced == audit.matches


# ─── Findings are advisory, by construction ──────────────────────────────────


def test_a_finding_cannot_claim_to_be_blocking():
    # `deckcheck` is a review surfaced at a checkpoint. The run's one required gate
    # is on `deck`, and a falsification pass that could halt a run would have to be
    # right about a target's financial statements.
    with pytest.raises(ValueError, match="review surfaced at a checkpoint"):
        CheckFinding(slide=1, figure="Revenue", verdict="contradicted", detail="x",
                     severity="blocking")


def test_a_finding_defaults_to_advisory():
    assert CheckFinding(slide=1, figure="Revenue", verdict="confirmed",
                        detail="x").severity == SEVERITY_ADVISORY


def test_an_unknown_verdict_is_rejected():
    with pytest.raises(ValueError, match="unknown verdict"):
        CheckFinding(slide=1, figure="Revenue", verdict="probably fine", detail="x")


def test_slide_number_is_one_based_for_the_analyst():
    assert CheckFinding(slide=6, figure="Revenue", verdict="confirmed", detail="x").slide_number == 7


# ─── Agenda and report ───────────────────────────────────────────────────────


def test_the_agenda_carries_the_do_not_report_list(tmp_path: Path):
    # The constraint that keeps this review from being ignored: expected CapIQ
    # error values must be in front of the reader at the moment they would
    # otherwise report one, not only in the skill's prose.
    deck = _deck_with(tmp_path, "Revenue of US$589.8MM")
    agenda = render_agenda(audit_deck(deck, _ledger(("Revenue", 589.8, "US$MM")), library=None))
    assert "Not defects" in agenda
    for context in EXPECTED_ERROR_CONTEXTS:
        assert context in agenda


def test_the_agenda_names_untraced_figures_and_their_slides(tmp_path: Path):
    deck = _deck_with(tmp_path, "Market size of $4.2B")
    agenda = render_agenda(audit_deck(deck, ProvenanceLedger(), library=None))
    assert "### Untraced figures — 1" in agenda
    assert "$4.2B" in agenda


def test_the_agenda_shows_a_traced_figure_with_its_citation(tmp_path: Path):
    deck = _deck_with(tmp_path, "Revenue of US$589.8MM")
    agenda = render_agenda(audit_deck(deck, _ledger(("Revenue FY2025", 589.8, "US$MM")),
                                     library=None))
    assert "FY2025 10-K, Consolidated Statements of Operations" in agenda
    assert "financial-summary!B6" in agenda


def test_the_report_leads_with_the_verdicts_then_the_agenda(tmp_path: Path):
    deck = _deck_with(tmp_path, "Revenue of US$589.8MM and a market size of $4.2B")
    audit = audit_deck(deck, _ledger(("Revenue FY2025", 589.8, "US$MM")), library=None)
    findings = [
        CheckFinding(slide=0, figure="Revenue FY2025", verdict="confirmed",
                     detail="agrees with the cited statement", source="FY2025 10-K p. 61"),
        CheckFinding(slide=0, figure="market size $4.2B", verdict="unsupported",
                     detail="no record and not in the attached CIM"),
    ]
    report = render_report(audit, findings, company="Example Target Inc.",
                          provenance_path=tmp_path / "provenance.json")

    assert report.index("## Findings") < report.index("## Agenda")
    assert "Advisory review, not a gate" in report
    assert "**confirmed**" in report and "**unsupported**" in report
    assert all(v in report for v in VERDICTS), "the verdict tally names every verdict"


def test_the_report_says_so_when_nothing_could_be_faulted(tmp_path: Path):
    deck = _deck_with(tmp_path, "Revenue of US$589.8MM")
    audit = audit_deck(deck, _ledger(("Revenue", 589.8, "US$MM")), library=None)
    assert "no figure could be faulted" in render_report(audit, company="Example Target Inc.")


def test_write_report_creates_its_parent_directory(tmp_path: Path):
    deck = _deck_with(tmp_path, "Revenue of US$589.8MM")
    audit = audit_deck(deck, ProvenanceLedger(), library=None)
    path = write_report(tmp_path / "artefacts" / "deckcheck-Project Test.md", audit,
                        company="Example Target Inc.")
    assert path.is_file()
    assert path.read_text(encoding="utf-8").startswith("# Deck check — Example Target Inc.")
