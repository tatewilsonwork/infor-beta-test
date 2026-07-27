"""Phase B step 3: the converge loop, wired against the same regression artefacts.

`test_deck_contract.py` asserts the oracle SEES the historical defects. This file
asserts the loop ACTS on them — and, just as importantly, that it fails loudly
when it cannot, since shipping a shrunk-but-still-overflowing deck is what the
retired estimator did for thirteen releases.

Every repair here is measured off a real LibreOffice render, so these tests need
LibreOffice. There is deliberately no estimate-based fallback to test against.
"""

from __future__ import annotations

import shutil
from pathlib import Path

import pytest
from pptx import Presentation
from pptx.util import Emu, Pt

from deck_contract import default_library_path, verify_deck
from deck_repair import (
    REPAIRABLE_KINDS,
    TEXT_SCALES,
    ConvergeResult,
    DeckNotConvergedError,
    assert_converged,
    converge_deck,
)
from excel_to_powerpoint import find_soffice
from pptx_helpers import normal_autofit_scale

_FIXTURES = Path(__file__).parent / "fixtures"
_REGRESSIONS = _FIXTURES / "regressions"

needs_render = pytest.mark.skipif(
    find_soffice() is None, reason="the repair loop measures on a LibreOffice render"
)


def _copy(src: Path, tmp_path: Path) -> Path:
    """Work on a copy — `converge_deck` repairs in place."""
    dst = tmp_path / src.name
    shutil.copy(src, dst)
    return dst


def _geometric(findings):
    return [f for f in findings if f.blocking and f.kind in REPAIRABLE_KINDS]


# ─── The loop repairs the historical defects ─────────────────────────────────


@needs_render
def test_repairs_prl17_market_entry_table(tmp_path):
    """PRL17's 5.91"-rendered table, fixed by measurement rather than estimation.

    The v0.5.23 fix for this was a per-character Palatino width table plus per-row
    content-height floors. Here the loop renders the table at each candidate body
    size and picks the largest that actually fits — and the table's declared height
    is untouched, because the declaration was never the problem.
    """
    deck = _copy(_REGRESSIONS / "prl17-market-entry-table.pptx", tmp_path)
    before = Emu(
        next(s for s in Presentation(deck).slides[0].shapes if s.name == "Table 1215").height
    ).inches

    result = converge_deck(deck, out_dir=tmp_path / "qa")

    assert result.converged, result.summary() + "\n" + "\n".join(
        str(f) for f in result.unrepaired
    )
    assert result.iterations == 1, f"expected one measured pass, took {result.iterations}"
    assert _geometric(verify_deck(deck, vision=False, out_dir=tmp_path / "after")) == []

    after = Emu(
        next(s for s in Presentation(deck).slides[0].shapes if s.name == "Table 1215").height
    ).inches
    assert after == pytest.approx(before, abs=0.01), (
        "the table's declared height must not move — the repair is smaller text, not "
        "a smaller declaration"
    )


@needs_render
def test_repairs_prl18_risk_table_by_clamping_to_the_library_height(tmp_path):
    """PRL18 declared 5.360" against the library's 5.1715"; the clamp is the fix.

    Distinct from PRL17: nothing renders past anything here, the declaration itself
    is wrong. So the repair clamps and only steps the font if the clamped rows then
    render past their new bottom.
    """
    deck = _copy(_REGRESSIONS / "prl18-risk-table.pptx", tmp_path)

    result = converge_deck(deck, out_dir=tmp_path / "qa")

    assert result.converged, result.summary()
    table = next(s for s in Presentation(deck).slides[0].shapes if s.name == "Table 4")
    assert Emu(table.height).inches == pytest.approx(5.1715, abs=0.01)
    assert _geometric(verify_deck(deck, vision=False, out_dir=tmp_path / "after")) == []


# ─── The loop fails loudly rather than shipping a bad deck ───────────────────


@needs_render
def test_does_not_converge_when_the_content_is_over_budget(tmp_path):
    """PRL14's overview block cannot fit its box at any legible size.

    1,126 characters in a box declaring 0.58" of height, with the box never sized to
    the band (the pre-v0.5.23 defect). The loop spends its ladder, reports the
    remaining overflow honestly, and refuses to call it converged. The retired
    estimator floored at 70% and shipped.
    """
    deck = _copy(_REGRESSIONS / "prl14-overview-bullets.pptx", tmp_path)

    result = converge_deck(deck, out_dir=tmp_path / "qa")

    assert not result.converged
    assert result.iterations <= 3, "the loop must stay bounded"
    surviving = [f for f in result.unrepaired if f.shape == "TextBox 9"]
    assert surviving, f"expected the overview block to still be reported: {result.summary()}"

    # It did shrink as far as it could before giving up.
    box = next(s for s in Presentation(deck).slides[0].shapes if s.name == "TextBox 9")
    assert normal_autofit_scale(box) == pytest.approx(min(TEXT_SCALES))

    with pytest.raises(DeckNotConvergedError) as excinfo:
        assert_converged(deck, result)
    message = str(excinfo.value)
    assert "TextBox 9" in message
    assert "shorten the copy" in message, (
        f"the remedy for over-budget content is editorial, and the error should say so: "
        f"{message}"
    )


def test_assert_converged_is_silent_on_a_converged_result(tmp_path):
    assert_converged(tmp_path / "any.pptx", ConvergeResult(converged=True, iterations=0)) is None


# ─── Scope: what the loop must NOT touch ─────────────────────────────────────


@needs_render
def test_string_findings_are_not_treated_as_repairable(tmp_path):
    """The pitch fixture has 12 blocking findings and needs zero repairs.

    All twelve are `[x]` tokens and unsubstituted currency letters — real defects,
    but not ones a font size fixes. The loop must recognise that its remedies do
    not apply and converge immediately rather than shrinking innocent shapes.
    """
    deck = _copy(_FIXTURES / "pitch-deck.pptx", tmp_path)
    before = Presentation(deck).slides[6]
    before_scale = normal_autofit_scale(
        next(s for s in before.shapes if s.name == "TextBox 9")
    )

    result = converge_deck(deck, out_dir=tmp_path / "qa")

    assert result.converged, result.summary()
    assert result.iterations == 0
    assert result.actions == []
    assert len(result.blocking) == 12, [str(f) for f in result.blocking]
    assert all(f.kind not in REPAIRABLE_KINDS for f in result.blocking)
    after_scale = normal_autofit_scale(
        next(s for s in Presentation(deck).slides[6].shapes if s.name == "TextBox 9")
    )
    assert after_scale == before_scale, "a converged shape must not be re-scaled"


@needs_render
def test_vision_findings_never_enter_the_loop(tmp_path):
    """Advisory findings are non-deterministic and belong at the `deck` checkpoint.

    Run with the vision tier ON: it must contribute findings to the result (so the
    checkpoint can show them) while contributing nothing to the repair decision.
    """
    deck = _copy(_FIXTURES / "pitch-deck.pptx", tmp_path)

    result = converge_deck(deck, out_dir=tmp_path / "qa", vision=True)

    assert result.converged
    assert result.actions == []
    assert any(f.kind == "vision-review" for f in result.advisory), (
        "the vision agenda must still reach the caller"
    )
    assert all(not f.blocking for f in result.advisory)


@needs_render
def test_converging_an_already_converged_deck_changes_nothing(tmp_path):
    """Idempotence: the stage can be re-run without cumulative shrinking.

    A loop that shaved a point off every pass would degrade a deck each time the
    stage was retried.
    """
    deck = _copy(_REGRESSIONS / "prl18-risk-table.pptx", tmp_path)
    assert converge_deck(deck, out_dir=tmp_path / "qa1").converged
    first = deck.read_bytes()

    second = converge_deck(deck, out_dir=tmp_path / "qa2")
    assert second.converged
    assert second.iterations == 0
    assert second.actions == []
    assert deck.read_bytes() == first, "a second pass must not modify a converged deck"


@needs_render
def test_table_repair_leaves_the_header_row_alone(tmp_path):
    """The library's 12 pt Considerations header must survive a body-font step.

    Shrinking every cell uniformly is what made an earlier revision "render
    noticeably smaller than the template" (v0.5.14).
    """
    deck = _copy(_REGRESSIONS / "prl18-risk-table.pptx", tmp_path)
    table = next(s for s in Presentation(deck).slides[0].shapes if s.name == "Table 4").table
    header_before = [
        run.font.size
        for cell in table.rows[0].cells
        for para in cell.text_frame.paragraphs
        for run in para.runs
    ]

    converge_deck(deck, out_dir=tmp_path / "qa")

    table = next(s for s in Presentation(deck).slides[0].shapes if s.name == "Table 4").table
    header_after = [
        run.font.size
        for cell in table.rows[0].cells
        for para in cell.text_frame.paragraphs
        for run in para.runs
    ]
    assert header_after == header_before
    assert all(size is None or size >= Pt(12) for size in header_after)


# ─── The library is the reference, and satisfies its own contract ────────────


@needs_render
def test_blank_library_needs_no_repair(tmp_path):
    """The geometric baseline cannot need repairing against itself."""
    deck = _copy(default_library_path(), tmp_path)

    result = converge_deck(deck, out_dir=tmp_path / "qa")

    assert result.converged
    assert result.actions == []
    assert deck.read_bytes() == default_library_path().read_bytes()
