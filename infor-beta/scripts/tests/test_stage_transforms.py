"""Tests for the Phase F in-process stage transforms (scripts/stage_transforms.py).

**These are the first executable tests of these four call sites.** Until Phase F
each transform was a sub-agent whose SKILL.md ended in a "Reference command" — a
fenced Python block the dispatched model was asked to reproduce. No test ran those
blocks, because a SKILL.md is prose; the only thing that ever executed them was a
live run. So deleting the four SKILL.md files removes no coverage, and porting
their reference commands into `stage_transforms.py` is what makes them testable at
all. What is asserted below is exactly what the prose used to promise:

- the classification itself (which skills are transforms, and that nothing else
  drifted into the registry — `deckcheck` above all, which reads renders and
  argues about figures and must stay a sub-agent);
- that no transform left a SKILL.md or a skill directory behind, and that no
  shipped plan references a skill that no longer exists;
- the argument fidelity of each port — the same call, with the same arguments,
  from the same `StageIO`;
- the two things the deleted prose made *mandatory* and a naive deletion would
  have dropped silently: the deck's vision review, and `financial-charts` chaining
  the pie onto the deck the FS-chart step wrote.

The assemblers and chart orchestrators themselves are covered in depth by
`test_slide_library_poc.py`, `test_earnings_update_assembler.py` and
`test_financial_charts.py`; nothing here duplicates them.
"""

from __future__ import annotations

import json
from pathlib import Path

import pytest
import yaml

import financial_charts
import stage_transforms
from deal_init import save_deal_context
from excel_to_powerpoint import find_soffice
from run_log import create_run_dir, stage_dir, write_stage_inputs
from schemas import Company, DealContext, Plan, SlidePlan
from stage_io import stage_io
from stage_transforms import StageTransformError, is_transform, run_transform

from tests.test_earnings_update_assembler import _sample_content as _earnings_content
from tests.test_slide_library_poc import _sample_content as _pitch_content

PLUGIN_ROOT = Path(__file__).resolve().parents[2]
_PLANS_DIR = PLUGIN_ROOT / "plans"
_SKILLS_DIR = PLUGIN_ROOT / "skills"

#: The Phase F classification, written out so a change to the registry has to be a
#: deliberate change to this list too.
_TRANSFORM_SKILLS = {
    "pitch-wireframe",
    "earningsupdate-wireframe",
    "deck-assembler",
    "financial-charts",
}
_JUDGMENT_SKILLS = {
    "captable",
    "comps",
    "precedents",
    "ownership",
    "financial-summary",
    "ltm-metrics",
    "pitch-content",
    "earningsupdate-content",
    "deckcheck",
}


def _plan(name: str) -> Plan:
    return Plan.model_validate(yaml.safe_load((_PLANS_DIR / name).read_text(encoding="utf-8")))


def _shipped_plans() -> list[Plan]:
    return [_plan("pitch.yaml"), _plan("earnings-update.yaml")]


def _io(tmp_path: Path, stage_id: str, inputs: dict):
    """A real `StageIO` for `stage_id`, built the way `run_transforms` builds one.

    A deal directory with `deal.json` (so `io.deal_dir` / `io.artefacts_dir`
    resolve by walking up, as they do in production), a run directory, and the
    stage's `inputs.json` written through the same serializer `prepare_wave` uses.
    """
    deal_dir = tmp_path / "Project Sample"
    save_deal_context(
        DealContext(
            codename="Project Sample",
            deal_dir=deal_dir,
            deliverable_type="pitch",
            subject_company=Company(legal_name="SampleCo Ltd.", ticker="TSX:SMPL"),
        )
    )
    run_dir = create_run_dir(deal_dir, "2026-07-28-test01")
    inputs_path = write_stage_inputs(run_dir, stage_id, inputs)
    return stage_io(
        [
            "test",
            str(PLUGIN_ROOT),
            str(inputs_path),
            str(stage_dir(run_dir, stage_id) / "outputs.json"),
        ]
    )


# ─── The classification ──────────────────────────────────────────────────────


def test_the_registry_holds_exactly_the_four_transforms():
    assert set(stage_transforms.TRANSFORMS) == _TRANSFORM_SKILLS


def test_every_judgment_skill_stays_out_of_the_registry():
    """`deckcheck` is the one that matters most.

    Phase G's falsification pass reads rendered PNGs, provenance records and
    source filings and argues about whether a figure is true. It is judgment by
    construction, and running it in-process would mean a Python function deciding
    whether a target's financial statements support a number on a slide.
    """
    for skill in _JUDGMENT_SKILLS:
        assert not is_transform(skill), f"{skill} must stay a dispatched sub-agent"


def test_every_shipped_stage_is_classified_exactly_once():
    for plan in _shipped_plans():
        for stage in plan.stages:
            transform = is_transform(stage.skill)
            has_skill_doc = (_SKILLS_DIR / stage.skill / "SKILL.md").is_file()
            assert transform != has_skill_doc, (
                f"{stage.skill}: a stage is either an in-process transform (no SKILL.md) "
                f"or a dispatched skill (with one) — found transform={transform}, "
                f"SKILL.md={has_skill_doc}"
            )


def test_no_shipped_plan_references_a_deleted_skill():
    """The Phase D failure mode, one phase on: a plan naming a skill that is gone.

    A judgment stage whose SKILL.md was deleted would dispatch a sub-agent with no
    instructions; the run would not fail until the wave collected nothing.
    """
    for plan in _shipped_plans():
        for stage in plan.stages:
            if is_transform(stage.skill):
                continue
            assert (_SKILLS_DIR / stage.skill / "SKILL.md").is_file(), (
                f"plan stage {stage.id!r} dispatches {stage.skill!r}, which has no SKILL.md"
            )


def test_no_transform_left_a_skill_directory_behind():
    """Phase D's pattern (workbook-aggregator): the doc and its directory both go.

    A leftover SKILL.md would be advertised in the plugin's skill list and
    invocable by name, so an analyst could run a stage the conductor now owns —
    against an `inputs.json` that no longer exists.
    """
    for skill in _TRANSFORM_SKILLS:
        assert not (_SKILLS_DIR / skill).exists(), f"skills/{skill}/ should have been deleted"


# ─── What the reclassification must not have reordered ───────────────────────


def test_the_reclassification_left_the_shipped_orderings_intact():
    """Three orderings this phase had to preserve, asserted as properties.

    `test_pitch_plan_waves` / `test_earnings_update_plan_waves` pin the full wave
    lists; this states *why* those lists matter, so a future plan edit that keeps
    them well-formed but breaks one of these fails with the reason attached.

    Who executes a stage is not an ordering fact — every edge still comes from a
    `$stages` reference — so all three hold for the same reason they did before:
    `compute_waves` never saw the transform registry.
    """
    from plan_schedule import compute_waves

    for plan in _shipped_plans():
        waves = compute_waves(plan)

        # 1. `deck` sits ALONE in its wave. Every other stage either feeds the
        #    assembly or consumes the assembled file, so a wave-mate would be a
        #    stage running concurrently with the deck it reads or edits —
        #    `financial-charts` mutates that file in place.
        deck_wave = next(w for w in waves if "deck" in w)
        assert deck_wave == ["deck"], f"{plan.deliverable_type}: deck must be alone in its wave"

        # 2. Nothing gates. `deck` carried the plugin's only `required` checkpoint
        #    until v0.5.49; the run is autonomous now, and the deck's QA (the
        #    converge loop, `vision_review_path`, `deckcheck`) asks no questions.
        gates = [s.id for s in plan.stages if str(s.checkpoint) == "required"]
        assert gates == [], f"{plan.deliverable_type}: expected no gates, got {gates}"

        # 3. `deckcheck` runs LAST, on the finished artefact — after `deck` and,
        #    for pitch, after `financial-charts`, whose charts carry figures a
        #    review of the pre-chart deck would never see.
        assert waves[-1] == ["deckcheck"], (
            f"{plan.deliverable_type}: deckcheck must be the final wave, alone"
        )
        deck_index = waves.index(deck_wave)
        assert deck_index < len(waves) - 1, "deck is not the last wave — deckcheck follows it"


def test_run_transform_refuses_a_judgment_skill(tmp_path: Path):
    io = _io(tmp_path, "comps", {"ticker": "TSX:SMPL"})
    with pytest.raises(StageTransformError, match="must be.*dispatched"):
        run_transform("comps", io)


# ─── Wireframes ──────────────────────────────────────────────────────────────


def test_pitch_wireframe_transform_writes_the_slide_plan(tmp_path: Path):
    io = _io(
        tmp_path,
        "wireframe",
        {
            "company": Company(legal_name="SampleCo Ltd.", ticker="TSX:SMPL"),
            "client_name": "SampleCo Ltd.",
            "financial_metric_count": 8,
            "include_investment_highlights": False,
        },
    )
    outputs = run_transform("pitch-wireframe", io)

    path = Path(outputs["slide_plan_path"])
    assert path == io.stage_dir / "slide_plan.json"
    plan = SlidePlan.model_validate_json(path.read_text(encoding="utf-8"))
    assert plan.deliverable_type == "pitch"
    ids = [s.library_entry_id for s in plan.slides]
    # The deck-spec options reached the builder: 8 metrics -> two FS slides, and
    # "omit" drops Key Investment Highlights entirely.
    assert ids.count("financial-summary") == 2
    assert "key-investment-highlights" not in ids


def test_pitch_wireframe_transform_defaults_the_slide_mix(tmp_path: Path):
    """Unsupplied optional plan inputs arrive as `null` and must read as "absent".

    `prepare_wave` resolves an optional plan input the analyst did not supply to
    `None`, so the transform receives explicit nulls rather than missing keys —
    which is the shape the builder's "omitted → default" rules have to survive.
    """
    io = _io(
        tmp_path,
        "wireframe",
        {
            "company": Company(legal_name="SampleCo Ltd."),
            "section_labels": None,
            "current_section": None,
            "market_entry_target_count": None,
            "financial_metric_count": None,
            "include_investment_highlights": None,
        },
    )
    plan = SlidePlan.model_validate_json(
        Path(run_transform("pitch-wireframe", io)["slide_plan_path"]).read_text(encoding="utf-8")
    )
    ids = [s.library_entry_id for s in plan.slides]
    assert ids.count("financial-summary") == 1  # default: one FS slide
    assert ids.count("market-entry-targets") == 4  # default: 8 targets, two per slide
    assert "key-investment-highlights" in ids  # default: included


def test_earnings_update_wireframe_transform_writes_the_fixed_five(tmp_path: Path):
    io = _io(
        tmp_path,
        "wireframe",
        {
            "company": Company(legal_name="SampleCo", ticker="TSX:SMPL"),
            "ticker": "TSX:SMPL",
            "reporting_quarter": "Q4 2025",
            "comparison_quarter": "Q4 2024",
        },
    )
    outputs = run_transform("earningsupdate-wireframe", io)

    plan = SlidePlan.model_validate_json(
        Path(outputs["slide_plan_path"]).read_text(encoding="utf-8")
    )
    assert plan.deliverable_type == "earnings-update"
    assert [s.library_entry_id for s in plan.slides] == [
        "earnings-update-cover",
        "earnings-update-company-overview",
        "earnings-update-earnings-summary",
        "earnings-update-disclaimer",
        "earnings-update-contact",
    ]
    assert "Q4 2025" in plan.slides[2].title


# ─── Deck assembly ───────────────────────────────────────────────────────────


def _deck_inputs(tmp_path: Path, io, *, pitch: bool) -> dict:
    """Write the two typed artefacts a deck stage consumes and return its inputs."""
    content = (_pitch_content() if pitch else _earnings_content()).model_dump_json()
    content_path = io.stage_dir / "content.json"
    content_path.write_text(content, encoding="utf-8")

    wireframe_io = _io(tmp_path / "wf", "wireframe", {"company": Company(legal_name="SampleCo")})
    if pitch:
        slide_plan_path = run_transform("pitch-wireframe", wireframe_io)["slide_plan_path"]
    else:
        wireframe_io.inputs.update(reporting_quarter="Q4 2025", comparison_quarter="Q4 2024")
        slide_plan_path = run_transform("earningsupdate-wireframe", wireframe_io)[
            "slide_plan_path"
        ]
    return {
        "slide_plan_path": slide_plan_path,
        "content_bundle_path": str(content_path),
        "template_name": "INFOR Slide Library.pptx",
    }


@pytest.mark.parametrize("pitch", [True, False], ids=["pitch", "earnings-update"])
def test_deck_transform_calls_the_right_assembler_with_the_envelope_arguments(
    tmp_path: Path, monkeypatch, pitch: bool
):
    """The port's fidelity: same assembler, same arguments, same output directory.

    Dispatch is on the SlidePlan's own `deliverable_type`, not on the plan, so the
    two deliverables cannot be crossed by a mis-set plan input. The deck lands in
    `<deal_dir>/artefacts/` — derived from the inputs path, which is what the
    Phase E handoff replaced a silently-unset `DEAL_DIR` with.
    """
    io = _io(tmp_path, "deck", {})
    io.inputs.update(_deck_inputs(tmp_path, io, pitch=pitch))
    seen = {}

    def _fake(**kwargs):
        seen.update(kwargs)
        deck = Path(kwargs["output_dir"]) / "Deck.pptx"
        deck.write_bytes(b"")
        return deck

    import earnings_update_assembler
    import pitch_deck_assembler

    monkeypatch.setattr(pitch_deck_assembler, "assemble_pitch_deck", _fake)
    monkeypatch.setattr(earnings_update_assembler, "assemble_earnings_update_deck", _fake)
    monkeypatch.setattr(stage_transforms, "_write_vision_review", lambda io, deck: None)

    outputs = run_transform("deck-assembler", io)

    assert outputs["deck_path"] == str(io.artefacts_dir / "Deck.pptx")
    assert seen["slide_plan_path"] == io.inputs["slide_plan_path"]
    assert seen["content_path"] == io.inputs["content_bundle_path"]
    assert seen["template_path"] == PLUGIN_ROOT / "templates" / "INFOR Slide Library.pptx"
    assert seen["output_dir"] == io.artefacts_dir
    assert seen["captable_workbook_path"] is None  # absent input, not a KeyError
    if pitch:
        assert seen["ownership_workbook_path"] is None
        assert seen["financial_metric_labels"] is None
    else:
        assert "ownership_workbook_path" not in seen


def test_deck_transform_rejects_an_unknown_deliverable_type(tmp_path: Path):
    io = _io(tmp_path, "deck", {})
    io.inputs.update(_deck_inputs(tmp_path, io, pitch=True))
    plan_path = Path(io.inputs["slide_plan_path"])
    plan_path.write_text(
        plan_path.read_text(encoding="utf-8").replace('"pitch"', '"overview"', 1), encoding="utf-8"
    )
    with pytest.raises(StageTransformError, match="unsupported deliverable_type"):
        run_transform("deck-assembler", io)


@pytest.mark.skipif(
    find_soffice() is None,
    reason="LibreOffice is the only render backend; the assembler's converge loop needs it",
)
def test_deck_transform_assembles_and_writes_the_vision_review(tmp_path: Path):
    """End to end on the real assembler, including the converge loop.

    Both halves of what this stage owes the analyst: a deck in the deal's artefacts
    directory, and a written review of its slides they can open. Since v0.5.49 there
    is no approval dialog to hold that review behind, which makes writing it the
    only thing that gets it read. Earnings update rather than pitch — five slides
    against nineteen, for the same wiring.
    """
    io = _io(tmp_path, "deck", {})
    io.inputs.update(_deck_inputs(tmp_path, io, pitch=False))

    outputs = run_transform("deck-assembler", io)

    deck = Path(outputs["deck_path"])
    assert deck.is_file() and deck.parent == io.artefacts_dir

    review = Path(outputs["vision_review_path"])
    assert review == io.stage_dir / stage_transforms.VISION_REVIEW_NAME
    body = review.read_text(encoding="utf-8")
    assert deck.name in body
    # The reading half of deck QA, which the deleted SKILL.md made mandatory: the
    # renders are on disk and the note says what to look for in them.
    assert "text drawn over other text" in body
    assert list((io.stage_dir / "vision" / "slides").glob("*.png"))


def test_vision_review_records_the_reason_it_could_not_run(tmp_path: Path):
    """A broken review must not fail the stage — but it must not vanish either.

    The deck exists and the run carries on, so the reason is written into the file
    `vision_review_path` points at, naming the deck to read by hand. Driven with an
    unreadable .pptx, which is the shape of every way this can fail.
    """
    io = _io(tmp_path, "deck", {})
    deck = io.artefacts_dir / "Deck.pptx"
    deck.write_bytes(b"not a pptx")

    path = Path(stage_transforms._write_vision_review(io, deck))
    assert path == io.stage_dir / stage_transforms.VISION_REVIEW_NAME
    body = path.read_text(encoding="utf-8")
    assert "could not be built" in body
    assert "by hand before it goes out" in body
    assert str(deck) in body


# ─── Financial Summary charts ────────────────────────────────────────────────


def test_financial_charts_transform_chains_the_pie_onto_the_charted_deck(
    tmp_path: Path, monkeypatch
):
    """The one piece of real logic in this port, and the one the prose insisted on.

    Both helpers modify the deck and return the path they wrote; the pie step must
    run against the FS step's output, not the original `deck_path`. Getting that
    wrong loses the FS charts silently — the same class of bug as re-assembling.
    """
    calls = []

    def _fs(*, deck_path, deal_workbook, **kw):
        calls.append(("fs", deck_path))
        return Path(deck_path)

    def _pie(*, deck_path, deal_workbook, **kw):
        calls.append(("pie", deck_path))
        return Path(deck_path)

    monkeypatch.setattr(financial_charts, "render_financial_summary_charts_into_deck", _fs)
    monkeypatch.setattr(financial_charts, "render_ltm_revenue_pie_into_deck", _pie)
    monkeypatch.setattr(stage_transforms, "_render_chart_qa", lambda io, deck: None)

    deck = str(tmp_path / "Deck.pptx")
    io = _io(tmp_path, "financial-charts", {"deck_path": deck, "deal_workbook": "/deals/wb.xlsx"})
    outputs = run_transform("financial-charts", io)

    assert [c[0] for c in calls] == ["fs", "pie"]
    assert calls[1][1] == deck  # the path the FS step handed back, not the input
    assert outputs == {
        "deck_path": deck,
        "charts_inserted": True,
        "pie_inserted": True,
        "chart_qa_dir": None,
    }


def test_financial_charts_transform_reports_a_skip_instead_of_hiding_it(
    tmp_path: Path, monkeypatch
):
    """A `None` from either helper means the placeholders stayed.

    In the dispatched form this depended on the sub-agent remembering to "say so
    explicitly"; the booleans are declared plan outputs now, so the wave-boundary
    surface carries the skip whether anyone mentions it or not.
    """
    monkeypatch.setattr(
        financial_charts, "render_financial_summary_charts_into_deck", lambda **kw: None
    )
    monkeypatch.setattr(financial_charts, "render_ltm_revenue_pie_into_deck", lambda **kw: None)
    monkeypatch.setattr(stage_transforms, "_render_chart_qa", lambda io, deck: None)

    deck = str(tmp_path / "Deck.pptx")
    io = _io(tmp_path, "financial-charts", {"deck_path": deck, "deal_workbook": "/deals/wb.xlsx"})
    outputs = run_transform("financial-charts", io)

    assert outputs["charts_inserted"] is False and outputs["pie_inserted"] is False
    assert outputs["deck_path"] == deck  # unchanged, not None


def test_financial_charts_transform_cannot_dispatch_a_skill():
    """The standing rule, now structural rather than an allow-list.

    `financial-charts` runs after the deck is assembled, so re-assembling would
    revert every filled table to a placeholder. Its SKILL.md excluded `Task` from
    its allow-list to make that impossible; an in-process transform has no tools at
    all, and this locks the module against growing a dispatch of its own.
    """
    source = (PLUGIN_ROOT / "scripts" / "stage_transforms.py").read_text(encoding="utf-8")
    for banned in ("Task(", "subprocess", "matplotlib", "plotly"):
        assert banned not in source, f"stage_transforms must not reach for {banned}"


# ─── Outputs contract ────────────────────────────────────────────────────────


def test_every_transform_emits_every_output_its_plan_declares(tmp_path: Path, monkeypatch):
    """`complete_wave` fails a stage that omits a declared output name.

    So the registry and the plans have to agree on the *keys*, not just the stage
    ids — a transform that returns a dict missing one declared name halts the run
    at its own wave boundary. Checked against the shipped plans' declarations, with
    the expensive bodies stubbed: this is about the contract, not the artefacts.
    """
    monkeypatch.setattr(stage_transforms, "_write_vision_review", lambda io, deck: None)
    monkeypatch.setattr(stage_transforms, "_render_chart_qa", lambda io, deck: None)
    monkeypatch.setattr(
        financial_charts, "render_financial_summary_charts_into_deck", lambda **kw: None
    )
    monkeypatch.setattr(financial_charts, "render_ltm_revenue_pie_into_deck", lambda **kw: None)

    import earnings_update_assembler
    import pitch_deck_assembler

    def _fake(**kwargs):
        deck = Path(kwargs["output_dir"]) / "Deck.pptx"
        deck.write_bytes(b"")
        return deck

    monkeypatch.setattr(pitch_deck_assembler, "assemble_pitch_deck", _fake)
    monkeypatch.setattr(earnings_update_assembler, "assemble_earnings_update_deck", _fake)

    checked = set()
    for plan in _shipped_plans():
        for stage in plan.stages:
            if not is_transform(stage.skill):
                continue
            io = _io(tmp_path / f"{plan.deliverable_type}-{stage.id}", stage.id, {})
            pitch = plan.deliverable_type == "pitch"
            if stage.skill.endswith("wireframe"):
                io.inputs.update(
                    company=Company(legal_name="SampleCo"),
                    reporting_quarter="Q4 2025",
                    comparison_quarter="Q4 2024",
                )
            elif stage.skill == "deck-assembler":
                io.inputs.update(_deck_inputs(tmp_path / f"{plan.deliverable_type}-deck-in", io, pitch=pitch))
            else:
                io.inputs.update(deck_path="/deals/Deck.pptx", deal_workbook="/deals/wb.xlsx")

            outputs = run_transform(stage.skill, io)
            missing = [spec.name for spec in stage.outputs if spec.name not in outputs]
            assert not missing, (
                f"{plan.deliverable_type}/{stage.id} ({stage.skill}) omits declared "
                f"output(s) {missing}; `complete_wave` would fail the stage"
            )
            checked.add(stage.skill)

    assert checked == _TRANSFORM_SKILLS


def test_a_transform_writes_json_serializable_outputs(tmp_path: Path):
    """`StageIO.write` is the same serializer the sub-agent used, so a `Path` in a
    returned dict must survive the round trip the conductor reads back."""
    io = _io(tmp_path, "wireframe", {"company": Company(legal_name="SampleCo")})
    io.write(run_transform("pitch-wireframe", io))
    assert "slide_plan_path" in json.loads(io.outputs_path.read_text(encoding="utf-8"))
