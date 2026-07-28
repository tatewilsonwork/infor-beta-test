"""Unit tests for the plan_schedule wave scheduler."""

from pathlib import Path

import pytest
import yaml

from plan_refs import iter_input_strings, parse_ref
from plan_schedule import PlanCycleError, compute_waves, stage_dependencies
from schemas import OutputSpec, Plan, Stage

_PLANS_DIR = Path(__file__).resolve().parents[2] / "plans"


def _load_plan(name: str) -> Plan:
    return Plan.model_validate(yaml.safe_load((_PLANS_DIR / name).read_text(encoding="utf-8")))


def _stage(id, skill=None, inputs=None):
    return Stage(
        id=id,
        skill=skill or id,
        inputs=inputs or {},
        outputs=[OutputSpec(name="out", type="str")],
    )


# --- real shipped plans -----------------------------------------------------


def test_pitch_plan_waves():
    """The pitch plan schedules 10 stages into 6 dependency waves.

    Wave 0 overlaps the research-heavy roots (financial-summary / comps /
    precedents / wireframe). financial-summary precedes ltm-metrics because it
    selects the deck's metrics and tells ltm-metrics which extra ones to bridge.
    `financial-charts` runs strictly last: it edits the assembled deck.

    Phase D removed the `workbook-aggregation` wave that used to sit between
    `deck` and `financial-charts`. The deal owns one workbook from stage one, so
    there is nothing to consolidate and the chart stage's LTM links already
    resolve.
    """
    waves = compute_waves(_load_plan("pitch.yaml"))
    assert waves == [
        ["wireframe", "financial-summary", "comps", "precedents"],
        ["content", "ltm-metrics"],
        ["captable"],
        ["ownership"],
        ["deck"],
        ["financial-charts"],
    ]


def test_earnings_update_plan_waves():
    waves = compute_waves(_load_plan("earnings-update.yaml"))
    assert waves == [
        ["wireframe", "ltm-metrics"],
        ["content", "captable"],
        ["deck"],
    ]


def test_financial_charts_depends_on_deck():
    """`financial-charts` edits the assembled deck, so it must depend on `deck`.

    That is now its ONLY ordering constraint — before Phase D it also had to
    follow `workbook-aggregation`, because the `financial-summary` LTM links
    resolved only in the combined workbook. Lock the remaining edge so a plan
    edit cannot reorder the stage ahead of the deck it mutates.
    """
    deps = stage_dependencies(_load_plan("pitch.yaml"))
    assert "deck" in deps["financial-charts"]


def test_required_deck_gate_precedes_final_artefact_waves():
    """Both shipped plans mark `deck` as the `required` pre-delivery checkpoint
    (v0.5.31). Checkpoints are evaluated at the wave boundary and only stop
    DOWNSTREAM waves, so the gate is only real if `deck` is scheduled in an
    earlier wave than the final-artefact stages — lock that here so a plan edit
    can't silently move charts alongside (or ahead of) the gate."""
    for name in ("pitch.yaml", "earnings-update.yaml"):
        plan = _load_plan(name)
        checkpoints = {s.id: s.checkpoint for s in plan.stages}
        assert checkpoints["deck"] == "required", name
        assert all(m == "informational" for sid, m in checkpoints.items() if sid != "deck"), name
        waves = compute_waves(plan)
        wave_index = {sid: i for i, wave in enumerate(waves) for sid in wave}
        if "financial-charts" in wave_index:
            assert wave_index["deck"] < wave_index["financial-charts"], name
        else:
            # Since Phase D the earnings-update plan ENDS at `deck`, so its gate
            # has no downstream wave to hold; it still fires before delivery.
            assert wave_index["deck"] == len(waves) - 1, name


def test_every_stage_scheduled_exactly_once():
    for name in ("pitch.yaml", "earnings-update.yaml"):
        plan = _load_plan(name)
        scheduled = [sid for wave in compute_waves(plan) for sid in wave]
        assert sorted(scheduled) == sorted(s.id for s in plan.stages)
        assert len(scheduled) == len(set(scheduled))  # no stage twice


def test_waves_respect_dependencies():
    """Each stage appears in a strictly later wave than all of its deps."""
    plan = _load_plan("pitch.yaml")
    deps = stage_dependencies(plan)
    wave_index = {sid: i for i, wave in enumerate(compute_waves(plan)) for sid in wave}
    for sid, sid_deps in deps.items():
        for dep in sid_deps:
            assert wave_index[dep] < wave_index[sid], f"{sid} must follow {dep}"


# --- dependency derivation --------------------------------------------------


def test_data_edges_from_references():
    plan = Plan(
        deliverable_type="pitch",
        description="x",
        stages=[
            _stage("a"),
            _stage("b", inputs={"x": "$stages.a.out"}),
        ],
    )
    assert stage_dependencies(plan) == {"a": set(), "b": {"a"}}


def test_nested_references_are_found():
    """References buried in a sub-dict/list (e.g. the aggregator's `workbooks:`)
    still produce an edge."""
    plan = Plan(
        deliverable_type="pitch",
        description="x",
        stages=[
            _stage("a"),
            _stage("b"),
            _stage("c", inputs={"workbooks": {"a": "$stages.a.out", "b": "$stages.b.out"}}),
        ],
    )
    assert stage_dependencies(plan)["c"] == {"a", "b"}


def test_unknown_stage_reference_is_ignored():
    """The scheduler drops the edge rather than crashing on a typo'd `$stages.<id>`.

    This leniency is defense-in-depth only: since v0.5.30 the load-time pre-flight
    (`plan_refs.validate_plan_references`, run by `conductor_cli.load_plan` and the
    conductor's Step 3) rejects such plans before they ever reach the scheduler.
    A hand-built plan that skipped the pre-flight still schedules; the bad ref is
    then rejected by `plan_refs.resolve_refs` at dispatch time."""
    plan = Plan(
        deliverable_type="pitch",
        description="x",
        stages=[_stage("a", inputs={"x": "$stages.ghost.out"})],
    )
    assert stage_dependencies(plan) == {"a": set()}


def test_non_stage_references_do_not_create_edges():
    plan = Plan(
        deliverable_type="pitch",
        description="x",
        stages=[
            _stage("a", inputs={"q": "$plan_inputs.reporting_quarter", "c": "$deal.codename"}),
        ],
    )
    assert stage_dependencies(plan) == {"a": set()}


# --- no forced barrier (Phase D) --------------------------------------------


def test_no_stage_gets_a_dependency_it_does_not_reference():
    """Every edge is a real reference — there is no hardcoded barrier any more.

    Replaces three tests that pinned the `workbook-aggregator` barrier: it
    depended on every stage EXCEPT its own downstream consumers (excluding them
    was what kept a post-aggregation stage from forming a cycle), and was forced
    alone into the final wave. That rule, and the filesystem side-effect it
    modelled — the aggregator merging and DELETING standalone workbooks the deck
    had already read — are gone with the aggregator.
    """
    for name in ("pitch.yaml", "earnings-update.yaml"):
        plan = _load_plan(name)
        deps = stage_dependencies(plan)
        for stage in plan.stages:
            referenced = set()
            for text in iter_input_strings(stage.inputs):
                parsed = parse_ref(text)
                if parsed and parsed[0] == "stages":
                    referenced.add(parsed[1][0])
            extra = deps[stage.id] - referenced
            assert not extra, (
                f"{name}: {stage.id} has dependencies it never references: {extra}"
            )


# --- cycle detection --------------------------------------------------------


def test_cycle_raises():
    plan = Plan(
        deliverable_type="pitch",
        description="x",
        stages=[
            _stage("a", inputs={"x": "$stages.b.out"}),
            _stage("b", inputs={"x": "$stages.a.out"}),
        ],
    )
    with pytest.raises(PlanCycleError):
        compute_waves(plan)


def test_single_stage_plan_is_one_wave():
    plan = Plan(deliverable_type="pitch", description="x", stages=[_stage("only")])
    assert compute_waves(plan) == [["only"]]
