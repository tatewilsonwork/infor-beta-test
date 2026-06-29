"""Unit tests for the plan_schedule wave scheduler."""

from pathlib import Path

import pytest
import yaml

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
    """The pitch plan schedules 11 stages into 7 dependency waves.

    Wave 0 overlaps the research-heavy roots (financial-summary / comps /
    precedents / wireframe). financial-summary precedes ltm-metrics because it
    selects the deck's metrics and tells ltm-metrics which extra ones to bridge.
    `workbook-aggregation` is alone in its wave (it consolidates + deletes the
    source workbooks), then `financial-charts` runs strictly last: it charts the
    combined workbook (where the LTM links resolve) and edits the assembled deck.
    """
    waves = compute_waves(_load_plan("pitch.yaml"))
    assert waves == [
        ["wireframe", "financial-summary", "comps", "precedents"],
        ["content", "ltm-metrics"],
        ["captable"],
        ["ownership"],
        ["deck"],
        ["workbook-aggregation"],
        ["financial-charts"],
    ]


def test_earnings_update_plan_waves():
    waves = compute_waves(_load_plan("earnings-update.yaml"))
    assert waves == [
        ["wireframe", "ltm-metrics"],
        ["content", "captable"],
        ["deck"],
        ["workbook-aggregation"],
    ]


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
    """A typo'd `$stages.<id>` is left for the resolver to reject — the scheduler
    drops the edge rather than crashing."""
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


# --- aggregator barrier -----------------------------------------------------


def test_aggregator_depends_on_everything():
    """A workbook-aggregator stage depends on all others even with no explicit
    reference to some of them — the hardcoded final-barrier rule."""
    plan = Plan(
        deliverable_type="pitch",
        description="x",
        stages=[
            _stage("a"),
            _stage("b"),
            _stage("agg", skill="workbook-aggregator", inputs={"w": "$stages.a.out"}),
        ],
    )
    deps = stage_dependencies(plan)
    assert deps["agg"] == {"a", "b"}  # 'b' added by the barrier, not a reference
    waves = compute_waves(plan)
    assert waves[-1] == ["agg"]  # strictly last, alone


def test_aggregator_runs_after_a_stage_it_does_not_reference():
    """Regression for the deck/aggregator side-effect ordering: the aggregator
    must follow `deck` even though it never references deck's output."""
    plan = _load_plan("pitch.yaml")
    wave_index = {sid: i for i, wave in enumerate(compute_waves(plan)) for sid in wave}
    assert wave_index["workbook-aggregation"] > wave_index["deck"]


def test_post_aggregation_consumer_does_not_cycle():
    """A stage that consumes the combined workbook (e.g. `financial-charts`) runs
    AFTER the aggregator. The barrier must not force the aggregator to depend on
    its own consumer (that would be a cycle)."""
    plan = Plan(
        deliverable_type="pitch",
        description="x",
        stages=[
            _stage("a"),
            _stage("deck", inputs={"x": "$stages.a.out"}),
            _stage("agg", skill="workbook-aggregator", inputs={"w": "$stages.a.out"}),
            _stage("charts", inputs={"wb": "$stages.agg.out", "d": "$stages.deck.out"}),
        ],
    )
    deps = stage_dependencies(plan)
    # The aggregator depends on producers (a, deck) but NOT on its consumer.
    assert "charts" not in deps["agg"]
    assert {"a", "deck"} <= deps["agg"]
    # No cycle; charts is strictly last.
    waves = compute_waves(plan)
    assert waves[-1] == ["charts"]
    wave_index = {sid: i for i, wave in enumerate(waves) for sid in wave}
    assert wave_index["charts"] > wave_index["agg"]


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
