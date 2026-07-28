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
    """The pitch plan schedules 11 stages into 7 dependency waves.

    Wave 0 overlaps the research-heavy roots (financial-summary / comps /
    precedents / wireframe). financial-summary precedes ltm-metrics because it
    selects the deck's metrics and tells ltm-metrics which extra ones to bridge.
    `financial-charts` edits the assembled deck, and Phase G's `deckcheck` audits
    what `financial-charts` produced — so the two tail waves are strictly ordered.

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
        ["deckcheck"],
    ]


def test_earnings_update_plan_waves():
    waves = compute_waves(_load_plan("earnings-update.yaml"))
    assert waves == [
        ["wireframe", "ltm-metrics"],
        ["content", "captable"],
        ["deck"],
        ["deckcheck"],
    ]


def test_deckcheck_audits_the_final_artefact_not_the_assembled_one():
    """In the pitch plan `deckcheck` must depend on `financial-charts`, not `deck`.

    Two reasons, and both are load-bearing. The charts land *after* assembly, so a
    review of `deck`'s output would miss every figure they carry. And
    `financial-charts` edits the deck in place — a render running alongside it would
    race the file it is reading. The earnings-update plan has no chart stage, so
    there `deck` is the final artefact.
    """
    pitch = stage_dependencies(_load_plan("pitch.yaml"))
    assert pitch["deckcheck"] == {"financial-charts"}
    assert stage_dependencies(_load_plan("earnings-update.yaml"))["deckcheck"] == {"deck"}


def test_financial_charts_depends_on_deck():
    """`financial-charts` edits the assembled deck, so it must depend on `deck`.

    That is now its ONLY ordering constraint — before Phase D it also had to
    follow `workbook-aggregation`, because the `financial-summary` LTM links
    resolved only in the combined workbook. Lock the remaining edge so a plan
    edit cannot reorder the stage ahead of the deck it mutates.
    """
    deps = stage_dependencies(_load_plan("pitch.yaml"))
    assert "deck" in deps["financial-charts"]


def test_required_deck_gate_precedes_every_later_wave():
    """Both shipped plans mark `deck` as the `required` pre-delivery checkpoint
    (v0.5.31), and nothing else is a gate. Checkpoints are evaluated at the wave
    boundary and only stop DOWNSTREAM waves, so the gate is only real if `deck`
    sits alone in its wave with everything else after it — lock that here so a plan
    edit can't silently move a later stage alongside (or ahead of) the gate.

    Phase G's `deckcheck` is deliberately in that downstream set and deliberately
    `informational`: a falsification pass reports on the artefact the analyst has
    already approved. Making it a second gate would mean halting a run on a claim
    about the target's financial statements.
    """
    for name in ("pitch.yaml", "earnings-update.yaml"):
        plan = _load_plan(name)
        checkpoints = {s.id: s.checkpoint for s in plan.stages}
        assert checkpoints["deck"] == "required", name
        assert all(m == "informational" for sid, m in checkpoints.items() if sid != "deck"), name

        waves = compute_waves(plan)
        wave_index = {sid: i for i, wave in enumerate(waves) for sid in wave}
        assert waves[wave_index["deck"]] == ["deck"], f"{name}: the gate has wave-mates"
        held = [sid for sid, i in wave_index.items() if i > wave_index["deck"]]
        assert "deckcheck" in held, f"{name}: the review is not behind the gate"


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
    (`plan_refs.validate_plan_references`, run by `conductor.load_plan` and the
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


# --- doc drift lock (Phase E) ------------------------------------------------


def _assert_claims_match(claims: dict[str, tuple[int, int]], where: str) -> None:
    assert set(claims) == {"pitch.yaml", "earnings-update.yaml"}, (
        f"{where}: expected a wave claim for each shipped plan, found {sorted(claims)}"
    )
    for name, (stages, waves) in claims.items():
        plan = _load_plan(name)
        assert (len(plan.stages), len(compute_waves(plan))) == (stages, waves), (
            f"{where} says {name} is {stages} stages / {waves} waves; the scheduler "
            f"returns {len(plan.stages)} / {len(compute_waves(plan))}"
        )


def test_readme_wave_counts_match_the_scheduler():
    """Every "N stages … M dependency waves" claim in the README is checked.

    Phase D changed both plans' shapes (pitch 11/7 -> 10/6, earnings update 6/4 ->
    5/3) by deleting a stage, and the numbers live in prose that nothing verified.
    (Phase G added `deckcheck` and put both back: pitch 11/7, earnings update 6/4.)
    The conductor SKILL.md used to carry a hardcoded wave list too; it now posts
    `conductor.plan_overview(run_dir).narration()` instead, so there is no second
    copy left to check.
    """
    import re

    readme = (Path(__file__).resolve().parents[3] / "README.md").read_text(encoding="utf-8")
    pattern = re.compile(
        r"`(?P<plan>[a-z-]+\.yaml)`\*\*[^\n]*?(?P<stages>\d+) stages?"
        r"[^\n]*?(?P<waves>\d+) dependency waves"
    )
    claims = {m.group("plan"): (int(m.group("stages")), int(m.group("waves"))) for m in pattern.finditer(readme)}
    _assert_claims_match(claims, "README.md")


def test_readme_dispatch_counts_match_the_transform_registry():
    """The README's "(N dispatched, M in-process)" split is checked too.

    Phase F's whole claim is a number — how many sub-agents a run costs — and the
    classification behind it lives in one dict. A skill moved between transform and
    judgment without the README following would leave the plugin's headline
    description of its own cost wrong, which is the drift the wave-count locks above
    exist for. Same parser shape, one release later.
    """
    import re

    import stage_transforms

    readme = (Path(__file__).resolve().parents[3] / "README.md").read_text(encoding="utf-8")
    pattern = re.compile(
        r"`(?P<plan>[a-z-]+\.yaml)`\*\*[^\n]*?\((?P<dispatched>\d+) dispatched, "
        r"(?P<transforms>\d+) in-process\)"
    )
    claims = {
        m.group("plan"): (int(m.group("dispatched")), int(m.group("transforms")))
        for m in pattern.finditer(readme)
    }
    assert set(claims) == {"pitch.yaml", "earnings-update.yaml"}, (
        f"README.md: expected a dispatch split for each shipped plan, found {sorted(claims)}"
    )
    for name, (dispatched, transforms) in claims.items():
        plan = _load_plan(name)
        kinds = [stage_transforms.is_transform(s.skill) for s in plan.stages]
        assert (kinds.count(False), kinds.count(True)) == (dispatched, transforms), (
            f"README.md says {name} is {dispatched} dispatched / {transforms} in-process; "
            f"the registry gives {kinds.count(False)} / {kinds.count(True)}"
        )


def test_contributor_brief_wave_counts_match_the_scheduler():
    """CLAUDE.md carries the same two numbers, in its own wording.

    Found while adding `deckcheck`: the brief's "Conductor plans" line said
    "earnings-update.yaml (5 stages / 3 waves)" and nothing checked it, which is the
    exact shape of the drift the README lock was written for. One more parser is
    cheaper than a second stale copy of a number the scheduler already knows.
    """
    import re

    brief = (Path(__file__).resolve().parents[3] / "CLAUDE.md").read_text(encoding="utf-8")
    pattern = re.compile(
        r"`(?P<plan>[a-z-]+\.yaml)`[^\n]*?\((?:[^()\n]*?, )?(?P<stages>\d+) stages? / "
        r"(?P<waves>\d+) waves\)"
    )
    claims = {m.group("plan"): (int(m.group("stages")), int(m.group("waves"))) for m in pattern.finditer(brief)}
    _assert_claims_match(claims, "CLAUDE.md")
