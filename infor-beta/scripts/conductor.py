"""The conductor, as code — plan load, wave prep, collect, checkpoints, summary.

Phase E finishes the trip `conductor_cli` started. Everything mechanical about
running a plan is a function with a return value here, not a paragraph of SKILL.md
prose the model may skip on turn 40 of a long run. What is left to the model is
exactly four things:

1. the intake conversation (deal-init + the locked deck-spec dialogs),
2. issuing the `Task` calls :func:`prepare_wave` hands back,
3. the checkpoint conversation :func:`complete_wave` hands back, and
4. the end-of-run summary, off the skeleton :func:`render_run_summary` returns.

    conductor.py plan          <run_dir>         # deliverable, stages, wave schedule
    conductor.py prepare-wave  <run_dir> <wave>  # resolve + write inputs, print envelopes
    conductor.py complete-wave <run_dir> <wave>  # read + validate outputs, print checkpoints
    conductor.py summary       <run_dir>         # write summary.md, print it

`<wave>` is 1-based, matching the conductor's "wave 1 / wave 2 / …" narration.

State is reconstructed from the run directory, which the conductor already
populates: `<run_dir>/plan.yaml` (the frozen plan snapshot), the deal's
`deal.json` (two levels up: `<deal_dir>/runs/<run-id>/`), an optional
`<run_dir>/plan_inputs.json` (the analyst-collected plan inputs; absent ⇒ `{}`),
and each prior stage's `stages/<id>/outputs.json`. Preparing wave *n* reads the
outputs of every earlier wave so `$stages.*` references resolve. Nothing is held
in the model's head between turns, so a resumed or retried wave behaves exactly
like a first attempt.

ALL serialization routes through `run_log._json_default`, which handles pydantic
models (`model_dump(mode="json")`), `Path`, and sets — so the `json.dumps`-on-a-
`Company` crash that killed a live run cannot recur.

This module owns no banking logic and makes no Agent calls. Sub-agent dispatch
stays with the model, because a `Task` call is not something a Python function can
issue; :func:`run_wave` composes the whole round trip for callers that *can*
dispatch programmatically (the tests here, and Phase F's in-process transforms).
"""

from __future__ import annotations

import argparse
import json
import os
import sys
from collections.abc import Callable, Mapping, Sequence
from dataclasses import dataclass, field
from pathlib import Path

import yaml

from deal_init import load_deal_context
from plan_refs import resolve_refs, validate_plan_references
from plan_schedule import compute_waves
from run_log import (
    _json_default,
    read_stage_outputs,
    stage_dir,
    write_stage_inputs,
    write_summary,
)
from schemas import Plan

PLAN_INPUTS_NAME = "plan_inputs.json"

#: Resolved inputs are rendered into the dispatch envelope verbatim up to this
#: size, so a sub-agent can see what it was given without reading a file. Past
#: it (an analyst's pasted notes, a long filing list) the envelope names the keys
#: and points at `inputs.json`, which is written either way.
INLINE_INPUTS_MAX_CHARS = 6000


# ---------------------------------------------------------------------------
# Run-state loading
# ---------------------------------------------------------------------------
def load_plan(run_dir: Path | str) -> Plan:
    """Load + validate the frozen plan snapshot at `<run_dir>/plan.yaml`.

    Validation is two-layered: the pydantic `Plan` schema (shape), then
    `plan_refs.validate_plan_references` (reference pre-flight — every
    `$stages.<id>` names a real stage, every `$stages.<id>.<name>` a declared
    output of that stage, every `$plan_inputs.<name>` a declared plan input).
    A typo'd reference is thus rejected here, at load, rather than mid-run when
    it fails to resolve. Every entry point in this module loads through this.
    """
    path = Path(run_dir) / "plan.yaml"
    if not path.exists():
        raise FileNotFoundError(f"no plan snapshot at {path} — was the run dir created?")
    plan = Plan.model_validate(yaml.safe_load(path.read_text(encoding="utf-8")))
    validate_plan_references(plan)
    return plan


def load_plan_inputs(run_dir: Path | str) -> dict:
    """Load the analyst-collected plan inputs from `<run_dir>/plan_inputs.json`.

    Returns `{}` when the file is absent — a plan with no required inputs (or whose
    references only hit `$deal` / `$stages`) needs none, and a missing *required*
    input still surfaces as a `ReferenceResolutionError` at resolve time.
    """
    path = Path(run_dir) / PLAN_INPUTS_NAME
    if not path.exists():
        return {}
    return json.loads(path.read_text(encoding="utf-8"))


def write_plan_inputs(run_dir: Path | str, plan_inputs: Mapping) -> Path:
    """Persist the analyst-collected plan inputs for the per-wave driver to read."""
    path = Path(run_dir) / PLAN_INPUTS_NAME
    path.write_text(
        json.dumps(dict(plan_inputs), indent=2, sort_keys=True, default=_json_default) + "\n",
        encoding="utf-8",
    )
    return path


def _deal_dir_for_run(run_dir: Path) -> Path:
    """`<deal_dir>/runs/<run-id>/` → `<deal_dir>` (two levels up)."""
    return run_dir.resolve().parent.parent


def default_plugin_root() -> Path:
    """`CLAUDE_PLUGIN_ROOT`, or the plugin root inferred from this file's location."""
    env = os.environ.get("CLAUDE_PLUGIN_ROOT")
    if env:
        return Path(env)
    return Path(__file__).resolve().parent.parent  # scripts/.. == infor-beta/


def _stage_by_id(plan: Plan, stage_id: str):
    for s in plan.stages:
        if s.id == stage_id:
            return s
    raise KeyError(f"stage {stage_id!r} not found in plan")


def _wave_stage_ids(plan: Plan, wave: int) -> list[str]:
    """Return the stage ids in 1-based `wave`, validating the wave number."""
    waves = compute_waves(plan)
    if wave < 1 or wave > len(waves):
        raise IndexError(
            f"wave {wave} out of range — the plan schedules {len(waves)} wave(s) "
            f"(1..{len(waves)})"
        )
    return waves[wave - 1]


# ---------------------------------------------------------------------------
# Plan overview
# ---------------------------------------------------------------------------
@dataclass(frozen=True)
class PlanOverview:
    """Everything the conductor tells the analyst before dispatching anything."""

    deliverable_type: str
    description: str
    codename: str
    run_dir: Path
    waves: tuple[tuple[str, ...], ...]
    stage_skills: dict[str, str]
    required_plan_inputs: tuple[str, ...]
    optional_plan_inputs: tuple[str, ...]
    required_checkpoints: tuple[str, ...]

    @property
    def stage_count(self) -> int:
        return sum(len(w) for w in self.waves)

    @property
    def wave_count(self) -> int:
        return len(self.waves)

    def narration(self) -> str:
        """The one-paragraph plan summary + wave schedule, ready to post."""
        stages = ", ".join(
            f"`{sid}` (skill: `{self.stage_skills[sid]}`)" for w in self.waves for sid in w
        )
        schedule = " → ".join(
            f"[{i}] " + ", ".join(f"`{sid}`" for sid in w) + (" (parallel)" if len(w) > 1 else "")
            for i, w in enumerate(self.waves, 1)
        )
        required = ", ".join(f"`{n}`" for n in self.required_plan_inputs) or "none"
        gates = ", ".join(f"`{n}`" for n in self.required_checkpoints) or "none"
        return (
            f"Plan for `{self.deliverable_type}` has {self.stage_count} stages: {stages}. "
            f"Plan inputs required: {required}. Required checkpoints: {gates}.\n\n"
            f"{self.wave_count} waves: {schedule}."
        )


def plan_overview(run_dir: Path | str) -> PlanOverview:
    """Load the frozen plan and derive everything the analyst is told up front."""
    run_dir = Path(run_dir)
    plan = load_plan(run_dir)
    ctx = load_deal_context(_deal_dir_for_run(run_dir))
    return PlanOverview(
        deliverable_type=str(plan.deliverable_type),
        description=plan.description,
        codename=getattr(ctx, "codename", ""),
        run_dir=run_dir,
        waves=tuple(tuple(w) for w in compute_waves(plan)),
        stage_skills={s.id: s.skill for s in plan.stages},
        required_plan_inputs=tuple(s.name for s in plan.plan_inputs if s.required),
        optional_plan_inputs=tuple(s.name for s in plan.plan_inputs if not s.required),
        required_checkpoints=tuple(s.id for s in plan.stages if s.checkpoint == "required"),
    )


# ---------------------------------------------------------------------------
# Stage-envelope rendering
# ---------------------------------------------------------------------------
def _extract_template(md: str) -> str:
    """Pull the fenced template block out of `stage-envelope.md` ("## Template")."""
    marker = md.index("## Template")
    open_fence = md.index("```", marker)
    body_start = md.index("\n", open_fence) + 1
    close_fence = md.index("```", body_start)
    return md[body_start:close_fence].rstrip("\n")


def _subject_company_summary(ctx) -> str:
    sc = getattr(ctx, "subject_company", None)
    if sc is None:
        return "n/a"
    ticker = getattr(sc, "ticker", None)
    name = getattr(sc, "legal_name", "n/a")
    return f"{name} ({ticker})" if ticker else f"{name} (private)"


def _declared_outputs_block(stage) -> str:
    if not stage.outputs:
        return "- (no named outputs declared)"
    return "\n".join(f"- {o.name}: {o.type}" for o in stage.outputs)


def _resolved_inputs_block(resolved: Mapping) -> str:
    """Render the stage's resolved inputs into the prompt body.

    Inline JSON while it is small enough to be worth reading in the prompt;
    otherwise the key list plus a pointer at `inputs.json`, which is written
    either way. The point of inlining is that nothing about the handoff depends
    on state surviving between the sub-agent's tool calls.
    """
    if not resolved:
        return "```json\n{}\n```\n\n(this stage declares no inputs)"
    body = json.dumps(dict(resolved), indent=2, sort_keys=True, default=_json_default)
    if len(body) <= INLINE_INPUTS_MAX_CHARS:
        return f"```json\n{body}\n```"
    keys = ", ".join(f"`{k}`" for k in sorted(resolved))
    return f"(too large to inline — {len(body):,} characters across keys: {keys}. Read the file.)"


def render_stage_envelope(
    plan: Plan,
    stage,
    *,
    ctx,
    run_dir: Path,
    plugin_root: Path,
    resolved_inputs: Mapping,
) -> str | None:
    """Render the conductor stage envelope for one stage (absolute paths baked in).

    Returns the rendered prompt, or None if the envelope template can't be found
    (preparation still writes inputs.json — only the printed prompt is skipped).
    """
    template_path = plugin_root / "skills" / "conductor" / "references" / "stage-envelope.md"
    if not template_path.exists():
        return None
    template = _extract_template(template_path.read_text(encoding="utf-8"))
    sdir = stage_dir(run_dir, stage.id)
    deal_dir = _deal_dir_for_run(run_dir)
    subs = {
        "{{stage_id}}": stage.id,
        "{{deliverable_type}}": str(plan.deliverable_type),
        "{{codename}}": getattr(ctx, "codename", ""),
        "{{deal_dir}}": str(deal_dir),
        "{{subject_company_summary}}": _subject_company_summary(ctx),
        "{{plugin_root}}": str(plugin_root),
        "{{skill_name}}": stage.skill,
        "{{stage_inputs_path}}": str((sdir / "inputs.json").resolve()),
        "{{stage_outputs_path}}": str((sdir / "outputs.json").resolve()),
        "{{declared_outputs_block}}": _declared_outputs_block(stage),
        "{{resolved_inputs_block}}": _resolved_inputs_block(resolved_inputs),
    }
    rendered = template
    for token, value in subs.items():
        rendered = rendered.replace(token, value)
    return rendered


# ---------------------------------------------------------------------------
# prepare_wave
# ---------------------------------------------------------------------------
@dataclass(frozen=True)
class PreparedStage:
    """One stage of a wave, resolved and ready to dispatch."""

    stage_id: str
    skill: str
    checkpoint: str
    inputs: dict
    inputs_path: Path
    outputs_path: Path
    prompt: str | None  # the rendered stage envelope — pass to `Task` verbatim


@dataclass(frozen=True)
class WaveDispatch:
    """What the model needs in order to issue one wave's `Task` calls."""

    wave: int
    total_waves: int
    stages: tuple[PreparedStage, ...]

    @property
    def prompts(self) -> list[str]:
        """The rendered envelopes, in dispatch order. Issue them in ONE message."""
        return [s.prompt for s in self.stages if s.prompt is not None]

    def narration(self) -> str:
        ids = ", ".join(f"`{s.stage_id}`" for s in self.stages)
        parallel = " (dispatched in parallel)" if len(self.stages) > 1 else ""
        return f"Wave {self.wave} of {self.total_waves}: {ids}{parallel}."


def prepare_wave(
    run_dir: Path | str,
    wave: int,
    *,
    plugin_root: Path | str | None = None,
) -> WaveDispatch:
    """Resolve references and write inputs.json for every stage in `wave` (1-based).

    Reads prior waves' outputs so `$stages.*` references resolve; lets unsupplied
    optional plan inputs resolve to None (via `optional_plan_inputs`); writes each
    stage's inputs.json through `run_log.write_stage_inputs` (pydantic/`Path`-safe);
    and renders each stage's dispatch envelope with the absolute handoff paths and
    the resolved inputs baked into the prompt body.
    """
    run_dir = Path(run_dir)
    plan = load_plan(run_dir)
    ctx = load_deal_context(_deal_dir_for_run(run_dir))
    plan_inputs = load_plan_inputs(run_dir)
    optional = {spec.name for spec in plan.plan_inputs if not spec.required}
    proot = Path(plugin_root) if plugin_root is not None else default_plugin_root()

    waves = compute_waves(plan)
    stage_ids = _wave_stage_ids(plan, wave)

    # Every `$stages.*` reference a wave member needs was produced by an earlier
    # wave — the scheduler guarantees it.
    stage_outputs: dict[str, dict] = {}
    for earlier in waves[: wave - 1]:
        for sid in earlier:
            stage_outputs[sid] = read_stage_outputs(run_dir, sid)

    prepared: list[PreparedStage] = []
    for sid in stage_ids:
        stage = _stage_by_id(plan, sid)
        resolved = resolve_refs(
            stage.inputs,
            plan_inputs=plan_inputs,
            deal_context=ctx,
            stage_outputs=stage_outputs,
            optional_plan_inputs=optional,
        )
        inputs_path = write_stage_inputs(run_dir, sid, resolved)
        prepared.append(
            PreparedStage(
                stage_id=sid,
                skill=stage.skill,
                checkpoint=str(stage.checkpoint),
                inputs=resolved,
                inputs_path=inputs_path,
                outputs_path=(stage_dir(run_dir, sid) / "outputs.json").resolve(),
                prompt=render_stage_envelope(
                    plan,
                    stage,
                    ctx=ctx,
                    run_dir=run_dir,
                    plugin_root=proot,
                    resolved_inputs=resolved,
                ),
            )
        )
    return WaveDispatch(wave=wave, total_waves=len(waves), stages=tuple(prepared))


# ---------------------------------------------------------------------------
# complete_wave
# ---------------------------------------------------------------------------
def _missing_declared_outputs(stage, outputs) -> list[str]:
    """Declared output NAMES absent from a stage's outputs.json.

    Presence-only: a declared output carrying `null` is legal (the v0.5.21
    contract requires e.g. ltm-metrics to emit null — never omit —
    `ltm_revenue` / `ltm_adj_ebitda`), and extra undeclared keys are allowed.
    The `type` labels are not checked. A non-dict outputs.json (e.g. a bare
    JSON list) misses every declared name.
    """
    present = outputs.keys() if isinstance(outputs, dict) else ()
    return [spec.name for spec in stage.outputs if spec.name not in present]


@dataclass(frozen=True)
class StageResult:
    """One stage's collected + validated outputs."""

    stage_id: str
    skill: str
    checkpoint: str
    ok: bool
    outputs: dict | None = None
    error: str | None = None


#: The `required`-checkpoint dialog, code-owned like every other analyst-facing
#: question since v0.5.27 — the same two options on every gate of every plan, so
#: an approval means the same thing every time.
APPROVE_LABEL = "Approve — continue the run"
HALT_LABEL = "Halt the run"


@dataclass(frozen=True)
class Checkpoint:
    """What to surface at a wave boundary for one stage, per its checkpoint mode."""

    stage_id: str
    mode: str
    surface: str
    question: dict | None = None  # `AskUserQuestion` payload — `required` only
    fallback_prompt: str | None = None  # plain text, for surfaces without the tool


@dataclass(frozen=True)
class WaveOutcome:
    """The result of one wave, plus the conversation the model owes the analyst."""

    wave: int
    total_waves: int
    results: tuple[StageResult, ...]
    checkpoints: tuple[Checkpoint, ...]
    is_final: bool
    failures: tuple[str, ...] = field(default_factory=tuple)

    @property
    def ok(self) -> bool:
        return not self.failures

    @property
    def gate(self) -> Checkpoint | None:
        """The `required` checkpoint to put to the analyst, if this wave has one.

        Wave-boundary semantics: a gate holds the *downstream* waves, never its
        own wave-mates, which have already run by the time it is evaluated.
        """
        return next((c for c in self.checkpoints if c.mode == "required"), None)

    @property
    def halt(self) -> bool:
        """True when the run must stop here regardless of the analyst's answer."""
        return bool(self.failures)

    def narration(self) -> str:
        lines = [c.surface for c in self.checkpoints if c.surface]
        return "\n\n".join(lines)


def _format_outputs_inline(outputs: Mapping | None) -> str:
    if not outputs:
        return "(none)"
    return ", ".join(f"`{k}={v}`" for k, v in sorted(outputs.items()))


def _format_outputs_bullets(outputs: Mapping | None) -> str:
    if not outputs:
        return "- (no outputs declared)"
    return "\n".join(f"- `{k}`: {v}" for k, v in sorted(outputs.items()))


def _checkpoint_for(result: StageResult, *, is_final: bool) -> Checkpoint:
    """Build one stage's analyst-facing payload, per `references/checkpoint-behaviour.md`."""
    if not result.ok:
        # A failed stage always surfaces, whatever its mode — the run stops here,
        # and `silent` was never meant to hide a failure.
        return Checkpoint(
            stage_id=result.stage_id,
            mode=result.checkpoint,
            surface=(
                f"Stage `{result.stage_id}` (`{result.skill}`) FAILED: {result.error}\n"
                f"The run stops here; the partial run on disk is preserved."
            ),
        )

    if result.checkpoint == "silent":
        return Checkpoint(stage_id=result.stage_id, mode="silent", surface="")

    if result.checkpoint == "required":
        held = "delivery" if is_final else "the remaining waves"
        surface = (
            f"Stage `{result.stage_id}` (`{result.skill}`) finished. Outputs:\n"
            f"{_format_outputs_bullets(result.outputs)}\n\n"
            f"Review the file(s) above, then answer the approval dialog."
        )
        return Checkpoint(
            stage_id=result.stage_id,
            mode="required",
            surface=surface,
            question={
                "question": (
                    f"Stage `{result.stage_id}` finished and is holding {held}. "
                    f"Approve it, or halt the run?"
                ),
                "header": "Checkpoint",
                "multiSelect": False,
                "options": [
                    {
                        "label": APPROVE_LABEL,
                        "description": f"Continue — release {held} of this plan.",
                    },
                    {
                        "label": HALT_LABEL,
                        "description": (
                            "Stop here. The partial run is preserved on disk; "
                            "re-running starts a new run id."
                        ),
                    },
                ],
            },
            fallback_prompt=(
                f"{surface}\n\nReply `approve` to continue the run, or `stop` to halt it."
            ),
        )

    return Checkpoint(
        stage_id=result.stage_id,
        mode="informational",
        surface=(
            f"Stage `{result.stage_id}` (`{result.skill}`) finished. "
            f"Outputs: {_format_outputs_inline(result.outputs)}. Proceeding."
        ),
    )


def complete_wave(run_dir: Path | str, wave: int) -> WaveOutcome:
    """Read + validate each stage's outputs.json for `wave`, and build the checkpoints.

    A stage is reported `ok=False` — and the run must stop rather than starting
    the next wave — when it wrote no outputs.json, wrote one that isn't valid JSON
    (a sub-agent truncating the file must not crash the driver), wrote one
    carrying an `error` key, or omitted a declared output name (null values pass,
    extras allowed).
    """
    run_dir = Path(run_dir)
    plan = load_plan(run_dir)
    waves = compute_waves(plan)
    stage_ids = _wave_stage_ids(plan, wave)
    is_final = wave == len(waves)

    results: list[StageResult] = []
    for sid in stage_ids:
        stage = _stage_by_id(plan, sid)
        common = {"stage_id": sid, "skill": stage.skill, "checkpoint": str(stage.checkpoint)}
        try:
            outputs = read_stage_outputs(run_dir, sid)
        except FileNotFoundError as exc:
            results.append(StageResult(**common, ok=False, error=str(exc)))
            continue
        except json.JSONDecodeError as exc:
            results.append(
                StageResult(
                    **common,
                    ok=False,
                    error=f"stage {sid!r} wrote malformed (non-JSON) outputs.json: {exc}",
                )
            )
            continue
        if isinstance(outputs, dict) and "error" in outputs:
            results.append(StageResult(**common, ok=False, outputs=outputs, error=outputs["error"]))
            continue
        missing = _missing_declared_outputs(stage, outputs)
        if missing:
            results.append(
                StageResult(
                    **common,
                    ok=False,
                    outputs=outputs,
                    error=(
                        f"stage {sid!r} outputs.json is missing declared output(s): "
                        f"{', '.join(missing)} (null values are fine — omitting the key is not)"
                    ),
                )
            )
            continue
        results.append(StageResult(**common, ok=True, outputs=outputs))

    return WaveOutcome(
        wave=wave,
        total_waves=len(waves),
        results=tuple(results),
        checkpoints=tuple(_checkpoint_for(r, is_final=is_final) for r in results),
        is_final=is_final,
        failures=tuple(f"{r.stage_id}: {r.error}" for r in results if not r.ok),
    )


# ---------------------------------------------------------------------------
# run_wave — the composed round trip
# ---------------------------------------------------------------------------
def run_wave(
    run_dir: Path | str,
    wave: int,
    dispatch: Callable[[WaveDispatch], object],
    *,
    plugin_root: Path | str | None = None,
) -> WaveOutcome:
    """Prepare a wave, dispatch it, collect it — the whole trip in one call.

    `dispatch` receives the :class:`WaveDispatch` and must not return until every
    stage in the wave has finished writing its outputs.json.

    The conductor *skill* cannot pass a callback: its "dispatch" is a set of `Task`
    tool calls, which no Python function can issue. It calls the two halves
    instead — :func:`prepare_wave`, then its `Task` calls, then
    :func:`complete_wave` — which is this same sequence with the model standing in
    for the callback. This entry point is for callers that *can* dispatch
    programmatically: the tests here, and Phase F's in-process transform stages.
    """
    prepared = prepare_wave(run_dir, wave, plugin_root=plugin_root)
    dispatch(prepared)
    return complete_wave(run_dir, wave)


# ---------------------------------------------------------------------------
# Run summary
# ---------------------------------------------------------------------------
_ARTEFACT_SUFFIXES = {".pptx", ".xlsx", ".docx", ".pdf", ".png"}


def _looks_like_artefact(value: object) -> bool:
    return isinstance(value, str) and Path(value).suffix.lower() in _ARTEFACT_SUFFIXES


def render_run_summary(run_dir: Path | str, *, notes: Sequence[str] = ()) -> str:
    """Compose the mechanical part of the end-of-run summary as markdown.

    Every stage across every wave, its status, its outputs, and the artefact paths
    the analyst can open right now. `notes` are appended verbatim under "Manual
    next steps" — that is where the model puts what only it knows (a sub-skill's
    "refresh the Capital IQ connector in the cap table", an abort reason, a caveat).
    """
    run_dir = Path(run_dir)
    plan = load_plan(run_dir)
    ctx = load_deal_context(_deal_dir_for_run(run_dir))

    lines = [
        f"# Run summary — {plan.deliverable_type}",
        "",
        f"- **Deal**: {getattr(ctx, 'codename', '')}",
        f"- **Run id**: {run_dir.name}",
        f"- **Run directory**: {run_dir}",
        f"- **Plan**: {plan.description}",
        "",
        "## Stages",
        "",
    ]
    artefacts: list[str] = []
    for wave in range(1, len(compute_waves(plan)) + 1):
        for r in complete_wave(run_dir, wave).results:
            status = "ok" if r.ok else f"FAILED — {r.error}"
            lines.append(f"- **wave {wave}** `{r.stage_id}` (`{r.skill}`): {status}")
            for key, value in sorted((r.outputs or {}).items()):
                lines.append(f"  - `{key}`: {value}")
                if _looks_like_artefact(value):
                    artefacts.append(str(value))

    lines += ["", "## Artefacts", ""]
    lines += [f"- {a}" for a in dict.fromkeys(artefacts)] or ["- (none produced)"]
    lines += ["", "## Manual next steps", ""]
    lines += [f"- {n}" for n in notes] or ["- (none)"]
    lines += ["", f"Full per-stage detail: `{run_dir / 'stages'}`.", ""]
    return "\n".join(lines)


def write_run_summary(run_dir: Path | str, *, notes: Sequence[str] = ()) -> Path:
    """Render the run summary and persist it as `<run_dir>/summary.md`."""
    return write_summary(run_dir, render_run_summary(run_dir, notes=notes))


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------
def _print_prepared(dispatch: WaveDispatch) -> None:
    print(dispatch.narration())
    print()
    for entry in dispatch.stages:
        print(f"=== stage `{entry.stage_id}` (skill: {entry.skill}) ===")
        print(f"inputs.json : {entry.inputs_path}")
        print(f"outputs.json: {entry.outputs_path}")
        if entry.prompt is not None:
            print("--- dispatch envelope ---")
            print(entry.prompt)
        else:
            print("(stage-envelope template not found; inputs.json written anyway)")
        print()


def _print_outcome(outcome: WaveOutcome) -> int:
    for r in outcome.results:
        if r.ok:
            keys = sorted(r.outputs.keys()) if isinstance(r.outputs, dict) else []
            print(f"[ok]   {r.stage_id}: outputs {keys}")
        else:
            print(f"[FAIL] {r.stage_id}: {r.error}")
    print()
    for c in outcome.checkpoints:
        if not c.surface:
            continue
        print(f"--- checkpoint `{c.stage_id}` ({c.mode}) ---")
        print(c.surface)
        if c.question is not None:
            print("--- AskUserQuestion payload ---")
            print(json.dumps([c.question], indent=2, ensure_ascii=False))
        print()
    return 1 if outcome.failures else 0


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(prog="conductor", description=__doc__)
    sub = parser.add_subparsers(dest="command", required=True)

    p_plan = sub.add_parser("plan", help="print the plan summary + wave schedule")
    p_plan.add_argument("run_dir")

    p_prep = sub.add_parser("prepare-wave", help="resolve + write inputs.json for a wave")
    p_prep.add_argument("run_dir")
    p_prep.add_argument("wave", type=int)
    p_prep.add_argument("--plugin-root", default=None)

    p_done = sub.add_parser("complete-wave", help="read + validate a wave's outputs")
    p_done.add_argument("run_dir")
    p_done.add_argument("wave", type=int)

    p_sum = sub.add_parser("summary", help="write + print the end-of-run summary")
    p_sum.add_argument("run_dir")
    p_sum.add_argument("--note", action="append", default=[])

    args = parser.parse_args(argv)

    if args.command == "plan":
        print(plan_overview(args.run_dir).narration())
        return 0
    if args.command == "prepare-wave":
        _print_prepared(prepare_wave(args.run_dir, args.wave, plugin_root=args.plugin_root))
        return 0
    if args.command == "complete-wave":
        return _print_outcome(complete_wave(args.run_dir, args.wave))
    if args.command == "summary":
        path = write_run_summary(args.run_dir, notes=args.note)
        print(path.read_text(encoding="utf-8"))
        return 0
    return 2


if __name__ == "__main__":  # pragma: no cover
    sys.exit(main())
