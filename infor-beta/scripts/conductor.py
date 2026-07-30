"""The conductor, as code — plan load, wave prep, collect, checkpoints, summary.

Phase E finishes the trip `conductor_cli` started. Everything mechanical about
running a plan is a function with a return value here, not a paragraph of SKILL.md
prose the model may skip on turn 40 of a long run. What is left to the model is
exactly four things:

1. the intake conversation — since v0.5.51 a single `AskUserQuestion` call on a
   slash-command run (`deck_spec.render_run_dialogs`, deal-init's questions
   merged with the deliverable's), then one attachment request and one echo,
2. issuing the `Task` calls :func:`prepare_wave` hands back — for the **judgment**
   stages only, since Phase F the driver runs the deterministic ones itself
   (:func:`run_transforms`; the classification lives in `stage_transforms`),
3. the checkpoint conversation :func:`complete_wave` hands back, and
4. the end-of-run summary, off the skeleton :func:`render_run_summary` returns.

    conductor.py plan           <run_dir>         # deliverable, stages, wave schedule
    conductor.py prepare-wave   <run_dir> <wave>  # resolve + write inputs, print envelopes
    conductor.py run-transforms <run_dir> <wave>  # execute the wave's in-process stages
    conductor.py complete-wave  <run_dir> <wave>  # read + validate outputs, print checkpoints
    conductor.py summary        <run_dir> --notes-file <file>   # write summary.md, print it

`summary` takes its notes from a **file** (`-` for stdin), not from an argv string:
notes are dollar-dense by nature and an argv string reaches Python through a shell,
which expands every `$34` / `$MM` in it to nothing. That silently destroyed every
figure in one run's notes. `--note` survives for a one-liner, and either way
`suspect_currency` warns loudly on a note that looks mangled rather than writing it
quietly into the analyst-facing artefact.

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
dispatch programmatically (the tests here).

Phase F took the deterministic stages out of that loop entirely. A stage whose
skill is in `stage_transforms.TRANSFORMS` is executed here, in-process, by
:func:`run_transforms` — no envelope, no `Task`, no sub-agent context. Everything
else about it is unchanged: it still writes `inputs.json` and `outputs.json`, still
carries its `$stages` edges into the wave schedule, and still reports through
:func:`complete_wave`, so a `required` gate downstream of it behaves identically.
"""

from __future__ import annotations

import argparse
import contextlib
import json
import os
import re
import sys
import traceback
from collections.abc import Callable, Mapping, Sequence
from dataclasses import dataclass, field
from io import StringIO
from pathlib import Path

import yaml

import stage_transforms
from deal_init import load_deal_context
from plan_refs import resolve_refs, validate_plan_references
from plan_schedule import compute_waves
from run_log import (
    _json_default,
    read_stage_outputs,
    stage_dir,
    write_stage_inputs,
    write_stage_log,
    write_summary,
)
from schemas import Plan
from stage_io import stage_io

PLAN_INPUTS_NAME = "plan_inputs.json"

#: A stage's execution kind, decided by `stage_transforms.TRANSFORMS` and nothing
#: else. `transform` means the driver calls a function; `judgment` means the model
#: dispatches a sub-agent. The plan YAML carries no annotation for this — one
#: registry, so a plan and the classification cannot disagree.
KIND_TRANSFORM = "transform"
KIND_JUDGMENT = "judgment"

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
    stage_kinds: dict[str, str]
    required_plan_inputs: tuple[str, ...]
    optional_plan_inputs: tuple[str, ...]
    required_checkpoints: tuple[str, ...]

    @property
    def stage_count(self) -> int:
        return sum(len(w) for w in self.waves)

    @property
    def wave_count(self) -> int:
        return len(self.waves)

    @property
    def transform_stages(self) -> tuple[str, ...]:
        """Stage ids the driver runs in-process, in plan order."""
        return tuple(
            sid for w in self.waves for sid in w if self.stage_kinds[sid] == KIND_TRANSFORM
        )

    @property
    def judgment_stages(self) -> tuple[str, ...]:
        """Stage ids the model dispatches as sub-agents, in plan order."""
        return tuple(sid for w in self.waves for sid in w if self.stage_kinds[sid] == KIND_JUDGMENT)

    @property
    def dispatch_count(self) -> int:
        """How many `Task` calls the whole run costs. Derived, never written down."""
        return len(self.judgment_stages)

    def narration(self) -> str:
        """The one-paragraph plan summary + wave schedule, ready to post."""
        stages = ", ".join(
            f"`{sid}` (skill: `{self.stage_skills[sid]}`"
            + (", in-process" if self.stage_kinds[sid] == KIND_TRANSFORM else "")
            + ")"
            for w in self.waves
            for sid in w
        )
        schedule = " → ".join(
            f"[{i}] " + ", ".join(f"`{sid}`" for sid in w) + (" (parallel)" if len(w) > 1 else "")
            for i, w in enumerate(self.waves, 1)
        )
        required = ", ".join(f"`{n}`" for n in self.required_plan_inputs) or "none"
        gates = ", ".join(f"`{n}`" for n in self.required_checkpoints) or "none"
        return (
            f"Plan for `{self.deliverable_type}` has {self.stage_count} stages: {stages}. "
            f"{self.dispatch_count} are dispatched as sub-agents; "
            f"{len(self.transform_stages)} run in-process. "
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
        stage_kinds={
            s.id: KIND_TRANSFORM if stage_transforms.is_transform(s.skill) else KIND_JUDGMENT
            for s in plan.stages
        },
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
    """One stage of a wave, resolved and ready to run.

    `kind` decides *how*: a `judgment` stage carries a rendered `prompt` for the
    model to pass to `Task`; a `transform` carries `prompt=None` because there is
    nothing to dispatch — :func:`run_transforms` calls it directly.
    """

    stage_id: str
    skill: str
    checkpoint: str
    kind: str
    inputs: dict
    inputs_path: Path
    outputs_path: Path
    prompt: str | None  # the rendered stage envelope — pass to `Task` verbatim


@dataclass(frozen=True)
class WaveDispatch:
    """One wave, resolved: what the driver runs and what the model dispatches."""

    wave: int
    total_waves: int
    stages: tuple[PreparedStage, ...]
    plugin_root: Path

    @property
    def transforms(self) -> tuple[PreparedStage, ...]:
        """The stages :func:`run_transforms` executes in-process."""
        return tuple(s for s in self.stages if s.kind == KIND_TRANSFORM)

    @property
    def judgment(self) -> tuple[PreparedStage, ...]:
        """The stages the model dispatches as sub-agents."""
        return tuple(s for s in self.stages if s.kind == KIND_JUDGMENT)

    @property
    def prompts(self) -> list[str]:
        """The rendered envelopes, in dispatch order. Issue them in ONE message.

        Judgment stages only — a transform has no envelope, so a wave of nothing
        but transforms yields an empty list and the model issues no `Task` at all.
        """
        return [s.prompt for s in self.judgment if s.prompt is not None]

    def narration(self) -> str:
        parts = []
        if self.judgment:
            ids = ", ".join(f"`{s.stage_id}`" for s in self.judgment)
            parts.append(f"{ids} dispatched" + (" in parallel" if len(self.judgment) > 1 else ""))
        if self.transforms:
            ids = ", ".join(f"`{s.stage_id}`" for s in self.transforms)
            parts.append(f"{ids} run in-process")
        return f"Wave {self.wave} of {self.total_waves}: {'; '.join(parts)}."


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
    and renders each **judgment** stage's dispatch envelope with the absolute
    handoff paths and the resolved inputs baked into the prompt body.

    A transform stage is prepared identically — same resolution, same inputs.json —
    and simply carries no envelope. Preparing runs nothing: :func:`run_transforms`
    is the separate, explicit step, so "resolve the wave" stays free of side
    effects and a forgotten call halts loudly in :func:`complete_wave` (the stage
    wrote no outputs.json) rather than silently skipping work.
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
        transform = stage_transforms.is_transform(stage.skill)
        prepared.append(
            PreparedStage(
                stage_id=sid,
                skill=stage.skill,
                checkpoint=str(stage.checkpoint),
                kind=KIND_TRANSFORM if transform else KIND_JUDGMENT,
                inputs=resolved,
                inputs_path=inputs_path,
                outputs_path=(stage_dir(run_dir, sid) / "outputs.json").resolve(),
                prompt=None
                if transform
                else render_stage_envelope(
                    plan,
                    stage,
                    ctx=ctx,
                    run_dir=run_dir,
                    plugin_root=proot,
                    resolved_inputs=resolved,
                ),
            )
        )
    return WaveDispatch(
        wave=wave, total_waves=len(waves), stages=tuple(prepared), plugin_root=proot
    )


# ---------------------------------------------------------------------------
# run_transforms — the driver's own half of a wave
# ---------------------------------------------------------------------------
@dataclass(frozen=True)
class TransformResult:
    """One in-process stage's execution, before `complete_wave` validates it."""

    stage_id: str
    skill: str
    ok: bool
    outputs: dict | None = None
    error: str | None = None
    #: `stages/<id>/log.txt` — the transform's own stdout + stderr. `None` only when
    #: the run directory refused the write.
    log_path: Path | None = None


class _Tee:
    """Fan one stream to the real console and to a stage's transcript.

    A tee rather than a redirect because both readers are real. The analyst (and
    the model driving the run) needs the converge loop's progress *live* — a deck
    that is 40s into a repair pass and a deck that has hung look identical
    otherwise — and the run directory needs it *afterwards*: the only record that a
    real run's repair loop shrank slide 7 to 85%, or that it finished on "18 blocking
    / 4 advisory finding(s)", was stdout in a shell nobody kept.

    Unknown attributes delegate to the wrapped stream, so anything that probes
    `encoding` / `isatty()` / `fileno()` sees the console it would have seen.
    """

    def __init__(self, buffer: StringIO, stream) -> None:
        self._buffer = buffer
        self._stream = stream

    def write(self, text: str) -> int:
        self._buffer.write(text)
        with contextlib.suppress(Exception):  # a closed console must not fail a stage
            self._stream.write(text)
        return len(text)

    def flush(self) -> None:
        with contextlib.suppress(Exception):
            self._stream.flush()

    def __getattr__(self, name: str):
        return getattr(self._stream, name)


def _write_transform_log(io, stage: PreparedStage, transcript: StringIO, *, error: str | None) -> Path | None:
    """Persist a transform's output as `stages/<id>/log.txt` and return the path.

    The **same file** `write_stage_log` puts a dispatched stage's transcript in. A
    pitch run's directory used to hold eight transcripts and four silences, and the
    four were the deterministic stages — the ones whose whole job is making
    decisions a reader might want to check. Written unconditionally, including for a
    stage that printed nothing: an empty transcript is the fact "it ran and said
    nothing", where an absent file is "we have no idea".

    Writing the log must never be what fails a stage — the deal directory can be a
    cloud-synced mount — so an `OSError` costs the transcript and nothing else.
    """
    body = transcript.getvalue()
    header = f"=== transform `{stage.stage_id}` (skill: {stage.skill}) ===\n"
    if error is not None:
        if body and not body.endswith("\n"):
            body += "\n"
        body += f"\n=== FAILED: {error} ===\n{traceback.format_exc()}"
    try:
        return write_stage_log(io.run_dir, io.stage_id, header + body)
    except OSError as exc:  # noqa: BLE001 — reported, never fatal
        print(f"conductor: could not write {stage.stage_id}'s transcript: {exc}", file=sys.stderr)
        return None


def _log_tail(path: Path | None, *, lines: int = 12) -> str:
    """The last few lines of a transform's transcript, for the CLI's failure print."""
    if path is None:
        return ""
    try:
        tail = path.read_text(encoding="utf-8").splitlines()[-lines:]
    except OSError:
        return ""
    return "\n".join(f"    | {line}" for line in tail)


def run_transforms(dispatch: WaveDispatch) -> tuple[TransformResult, ...]:
    """Execute every transform stage in `dispatch`, writing each outputs.json.

    Transforms in one wave are independent by construction — the scheduler put
    them there because nothing in the wave references anything else in it — so
    they run in declaration order and the order does not matter.

    **A raising transform is a stage failure, not a crash.** The exception is
    written as `{"error": …}` to the stage's outputs.json, exactly as a sub-agent's
    `io.fail(...)` would, so `complete_wave` reports `ok=False`, the checkpoint
    surfaces the reason whatever the stage's mode, and the run halts before the
    next wave. That is what keeps a `DeckNotConvergedError` — the one failure the
    deck stage is expected to be able to produce — a legible halt with the shape
    and the depth in the message rather than a traceback out of the driver.

    Reads each stage's inputs back off disk through `stage_io.stage_io`, the same
    entry point the dispatched form used, so a transform sees byte-identical
    inputs to what a sub-agent would have read.

    **Every transform leaves a transcript**, at `stages/<id>/log.txt` — the same
    file `write_stage_log` puts a sub-agent's in. Until v0.5.51 the eight judgment
    stages of a pitch run each had one and the four transforms had none, so the only
    record of the decisions the converge loop *makes* — which shape it shrank, to
    what scale, how many findings survived — was stdout in a shell. `converge_deck`
    prints to stderr by default and takes no other route, so the capture is a tee
    around the call (:class:`_Tee`) rather than anything the transforms know about:
    no contract changes, and a transform that starts printing more is covered by
    construction.
    """
    results: list[TransformResult] = []
    for stage in dispatch.transforms:
        io = stage_io(
            [
                "run_transforms",
                str(dispatch.plugin_root),
                str(stage.inputs_path),
                str(stage.outputs_path),
            ]
        )
        transcript = StringIO()
        try:
            with (
                contextlib.redirect_stdout(_Tee(transcript, sys.stdout)),
                contextlib.redirect_stderr(_Tee(transcript, sys.stderr)),
            ):
                outputs = stage_transforms.run_transform(stage.skill, io)
        except Exception as exc:  # noqa: BLE001 — every failure becomes a stage failure
            error = f"{type(exc).__name__}: {exc}"
            io.fail(error)
            results.append(
                TransformResult(
                    stage_id=stage.stage_id,
                    skill=stage.skill,
                    ok=False,
                    error=error,
                    log_path=_write_transform_log(io, stage, transcript, error=error),
                )
            )
            continue
        io.write(outputs)
        results.append(
            TransformResult(
                stage_id=stage.stage_id,
                skill=stage.skill,
                ok=True,
                outputs=dict(outputs),
                log_path=_write_transform_log(io, stage, transcript, error=None),
            )
        )
    return tuple(results)


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


def _declared_paths(stage, outputs) -> tuple[tuple[str, str], ...]:
    """`(name, value)` for every declared **Path** output the stage actually produced.

    Declaration order, and a `null` is dropped — the pitch plan's `ownership` stage
    legitimately emits `workbook_path: null`, and telling an analyst to open `None`
    is worse than saying nothing.

    Read off the plan's own `type:` labels rather than by sniffing values for a
    suffix, because the point is what the stage *declared*: `vision_review_path` is a
    `.md` file and `chart_qa_dir` a directory, and both are things to open.
    """
    if not isinstance(outputs, dict):
        return ()
    return tuple(
        (spec.name, str(outputs[spec.name]))
        for spec in stage.outputs
        if spec.type == "Path" and outputs.get(spec.name) is not None
    )


@dataclass(frozen=True)
class StageResult:
    """One stage's collected + validated outputs."""

    stage_id: str
    skill: str
    checkpoint: str
    ok: bool
    outputs: dict | None = None
    error: str | None = None
    #: The declared `Path` outputs, `(name, value)` in declaration order. Named in
    #: the checkpoint surface so a file the analyst is meant to read cannot go
    #: unmentioned — see :func:`_paths_block`.
    paths: tuple[tuple[str, str], ...] = ()


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

    @property
    def path_outputs(self) -> tuple[tuple[str, str, str], ...]:
        """`(stage_id, output name, value)` for every declared `Path` this wave wrote."""
        return tuple(
            (r.stage_id, name, value) for r in self.results for name, value in r.paths
        )

    def narration(self) -> str:
        """The wave boundary, ready to post — **including every file it produced**.

        Each stage's `Path` outputs are named here (`_paths_block`), so posting this
        verbatim is what tells the analyst where the deck, the written vision review
        and the reviews are. That used to be an instruction in the conductor's
        SKILL.md, and an instruction that has to be remembered on turn 40 of a long
        run is one that eventually is not.
        """
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


def _paths_block(result: StageResult) -> str:
    """The stage's `Path` outputs, as an "Open:" list, or "" when it produced none.

    The driver names these because the alternative was an instruction the model had
    to remember. The conductor's Step 5 used to say "name that path in the surface
    so the analyst can open it while the run continues" — and on a real pitch run
    the wave-5 boundary named no path at all, so the deck's written vision review
    existed, was 19 KB, and was never mentioned to the analyst who was meant to read
    it. The driver already holds every stage's outputs and the plan's declared types,
    so nothing about that needed remembering.
    """
    if not result.paths:
        return ""
    return "Open:\n" + "\n".join(f"- `{name}`: {value}" for name, value in result.paths)


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

    paths = _paths_block(result)

    if result.checkpoint == "required":
        held = "delivery" if is_final else "the remaining waves"
        surface = (
            f"Stage `{result.stage_id}` (`{result.skill}`) finished. Outputs:\n"
            f"{_format_outputs_bullets(result.outputs)}\n\n"
            + (f"{paths}\n\n" if paths else "")
            + "Review the file(s) above, then answer the approval dialog."
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
            + (f"\n{paths}" if paths else "")
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
        results.append(
            StageResult(
                **common, ok=True, outputs=outputs, paths=_declared_paths(stage, outputs)
            )
        )

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
    """Prepare a wave, run its transforms, dispatch the rest, collect it.

    `dispatch` receives the :class:`WaveDispatch` and must not return until every
    **judgment** stage in the wave has finished writing its outputs.json. The
    transforms have already run by then — :func:`run_transforms` is called first,
    so a wave of nothing but transforms passes the callback an empty
    `dispatch.prompts` and needs nothing from it.

    The conductor *skill* cannot pass a callback: its "dispatch" is a set of `Task`
    tool calls, which no Python function can issue. It calls the three parts
    instead — :func:`prepare_wave`, :func:`run_transforms`, its `Task` calls, then
    :func:`complete_wave` — which is this same sequence with the model standing in
    for the callback.
    """
    prepared = prepare_wave(run_dir, wave, plugin_root=plugin_root)
    run_transforms(prepared)
    dispatch(prepared)
    return complete_wave(run_dir, wave)


# ---------------------------------------------------------------------------
# Run summary
# ---------------------------------------------------------------------------
_ARTEFACT_SUFFIXES = {".pptx", ".xlsx", ".docx", ".pdf", ".png"}


def _looks_like_artefact(value: object) -> bool:
    return isinstance(value, str) and Path(value).suffix.lower() in _ARTEFACT_SUFFIXES


#: A currency letter sitting **directly** on a digit — `US50.0`, `C4.97` — with the
#: `$` that belongs between them gone. In banking notes that never occurs by
#: accident, and it has exactly one common cause: the note travelled through a
#: double-quoted shell command, where `$34` / `$150` / `$MM` are variables that
#: expand to nothing. The digits are captured so the warning can quote what it found.
CORRUPT_CURRENCY = re.compile(r"\b(?:US|C)\d[\d.,]*")


def suspect_currency(text: str) -> tuple[str, ...]:
    """Fragments of `text` that look like a dollar sign was eaten, in order.

    A **warning**, not a verdict, and not a complete one either. It catches five of
    the six mangled fragments in the run that prompted it — "the C4.97 share price",
    "+US50.0MM", "-US6.7MM", "+US3.3MM", "~C17MM", "~US00MM+" — and misses the sixth,
    "an All figures in C footnote", because `C$MM` collapses to a letter with no digit
    after it and nothing distinguishes that from prose. It can also be wrong in the
    other direction: `C4` is how one refers to a cell. Both are worth it, because the
    failure is silent and lands in the artefact the analyst reads.

    Deliberately not a guess at the repair: `US50.0MM` could have been `US$50.0MM`
    or `US$150.0MM` (it was the latter), and only the note's author knows which.
    """
    return tuple(CORRUPT_CURRENCY.findall(text or ""))


def corrupted_notes(
    notes: Sequence[str],
) -> tuple[tuple[str, tuple[str, ...]], ...]:
    """`(note, suspect fragments)` for each note that looks shell-mangled."""
    found = ((note, suspect_currency(note)) for note in notes)
    return tuple((note, fragments) for note, fragments in found if fragments)


def read_notes(source: Path | str) -> list[str]:
    """Read summary notes from a file — one per line — or from stdin for `-`.

    **This is the way to pass notes**, and the CLI documents it as such. The
    alternative is an argv string, and an argv string reaches Python through a shell:
    `python3 -c "... notes=['+US$150.0MM'] ..."` inside a double-quoted command
    arrives as `+US.0MM`, silently, in a client-facing artefact.

    A leading `- ` is stripped so a pasted bullet list works, and blank lines are
    dropped so a trailing newline is not a note.
    """
    text = sys.stdin.read() if str(source) == "-" else Path(source).read_text(encoding="utf-8")
    notes = []
    for line in text.splitlines():
        stripped = line.strip()
        if stripped.startswith("- "):
            stripped = stripped[2:].strip()
        if stripped:
            notes.append(stripped)
    return notes


def _corruption_warning(notes: Sequence[str]) -> list[str]:
    """The in-document warning for shell-mangled notes, or `[]` when they look fine."""
    suspects = corrupted_notes(notes)
    if not suspects:
        return []
    quoted = ", ".join(f"`{f}`" for _, fragments in suspects for f in fragments)
    return [
        "",
        f"> **⚠ The note text above looks shell-mangled: {quoted}.** A currency letter "
        "sits directly on a digit with no `$` — most likely the notes were passed as an "
        "argv string through a double-quoted shell command, where `$150` / `$MM` expand "
        "to nothing before Python sees them. The original figures cannot be recovered "
        "from this file. Re-run "
        "`python conductor.py summary <run_dir> --notes-file <file>` (or `--notes-file -` "
        "for stdin) with the intended text.",
    ]


def render_run_summary(run_dir: Path | str, *, notes: Sequence[str] = ()) -> str:
    """Compose the mechanical part of the end-of-run summary as markdown.

    Every stage across every wave, its status, its outputs, and the artefact paths
    the analyst can open right now. `notes` are appended verbatim under "Manual
    next steps" — that is where the model puts what only it knows (a sub-skill's
    "refresh the Capital IQ connector in the cap table", an abort reason, a caveat).

    Verbatim, but not unexamined: a note whose dollar signs a shell has eaten gets a
    warning printed beside it (:func:`suspect_currency`). Banking notes are
    dollar-dense, the corruption is silent, and this file is the analyst-facing
    record — one real run wrote "proceeds +US50.0MM ... dividend -US6.7MM" for
    +US$150.0MM and -US$66.7MM and read as if those were the figures.
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
    lines += _corruption_warning(notes)
    lines += ["", f"Full per-stage detail: `{run_dir / 'stages'}`.", ""]
    return "\n".join(lines)


def write_run_summary(run_dir: Path | str, *, notes: Sequence[str] = ()) -> Path:
    """Render the run summary and persist it as `<run_dir>/summary.md`.

    Warns on stderr as well as in the document when a note looks shell-mangled —
    the writer is who can still fix it, and they are looking at a terminal.
    """
    for note, fragments in corrupted_notes(notes):
        print(
            f"conductor: WARNING — note looks shell-mangled ({', '.join(fragments)}): "
            f"{note!r}\n"
            f"  A currency letter with no `$` before its digits. Pass notes with "
            f"--notes-file (or read_notes) instead of an argv string, and re-run the "
            f"summary.",
            file=sys.stderr,
        )
    return write_summary(run_dir, render_run_summary(run_dir, notes=notes))


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------
def _print_prepared(dispatch: WaveDispatch) -> None:
    print(dispatch.narration())
    print()
    for entry in dispatch.stages:
        print(f"=== stage `{entry.stage_id}` (skill: {entry.skill}, {entry.kind}) ===")
        print(f"inputs.json : {entry.inputs_path}")
        print(f"outputs.json: {entry.outputs_path}")
        if entry.kind == KIND_TRANSFORM:
            print("(in-process transform — no sub-agent; run `run-transforms` for this wave)")
        elif entry.prompt is not None:
            print("--- dispatch envelope ---")
            print(entry.prompt)
        else:
            print("(stage-envelope template not found; inputs.json written anyway)")
        print()


def _print_transforms(results: tuple[TransformResult, ...]) -> int:
    if not results:
        print("(this wave has no in-process transform stages)")
        return 0
    for r in results:
        if r.ok:
            keys = sorted(r.outputs or {})
            print(f"[ok]   {r.stage_id} ({r.skill}): outputs {keys}")
        else:
            print(f"[FAIL] {r.stage_id} ({r.skill}): {r.error}")
        if r.log_path is not None:
            print(f"       transcript: {r.log_path}")
        if not r.ok:
            # The tail as well as the path: a failure the reader has to open a file
            # to understand is one they will act on with less than they had.
            tail = _log_tail(r.log_path)
            if tail:
                print(tail)
    return 1 if any(not r.ok for r in results) else 0


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

    p_run = sub.add_parser("run-transforms", help="execute a wave's in-process stages")
    p_run.add_argument("run_dir")
    p_run.add_argument("wave", type=int)
    p_run.add_argument("--plugin-root", default=None)

    p_done = sub.add_parser("complete-wave", help="read + validate a wave's outputs")
    p_done.add_argument("run_dir")
    p_done.add_argument("wave", type=int)

    p_sum = sub.add_parser("summary", help="write + print the end-of-run summary")
    p_sum.add_argument("run_dir")
    p_sum.add_argument(
        "--notes-file",
        default=None,
        help=(
            "PREFERRED: read notes from a file, one per line ('-' for stdin). An argv "
            "string reaches Python through a shell, which eats every `$` in it — a run's "
            "summary said 'proceeds +US50.0MM' for +US$150.0MM and nothing complained."
        ),
    )
    p_sum.add_argument(
        "--note",
        action="append",
        default=[],
        help="a single note (repeatable). Shell-quoting is YOURS to get right; prefer --notes-file.",
    )

    args = parser.parse_args(argv)

    if args.command == "plan":
        print(plan_overview(args.run_dir).narration())
        return 0
    if args.command == "prepare-wave":
        _print_prepared(prepare_wave(args.run_dir, args.wave, plugin_root=args.plugin_root))
        return 0
    if args.command == "run-transforms":
        dispatch = prepare_wave(args.run_dir, args.wave, plugin_root=args.plugin_root)
        return _print_transforms(run_transforms(dispatch))
    if args.command == "complete-wave":
        return _print_outcome(complete_wave(args.run_dir, args.wave))
    if args.command == "summary":
        notes = (read_notes(args.notes_file) if args.notes_file else []) + list(args.note)
        path = write_run_summary(args.run_dir, notes=notes)
        print(path.read_text(encoding="utf-8"))
        return 0
    return 2


if __name__ == "__main__":  # pragma: no cover
    sys.exit(main())
