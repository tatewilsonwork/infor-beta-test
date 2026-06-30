"""Thin conductor driver — per-wave prep / collect helpers.

The conductor (a Claude Code skill) dispatches a plan's stages in dependency
**waves** (`plan_schedule.compute_waves`). Driving that by hand — resolving each
stage's references, writing every `inputs.json`, rendering each stage envelope,
then reading back every `outputs.json` — is repetitive boilerplate, and a stray
`json.dumps` over a pydantic object (e.g. a `Company` resolved from
`$deal.subject_company`) blew up a live run. This module collapses that loop into
two commands and routes ALL serialization through `run_log._json_default`, which
already handles pydantic models (`model_dump(mode="json")`), `Path`, and sets, so
the `json.dumps`-on-`Company` class of error cannot recur.

    conductor_cli.py prep-wave    <run_dir> <wave>   # resolve + write inputs, print envelopes
    conductor_cli.py collect-wave <run_dir> <wave>   # read + validate each stage's outputs

`<wave>` is 1-based, matching the conductor's "wave 1 / wave 2 / …" narration.

State is reconstructed from the run directory, which the conductor already
populates: `<run_dir>/plan.yaml` (the frozen plan snapshot), the deal's
`deal.json` (two levels up: `<deal_dir>/runs/<run-id>/`), an optional
`<run_dir>/plan_inputs.json` (the analyst-collected plan inputs; absent ⇒ `{}`),
and each prior stage's `stages/<id>/outputs.json`. `prep-wave` for wave *n* reads
the outputs of every earlier wave so `$stages.*` references resolve.

This is a convenience driver, not a new authority: it owns no banking logic and
makes no Agent calls. The conductor still dispatches the printed envelopes via the
`Task` tool and runs the checkpoint behaviour itself.
"""

from __future__ import annotations

import argparse
import json
import os
import sys
from pathlib import Path

import yaml

from deal_init import load_deal_context
from plan_refs import resolve_refs
from plan_schedule import compute_waves
from run_log import _json_default, read_stage_outputs, stage_dir, write_stage_inputs
from schemas import Plan

PLAN_INPUTS_NAME = "plan_inputs.json"


# ---------------------------------------------------------------------------
# Run-state loading
# ---------------------------------------------------------------------------
def load_plan(run_dir: Path | str) -> Plan:
    """Load + validate the frozen plan snapshot at `<run_dir>/plan.yaml`."""
    path = Path(run_dir) / "plan.yaml"
    if not path.exists():
        raise FileNotFoundError(f"no plan snapshot at {path} — was the run dir created?")
    return Plan.model_validate(yaml.safe_load(path.read_text(encoding="utf-8")))


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


def write_plan_inputs(run_dir: Path | str, plan_inputs: dict) -> Path:
    """Persist the analyst-collected plan inputs for the per-wave driver to read."""
    path = Path(run_dir) / PLAN_INPUTS_NAME
    path.write_text(
        json.dumps(plan_inputs, indent=2, sort_keys=True, default=_json_default) + "\n",
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


# ---------------------------------------------------------------------------
# Stage-envelope rendering
# ---------------------------------------------------------------------------
def _stage_by_id(plan: Plan, stage_id: str):
    for s in plan.stages:
        if s.id == stage_id:
            return s
    raise KeyError(f"stage {stage_id!r} not found in plan")


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


def render_stage_envelope(
    plan: Plan,
    stage,
    *,
    ctx,
    run_dir: Path,
    plugin_root: Path,
) -> str | None:
    """Render the conductor stage envelope for one stage (absolute paths baked in).

    Returns the rendered prompt, or None if the envelope template can't be found
    (prep-wave still writes inputs.json — only the printed prompt is skipped).
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
    }
    rendered = template
    for token, value in subs.items():
        rendered = rendered.replace(token, value)
    return rendered


# ---------------------------------------------------------------------------
# prep-wave / collect-wave
# ---------------------------------------------------------------------------
def _wave_stage_ids(plan: Plan, wave: int) -> list[str]:
    """Return the stage ids in 1-based `wave`, validating the wave number."""
    waves = compute_waves(plan)
    if wave < 1 or wave > len(waves):
        raise IndexError(
            f"wave {wave} out of range — the plan schedules {len(waves)} wave(s) "
            f"(1..{len(waves)})"
        )
    return waves[wave - 1]


def _prior_stage_outputs(plan: Plan, wave: int, run_dir: Path) -> dict[str, dict]:
    """Read outputs.json for every stage in waves before `wave` (1-based)."""
    waves = compute_waves(plan)
    outputs: dict[str, dict] = {}
    for earlier in waves[: wave - 1]:
        for sid in earlier:
            outputs[sid] = read_stage_outputs(run_dir, sid)
    return outputs


def prep_wave(
    run_dir: Path | str,
    wave: int,
    *,
    plugin_root: Path | str | None = None,
) -> list[dict]:
    """Resolve references and write inputs.json for every stage in `wave` (1-based).

    Reads prior waves' outputs so `$stages.*` references resolve; lets unsupplied
    optional plan inputs resolve to None (via `optional_plan_inputs`); writes each
    stage's inputs.json through `run_log.write_stage_inputs` (pydantic/`Path`-safe);
    and renders each stage's dispatch envelope. Returns one dict per stage with its
    id, skill, resolved inputs, the inputs/outputs paths, and the rendered prompt.
    """
    run_dir = Path(run_dir)
    plan = load_plan(run_dir)
    ctx = load_deal_context(_deal_dir_for_run(run_dir))
    plan_inputs = load_plan_inputs(run_dir)
    optional = {spec.name for spec in plan.plan_inputs if not spec.required}
    proot = Path(plugin_root) if plugin_root is not None else default_plugin_root()

    stage_ids = _wave_stage_ids(plan, wave)
    stage_outputs = _prior_stage_outputs(plan, wave, run_dir)

    prepared: list[dict] = []
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
        envelope = render_stage_envelope(
            plan, stage, ctx=ctx, run_dir=run_dir, plugin_root=proot
        )
        prepared.append(
            {
                "stage_id": sid,
                "skill": stage.skill,
                "inputs": resolved,
                "inputs_path": str(inputs_path),
                "outputs_path": str((stage_dir(run_dir, sid) / "outputs.json").resolve()),
                "envelope": envelope,
            }
        )
    return prepared


def collect_wave(run_dir: Path | str, wave: int) -> list[dict]:
    """Read + validate each stage's outputs.json for `wave` (1-based).

    Returns one dict per stage: `{stage_id, ok, outputs|error}`. A stage that
    wrote no outputs.json, or wrote one carrying an `error` key, is reported with
    `ok=False` (the conductor then halts rather than starting the next wave).
    """
    run_dir = Path(run_dir)
    plan = load_plan(run_dir)
    stage_ids = _wave_stage_ids(plan, wave)

    results: list[dict] = []
    for sid in stage_ids:
        try:
            outputs = read_stage_outputs(run_dir, sid)
        except FileNotFoundError as exc:
            results.append({"stage_id": sid, "ok": False, "error": str(exc)})
            continue
        if isinstance(outputs, dict) and "error" in outputs:
            results.append(
                {"stage_id": sid, "ok": False, "error": outputs["error"], "outputs": outputs}
            )
        else:
            results.append({"stage_id": sid, "ok": True, "outputs": outputs})
    return results


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------
def _print_prep(prepared: list[dict]) -> None:
    for entry in prepared:
        print(f"=== stage `{entry['stage_id']}` (skill: {entry['skill']}) ===")
        print(f"inputs.json : {entry['inputs_path']}")
        print(f"outputs.json: {entry['outputs_path']}")
        if entry["envelope"] is not None:
            print("--- dispatch envelope ---")
            print(entry["envelope"])
        else:
            print("(stage-envelope template not found; inputs.json written anyway)")
        print()


def _print_collect(results: list[dict]) -> int:
    failures = 0
    for r in results:
        if r["ok"]:
            keys = sorted(r["outputs"].keys()) if isinstance(r.get("outputs"), dict) else []
            print(f"[ok]   {r['stage_id']}: outputs {keys}")
        else:
            failures += 1
            print(f"[FAIL] {r['stage_id']}: {r['error']}")
    return failures


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(prog="conductor_cli", description=__doc__)
    sub = parser.add_subparsers(dest="command", required=True)

    p_prep = sub.add_parser("prep-wave", help="resolve + write inputs.json for a wave")
    p_prep.add_argument("run_dir")
    p_prep.add_argument("wave", type=int)
    p_prep.add_argument("--plugin-root", default=None)

    p_collect = sub.add_parser("collect-wave", help="read + validate a wave's outputs")
    p_collect.add_argument("run_dir")
    p_collect.add_argument("wave", type=int)

    args = parser.parse_args(argv)

    if args.command == "prep-wave":
        prepared = prep_wave(args.run_dir, args.wave, plugin_root=args.plugin_root)
        _print_prep(prepared)
        return 0
    if args.command == "collect-wave":
        results = collect_wave(args.run_dir, args.wave)
        failures = _print_collect(results)
        return 1 if failures else 0
    return 2


if __name__ == "__main__":  # pragma: no cover
    sys.exit(main())
