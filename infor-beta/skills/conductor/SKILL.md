---
name: conductor
description: >
  Use this skill when the user wants to build a complete INFOR deliverable end-to-end —
  pitch, earnings update, or overview deck — instead of
  invoking individual skills. Activates on "build a <deliverable>", "kick off
  <deliverable>", "conductor", "/conductor", "orchestrate", or any request that names a
  deliverable type rather than a single workflow step. The conductor handles deal-init
  (one G7 prompt per codename), loads the plan YAML for the deliverable, collects
  plan-specific inputs, dispatches each stage to its skill via the Agent tool with a
  file-based input / output handoff, and emits a run log under
  ~/Documents/INFOR Deals/<codename>/runs/<run-id>/.
version: 0.5.12
allowed-tools: [Read, Write, Bash, Glob, Task]
---

# Conductor — Workflow

The conductor is a thin orchestrator: **dumb about banking, smart about orchestration**. It never produces a deliverable directly. It loads a plan, runs each stage as a sub-agent, and accounts for what happened.

The architectural backbone — DealContext schema, codename rules, deliverable types, three checkpoint modes — is locked in Obsidian note `12 — Locked Decisions`. Re-read note 12 H1–H8 before changing this skill's behaviour.

**Detailed references** (loaded on demand):
- [`references/plan-schema.md`](references/plan-schema.md) — Plan / Stage YAML schema and reference-string semantics (`$plan_inputs`, `$deal`, `$stages`).
- [`references/stage-envelope.md`](references/stage-envelope.md) — the prompt template rendered per stage and passed to the Agent tool.
- [`references/checkpoint-behaviour.md`](references/checkpoint-behaviour.md) — what `required` / `informational` / `silent` mean operationally.

Helpers live at `${CLAUDE_PLUGIN_ROOT:-./infor-beta}/scripts/`:

```python
import sys, os
sys.path.insert(0, os.environ.get("CLAUDE_PLUGIN_ROOT", "./infor-beta") + "/scripts")

from schemas import Plan, Stage, DealContext, InputSpec
from codename import resolve, find_existing, disambiguate
from deal_init import render_init_prompt, load_or_locate_deal, save_deal_context, load_deal_context
from plan_refs import resolve_refs
from run_log import (
    make_run_id, create_run_dir, write_plan_snapshot,
    write_stage_inputs, read_stage_outputs, write_stage_log, write_summary,
)
```

---

## Workflow Steps

### Step 1 — Identify deliverable + codename

Read the analyst's request. Extract:
- **Deliverable type**: `pitch` / `earnings-update` / `overview` / `one-off-skill`. If ambiguous, ask before continuing. (`overview` is a stub plan — not yet implemented; say so if selected.)
- **Codename**: the analyst's `Project <target>` string. If absent, ask for it.

If the analyst is invoking a **one-off skill** (no deliverable plan needed), say so and stop — the conductor only orchestrates plans. Direct skill invocation is for one-offs.

### Step 2 — Deal-init (once per deal)

Call `load_or_locate_deal(codename)`:
- If it returns `(ctx, deal_dir)` with `ctx` non-None — the deal exists, the analyst has worked on it before. Confirm with the analyst:
  > "I already have `<codename>` at `<deal_dir>` with `<inventory of facts/, filings/, artefacts/, prior runs>`. Continue this deal? [y/n]"
  - If `n`, call `disambiguate(deals_root, codename)` and present 1–4 alternatives; ask the analyst to pick or type their own.
  - If `y`, proceed to Step 3.
- If it returns `(None, deal_dir)` — fresh deal. Render the G7 prompt with `render_init_prompt()` and ask the analyst to answer all seven items. Construct a `DealContext`, then `save_deal_context(ctx)` to persist `<deal_dir>/deal.json` and bootstrap `facts/`, `filings/`, `artefacts/`, `runs/`.

**Filings handling:** if the analyst attaches files at the G7 prompt, save them under `<deal_dir>/filings/` with descriptive names and append matching `Filing` entries to `ctx.filings`. Re-save the DealContext.

### Step 3 — Load the plan

Plans live at `${CLAUDE_PLUGIN_ROOT}/plans/<deliverable>.yaml`. Resolve and read the YAML; validate by constructing a `Plan` pydantic model. If validation fails, surface the error and stop — do not attempt to run a partially-valid plan.

Read the plan's `description` and present a one-paragraph summary to the analyst:
> "Plan for `<deliverable>` has N stages: `<stage_id_1>` (skill: `<skill_1>`), `<stage_id_2>` (skill: `<skill_2>`), …. Plan inputs required: `<list>`. Checkpoints: `<list of required checkpoints>`."

### Step 4 — Collect plan inputs

For each `InputSpec` in `plan.plan_inputs` where `required is True`, prompt the analyst (in a single message — ask for all of them at once). For optional inputs, ask only if the analyst's initial message did not already supply them.

Store the collected values as a plain dict `plan_inputs: dict[str, Any]`. Validate types informally — exact pydantic-validation of plan-input values is deferred; in v1 the type field is documentation.

### Step 5 — Create the run directory

```python
run_id = make_run_id(plan.deliverable_type)
run_dir = create_run_dir(ctx.deal_dir, run_id)
write_plan_snapshot(run_dir, plan_yaml_text)  # frozen snapshot of the plan
```

Tell the analyst the run id and its path. Subsequent log files all land under that directory.

### Step 6 — Dispatch each stage in order

Maintain an in-memory `stage_outputs: dict[str, dict[str, Any]]` keyed by stage id. For each stage in `plan.stages` (declaration order — sequential, no parallel in v1):

1. **Resolve inputs.** Call `resolve_refs(stage.inputs, plan_inputs=..., deal_context=ctx, stage_outputs=stage_outputs)`.
2. **Persist inputs.** `write_stage_inputs(run_dir, stage.id, resolved_inputs)`.
3. **Render envelope.** Load `references/stage-envelope.md`, substitute its placeholders, prepare the prompt for the Agent tool.
4. **Dispatch via `Task`** (the Agent tool). Pass it the rendered envelope. Set environment variables `STAGE_INPUTS` and `STAGE_OUTPUTS` (absolute paths to `inputs.json` and `outputs.json`), and `DEAL_DIR` (absolute path to the deal directory). Wait for the sub-agent to finish.
5. **Read outputs.** `outputs = read_stage_outputs(run_dir, stage.id)`. This raises `FileNotFoundError` if the sub-skill failed to write outputs.json — in that case, surface the failure to the analyst and stop. Do not proceed past a stage that didn't write structured outputs.
6. **Capture log.** Take the sub-agent's reply transcript and persist it via `write_stage_log(run_dir, stage.id, transcript)`.
7. **Update state.** `stage_outputs[stage.id] = outputs`.
8. **Checkpoint.** Run the checkpoint behaviour for `stage.checkpoint` per `references/checkpoint-behaviour.md` (`required` halts, `informational` summarises and continues, `silent` does nothing).

### Step 7 — Emit summary

Compose an analyst-readable end-of-run summary and write it with `write_summary(run_dir, summary_md)`. The summary should include:

- Deliverable, codename, run id, total runtime if available.
- One line per stage: status (ok/failed), notable outputs (paths).
- All artefact paths the analyst can open right now.
- Any manual next steps surfaced by individual sub-skills (e.g. "refresh the Capital IQ connector in the cap table workbook").
- Pointer to the run directory for full per-stage detail.

Then post the summary back to the analyst.

---

## Stop conditions

The conductor halts (with a clear error message and the partial-run state preserved on disk) when:

- The deliverable type or codename is missing and the analyst declines to supply it.
- An existing deal is detected and the analyst declines to continue or rename.
- Plan YAML fails pydantic validation.
- A reference cannot be resolved (e.g. `$stages.x.y` where stage `x` hasn't run yet).
- A sub-skill fails to write `outputs.json`.
- A `required` checkpoint is rejected by the analyst.

Never silently skip a stage. Never proceed past a missing output. Never overwrite an existing `outputs.json` (a re-run uses a new `run_id`).

## What the conductor does not do

- Produce slides, models, or copy. That is each sub-skill's job.
- Make banking decisions. Voice, brand, and source-trust rules live in each stage skill's own SKILL.md / references and its allow-list — not in the conductor.
- Parallelise stages. v1 is sequential. Parallel + DAG dependencies arrive when Phase 3's deck-assembler + slide library justify the complexity.
- Emit telemetry beyond per-stage transcripts. `meta.json` (model, tokens, latency) is Phase 5.
