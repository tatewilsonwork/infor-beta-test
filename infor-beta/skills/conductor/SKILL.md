---
name: conductor
description: >
  Use this skill when the user wants to build a complete INFOR deliverable end-to-end —
  pitch, earnings update, or overview deck — instead of
  invoking individual skills. Activates on "build a <deliverable>", "kick off
  <deliverable>", "conductor", "/conductor", "orchestrate", the /pitch and
  /earnings-update commands (which preset the deliverable type + subject company), or
  any request that names a deliverable type rather than a single workflow step. The
  conductor handles deal-init (one set of G7 dialogs per codename), loads the plan YAML
  for the deliverable, collects plan-specific inputs via the locked interactive deck-spec
  dialogs (AskUserQuestion), dispatches each stage to its skill via the Agent tool with a
  file-based input / output handoff, and emits a run log under
  ~/Documents/INFOR Deals/<codename>/runs/<run-id>/.
version: 0.5.27
allowed-tools: [Read, Write, Bash, Glob, Task, AskUserQuestion]
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
from deal_init import (
    render_init_dialogs, render_init_filings_note, render_init_prompt,
    load_or_locate_deal, save_deal_context, load_deal_context,
)
from deck_spec import (
    render_deck_spec_dialogs, render_deck_spec_defaults, render_deck_spec_documents_note,
    render_deck_spec_prompt,  # text fallback when the interactive UI is unavailable
    default_presentation_date, prior_year_quarter,
    metric_count_from_slides, market_entry_targets_from_slides,
    NO_NOTES_ANALYST_NOTES,
    PITCH_DIALOG_PLAN_INPUTS, EARNINGS_UPDATE_DIALOG_PLAN_INPUTS,
)
from plan_refs import resolve_refs
from plan_schedule import compute_waves
from run_log import (
    make_run_id, create_run_dir, write_plan_snapshot,
    write_stage_inputs, read_stage_outputs, write_stage_log, write_summary,
)
# Optional thin driver that collapses the per-wave boilerplate (resolve refs,
# write every inputs.json, render envelopes / read + validate every outputs.json):
from conductor_cli import prep_wave, collect_wave, write_plan_inputs
```

---

## Workflow Steps

### Step 1 — Identify deliverable + codename

Read the analyst's request. Extract:
- **Deliverable type**: `pitch` / `earnings-update` / `overview` / `one-off-skill`. If ambiguous, ask before continuing. (`overview` is a stub plan — not yet implemented; say so if selected.)
- **Codename**: the analyst's `Project <target>` string. If absent, ask for it.

**Slash-command entry.** The plugin ships `/pitch <company name>` and `/earnings-update <company name>` commands that route here with the deliverable type AND the subject company name pre-answered. When entered that way, do not re-ask either — only the codename (and the remaining G7 items) still need collecting.

If the analyst is invoking a **one-off skill** (no deliverable plan needed), say so and stop — the conductor only orchestrates plans. Direct skill invocation is for one-offs.

### Step 2 — Deal-init (once per deal)

> **Interactive UI.** Every analyst-facing question in Steps 1–2 and 4 — and every `required` checkpoint — goes through the **`AskUserQuestion` tool** (clickable options + an automatic "Other" free-text box), never a numbered text block the analyst answers by typing item numbers. The dialog payloads are code-owned (`render_init_dialogs` / `render_deck_spec_dialogs`) so every run asks the same questions with the same options — render them **verbatim**: do not paraphrase, reorder, re-option, or invent extra questions. Only two things stay plain text: attachment checklists (files cannot come through a dialog) and pure free-text facts with nothing to suggest (the subject company name). If the `AskUserQuestion` tool is unavailable on the current surface, fall back to the locked text prompts (`render_init_prompt()` / `render_deck_spec_prompt(...)`) — same items, same order.

Call `load_or_locate_deal(codename)`:
- If it returns `(ctx, deal_dir)` with `ctx` non-None — the deal exists, the analyst has worked on it before. Confirm via `AskUserQuestion` (one question):
  > "I already have `<codename>` at `<deal_dir>` with `<inventory of facts/, filings/, artefacts/, prior runs>`. Continue this deal?" — options **"Continue `<codename>`"** / **"Different deal"**.
  - On "Different deal", call `disambiguate(deals_root, codename)` and present the 1–4 alternatives as another `AskUserQuestion` (one option per alternative; the analyst types their own via Other).
  - On "Continue", proceed to Step 3.
- If it returns `(None, deal_dir)` — fresh deal. Render the G7 dialogs with `render_init_dialogs(include_deliverable=<True only when Step 1 could not determine the deliverable>)` — one `AskUserQuestion` call per dialog, payload verbatim — **dropping any question whose answer is already preset** (the codename from Step 1; the deliverable + subject company from a slash command). Alongside the dialogs:
  - Post `render_init_filings_note()` as plain text — the filings are attachments and cannot come through a dialog.
  - If the subject company name is not preset, ask for it as a plain chat question (pure free text — it has no dialog).
  - Codename "Propose one for me" → propose a `Project <single word>` and confirm it before creating the deal directory. "Public — I'll give the ticker" with no ticker in the Other text → ask for ticker + exchange as a plain follow-up. Sector "Infer from the web — I'll confirm" → research it, verify by web search, and confirm the one-liner with the analyst.

  Construct a `DealContext`, then `save_deal_context(ctx)` to persist `<deal_dir>/deal.json` and bootstrap `facts/`, `filings/`, `artefacts/`, `runs/`.

**Filings handling:** if the analyst attaches files at deal-init, save them under `<deal_dir>/filings/` with descriptive names and append matching `Filing` entries to `ctx.filings`. Re-save the DealContext.

### Step 3 — Load the plan

Plans live at `${CLAUDE_PLUGIN_ROOT}/plans/<deliverable>.yaml`. Resolve and read the YAML; validate by constructing a `Plan` pydantic model. If validation fails, surface the error and stop — do not attempt to run a partially-valid plan.

Read the plan's `description` and present a one-paragraph summary to the analyst:
> "Plan for `<deliverable>` has N stages: `<stage_id_1>` (skill: `<skill_1>`), `<stage_id_2>` (skill: `<skill_2>`), …. Plan inputs required: `<list>`. Checkpoints: `<list of required checkpoints>`."

### Step 4 — Collect plan inputs (deck-spec dialogs)

For `pitch` and `earnings-update`, collect the deck spec through the **locked interactive dialogs**. Only the judgement items are asked; everything with a sensible default is defaulted and echoed for override. In order:

1. **Compute the defaults.** Six pitch inputs (two for earnings-update) are never asked:
   - `client_name` ← the deal-init subject company name,
   - `presentation_date` ← `default_presentation_date(date.today())` (spelled-out month + year),
   - `reporting_quarter` ← the latest attached interim filing's fiscal quarter — infer it from the statements themselves (fiscal quarter labels depend on the company's fiscal calendar, not the calendar date),
   - `comparison_quarter` ← `prior_year_quarter(reporting_quarter)`,
   - `financial_metric_count` and `section_labels` (pitch only) ← left OUT of `plan_inputs`; the wireframe applies its own defaults (one Financial Summary slide / 4 metrics; standard section labels).
2. **Post the defaults echo** — `render_deck_spec_defaults(plan.deliverable_type, client_name=…, presentation_date=…, reporting_quarter=…, comparison_quarter=…)` verbatim, so the analyst can override any default by replying. Override answers convert deterministically: "2 Financial Summary slides" → `financial_metric_count = metric_count_from_slides(2)`; a replacement reporting quarter re-derives the comparison quarter via `prior_year_quarter` unless the analyst gave both.
3. **Render the dialogs** — `render_deck_spec_dialogs(plan.deliverable_type)`: one `AskUserQuestion` call per dialog, payload verbatim (see the Step 2 interactive-UI rules). Never re-ask a G7 item. If the analyst's earlier messages already answered a dialog item, drop just that question from the payload and note "(from your message: …)" — do not skip whole dialogs.
4. **Post the documents note** — `render_deck_spec_documents_note(plan.deliverable_type)` as plain text (the G7 filings, SEDI PDF, Bloomberg export, EEO snip, and CIM are attachments — they cannot come through a dialog).

Map every dialog answer onto `plan_inputs` with the module's header tables (`PITCH_DIALOG_PLAN_INPUTS` / `EARNINGS_UPDATE_DIALOG_PLAN_INPUTS`). Conversions are deterministic, never improvised:

- **Notes** → `analyst_notes`: "I'll paste notes in my next message" → wait for the notes, use them verbatim; "Draft from the attached filings + web" → the code-owned literal `NO_NOTES_ANALYST_NOTES`; Other → the typed text.
- **Targets** → `market_entry_target_count = market_entry_targets_from_slides(n_slides)` (2 per slide, at most 4 slides).
- **Highlights** → `include_investment_highlights = False` only on "Omit"; any include variant leaves the input unset (analyst-dictated highlight copy belongs in the analyst notes).
- **CIM / EEO snip** → "Attached in this chat" → save the attachment under `<deal_dir>/filings/` and use the saved path; "None" → leave unset; Other → the given path.
- **Valuation / Risk notes** → "None" → leave unset; "I'll provide…" with no text supplied → collect it as a plain follow-up; Other → the typed text.

An answer that just accepts a default is left OUT of `plan_inputs` (see the optional-input rule below); the computed defaults for the four *required* pitch inputs (`client_name`, `presentation_date`, `reporting_quarter`, `comparison_quarter` — quarters only for earnings-update) are always supplied. Files the analyst attaches at the deck spec (SEDI PDF, Bloomberg export, EEO snip, CIM) are saved under `<deal_dir>/filings/` exactly like G7 attachments.

**Text fallback:** if `AskUserQuestion` is unavailable, render `render_deck_spec_prompt(plan.deliverable_type)` verbatim in a single message instead — it asks the same items (numbered; map answers via `PITCH_ITEM_PLAN_INPUTS` / `EARNINGS_UPDATE_ITEM_PLAN_INPUTS`) and lists the same defaults; `"defaults"` accepts every bracketed default.

For any other deliverable (no questionnaire — the deck-spec renderers raise `ValueError`), fall back to the generic collection: for each `InputSpec` in `plan.plan_inputs` where `required is True`, prompt the analyst in a single message; for optional inputs, ask only if the analyst's initial message did not already supply them.

Store the collected values as a plain dict `plan_inputs: dict[str, Any]`. Validate types informally — exact pydantic-validation of plan-input values is deferred; in v1 the type field is documentation.

**Optional inputs the analyst didn't supply must stay OUT of `plan_inputs` — do not pre-seed them with `None` or any placeholder.** Instead compute the set of optional plan-input names once and pass it to every `resolve_refs` call in Step 6a:

```python
optional_plan_inputs = {spec.name for spec in plan.plan_inputs if not spec.required}
```

A stage that references an unsupplied optional input (`$plan_inputs.<name>`) then resolves to `None` automatically; a missing *required* input — or any missing `$deal.*` / `$stages.*` reference — still raises and halts the run.

### Step 5 — Create the run directory

```python
run_id = make_run_id(plan.deliverable_type)
run_dir = create_run_dir(ctx.deal_dir, run_id)
write_plan_snapshot(run_dir, plan_yaml_text)  # frozen snapshot of the plan
```

Tell the analyst the run id and its path. Subsequent log files all land under that directory.

### Step 6 — Dispatch stages wave-by-wave

Stages run in **dependency waves**, not one at a time. Compute the schedule once:

```python
waves = compute_waves(plan)   # list[list[stage_id]] in execution order
```

Each wave is a list of stage ids with **no dependency between them**, so the whole wave is dispatched **concurrently** and the conductor waits for it to finish before starting the next. Dependencies are auto-derived from the `$stages.<id>.<name>` references already in each stage's inputs — the references *are* the DAG, there is no `depends_on` field. The `workbook-aggregator` stage carries one hardcoded extra edge: it depends on **every stage except its own downstream consumers** (it consolidates and deletes the individual workbooks, so nothing that produces or reads them may run alongside it) — so it is always alone in its wave, but a post-aggregation stage that consumes its output (the pitch plan's `financial-charts`) runs in a later wave after it. Tell the analyst the wave plan up front, e.g. for pitch:

> 7 waves: [1] `wireframe`, `financial-summary`, `comps`, `precedents` (parallel) → [2] `content`, `ltm-metrics` → [3] `captable` → [4] `ownership` → [5] `deck` → [6] `workbook-aggregation` → [7] `financial-charts`.

> **Driver shortcut.** Rather than hand-coding 6a/6c per wave, persist the collected `plan_inputs` once with `write_plan_inputs(run_dir, plan_inputs)` and drive each wave with the `conductor_cli` helpers: `prep_wave(run_dir, n)` resolves every reference (passing `optional_plan_inputs` for you), writes each stage's `inputs.json` through the pydantic/`Path`-safe encoder, and returns the rendered dispatch envelopes; after the wave, `collect_wave(run_dir, n)` reads + validates every `outputs.json`. (Equivalently, `python ${CLAUDE_PLUGIN_ROOT}/scripts/conductor_cli.py prep-wave <run_dir> <n>` / `collect-wave <run_dir> <n>`.) You still issue the `Task` calls and run the checkpoints yourself. The manual steps below are the contract these helpers implement.

Maintain an in-memory `stage_outputs: dict[str, dict[str, Any]]` keyed by stage id. For each wave, in order:

**6a — Prepare every stage in the wave.** For each stage id in the wave, look up its `Stage` and:
1. **Resolve inputs.** `resolve_refs(stage.inputs, plan_inputs=plan_inputs, deal_context=ctx, stage_outputs=stage_outputs, optional_plan_inputs=optional_plan_inputs)`. Every `$stages.*` reference a wave member needs is already in `stage_outputs` — the scheduler guarantees its producer ran in an earlier wave. Passing `optional_plan_inputs` (from Step 4) lets an unsupplied optional plan input resolve to `None` instead of crashing the run.
2. **Persist inputs.** `write_stage_inputs(run_dir, stage.id, resolved_inputs)`.
3. **Render envelope.** Load `references/stage-envelope.md`, substitute its placeholders with the **absolute** paths for *that stage* (`{{stage_inputs_path}}` → its `inputs.json`, `{{stage_outputs_path}}` → its `outputs.json`, `{{deal_dir}}`, `{{plugin_root}}`), and prepare the prompt for the Agent tool. The rendered prompt carries the `export STAGE_INPUTS=… / STAGE_OUTPUTS=… / DEAL_DIR=… / CLAUDE_PLUGIN_ROOT=…` block as the sub-agent's first step.

**6b — Dispatch the whole wave concurrently.** Issue one `Task` (Agent) call per stage **in a single message** so they run in parallel, passing each stage's **rendered prompt** (the only channel — the `Task`/`Agent` tool has no parameter for environment variables, so the handoff paths must live in the prompt body, not be set as env vars on the call). Each sub-agent exports those paths itself before running its SKILL.md commands. Wait for every sub-agent in the wave to finish before continuing. (A single-stage wave is just one `Task` call — same as the old sequential behaviour.)

**6c — Collect the wave.** After all of the wave's sub-agents return, for each stage id in the wave (in listed order):
4. **Read outputs.** `outputs = read_stage_outputs(run_dir, stage.id)`. This raises `FileNotFoundError` if the sub-skill failed to write outputs.json — surface the failure to the analyst and stop. Do not start the next wave if any stage in this one produced no structured outputs.
5. **Capture log.** Persist the sub-agent's reply transcript via `write_stage_log(run_dir, stage.id, transcript)`.
6. **Update state.** `stage_outputs[stage.id] = outputs`.

**6d — Checkpoint the wave.** Once the wave's state is updated, run the checkpoint behaviour (`references/checkpoint-behaviour.md`) for each stage in the wave, in listed order — `informational` summarises (batch the wave's outputs into one surface), `silent` does nothing, `required` halts and asks for approval via `AskUserQuestion` ("Approve — continue the run" / "Halt the run"). If the analyst rejects any `required` checkpoint, halt before dispatching the next wave.

> **`required` checkpoints and parallelism.** A `required` gate is evaluated at its **wave boundary**, after every stage in that wave has already run — so it cannot stop its own wave-mates, only the downstream waves. Today every shipped plan's checkpoints are `informational`, so behaviour is unchanged. If a future plan needs a gate to stop work *before* an expensive stage starts, give that stage a dependency so it lands in a later wave (the scheduler will serialise it behind the gate).

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
- Invent or infer dependencies beyond the references. The wave schedule comes purely from the `$stages.*` references in each stage's inputs plus the hardcoded aggregator barrier (`workbook-aggregator` depends on every stage except its own downstream consumers; `plan_schedule.compute_waves`). The conductor will not add edges, reorder, or parallelise beyond what those imply — and it never dispatches a stage before everything it references has produced outputs.
- Emit telemetry beyond per-stage transcripts. `meta.json` (model, tokens, latency) is Phase 5.
