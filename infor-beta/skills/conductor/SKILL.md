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
  dialogs (AskUserQuestion), runs each deterministic stage in-process and dispatches each
  judgment stage to its skill via the Agent tool with a file-based input / output handoff,
  and emits a run log under ~/Documents/INFOR Deals/<codename>/runs/<run-id>/.
allowed-tools: [Read, Write, Bash, Glob, Task, AskUserQuestion]
---

# Conductor — Workflow

The conductor is a thin orchestrator: **dumb about banking, smart about orchestration**. It never produces a deliverable directly.

Since Phase E the mechanics live in `scripts/conductor.py`, not in this file. **Your job is four things** — intake, issuing the `Task` calls the driver hands back, reporting each wave boundary, and the summary. Everything else is a function call with a return value, so it cannot be skipped on turn 40 of a long run.

Since v0.5.49 **no shipped plan asks the analyst to approve anything mid-run.** The intake is the last question; after it, run every wave to the end and report. Nothing to approve does not mean nothing is checked — the geometry converge loop, the written vision review and the `deckcheck` falsification pass all still run, and none of them needs an answer.

Since Phase F the second of those got smaller: a stage is either a **transform** (deterministic — the driver calls the function in-process) or **judgment** (research and drafting — a sub-agent with a real allow-list; `deckcheck` included). You dispatch the judgment stages only. `plan_overview(run_dir).narration()` reports the split for the plan in front of you; no number is written down here, because a written-down wave count went stale once already.

The architectural backbone — DealContext schema, codename rules, deliverable types, three checkpoint modes — is locked in Obsidian note `12 — Locked Decisions`. Re-read note 12 H1–H8 before changing this skill's behaviour.

**Detailed references** (loaded on demand):
- [`references/plan-schema.md`](references/plan-schema.md) — Plan / Stage YAML schema and reference-string semantics (`$plan_inputs`, `$deal`, `$stages`).
- [`references/stage-envelope.md`](references/stage-envelope.md) — the prompt template `conductor.py` renders per stage. **No environment variables:** the handoff paths are arguments.
- [`references/checkpoint-behaviour.md`](references/checkpoint-behaviour.md) — what `required` / `informational` / `silent` mean operationally.

```python
import sys, os
sys.path.insert(0, os.environ.get("CLAUDE_PLUGIN_ROOT", "./infor-beta") + "/scripts")

from codename import codename_from_company, disambiguate
from deal_init import (
    render_init_dialogs, render_init_filings_note, render_init_prompt,
    load_or_locate_deal, save_deal_context,
)
from deck_spec import (
    render_deck_spec_dialogs, render_deck_spec_defaults, render_deck_spec_documents_dialogs,
    render_deck_spec_documents_note, render_deck_spec_prompt,
    default_presentation_date, prior_year_quarter,
    metric_count_from_slides, market_entry_targets_from_slides, NO_NOTES_ANALYST_NOTES,
    PITCH_DIALOG_PLAN_INPUTS, EARNINGS_UPDATE_DIALOG_PLAN_INPUTS,
)
from run_log import make_run_id, create_run_dir, write_plan_snapshot, write_stage_log
from conductor import (      # the driver — everything mechanical
    plan_overview, prepare_wave, run_transforms, complete_wave,
    write_plan_inputs, write_run_summary, APPROVE_LABEL, HALT_LABEL,
)
```

> **Interactive UI.** Every analyst-facing question — Steps 1, 2 and 4, which since v0.5.49 are all of them — goes through **`AskUserQuestion`**, never a numbered text block. The payloads are code-owned (`render_init_dialogs` / `render_deck_spec_dialogs` / `render_deck_spec_documents_dialogs`; and `WaveOutcome.gate.question` if a plan ever carries a `required` checkpoint again): render them **verbatim** — do not paraphrase, reorder, re-option, or invent extra questions. Attachments get fixed **status** dialogs (attached / will-drop / none) because file bytes cannot come through a dialog; the file itself arrives via the chat input or an absolute path in Other, with the plain-text notes posted alongside as the checklist detail. Pure free-text facts with nothing to suggest (the subject company name) stay plain chat questions. If `AskUserQuestion` is unavailable, fall back to the locked text prompts (`render_init_prompt()` / `render_deck_spec_prompt(...)` / `Checkpoint.fallback_prompt`) — same items, same order.

## Step 1 — Deliverable + codename

Extract the **deliverable type** (`pitch` / `earnings-update` / `overview` / `one-off-skill`; ask if ambiguous — `overview` is a stub, say so if selected) and the **codename**. Never ask for the codename: use the analyst's `Project <target>` string if they typed one, else derive it silently with `codename_from_company(<subject company>)` and state it when announcing the deal directory (overridable in chat until the directory is created). `/pitch` and `/earnings-update` pre-answer the deliverable **and** the company — do not re-ask either. A **one-off skill** needs no plan: say so and stop.

## Step 2 — Deal-init (once per deal)

`load_or_locate_deal(codename)` → an existing deal means one `AskUserQuestion` ("Continue `<codename>`" / "Different deal"; on the latter, present `disambiguate(...)`'s 1–4 alternatives as another dialog). A fresh deal means `render_init_dialogs(include_deliverable=<True only when Step 1 could not determine it>)` — one call per dialog, verbatim, **dropping any question already answered** (the slash-command deliverable + company; Filings when files are already attached) — plus `render_init_filings_note()` as plain text and, if the company name is not preset, a plain chat question for it. "Public — I'll give the ticker" with no ticker → ask ticker + exchange as a follow-up. Sector "Infer from the web" → research, verify by web search, use the one-liner; no confirmation. Filings "Attached" → save now; "next message" → wait for it; "None for now" → proceed. Then build the `DealContext` and `save_deal_context(ctx)`.

**Filings handling:** save every attachment under `<deal_dir>/filings/` with a descriptive name, append matching `Filing` entries to `ctx.filings`, and re-save. This applies to deck-spec attachments (SEDI PDF, Bloomberg export, EEO snip, CIM) too.

## Step 3 — Plan + run directory

Read `${CLAUDE_PLUGIN_ROOT}/plans/<deliverable>.yaml`, then:

```python
run_id = make_run_id(plan.deliverable_type)
run_dir = create_run_dir(ctx.deal_dir, run_id)
write_plan_snapshot(run_dir, plan_yaml_text)   # frozen snapshot; the driver reads only this
```

Tell the analyst the run id and its path. `plan_overview(run_dir)` (Step 5) validates the snapshot in two layers — the pydantic `Plan` shape, then the reference pre-flight `validate_plan_references` — so a typo'd `$stages` / `$plan_inputs` reference is dead at load, not mid-run. Surface any error and stop; never run a partially-valid plan.

## Step 4 — Collect plan inputs (deck-spec dialogs)

For `pitch` and `earnings-update`, only the judgement items are asked; everything with a sensible default is defaulted and echoed for override. In order: **compute the defaults** (`client_name` ← the subject company; `presentation_date` ← `default_presentation_date(date.today())`; `reporting_quarter` ← inferred from the latest attached interim filing, whose fiscal labels depend on the company's fiscal calendar, not the calendar date; `comparison_quarter` ← `prior_year_quarter(...)`; `financial_metric_count` + `section_labels` ← left unset so the wireframe defaults apply) → **post `render_deck_spec_defaults(...)`** → **render `render_deck_spec_dialogs(...)`** → **render `render_deck_spec_documents_dialogs(...)`** → **post `render_deck_spec_documents_note(...)`**. Never re-ask a G7 item; if an earlier message already answered a dialog item, drop just that question and note "(from your message: …)".

Map answers with `PITCH_DIALOG_PLAN_INPUTS` / `EARNINGS_UPDATE_DIALOG_PLAN_INPUTS`. Conversions are deterministic, never improvised: **Notes** → `analyst_notes` (wait for pasted notes; "Draft from the attached filings + web" → the literal `NO_NOTES_ANALYST_NOTES`); **Targets** → `market_entry_target_count = market_entry_targets_from_slides(n)`; an override of "2 Financial Summary slides" → `metric_count_from_slides(2)`; **Highlights** → `include_investment_highlights = False` on "Omit" only; **CIM / EEO snip** → the saved path, or unset on "None"; **Valuation / Risk notes** → unset on "None", else the typed text (collect as a follow-up if promised but not supplied).

**Attachment-status answers are not plan inputs** — the consuming stages discover the saved files under `<deal_dir>/filings/`. "Will drop it next message" holds the run until it arrives; "None" proceeds with the placeholder path that stage already documents.

**Optional inputs the analyst didn't supply stay OUT of the dict — never pre-seed `None`.** The driver computes the optional set from the plan and resolves an unsupplied optional reference to `None` for you; a missing *required* input still halts. For a deliverable with no questionnaire (the renderers raise `ValueError`), prompt from `plan.plan_inputs` directly. Then persist once:

```python
write_plan_inputs(run_dir, plan_inputs)
```

## Step 5 — Run the waves

Stages run in dependency **waves** derived from the `$stages.*` references in each stage's inputs — the references *are* the DAG, and since Phase D that is the only rule (no hardcoded barriers). Post `plan_overview(run_dir).narration()` (stage list, plan inputs, gates, and the wave schedule) up front, then for each wave `n` in `1..overview.wave_count`:

1. **`dispatch = prepare_wave(run_dir, n)`** — resolves every reference, writes each `inputs.json`, and returns one `PreparedStage` per stage. A **judgment** stage carries a rendered `prompt`; a **transform** carries none.
2. **`run_transforms(dispatch)`** — executes the wave's in-process stages (`dispatch.transforms`) and writes their `outputs.json`. Returns one `TransformResult` each; a raising transform is recorded as a stage failure, not a crash. Call it for every wave — it is a no-op when the wave has none.
3. **Issue one `Task` call per stage in `dispatch.judgment`, all in a single message** so the wave runs concurrently, passing each `stage.prompt` **verbatim** as the prompt. `dispatch.prompts` is the same list. A wave of nothing but transforms needs **no `Task` call at all** — waves 5 and 6 of the pitch plan are exactly that. Wait for every sub-agent to return, then `write_stage_log(run_dir, stage_id, transcript)` for each.
4. **`outcome = complete_wave(run_dir, n)`** — reads and validates every `outputs.json`, whoever wrote it (missing, malformed, an `error` key, or a missing declared output name all fail), and builds the checkpoint payloads.
5. **Post `outcome.narration()` and go straight on to the next wave.** If `outcome.halt`, stop — a stage failed, so do not start the next wave. `outcome.gate` is `None` for every shipped plan; on the off chance a plan carries a `required` checkpoint, put `outcome.gate.question` to the analyst with `AskUserQuestion` and halt on `HALT_LABEL`.

**Transform vs judgment is not yours to decide.** `stage_transforms.TRANSFORMS` is the one place the classification lives; `prepare_wave` reads it. Never dispatch a `Task` for a stage in `dispatch.transforms`, and never hand-run a judgment stage's work yourself.

**Do not invent an approval pause.** No shipped stage is `required` (v0.5.49), so a wave boundary is a *report*, not a question: post the narration and dispatch the next wave in the same turn. Asking "shall I continue?" out of caution reinstates by hand the gate the plans deliberately dropped. The deck's QA is automated and unchanged — `deck_repair` converges the geometry from measured renders inside the assembler, and `deck`'s `vision_review_path` is the read-the-slides pass written to disk. **Name that path in the surface** so the analyst can open it while the run continues, and name it again in the summary.

## Step 6 — Summary

`write_run_summary(run_dir, notes=[...])` writes `summary.md` with every stage's status, its outputs, and the artefact paths. Pass as `notes` what only you know: manual next steps a sub-skill surfaced (e.g. "refresh the Capital IQ connector in the deal workbook"), an abort reason, or a caveat. Post it back to the analyst.

## Stop conditions

Halt — with a clear message, the partial run preserved on disk — when: the deliverable type or codename is missing and the analyst declines to supply it; an existing deal is detected and the analyst declines to continue or rename; the plan fails either validation layer; a reference cannot be resolved; or `complete_wave` reports any stage `ok=False` (a missing, malformed or `error`-carrying `outputs.json`, a missing declared output, or a transform that raised — a `DeckNotConvergedError` reaches you this way, naming the shape it could not fit). A **stage failure is not a checkpoint**: it halts the run whatever the plan's checkpoint modes say, and it is now the only thing that stops a run mid-flight, since no shipped stage is `required`. If some future plan does carry a `required` checkpoint, a rejected one halts too.

Never silently skip a stage. Never proceed past a missing output. Never overwrite an existing `outputs.json` — a re-run uses a new `run_id`.

## What the conductor does not do

- Produce slides, models, or copy — that is each sub-skill's job.
- Make banking decisions. Voice, brand, and source-trust rules live in each stage skill's own SKILL.md and its allow-list.
- Invent dependencies. The wave schedule comes purely from the references (`plan_schedule.compute_waves`); the conductor never adds edges, reorders, or dispatches a stage before everything it references has produced outputs.
- Emit telemetry beyond per-stage transcripts. `meta.json` (model, tokens, latency) is Phase 5.
