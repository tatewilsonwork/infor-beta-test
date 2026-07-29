---
name: conductor
description: >
  Use this skill when the user wants to build a complete INFOR deliverable end-to-end —
  pitch, earnings update, or overview deck — instead of
  invoking individual skills. Activates on "build a <deliverable>", "kick off
  <deliverable>", "conductor", "/conductor", "orchestrate", the /pitch and
  /earnings-update commands (which preset the deliverable type + subject company), or
  any request that names a deliverable type rather than a single workflow step. The
  conductor handles deal-init, collects the whole locked questionnaire — deal-init's
  questions plus the deliverable's — in ONE interactive AskUserQuestion call per
  slash-command run, loads the plan YAML
  for the deliverable, runs each deterministic stage in-process and dispatches each
  judgment stage to its skill via the Agent tool with a file-based input / output handoff,
  and emits a run log under <deals root>/<codename>/runs/<run-id>/, where the deals root is
  resolved and reported at deal-init (codename.resolve_deals_root) rather than assumed.
allowed-tools: [Read, Write, Bash, Glob, Task, AskUserQuestion]
---

# Conductor — Workflow

The conductor is a thin orchestrator: **dumb about banking, smart about orchestration**. It never produces a deliverable directly.

Since Phase E the mechanics live in `scripts/conductor.py`, not in this file. **Your job is four things** — intake, issuing the `Task` calls the driver hands back, reporting each wave boundary, and the summary. Everything else is a function call with a return value, so it cannot be skipped on turn 40 of a long run.

Since v0.5.49 **no shipped plan asks the analyst to approve anything mid-run**, and since v0.5.51 the intake is **one `AskUserQuestion` call** on a slash-command run. So a `/pitch` build is: one dialog, one attachment request, one pause, then every wave to the end and a report. Nothing to approve does not mean nothing is checked — the geometry converge loop, the written vision review and the `deckcheck` falsification pass all still run, and none of them needs an answer.

Since Phase F the second of those got smaller: a stage is either a **transform** (deterministic — the driver calls the function in-process) or **judgment** (research and drafting — a sub-agent with a real allow-list; `deckcheck` included). You dispatch the judgment stages only. `plan_overview(run_dir).narration()` reports the split for the plan in front of you; no number is written down here, because a written-down wave count went stale once already.

The architectural backbone — DealContext schema, codename rules, deliverable types, three checkpoint modes — is locked in Obsidian note `12 — Locked Decisions`. Re-read note 12 H1–H8 before changing this skill's behaviour.

**Detailed references** (loaded on demand):
- [`references/plan-schema.md`](references/plan-schema.md) — Plan / Stage YAML schema and reference-string semantics (`$plan_inputs`, `$deal`, `$stages`).
- [`references/stage-envelope.md`](references/stage-envelope.md) — the prompt template `conductor.py` renders per stage. **No environment variables:** the handoff paths are arguments.
- [`references/checkpoint-behaviour.md`](references/checkpoint-behaviour.md) — what `required` / `informational` / `silent` mean operationally.

```python
import sys, os
sys.path.insert(0, os.environ.get("CLAUDE_PLUGIN_ROOT", "./infor-beta") + "/scripts")

from codename import (
    codename_from_company, disambiguate,
    resolve_deals_root,   # WHERE the deals live — discovered, never hardcoded
    split_listing,        # "Open Text (TSX:OTEX)" -> ("Open Text", Listing("TSX", "OTEX"))
)
from deal_init import (
    INIT_DIALOG_FIELDS, INIT_DEFAULT_FIELDS,
    render_init_dialogs, render_init_prompt,   # deal-init's half alone — generic entry only
    load_or_locate_deal, save_deal_context,
)
from deck_spec import (
    render_run_dialogs,                     # THE dialog — ONE call per run
    render_run_attachment_request,          # THE attachment message — one per run
    render_run_defaults,                    # THE defaults echo — one per run
    render_deck_spec_dialogs, render_deck_spec_prompt,   # the deliverable's half alone
    default_presentation_date, prior_year_quarter,
    metric_count_from_slides, market_entry_targets_from_slides, NO_NOTES_ANALYST_NOTES,
    PITCH_DIALOG_PLAN_INPUTS, EARNINGS_UPDATE_DIALOG_PLAN_INPUTS,
    PITCH_ATTACHMENT_PLAN_INPUTS, EARNINGS_UPDATE_ATTACHMENT_PLAN_INPUTS,
)
from run_log import make_run_id, create_run_dir, write_plan_snapshot, write_stage_log
from conductor import (      # the driver — everything mechanical
    plan_overview, prepare_wave, run_transforms, complete_wave,
    write_plan_inputs, write_run_summary, APPROVE_LABEL, HALT_LABEL,
)
```

> **Interactive UI.** Every analyst-facing question goes through **`AskUserQuestion`**, never a numbered text block. The payloads are code-owned (`render_run_dialogs`; and `WaveOutcome.gate.question` if a plan ever carries a `required` checkpoint again): render them **verbatim** — do not paraphrase, reorder, re-option, or invent extra questions. Pure free-text facts with nothing to suggest (the subject company name) stay plain chat questions. If `AskUserQuestion` is unavailable, fall back to the locked text prompts (`render_init_prompt()` **then** `render_deck_spec_prompt(...)` — two prompts, each numbering its own half; plus `Checkpoint.fallback_prompt`) — same items, same order as the one dialog.
>
> **ONE dialog per run** (v0.5.51). `render_run_dialogs(<deliverable>)` returns a **single** `AskUserQuestion` payload: it merges deal-init's questions with the deliverable's, so `/pitch` asks four (`Listing`, `Notes`, `Targets`, `Highlights`) and `/earnings-update` asks one (`Listing`). Post it once, in Step 2, and never split it — a pitch run walked the analyst through three sequential dialogs through v0.5.50. **The 4-question cap is the tool's own**, so there is no headroom: if you find yourself wanting another question, the answer is a default the analyst overrides by replying, not a second dialog. **Generic entry is the one exception**: with no deliverable named, the `Deliverable` answer is what decides which deck spec exists, so there is nothing to merge — ask `render_init_dialogs(include_deliverable=True)`, then `render_deck_spec_dialogs(<answer>)`. Two calls, expected, not a defect.
>
> **Ask NOTHING about attachments.** No dialog, no status question, no "is it attached yet?" — not for the filings, not for the SEDI report, not for the CIM or the EEO snip. The dialog comes first; then you post `render_run_attachment_request(<deliverable>)` **once**, as plain text, and **wait once** for the analyst to drop the files into chat. That single pause replaced three dialogs and three pauses on a pitch run, each of which asked the analyst to assert something `<deal_dir>/filings/` already knew. Never hand-write the list or a bullet of it: it is generated from the same specs the dialogs are, so a document that is asked for is a document whose consequence-if-missing is stated. "None for now" and "Not applicable" have no dialog to live in — the analyst says it in chat, or you proceed after the drop with whatever arrived.
>
> **Defaults are echoed, never asked.** `render_run_defaults(...)` is **one** message covering deal-init's defaults and the deliverable's — the researched sector, the client name, the presentation date, the quarters, the valuation range, the risk notes, the slide-mix fallbacks. Post it once (Step 4), never one echo per spec, and never turn a defaulted item back into a question to be safe: the analyst overrides by replying.

## Step 1 — Deliverable + codename

Extract the **deliverable type** (`pitch` / `earnings-update` / `overview` / `one-off-skill`; ask if ambiguous — `overview` is a stub, say so if selected) and the **codename**. Never ask for the codename: use the analyst's `Project <target>` string if they typed one, else derive it silently with `codename_from_company(<subject company>)` and state it when announcing the deal directory (overridable in chat until the directory is created). `/pitch` and `/earnings-update` pre-answer the deliverable **and** the company — do not re-ask either. A **one-off skill** needs no plan: say so and stop.

**Split the listing hint off the company string first.** `/pitch Open Text (TSX:OTEX)` is the documented invocation and that parenthetical is a listing, not part of the name:

```python
company_name, listing = split_listing(company_arg)   # ("Open Text", Listing("TSX", "OTEX"))
codename = codename_from_company(company_arg)        # "Project Open Text" — splits it itself
```

Use `company_name` as `subject_company.legal_name`, and when `listing` is not None **the Listing question is already answered**: set `subject_company.exchange = listing.exchange` and `subject_company.ticker = listing.ticker` (`listing.capiq` is the Capital IQ `Exchange:Ticker` render), then drop `Listing` from Step 2's dialog via `omit=("Listing",)` and note "(from your message: TSX:OTEX)". A trailing parenthetical with **no colon** in it is part of the name — `"Acme (Canada)"` stays whole and `listing` is None, so the Listing question still gets asked. Never re-derive the codename from a string you have already stripped: `codename_from_company` takes the raw argument.

## Step 2 — Deal-init + the run's one dialog

**Resolve the deals root before anything else, and state it.** The E1 default `~/Documents/INFOR Deals` **does not exist on the production runtime** — real deals live under the mounted workspace folder (`$HOME/mnt/<mounted folder>/INFOR Deals/`, `$HOME` being `/sessions/<session>`), and both variable parts mean it cannot be hardcoded. Discover it, report it, then pass it to everything:

```python
root = resolve_deals_root()          # or resolve_deals_root(<path>) if the analyst named one
print(root.describe())               # "Deals root: … — discovered under $HOME/mnt/, holding 11 existing deals."
ctx_or_none, deal_dir = load_or_locate_deal(codename, deals_root=root)
```

`root.describe()` is a **reported decision**: post that line before you write anything, so a wrong root is caught by the analyst in one message rather than after a deck lands somewhere they cannot see. `resolve_deals_root` prefers a root that already **holds deals** over an empty one and returns the runners-up in `root.alternatives` — if the analyst names one of those, or any other path, re-resolve with `resolve_deals_root(<their path>)` and pass that. Never assume `~/Documents/INFOR Deals`, never build a deals path by string concatenation, and never call the `deals_root=`-taking helpers (`load_or_locate_deal`, `find_existing`, `disambiguate`, `resolve`) without passing the resolved root — the whole point of stating it is that every subsequent lookup agrees with it.

An existing deal means one `AskUserQuestion` ("Continue `<codename>`" / "Different deal"; on the latter, present `disambiguate(root, codename)`'s 1–4 alternatives as another dialog).

A fresh deal means **`render_run_dialogs(<deliverable>)` — a single `AskUserQuestion` call, verbatim**, carrying deal-init's questions *and* the deliverable's. This is the whole questionnaire: `/pitch` → `[Listing, Notes, Targets, Highlights]`, `/earnings-update` → `[Listing]`. Pass `omit=(...)` to drop any question an earlier message already answered, and note "(from your message: …)" for each one you drop. If the company name is not preset, ask it as a plain chat question — it is free text with nothing to suggest, so it is in no dialog. Then:

- **A `Listing` from Step 1's `split_listing`** → the question was already dropped; do not ask it again, and do not "confirm" the ticker.
- **"Public — I'll give the ticker" with no ticker** → ask ticker + exchange as a follow-up.
- **The sector is not asked** (v0.5.51). Research it, verify by web search, set `subject_company.sector` / `.industry` from the one-liner, and let the Step 4 echo report it — `INIT_DEFAULT_FIELDS` is the table. No confirmation, and no dialog question: its old question already defaulted to "Infer from the web".
- Build the `DealContext` and `save_deal_context(ctx)` — which creates the deal directory, so `filings/` exists before anything is dropped. Hold the deck-spec answers; they become `plan_inputs` in Step 4.

**Generic entry only** (Step 1 could not determine the deliverable): there is nothing to merge until `Deliverable` is answered, so ask `render_init_dialogs(include_deliverable=True)` and then `render_deck_spec_dialogs(<answer>)`. Two calls is correct here.

**Do not ask about the filings here, and do not post the request yet.** The G7 filings are a REQUIRED bullet of the one attachment request, which goes out in Step 4 — after every question has been answered — so the analyst answers once and then attaches everything in one go.

**Filings handling** (whenever files arrive, including a pre-attached message): save every attachment under `<deal_dir>/filings/` with a descriptive name, append matching `Filing` entries to `ctx.filings`, and re-save `deal.json`. This applies to the deliverable's documents (SEDI PDF, Bloomberg export, CIM, EEO snip) too — same directory, same `Filing` entries.

## Step 3 — Plan + run directory

Read `${CLAUDE_PLUGIN_ROOT}/plans/<deliverable>.yaml`, then:

```python
run_id = make_run_id(plan.deliverable_type)
run_dir = create_run_dir(ctx.deal_dir, run_id)
write_plan_snapshot(run_dir, plan_yaml_text)   # frozen snapshot; the driver reads only this
```

Tell the analyst the run id and its path. `plan_overview(run_dir)` (Step 5) validates the snapshot in two layers — the pydantic `Plan` shape, then the reference pre-flight `validate_plan_references` — so a typo'd `$stages` / `$plan_inputs` reference is dead at load, not mid-run. Surface any error and stop; never run a partially-valid plan.

## Step 4 — One attachment request, one wait, one defaults echo

Three moves, in this order. The questions are already answered — they were the single dialog in Step 2 — so nothing here asks anything.

1. **Post `render_run_attachment_request(<deliverable>)`** — one plain-text message, generated, listing every document the run needs under REQUIRED and OPTIONAL. It merges deal-init's filings with the deliverable's own, so this is the *only* place attachments are raised in the whole run.
2. **Wait here. Once.** There is no pause per document, because there is no dialog per document. Post the request and stop, on one turn, until the analyst's next message. Their reply may carry the files and any promised text (pasted notes, a valuation range, specific risks) together — take all of it from that one message. If they say a document does not exist or to skip it, proceed with what arrived; each bullet already told them what that costs. Save everything per Step 2's filings handling.
3. **Then compute the defaults and post `render_run_defaults(...)` — once, merged.** After the drop, not before: `reporting_quarter` is inferred from the **latest attached interim filing** (fiscal quarter labels depend on the company's fiscal calendar, not the calendar date, so never compute it from today's date), and `comparison_quarter` ← `prior_year_quarter(...)` follows it. `sector` ← the one-liner you researched in Step 2; `client_name` ← the subject company; `presentation_date` ← `default_presentation_date(date.today())`; `valuation_range`, `risk_notes`, `financial_metric_count` and `section_labels` ← left unset, so the content stage and the wireframe apply their own defaults. One echo covers deal-init's defaults and the deliverable's; never post two. If **no interim filing arrived**, the quarter cannot be inferred and it is a required input: ask for it as a plain chat question. Don't guess it.

**An override can arrive either side of the echo** — volunteered with the attachment drop, or as a reply to the echo itself — and converts exactly the way a dialog answer would. **Valuation range / Risk notes** → the typed text, else the key stays out of the dict entirely. Take an override from whichever message carried it; do not re-ask for one because the echo had not gone out yet, and do not wait for a reply to the echo (no shipped stage gates).

Map dialog answers with `PITCH_DIALOG_PLAN_INPUTS` / `EARNINGS_UPDATE_DIALOG_PLAN_INPUTS` (and `INIT_DIALOG_FIELDS` for deal-init's). Conversions are deterministic, never improvised: **Notes** → `analyst_notes` ("Draft from the attached filings + web" → the literal `NO_NOTES_ANALYST_NOTES`); **Targets** → `market_entry_target_count = market_entry_targets_from_slides(n)`; an override of "2 Financial Summary slides" → `metric_count_from_slides(2)`; **Highlights** → `include_investment_highlights = False` on "Omit" only.

**Resolve the path-carrying attachments from the files you saved**, through `PITCH_ATTACHMENT_PLAN_INPUTS` / `EARNINGS_UPDATE_ATTACHMENT_PLAN_INPUTS` (`{plan input: the document's label}`) — you named the files, so you know which is which. Two rules, and they differ:
- **Optional** (the pitch CIM → `cim_path`): no file, no key. Leave it **out** of the dict.
- **REQUIRED** (the EEO snip → `eeo_snip_path`): no file, no run. **Halt** before Step 5 and say plainly which document is missing and where it goes — "the Bloomberg EEO snip never arrived; drop it in chat and I'll restart the run", not a reference-resolution error naming `$plan_inputs.eeo_snip_path`. Never substitute a placeholder path, and never let it fall through to the driver.

Every other attachment carries no plan input at all — the consuming stage finds it under `<deal_dir>/filings/` itself, and works from the placeholder its own SKILL.md documents when it is not there.

**Optional inputs the analyst didn't supply stay OUT of the dict — never pre-seed `None`.** The driver computes the optional set from the plan and resolves an unsupplied optional reference to `None` for you; a missing *required* input still halts. For a deliverable with no questionnaire, prompt from `plan.plan_inputs` directly — only `render_deck_spec_dialogs` / `render_deck_spec_prompt` raise there. The three run-level renderers **degrade** instead: `render_run_dialogs` returns deal-init's questions alone, `render_run_attachment_request` the filings alone, `render_run_defaults` the sector alone, so the deal still gets its listing, its documents and its echo. Then persist once:

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

Halt — with a clear message, the partial run preserved on disk — when: the deliverable type or codename is missing and the analyst declines to supply it; an existing deal is detected and the analyst declines to continue or rename; **a REQUIRED attachment carrying a plan input never arrives** (name the document, not the reference); the plan fails either validation layer; a reference cannot be resolved; or `complete_wave` reports any stage `ok=False` (a missing, malformed or `error`-carrying `outputs.json`, a missing declared output, or a transform that raised — a `DeckNotConvergedError` reaches you this way, naming the shape it could not fit). A **stage failure is not a checkpoint**: it halts the run whatever the plan's checkpoint modes say, and once the run is under way it is the only thing that stops it, since no shipped stage is `required`. If some future plan does carry a `required` checkpoint, a rejected one halts too.

Never silently skip a stage. Never proceed past a missing output. Never overwrite an existing `outputs.json` — a re-run uses a new `run_id`.

## What the conductor does not do

- Produce slides, models, or copy — that is each sub-skill's job.
- Make banking decisions. Voice, brand, and source-trust rules live in each stage skill's own SKILL.md and its allow-list.
- Invent dependencies. The wave schedule comes purely from the references (`plan_schedule.compute_waves`); the conductor never adds edges, reorders, or dispatches a stage before everything it references has produced outputs.
- Emit telemetry beyond per-stage transcripts. `meta.json` (model, tokens, latency) is Phase 5.
