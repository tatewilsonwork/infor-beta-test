# Plan schema — reference

A plan is a YAML file at `infor-beta/plans/<deliverable>.yaml`. It declares the stages the conductor will run for a deliverable (earnings update, pitch, etc.) plus the plan-specific inputs the analyst must supply after deal-init.

The canonical pydantic models live in `infor-beta/scripts/schemas/plan.py` (`Plan`, `Stage`, `InputSpec`, `OutputSpec`). JSON Schemas are under `scripts/schemas/json/plan.schema.json`.

## Top-level Plan fields

```yaml
deliverable_type: earnings-update   # must be one of DealContext.DeliverableType
description: One-line summary shown by the conductor in its plan summary.
plan_inputs:
  - name: reporting_quarter
    type: str
    required: true
    description: Optional human-readable hint shown to the analyst.
stages:
  - id: ...
    skill: ...
    ...
```

- `deliverable_type` must be one of: `pitch`, `earnings-update`, `overview`, `one-off-skill`.
- `plan_inputs` are collected by the conductor AFTER deal-init (which always asks the G7 7-field block — never repeat G7 questions here).
- `stages` is a non-empty list; ids must be unique within the plan.

## Stage fields

```yaml
- id: content
  skill: earningsupdate-content
  inputs:
    company: $deal.subject_company
    ticker: $deal.subject_company.ticker
    reporting_quarter: $plan_inputs.reporting_quarter
    eeo_snip_path: $plan_inputs.eeo_snip_path
  outputs:
    - name: deck_path
      type: Path
  checkpoint: informational
```

- `id` — unique within the plan. Lowercase + hyphens by convention (`ltm-metrics`, `financial-charts`); must not contain a dot (it would break `$stages.<id>.<name>` parsing).
- `skill` — must match a skill directory under `infor-beta/skills/`. The conductor dispatches the named skill via the `Task` (Agent) tool.
- `inputs` — dict passed to the sub-skill. Values may be literals or **reference strings** (see below).
- `outputs` — names + types the sub-skill is expected to write to its `outputs.json`. The conductor reads `outputs.json` after the stage finishes and exposes the named outputs as `$stages.<id>.<name>` to later stages. **Declared names are enforced at collect time** (`conductor.complete_wave`): every declared name must be present as a key, or the stage is reported failed naming the missing key(s). Presence only — a `null` value passes (the contract requires e.g. ltm-metrics to emit null, never omit, `ltm_revenue`/`ltm_adj_ebitda`), extra undeclared keys are allowed, and the `type` label stays documentation (no value-type check). Malformed (non-JSON) `outputs.json` is likewise converted to a stage failure rather than crashing the driver.
- `checkpoint` — `required` (halt, wait for analyst approval), `informational` (summarise and continue — the default), or `silent` (no surface). See `checkpoint-behaviour.md`.

## Reference strings

Any input value that is a whole string starting with `$` and matching one of these forms is resolved at dispatch time:

| Prefix | Resolves to |
|---|---|
| `$plan_inputs.<name>` | Value the analyst supplied for that plan input |
| `$deal.<field>` | Field on the DealContext (`codename`, `deal_dir`, `deliverable_type`, `subject_company`, `filings`, `notes`); dotted access works (`$deal.subject_company.ticker`) |
| `$stages.<stage_id>.<output_name>` | Named output of an earlier stage in the same plan |

Mid-string interpolation is **not** supported (no `"hello $deal.codename"`). Each value is either a literal or one whole reference.

**References are validated at load time.** `plan_refs.validate_plan_references(plan)` runs on every conductor load path (Step 3 and `conductor.load_plan`) and rejects the plan — before any stage is dispatched — when a `$stages.<id>` names no stage in the plan, a `$stages.<id>.<name>` names an output the stage doesn't declare, a `$plan_inputs.<name>` names no declared plan input, or a whole-string `$…` value matches no reference form at all. `$deal.<field>` is the one form left to resolve time (it is checked against the live DealContext). The pre-flight shares its grammar with the resolver and the wave scheduler (`plan_refs.parse_ref` / `iter_input_strings` — one grammar, no drift), so a plan that loads is a plan whose stage references both schedule and resolve.

Unknown prefix or unresolvable field at dispatch time still raises an error and halts the conductor — the partial run on disk is preserved.

**References also determine execution order.** A `$stages.<id>.<name>` reference is both a value lookup *and* a dependency edge: the conductor groups stages into ordered **waves** (`plan_schedule.compute_waves`) so every referenced stage runs in an earlier wave, and dispatches the stages within a wave concurrently. Stages that reference none of each other run in parallel. To force stage B to wait for stage A, have B reference one of A's outputs.

## What plans should NOT contain

- `depends_on` / `parallel_with` — there is deliberately **no** dependency field. The conductor derives the execution DAG from the `$stages.<id>.<name>` references in each stage's inputs and dispatches independent stages in concurrent waves (`plan_schedule.compute_waves`). To make stage B wait for stage A, reference one of A's outputs in B's inputs. Since Phase D there are no non-reference edges: the one that existed — `workbook-aggregator` depending on every stage except its own downstream consumers, because it merged and DELETED the individual workbooks — went with the aggregator.
- `cost_class`, `expected_duration`, `gate_mode_default`, `idempotent` — per-skill operational metadata, not a plan concern. Deferred to a future phase (the prototyped typed skill manifest was removed in 0.5.6; revisit when per-skill metadata is actually needed).
- Anything that re-asks a G7 question — those are owned by deal-init.
- Free-form prompt text for the sub-agent. Each stage gets the standardised envelope from `stage-envelope.md`.
