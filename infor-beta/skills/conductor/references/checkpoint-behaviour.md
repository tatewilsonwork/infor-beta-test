# Checkpoint behaviour

Per locked decision A2 (note 12), every stage carries a `checkpoint` mode. The conductor's behaviour at each stage boundary depends on the mode:

## `required`

The conductor **halts** after the stage's `outputs.json` is read and waits for explicit analyst approval before dispatching the next stage.

Surface to the analyst:

> Stage `<stage_id>` (`<skill>`) finished. Outputs:
> - `deck_path`: /Users/tate/Documents/INFOR Deals/Project OpenText/artefacts/Earnings Update - OpenText.pptx
> - `workbook_path`: /Users/tate/Documents/INFOR Deals/Project OpenText/artefacts/OpenText - Capitalization Table.xlsx
>
> Review the file(s) above. Reply `approve` to continue to the next stage, or `stop` to halt the run.

If the analyst replies `stop`, the conductor writes a summary noting the abort point and exits cleanly. The partial run on disk is preserved — re-running the conductor creates a new `run_id` and starts over.

Use `required` for stages where the analyst's review is a real-world checkpoint, e.g.:
- buyer-list approval (D-series, note 12)
- the assembled deck before QA
- the final deck before it's shared externally

## `informational` (default)

The conductor surfaces a short summary of the stage's outputs and **proceeds to the next stage without waiting** for explicit approval. Analyst may interject if they want to halt, but the default behaviour is to continue.

Surface to the analyst:

> Stage `<stage_id>` (`<skill>`) finished. Outputs: `deck_path=...`, `workbook_path=...`. Proceeding to next stage.

Use `informational` for stages where the analyst wants visibility but the work is mechanical enough that a default-continue is appropriate. This is the v1 default per A2 (medium HITL — analyst confirms by *not* halting).

## `silent`

No surface. The conductor moves to the next stage without any message. Reserved for future autonomous mode — in Phase 2, `silent` is functionally equivalent to `informational` with the message suppressed.

Do not use `silent` in v1 unless the stage is genuinely uninteresting to the analyst (a no-op or a trivial mechanical translation). Defaults should err toward `informational`.

---

## Mode → run-summary impact

All three modes still produce the per-stage `log.txt`, `inputs.json`, `outputs.json` on disk. The mode only affects what the conductor says to the analyst in the chat surface between stages and at the end of the run.

The final `summary.md` lists every stage regardless of checkpoint mode.

---

## Migration to autonomous (Phase 2+ → eventually)

Per A2, the path to v2 autonomous mode is **configuration, not code**. To flip a plan from medium-HITL to autonomous:

1. Change every `required` checkpoint to `informational` (or `silent` for stages the analyst doesn't want to see).
2. Keep `required` only on the genuine legal/business gates that should always halt (buyer-list approval, final deck before external delivery).

No conductor code changes. No skill changes. The checkpoint mechanism is designed to live in the plan YAML so the same conductor runs both modes.
