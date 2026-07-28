# Checkpoint behaviour

Per locked decision A2 (note 12), every stage carries a `checkpoint` mode. Since v0.5.13 stages run in dependency **waves** (conductor step 6), so checkpoints are evaluated at the **wave boundary**, after every stage in the wave has already run: a `required` gate stops the downstream waves, not its own wave-mates. (To make a gate stop work *before* an expensive stage starts, give that stage a dependency so it lands in a later wave.) The conductor's behaviour at each wave boundary depends on the mode:

## `required`

The conductor **halts** after the stage's `outputs.json` is read and waits for explicit analyst approval before dispatching the next wave.

Surface the outputs and the approval question through the **`AskUserQuestion` tool** (one question; the UI adds an "Other" free-text box for comments automatically):

> Stage `<stage_id>` (`<skill>`) finished. Outputs:
> - `deck_path`: /Users/tate/Documents/INFOR Deals/Project OpenText/artefacts/Earnings Update - OpenText.pptx
> - `workbook_path`: /Users/tate/Documents/INFOR Deals/Project OpenText/artefacts/OpenText - Capitalization Table.xlsx
>
> Review the file(s) above, then answer the approval dialog — options **"Approve — continue the run"** / **"Halt the run"**.

(If `AskUserQuestion` is unavailable on the current surface, fall back to the plain-text form: "Reply `approve` to continue the run, or `stop` to halt it.")

If the analyst halts, the conductor writes a summary noting the abort point and exits cleanly. The partial run on disk is preserved — re-running the conductor creates a new `run_id` and starts over.

Use `required` for stages where the analyst's review is a real-world checkpoint, e.g.:
- buyer-list approval (D-series, note 12)
- the assembled deck before QA
- the final deck before it's shared externally

**Shipped usage:** both shipped plans (`pitch.yaml`, `earnings-update.yaml`) mark the `deck` stage `required` — the pre-delivery gate, and the plugin's only one. The assembled deck is built from untrusted external inputs (filings, exports, web data), so the analyst approves it before delivery. `deck` is scheduled alone in its wave in both plans, so the wave-boundary gate genuinely holds what follows: the charted deck for pitch, the `deckcheck` falsification pass for both. Every other shipped stage is `informational`, `deckcheck` included — an advisory pass that could halt a run would have to be right about a target's financial statements.

**A gate on an in-process stage behaves identically.** Since Phase F, `deck` is a transform: the driver assembles the deck itself instead of dispatching a sub-agent. The checkpoint is built from the stage's `outputs.json` by `complete_wave`, which cannot tell who wrote it, so the gate fires at the same boundary with the same locked dialog and holds the same downstream waves. Its `vision_review_path` output is the review of the slides the analyst is being asked to approve — surface it by name.

## `informational` (default)

The conductor surfaces a short summary of the stage's outputs and **proceeds to the next wave without waiting** for explicit approval (batch a wave's `informational` summaries into one surface). Analyst may interject if they want to halt, but the default behaviour is to continue.

Surface to the analyst:

> Stage `<stage_id>` (`<skill>`) finished. Outputs: `deck_path=...`, `workbook_path=...`. Proceeding.

Use `informational` for stages where the analyst wants visibility but the work is mechanical enough that a default-continue is appropriate. This is the v1 default per A2 (medium HITL — analyst confirms by *not* halting).

## `silent`

No surface. The conductor moves on without any message. Reserved for future autonomous mode — in Phase 2, `silent` is functionally equivalent to `informational` with the message suppressed.

Do not use `silent` in v1 unless the stage is genuinely uninteresting to the analyst (a no-op or a trivial mechanical translation). Defaults should err toward `informational`.

---

## Mode → run-summary impact

All three modes still produce the per-stage `log.txt`, `inputs.json`, `outputs.json` on disk. The mode only affects what the conductor says to the analyst in the chat surface between waves and at the end of the run.

The final `summary.md` lists every stage regardless of checkpoint mode.

---

## Migration to autonomous (Phase 2+ → eventually)

Per A2, the path to v2 autonomous mode is **configuration, not code**. To flip a plan from medium-HITL to autonomous:

1. Change every `required` checkpoint to `informational` (or `silent` for stages the analyst doesn't want to see).
2. Keep `required` only on the genuine legal/business gates that should always halt (buyer-list approval, final deck before external delivery).

No conductor code changes. No skill changes. The checkpoint mechanism is designed to live in the plan YAML so the same conductor runs both modes.
