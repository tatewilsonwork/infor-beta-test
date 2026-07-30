# Checkpoint behaviour

Per locked decision A2 (note 12), every stage carries a `checkpoint` mode. Since v0.5.13 stages run in dependency **waves** (conductor step 6), so checkpoints are evaluated at the **wave boundary**, after every stage in the wave has already run: a `required` gate stops the downstream waves, not its own wave-mates. (To make a gate stop work *before* an expensive stage starts, give that stage a dependency so it lands in a later wave.) The conductor's behaviour at each wave boundary depends on the mode:

**Since v0.5.49 every stage of every shipped plan is `informational`**, so a run never pauses for an approval — the A2 flip to autonomous has been made. `required` is documented first because it is the mode with behaviour to describe, not because anything uses it.

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

Reserve `required` for a stage whose approval is a real-world event that has to happen before the work continues — buyer-list approval (D-series, note 12) is the archetype: sending the list *is* the next step, and nobody but the analyst can authorise it. A gate is not the right tool for "check my work"; that is what an automated QA pass plus a written review is for.

**Shipped usage: none.** Since v0.5.49 no shipped plan carries a `required` checkpoint. The analyst gate on `deck` is gone, so a run goes from intake to `deckcheck` without pausing, and `plan_overview(...).narration()` reports `Required checkpoints: none` for every shipped plan. The mode itself is unchanged, still supported, and still tested (`test_the_required_gate_halts_the_run_when_the_analyst_rejects`) — it is simply unused.

What replaced the gate is not an absence of review but an automated one, none of which asks the analyst anything:

- `deck_repair`'s converge loop, inside the assembler, decides every font size and autofit scale from a **measured** render and raises `DeckNotConvergedError` if a shape will not fit.
- `deck`'s `vision_review_path` is the read-the-slides **checklist** written to disk — which slides to look at and why, every slide's render, and each rasterised picture at native resolution. It is an agenda, not a review: it asks the questions and answers none.
- `deckread` answers them, on the finished artefact, and reports what it **saw** — text over text, colliding shapes, faint contrast, label pileup, an illegible pasted range. Advisory by construction. It exists because the checklist above had no reader for four releases: nothing referenced it, so a defect sitting on a slide whose pictures it listed by name went out with a clean run.
- `deckcheck` attempts to **disprove** every figure on the finished artefact and reports what it could not confirm. Also advisory by construction; both reading stages could never have gated.

All four land in the run summary, so the analyst reads them on a finished deliverable instead of being stopped mid-run to approve an artefact the later waves have not touched yet. What still halts a run is a stage *failure* — `complete_wave` reporting `ok=False`, which is not a checkpoint at all.

**A gate on an in-process stage would behave identically.** Since Phase F, `deck` is a transform: the driver assembles the deck itself instead of dispatching a sub-agent. The checkpoint is built from the stage's `outputs.json` by `complete_wave`, which cannot tell who wrote it, so a gate on a transform fires at the same boundary with the same locked dialog and holds the same downstream waves. Nothing about removing the `deck` gate changed that.

## `informational` (default)

The conductor surfaces a short summary of the stage's outputs and **proceeds to the next wave without waiting** for explicit approval (batch a wave's `informational` summaries into one surface). Analyst may interject if they want to halt, but the default behaviour is to continue.

Surface to the analyst:

> Stage `<stage_id>` (`<skill>`) finished. Outputs: `deck_path=...`, `workbook_path=...`. Proceeding.

Use `informational` for stages where the analyst wants visibility but nothing about the stage needs an authorisation. Since v0.5.49 this is the mode **every** shipped stage carries: the analyst watches the run go by and reads the reviews at the end, and can still interject to stop it.

## `silent`

No surface. The conductor moves on without any message — `informational` with the message suppressed. A **failed** stage surfaces anyway (`_checkpoint_for` checks `ok` before the mode); `silent` was never meant to hide a failure.

Autonomous did **not** turn out to mean `silent`: v0.5.49 made every shipped stage `informational`, because a run nobody has to approve is exactly the run where the analyst needs the artefact paths going by. No stage uses `silent`, and none should unless it is genuinely uninteresting (a no-op or a trivial mechanical translation).

---

## Mode → run-summary impact

All three modes still produce the per-stage `log.txt`, `inputs.json`, `outputs.json` on disk. The mode only affects what the conductor says to the analyst in the chat surface between waves and at the end of the run.

The final `summary.md` lists every stage regardless of checkpoint mode.

---

## Autonomous — done, in v0.5.49

Per A2 the path from medium-HITL to autonomous was **configuration, not code**, and that is how it happened: three `checkpoint: required` lines became `informational` (`pitch.yaml`, `earnings-update.yaml`, `overview.yaml`) and a conductor run now goes end to end without waiting on an analyst. No conductor code changed, no skill changed, and the `required` branch of `_checkpoint_for` — plus `WaveOutcome.gate`, `APPROVE_LABEL` and `HALT_LABEL` — is still there for a future plan that has a genuine authorisation step to gate on.

The QA the gate used to sit in front of did **not** go with it: the converge loop, the written vision-review checklist, the `deckread` reading pass that answers it and the `deckcheck` falsification pass all still run, and none of them asks a question. See "Shipped usage" above.

Two things a run still stops for, neither of them a checkpoint: a stage **failure** (`complete_wave` reporting `ok=False` — a missing, malformed or `error`-carrying `outputs.json`, including a raising transform), and the **intake** before wave 1, where the analyst supplies the facts, the attachments and the deck spec.
