---
description: Build a full INFOR pitch deck for a company — conductor-orchestrated, end to end
argument-hint: <company name>
---

Build a full INFOR **pitch** deck for **$ARGUMENTS**, end to end, using this plugin's `conductor` skill (infor-beta). Invoke the conductor skill now and follow its workflow exactly, with these presets supplied by this command:

- **Deliverable type: `pitch`** (G7 item 2 — pre-answered, do not re-ask).
- **Subject company name: "$ARGUMENTS"** (G7 item 3 — pre-answered, do not re-ask). If no company name was given after the command, ask for it before anything else.
- Collect every remaining analyst input through the **interactive question UI** (`AskUserQuestion`), never as a numbered text block — the conductor skill's Steps 2 and 4 spell out the flow:
  - **Deal-init:** derive the codename silently — `codename.codename_from_company("$ARGUMENTS")`, analyst-overridable in chat — then render the `deal_init.render_init_dialogs()` dialogs verbatim (public/private + ticker, sector; the deliverable question is dropped — it is preset). Ask nothing about the filings here.
  - **Deck spec:** render the `deck_spec.render_deck_spec_dialogs("pitch")` dialogs verbatim (analyst notes, valuation range, risk notes, market-entry slides, Key Investment Highlights). Answers convert per the conductor skill's Step 4 (`market_entry_targets_from_slides`; "Omit" Key Investment Highlights → `include_investment_highlights: false`; the no-notes choice → `NO_NOTES_ANALYST_NOTES`).
  - **Attachments — one message, one wait, and no question about them.** Once the dialogs are answered, post `deck_spec.render_run_attachment_request("pitch")` as plain text (the G7 filings, the CIM, the SEDI PDF, the Bloomberg export — REQUIRED and OPTIONAL, each bullet stating what a missing file costs) and wait for the analyst to drop the files into chat. Save them under `<deal_dir>/filings/`, resolve `cim_path` from the saved CIM if one arrived (leave it unset otherwise), and only **then** infer the reporting quarter from the latest attached interim filing and post `deck_spec.render_deck_spec_defaults(...)` for override.
- Proceed through the pitch plan's stages wave by wave as the conductor skill directs, and finish with the run summary.

Do not skip the dialogs even if the analyst's message already contains some of the answers — drop just the pre-answered questions from the payloads, note "(from your message: …)", and ask the rest. If `AskUserQuestion` is unavailable on this surface, fall back to the locked text prompts (`render_init_prompt()` / `render_deck_spec_prompt("pitch")`).
