---
description: Build a full INFOR pitch deck for a company — conductor-orchestrated, end to end
argument-hint: <company name>
---

Build a full INFOR **pitch** deck for **$ARGUMENTS**, end to end, using this plugin's `conductor` skill (infor-beta). Invoke the conductor skill now and follow its workflow exactly, with these presets supplied by this command:

- **Deliverable type: `pitch`** (G7 item 2 — pre-answered, do not re-ask).
- **Subject company name: "$ARGUMENTS"** (G7 item 3 — pre-answered, do not re-ask). If no company name was given after the command, ask for it before anything else.
- Collect every remaining analyst input through the **interactive question UI** (`AskUserQuestion`), never as a numbered text block — the conductor skill's Steps 2 and 4 spell out the flow:
  - **Deal-init:** render the `deal_init.render_init_dialogs()` dialogs verbatim (codename, public/private + ticker, sector, notes; the deliverable question is dropped — it is preset), post `render_init_filings_note()` as plain text for the attachments.
  - **Deck spec:** compute the defaults (client name = the company; presentation date = current month + year; reporting/comparison quarters inferred from the latest interim filing), post `deck_spec.render_deck_spec_defaults(...)` so the analyst can override by replying, then render the `deck_spec.render_deck_spec_dialogs("pitch")` dialogs verbatim (analyst notes, CIM, valuation range, risk notes, market-entry slides, Key Investment Highlights) and post `render_deck_spec_documents_note("pitch")` for the attachments. Answers convert per the conductor skill's Step 4 (`market_entry_targets_from_slides`; "Omit" Key Investment Highlights → `include_investment_highlights: false`; the no-notes choice → `NO_NOTES_ANALYST_NOTES`).
- Proceed through the pitch plan's stages wave by wave as the conductor skill directs, and finish with the run summary.

Do not skip the dialogs even if the analyst's message already contains some of the answers — drop just the pre-answered questions from the payloads, note "(from your message: …)", and ask the rest. If `AskUserQuestion` is unavailable on this surface, fall back to the locked text prompts (`render_init_prompt()` / `render_deck_spec_prompt("pitch")`).
