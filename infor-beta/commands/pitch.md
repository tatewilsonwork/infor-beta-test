---
description: Build a full INFOR pitch deck for a company — conductor-orchestrated, end to end
argument-hint: <company name>
---

Build a full INFOR **pitch** deck for **$ARGUMENTS**, end to end, using this plugin's `conductor` skill (infor-beta). Invoke the conductor skill now and follow its workflow exactly, with these presets supplied by this command:

- **Deliverable type: `pitch`** (G7 item 2 — pre-answered, do not re-ask).
- **Subject company name: "$ARGUMENTS"** (G7 item 3 — pre-answered, do not re-ask). If no company name was given after the command, ask for it before anything else.
- Ask for the **codename** (G7 item 1). If the analyst hasn't got one, propose `Project <single word>` and confirm.
- Render the **G7 deal-init prompt** for the remaining items (public/private + ticker, sector, filings, notes) exactly as `deal_init.render_init_prompt()` emits it, minus the two pre-answered items.
- Then render the **pitch deck-spec questionnaire** verbatim — `deck_spec.render_deck_spec_prompt("pitch")` — in a single message, and map the answers to plan inputs per the conductor skill's Step 4 (slide-count answers convert via `deck_spec.metric_count_from_slides` / `deck_spec.market_entry_targets_from_slides`; "omit" Key Investment Highlights → `include_investment_highlights: false`).
- Proceed through the pitch plan's stages wave by wave as the conductor skill directs, and finish with the run summary.

Do not skip the questionnaire even if the analyst's message already contains some of the answers — render it once, pre-fill what is already known inline (marked "from your message"), and only ask for what is missing.
