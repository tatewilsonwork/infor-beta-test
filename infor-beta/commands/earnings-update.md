---
description: Build a full INFOR earnings-update deck for a company — conductor-orchestrated, end to end
argument-hint: <company name>
---

Build a full INFOR **earnings update** deck for **$ARGUMENTS**, end to end, using this plugin's `conductor` skill (infor-beta). Invoke the conductor skill now and follow its workflow exactly, with these presets supplied by this command:

- **Deliverable type: `earnings-update`** (G7 item 2 — pre-answered, do not re-ask).
- **Subject company name: "$ARGUMENTS"** (G7 item 3 — pre-answered, do not re-ask). If no company name was given after the command, ask for it before anything else.
- Ask for the **codename** (G7 item 1). If the analyst hasn't got one, propose `Project <single word>` and confirm.
- Render the **G7 deal-init prompt** for the remaining items (public/private + ticker, sector, filings, notes) exactly as `deal_init.render_init_prompt()` emits it, minus the two pre-answered items.
- Then render the **earnings-update deck-spec questionnaire** verbatim — `deck_spec.render_deck_spec_prompt("earnings-update")` — in a single message (reporting/comparison quarters + the Bloomberg EEO snip; the deck itself is the fixed 5-slide layout).
- Proceed through the earnings-update plan's stages wave by wave as the conductor skill directs, and finish with the run summary.

Do not skip the questionnaire even if the analyst's message already contains some of the answers — render it once, pre-fill what is already known inline (marked "from your message"), and only ask for what is missing.
