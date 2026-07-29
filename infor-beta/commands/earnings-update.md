---
description: Build a full INFOR earnings-update deck for a company — conductor-orchestrated, end to end
argument-hint: <company name>
---

Build a full INFOR **earnings update** deck for **$ARGUMENTS**, end to end, using this plugin's `conductor` skill (infor-beta). Invoke the conductor skill now and follow its workflow exactly, with these presets supplied by this command:

- **Deliverable type: `earnings-update`** (G7 item 2 — pre-answered, do not re-ask).
- **Subject company name: "$ARGUMENTS"** (G7 item 3 — pre-answered, do not re-ask). If no company name was given after the command, ask for it before anything else.
- Collect every remaining analyst input through the **interactive question UI** (`AskUserQuestion`), never as a numbered text block — the conductor skill's Steps 2 and 4 spell out the flow:
  - **Deal-init:** derive the codename silently — `codename.codename_from_company("$ARGUMENTS")`, analyst-overridable in chat — then render the `deal_init.render_init_dialogs()` dialogs verbatim (public/private + ticker, sector; the deliverable question is dropped — it is preset). Ask nothing about the filings here.
  - **Deck spec: there is nothing to ask.** `deck_spec.render_deck_spec_dialogs("earnings-update")` is empty — both quarters are defaulted and the EEO snip is an attachment — so render no dialog for it. The deck itself is the fixed 5-slide layout, no slide options.
  - **Attachments — one message, one wait, and no question about them.** Post `deck_spec.render_run_attachment_request("earnings-update")` as plain text (the G7 filings and the Bloomberg EEO snip, both REQUIRED) and wait for the analyst to drop the files into chat. Save them under `<deal_dir>/filings/`, then resolve `eeo_snip_path` from the saved snip. It is a **required** plan input: if it never arrives, halt and say which document is missing — never a placeholder path. Only after the drop, infer the reporting quarter from the latest attached interim filing, set comparison = `prior_year_quarter(...)`, and post `deck_spec.render_deck_spec_defaults("earnings-update", ...)` for override.
- Proceed through the earnings-update plan's stages wave by wave as the conductor skill directs, and finish with the run summary.

Do not skip the deal-init dialogs even if the analyst's message already contains some of the answers — drop just the pre-answered questions from the payloads, note "(from your message: …)", and ask the rest. If `AskUserQuestion` is unavailable on this surface, fall back to the locked text prompts (`render_init_prompt()` / `render_deck_spec_prompt("earnings-update")`).
