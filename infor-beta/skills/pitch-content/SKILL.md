---
name: pitch-content
description: Drafts the broad typed content bundle for the 16-slide INFOR slide-library POC deck from analyst notes and optional supporting sources.
version: 0.5.31
allowed-tools:
  - Read
  - Write
  - Bash
  - WebSearch
inputs:
  - name: analyst_notes
    type: str
    required: true
  - name: company
    type: Company
    required: true
  - name: presentation_date
    type: str
    required: true
  - name: slide_plan_path
    type: Path
    required: true
  - name: include_investment_highlights
    type: bool
    required: false
outputs:
  - name: content_bundle_path
    type: Path
side_effects:
  - kind: file_write
    description: Writes a typed PitchDeckContent JSON artefact to the stage output directory.
---

# pitch-content

Drafts the single broad `PitchDeckContent` handoff for the Phase 3 slide-library POC.

Scope:
- Executive Summary bullets: flexible count; choose main/sub-bullets.
- Public-company overview bullets: concise description of who the company is and what it does. **Budget ≈ 950 characters / ≤ 7 bullets total** — the box shares the overview slide with the LTM revenue pie, and the assembler shrinks over-long copy to fit the band (a fit is guaranteed, but past the budget the text renders noticeably smaller than the rest of the deck).
- Financial Summary metric labels: **not your job** — the `financial-summary` stage selects the metrics (four per FS slide) and emits their labels, so the deck reads them from that stage, not from this content bundle. Do not draft them here.
- Acquirer risks/mitigants: **draft five rows** — the slide-10 table has five body rows, so five consideration/mitigant pairs fill it (the schema allows 1–5; fewer leaves blank rows at the bottom, which is what we're avoiding). Each row has exactly three mitigants. Each mitigant should be **one very short sentence**, and the budget that matters is width: **aim for ≤ ~85 characters per mitigant** (the schema's hard cap is 160). The table is clamped to the template's fixed 5.18" height; a mitigant past ~85 chars wraps to a second line, and once the rows can't fit the assembler steps the body font down from the template's 10 pt — the table stays 5.18" but reads smaller than the rest of the deck.
- Comps takeaway: one sentence.
- Precedents takeaway: one sentence — the precedent-transactions slide's one-line takeaway (mirrors the comps takeaway; the chart stays a placeholder).
- Key investment highlights: up to 4 numbered quadrants, each a short header + **1–2 concise bullets** (the schema rejects a third — three crowded the quadrant boxes), plus an optional one-line tagline. Optional — omit to leave the slide's placeholders. **When the `include_investment_highlights` input is `false`** (the deck spec dropped the slide; the slide plan carries no `key-investment-highlights` entry), **draft none** — leave `investment_highlights` empty; anything drafted would be ignored because the slide does not exist. When the analyst dictated their own highlights in the notes, use those verbatim (tightened to the budget) rather than inventing new ones.
- Market-entry targets: optional `market_entry_market` (fills the title), `market_entry_row_labels`, and the target columns. **When the analyst does not specify how many acquisition targets they want, draft 8** (filling 4 slides, two targets per slide); otherwise draft the number requested, up to a maximum of **8**. Each target carries an optional `name` (the company name) plus `cells` that align 1:1 with the labels. Set `name` from the target's heading in the analyst notes — it labels the slide's logo box as `[<name> Logo]` (e.g. `[Kueski Logo]`) so the analyst knows which logo to drop in; omit it only when the company is unnamed (it then falls back to a generic `[Company Name Logo]`). The deck lays targets out **two per slide** (`ceil(N/2)` market-entry slides, titled `Potential <Market> Market Entry Targets (N of M)`).
  - `market_entry_row_labels` is a **fixed 12-row structure**, in this exact order:
    1. `Overview` — very short description of who the target is
    2. `Headquarters` — City, Country
    3. `Year Founded` — year
    4–10. **Seven industry-relevant metrics** — chosen for the market the acquirer is entering, and **identical across every target** (one shared label list drives all target slides, so the rows stay consistent)
    11. `Scale KPIs` — the best available benchmark of scale (revenue if available, else number of transactions / loans / customers / merchants, etc.)
    12. `Strategic Rationale` — why the target is a good acquisition for the Company
  - Rows 1–3 and 11–12 use those exact labels; only the middle seven are deck-specific. Each target's `cells` must supply all 12 values in order. The schema rejects any other row count or fixed-label text.
  - **Keep the seven metric labels short — ≤ ~18 characters.** The label column is 1.66" wide; a longer label (e.g. `Geographic Footprint`) no longer fits at 11 pt and the assembler steps its font down so it can't wrap the row taller. Prefer `Geography` over `Geographic Footprint`, `Capital Raised` over `Total Capital Raised`.
  - **Keep each cell tight.** The table is clamped to a fixed height; PowerPoint grows a row to fit its text, so verbose cells push the table off the slide. Aim for ≤ ~90 characters in the wordy rows (`Overview`, `Strategic Rationale`) — a crisp phrase, not a sentence — and a few words in the metric rows.

Required source:
- analyst notes

Optional sources:
- CIM / management deck
- valuation range
- public filings
- company website
- S&P Capital IQ snippets
- analyst risk notes

Do not write PowerPoint. Write `content_bundle_path` to `$STAGE_OUTPUTS`.
