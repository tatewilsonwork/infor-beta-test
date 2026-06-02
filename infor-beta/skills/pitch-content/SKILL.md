---
name: pitch-content
description: Drafts the broad typed content bundle for the 14-slide INFOR slide-library POC deck from analyst notes and optional supporting sources.
version: 0.5.6
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
- Public-company overview bullets: concise description of who the company is and what it does.
- Financial Summary metric labels: exactly four, **metric NAMES only** — e.g. `Revenue`, `Adjusted EBITDA`, `Combined Loan Balances`, `Adjusted Return on Equity`. Do **not** put the amount, currency, units, or YoY delta in the label (no `FY2025 Revenue: US$589.8MM (+31% YoY)`); the value is shown by the (placeholder) chart, not the tile. The schema rejects digits, `$`/`%`, and colons in these labels.
- Acquirer risks/mitigants: concise rows, exactly three short mitigants each.
- Comps takeaway: one sentence.
- Key investment highlights: up to 4 numbered quadrants, each a short header + 1–3 concise bullets, plus an optional one-line tagline. Optional — omit to leave the slide's placeholders.
- Market-entry targets: optional `market_entry_market` (fills the title), `market_entry_row_labels`, and up to **8** target columns whose `cells` align 1:1 with the labels. The deck lays targets out **two per slide** (`ceil(N/2)` market-entry slides, titled `Potential <Market> Market Entry Targets (N of M)`); target logos stay as deferred image placeholders.
  - `market_entry_row_labels` is a **fixed 12-row structure**, in this exact order:
    1. `Overview` — very short description of who the target is
    2. `Headquarters` — City, Country
    3. `Year Founded` — year
    4–10. **Seven industry-relevant metrics** — chosen for the market the acquirer is entering, and **identical across every target** (one shared label list drives all target slides, so the rows stay consistent)
    11. `Scale KPIs` — the best available benchmark of scale (revenue if available, else number of transactions / loans / customers / merchants, etc.)
    12. `Strategic Rationale` — why the target is a good acquisition for the Company
  - Rows 1–3 and 11–12 use those exact labels; only the middle seven are deck-specific. Each target's `cells` must supply all 12 values in order. The schema rejects any other row count or fixed-label text.

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
