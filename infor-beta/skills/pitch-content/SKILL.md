---
name: pitch-content
description: Drafts the broad typed content bundle for the 16-slide INFOR slide-library POC deck from analyst notes and optional supporting sources.
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
- **`figure_currency`: the ISO 4217 code every figure you write is stated in** — the target's filing reporting currency (`"USD"`, `"CAD"`, …), as a code, never a rendered token like `"US$MM"`. This is what labels the deck: the assembler puts a currency note on every slide your copy fills, so a slide's label is a property of the copy on it rather than one letter read off the cap table. **Do not convert anything.** The cap table converts to its own output currency and the overview slide names both plus the rate between them; a bundle that quietly restated some figures in a second currency would defeat that. If your sources genuinely disagree on currency, state the figures in the filing currency and say so in `manual_steps`.
- Executive Summary bullets: flexible count; choose main/sub-bullets.
- Public-company overview bullets: concise description of who the company is and what it does. **Budget ≈ 950 characters / ≤ 7 bullets total** — the box shares the overview slide with the LTM revenue pie, and the assembler shrinks over-long copy to fit the band (a fit is guaranteed, but past the budget the text renders noticeably smaller than the rest of the deck).
- Financial Summary metric labels: **not your job** — the `financial-summary` stage selects the metrics (four per FS slide) and emits their labels, so the deck reads them from that stage, not from this content bundle. Do not draft them here.
- Acquirer risks/mitigants: **draft five rows** — the slide-10 table has five body rows, so five consideration/mitigant pairs fill it (the schema allows 1–5; fewer leaves blank rows at the bottom, which is what we're avoiding). Each row has exactly three mitigants. Each mitigant should be **one very short sentence**, and the budget that matters is width: **aim for ≤ ~85 characters per mitigant** (the schema's hard cap is 160). The table is clamped to the template's fixed 5.18" height; a mitigant past ~85 chars wraps to a second line, and once the rows can't fit the assembler steps the body font down from the template's 10 pt — the table stays 5.18" but reads smaller than the rest of the deck.
- Comps takeaway: one sentence.
- Precedents takeaway: one sentence — the precedent-transactions slide's one-line takeaway (mirrors the comps takeaway; the chart stays a placeholder).
- Ownership takeaway: one sentence — the third of the takeaway trio, on the ownership slide. Write it from the ownership data the deck actually shows (the insider and institutional tables the `ownership` stage builds): who holds the register, how concentrated it is, and what that means for a transaction. A real sentence, not a hedge — "118 institutions hold 61.1% against 1.7% insider ownership, so a deal turns on institutional support" is the shape. Required: the slide ships a takeaway box, and before this field existed the delivered deck printed a bare `[x]` in it.
- Key investment highlights: up to 4 numbered quadrants, each a short header + **1–2 concise bullets** (the schema rejects a third — three crowded the quadrant boxes), plus an optional one-line tagline. Optional — omit to leave the slide's placeholders. **When the `include_investment_highlights` input is `false`** (the deck spec dropped the slide; the slide plan carries no `key-investment-highlights` entry), **draft none** — leave `investment_highlights` empty; anything drafted would be ignored because the slide does not exist. When the analyst dictated their own highlights in the notes, use those verbatim (tightened to the budget) rather than inventing new ones.
- Market-entry targets: optional `market_entry_market` (fills the title), `market_entry_row_labels`, and the target columns. **When the analyst does not specify how many acquisition targets they want, draft 8** (filling 4 slides, two targets per slide); otherwise draft the number requested, up to a maximum of **8**. Each target carries an optional `name` (the company name) plus `cells` that align 1:1 with the labels. Set `name` from the target's heading in the analyst notes — it labels the slide's logo box as `[Placeholder for <name> Logo]` (e.g. `[Placeholder for Kueski Logo]`) so the analyst knows which logo to drop in and the box still reads as something to clear; omit it only when the company is unnamed (it then falls back to a bare `[Placeholder for Logo]`). The deck lays targets out **two per slide** (`ceil(N/2)` market-entry slides, titled `Potential <Market> Market Entry Targets (N of M)`).
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
- Contact cards: **leave `contacts` empty unless the analyst named the deal team.** The Contact slide's cards are a declared input, and empty means "keep whatever the library ships filled" — INFOR's own card — while the deck DELETES the rest rather than shipping the template's `[x]`. Only populate `contacts` from names, titles, phone numbers and emails the analyst actually gave you, in the order they should read; never invent a colleague, and never copy one from an unrelated source. Anything you supply replaces the library's cards wholesale.

Required source:
- analyst notes

Optional sources:
- CIM / management deck
- valuation range
- public filings
- company website
- S&P Capital IQ snippets
- analyst risk notes

## Provenance — REQUIRED, and structured

**Every figure you write into the bundle carries a record.** You draft the deck's *prose*, and
the prose is full of numbers: the executive summary's ARR and revenue growth, the overview
bullets' headcount and market position, every one of the 24 cells in the market-entry targets'
metric rows. Those were the deck's least traceable figures precisely because this stage recorded
nothing — a run's ledger held 70 records and not one of them came from here, so `deckcheck` could
adjudicate 43% of what was on the slides and the executive summary's ARR "traced" to an unrelated
gross-profit figure that happened to be 0.01% away.

Build a `provenance.FigureSource` for each figure — the **filing**, the **statement/section** and
the **page** you read it on, or the **url** plus the **date** you read it for a web source — record
it in this stage's `ProvenanceLedger`, and write the fragment with `ledger.write(io.stage_dir)`
(never a shared file: wave-mates run concurrently, so a shared ledger is a read-modify-write race).
A **citation string is rejected**: it builds a record with the whole sentence in `filing` and no
statement or page, which reads like provenance and cannot be followed.

**Name where the figure lands, or it cannot be traced.** `deckcheck` joins a number on a slide to a
record by **identity**, not by value — a value agreement alone is reported as a coincidence, which
is what it usually is. So every record carries a `placement`: the slide (1-based) and the typed
`PitchDeckContent` field the figure sits in. Resolve the slide **from the slide plan you were
handed** (`slide_plan_path`) with `wireframe_common.slide_placement` — never write a slide number
down, because the deck's slide mix is a deck-spec option (one Financial Summary slide or two, one
market-entry slide or four, Key Investment Highlights present or not).

The field is the bundle path, exactly as it reads in the JSON:

| Figure on the deck | `library_entry_id` | `field` |
|---|---|---|
| an executive-summary bullet's figure | `executive-summary` | `executive_summary_bullets[1]` |
| an overview bullet's figure | `public-company-overview` | `company_overview_bullets[0]` |
| a highlight quadrant's figure | `key-investment-highlights` | `investment_highlights[2].bullets[0]` |
| a market-entry cell | `market-entry-targets` | `market_entry_targets[3].cells[10]` |
| an ownership-takeaway figure | `insider-ownership` | `ownership_takeaway` |

Market-entry targets are laid out **two per slide**, so a target's `occurrence` is
`index // 2` — `slide_placement(plan, "market-entry-targets", field, occurrence=i // 2)`.

**A figure you cannot source does not go on the deck.** Analyst-supplied figures (a valuation range,
a market size the analyst dictated) are legitimate — record them with the analyst notes as the
filing (`FigureSource(filing="Analyst notes (deal intake)")`) so the review can say where they came
from instead of calling them unsupported. If a number is recoverable from no source at all, leave it
out and say so in `manual_steps`.

```python
import sys
from pathlib import Path

sys.path.insert(0, str(Path(sys.argv[1]) / "scripts"))   # plugin root — arg 1

from stage_io import stage_io
from provenance import FigureSource, ProvenanceLedger
from wireframe_common import load_slide_plan, slide_placement

io = stage_io()
plan = load_slide_plan(io.inputs["slide_plan_path"])
ledger = ProvenanceLedger(stage=io.stage_id)   # the plan's stage id — this stage runs as `content`

# FORMAT ILLUSTRATION ONLY — obviously-synthetic placeholders showing the call
# shape; NEVER reuse them as data. Every real figure comes from a real source.
ledger.record(
    "ARR",
    sources=FigureSource(filing="Q1 2026 10-Q", statement="MD&A — key metrics", page=99),
    value=999.9, units="US$MM",
    placement=slide_placement(plan, "executive-summary", "executive_summary_bullets[0]"),
)
ledger.record(
    "Example Target Inc. — annual originations",
    sources=FigureSource(url="https://example.com/about", retrieved="<YYYY-MM-DD>"),
    value=99.9, units="US$MM",
    # target index 3 -> the second market-entry slide (two targets per slide)
    placement=slide_placement(plan, "market-entry-targets",
                              "market_entry_targets[3].cells[10]", occurrence=1),
)

ledger.write(io.stage_dir)   # this stage's fragment; `deckcheck` merges every stage's
```

Do not write PowerPoint. Write the bundle to the stage directory and hand back `content_bundle_path` in `outputs.json` — `io.write({"content_bundle_path": ...})`, with the three handoff paths taken from your dispatch envelope as command-line arguments (never environment variables).
