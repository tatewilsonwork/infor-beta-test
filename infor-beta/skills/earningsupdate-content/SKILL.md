---
name: earningsupdate-content
description: >
  Use this skill as the Phase 3 POC content stage for a quarterly earnings update. It consumes a
  typed SlidePlan plus source inputs and emits a strict EarningsUpdateContent JSON bundle for the
  deck-assembler stage. Activates inside the conductor plan stage `content`.
allowed-tools: [Read, Write, Bash, WebSearch, WebFetch]
---

# Earnings Update Content — Phase 3 POC

This conductor-stage skill fills the typed content handoff for the earnings-update deck. It does not write `.pptx` files.

## Conductor mode

Your dispatch envelope carries three paths — plugin root, `inputs.json`, `outputs.json` — passed **as arguments** to every command (`python <script.py> "<plugin_root>" "<inputs.json>" "<outputs.json>"`, read back by `stage_io()`). Nothing is exported; nothing is read from the environment.

Your resolved inputs carry `company`, `ticker`, `reporting_quarter`, `comparison_quarter`, `eeo_snip_path`, and `slide_plan_path`.

## Required output

Write a full `EarningsUpdateContent` JSON artefact to `content_bundle.json` in the stage directory (`io.stage_dir`), then `io.write({"content_bundle_path": ...})` — i.e. `outputs.json` holding:

```json
{
  "content_bundle_path": "/absolute/path/to/content_bundle.json"
}
```

The schema is exported at `scripts/schemas/json/earnings_update_content.schema.json` and imported from `schemas.EarningsUpdateContent`.

## Number & currency formatting (read first)

Every figure on this deck follows one house style. The footnote already scopes
the currency (`All figures in US$MM` / `C$MM`), so **never write a currency
code inline** — no `US$`, `C$`, `USD`, `CAD`. Use a plain `$`.

- **Dollar figures in prose** (overview bullets, business updates, performance
  summary): write `$XMM` for millions (no decimals — `$493MM`, `$352MM`) and
  `$X.XB` for a billion or more (one decimal — `$1.1B`, `$1.3B`). Never write
  out "million" / "billion" as words and never prefix a currency code.
  - Good: `Total revenue of $1,283MM, up 2.2% year-over-year` or `$1.3B`
  - Bad: `Total revenue of US$1,282.5 million`
- **KPI tile values** (`prior_value` / `current_value`): supply a **plain
  integer in millions** — just the number, no `$`, no `MM`, no commas required
  (e.g. `1283`, `438`, `493`). The assembler adds the `$` and the `MM`/`B` suffix
  and converts to billions automatically. Do not pre-format these.
- **Broker rows** (`reported` / `estimate` / `variance`): supply plain numerics
  (`1,283`, `(56)`) with no `$`; the assembler prefixes it.

**Never build the JSON with a regex / text substitution.** PowerShell `-replace`
and JS `.replace()` treat `$1`, `$4`, … in the *replacement* string as
capture-group backreferences and silently delete the `$` and the digit —
`US$1,057.8` becomes `US,057.8`. Write every value as a literal with the `Write`
tool. If you must transform text in a script, use Python (its `re` replacement
uses `\1`, not `$1`) and verify the `$` survived.

## Content rules carried forward from the monolith

- Company overview bullets: 6–8 bullets, each ≤220 chars, **560–820 chars total**, no terminal periods or semicolons. The text box sits directly above the "LTM Revenue Breakdown" header and pie, so overshooting this budget pushes the bullets down into that header. Aim for the middle of the range — enough to fill the column without crowding the pie title. (The assembler shrinks an over-budget block as a backstop, but that makes the overview type smaller than the rest of the deck, so stay within budget rather than relying on it.)
  - Use sentence-long or max two-sentence-long bullets that concisely explain what the company does and who they are.
  - Do **not** use bold `Header:` / `Topic:` prefix formatting for general overview bullets. Only use `bold_prefix` for true product / service segment names when the bullet is specifically walking through business segments.
- Business updates: 4–6 bullets, each ≤250 chars, ≤900 chars total, no terminal periods or semicolons.
- KPI rows: exactly 4 rows of currency/value metrics. `prior_value` / `current_value` are **plain integers in MM** (see the formatting section above — the assembler adds `$` and the `MM`/`B` suffix and converts to billions). The period is **not** in the box. Rate deltas in `%`, never bps; `delta_sign` is `1`, `0`, or `-1` and controls green/red formatting downstream.
  - Keep `name` short — it shares the tile with the value, so prefer **≤~24 characters** and abbreviate where helpful: `Adj. EBITDA`, `Non-GAAP Net Income`, `Total Revenue`, `Cloud Services & Subscriptions Rev.` (not `…Revenue`). The assembler auto-shrinks a long name's font so it still fits the tile, but a shorter label reads better — don't lean on the shrink for routine labels.
- Period label: the reporting and comparison quarters print in the mid-blue bar below the "Financial Highlights" title, **not** inside the metric boxes. Each metric box carries only the rounded value and the metric name.
- Broker rows: exactly 5 rows; no `N/A`, `NA`, or `-` cells; `variance_sign` is `1`, `0`, or `-1`. The assembler prefixes `$` onto the Reported, Bloomberg Estimate, and Variance values, so supply plain numerics (e.g. `1,234`, `(56)`) without a leading currency symbol.
  - Do **not** repeat the table's MM-currency scope in row labels. The table header already prints "Figures in {currency_short}", so write plain labels like `Revenue`, `Adj. EBITDA`, `Operating income`, `Free cashflow`. Only per-share metrics carry an inline unit such as `EPS (US$)` or `EPS (C$)` (the assembler defensively strips `(US$MM)` / `(C$MM)` / `(MM)` suffixes but does not strip non-MM markers).
- Management quotes: exactly 2 quotes; each ≤200 chars and ≤30 words.
  - Use abbreviated role titles on the `role` field: `CEO`, `CFO`, `Interim CEO`, `Executive VP and CFO`, `COO`, etc. Do not spell out "Chief Executive Officer" or "Chief Financial Officer" — abbreviations keep the quote attribution on one line at the template font size.
- Performance summary: ≤25 words and ≤150 chars.
- Currency footnote convention: the full code lives only in the footnote and the broker-table header (`C$MM`, `US$MM`). The `currency` field drives the footnote's `[x]$MM` token (the assembler substitutes the `C` / `US` letter), and `currency_short` drives the broker header. On-slide values are always plain `$` — see the formatting section above.

## Source expectations

Use the company filings, MD&A / press release, earnings call transcript if provided, and the Bloomberg EEO snip path supplied in stage inputs. If a required data point is not recoverable from the provided sources, ask the analyst rather than inventing it.

## Provenance — REQUIRED, and structured

**Every figure you write into the bundle carries a record**, on the same contract as
`financial-summary` and `ltm-metrics`. This deck is almost entirely figures you drafted — the four
KPI tiles, the five broker rows, the performance summary, every dollar figure in the overview and
business-update bullets — so a bundle with no records leaves the whole deck untraceable no matter
how well the workbook stages did their half.

Build a `provenance.FigureSource` per figure (the **filing**, the **statement/section**, the
**page**; or the Bloomberg EEO snip as the filing for a broker estimate), record it in this stage's
`ProvenanceLedger`, and write the fragment with `ledger.write(io.stage_dir)` — never a shared file,
because wave-mates run concurrently and a shared ledger is a read-modify-write race. A **citation
string is rejected**: it produces a record with the whole sentence in `filing` and no statement or
page, which reads like provenance and cannot be followed. The `sources` list on
`EarningsUpdateContent` is the *analyst-facing* note on the slide and is not a substitute — it is
prose, and nothing can join it to a figure.

**Name where the figure lands.** `deckcheck` joins by identity, not by value, so each record carries
a `placement` naming the slide (1-based) and the bundle field. Resolve the slide from the
`slide_plan_path` you were handed via `wireframe_common.slide_placement` rather than writing a number
down. A **variance** row is derived, not read: give it a `derivation` **and** `derived_from` refs to
the reported and estimate records it is the difference of, so the chain can be followed rather than
re-read.

```python
import sys
from pathlib import Path

sys.path.insert(0, str(Path(sys.argv[1]) / "scripts"))   # plugin root — arg 1

from stage_io import stage_io
from provenance import FigureRef, FigureSource, ProvenanceLedger
from wireframe_common import load_slide_plan, slide_placement

io = stage_io()
plan = load_slide_plan(io.inputs["slide_plan_path"])
ledger = ProvenanceLedger(stage=io.stage_id)   # this stage runs as the plan's `content` stage

# FORMAT ILLUSTRATION ONLY — obviously-synthetic placeholders; never reuse as data.
reported = ledger.record(
    "Total revenue — reported",
    sources=FigureSource(filing="Q1 2026 press release", statement="Consolidated results", page=9),
    value=999.9, units="US$MM",
    placement=slide_placement(plan, "earnings-update-earnings-summary", "broker_rows[0].reported"),
)
estimate = ledger.record(
    "Total revenue — Bloomberg estimate",
    sources=FigureSource(filing="Bloomberg EEO snip", statement="consensus, revenue"),
    value=999.9, units="US$MM",
    placement=slide_placement(plan, "earnings-update-earnings-summary", "broker_rows[0].estimate"),
)
ledger.record(
    "Total revenue — variance",
    value=0.0, units="US$MM",
    derivation="reported − Bloomberg estimate",
    derived_from=[FigureRef(figure=reported.figure), FigureRef(figure=estimate.figure)],
    placement=slide_placement(plan, "earnings-update-earnings-summary", "broker_rows[0].variance"),
)

ledger.write(io.stage_dir)   # this stage's fragment; `deckcheck` merges every stage's
```

Use the `library_entry_id` values from the slide plan you were given — `slide_placement` raises
naming the ids it does hold if you guess one the plan does not carry.

## Boundary

Do not assemble the deck, do not open PowerPoint, and do not produce the companion cap table. Those are separate conductor stages — and the deck stage is not a skill you could call in any case: the conductor assembles it in-process.
