---
name: earningsupdate-content
description: >
  Use this skill as the Phase 3 POC content stage for a quarterly earnings update. It consumes a
  typed SlidePlan plus source inputs and emits a strict EarningsUpdateContent JSON bundle for the
  deck-assembler stage. Activates inside the conductor plan stage `content`.
version: 0.5.20
allowed-tools: [Read, Write, Bash, WebSearch, WebFetch]
---

# Earnings Update Content — Phase 3 POC

This conductor-stage skill fills the typed content handoff for the earnings-update deck. It does not write `.pptx` files.

## Conductor mode

When invoked by the conductor, read:

- `$STAGE_INPUTS` — JSON with `company`, `ticker`, `reporting_quarter`, `comparison_quarter`, `eeo_snip_path`, and `slide_plan_path`
- `$STAGE_OUTPUTS` — path where this stage must write its structured handoff
- `$DEAL_DIR` — deal directory root

## Required output

Write a full `EarningsUpdateContent` JSON artefact to `content_bundle.json` in the stage directory, then write `$STAGE_OUTPUTS` as:

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

## Boundary

Do not call `deck-assembler`, do not open PowerPoint, and do not produce the companion cap table. Those are separate conductor stages.
