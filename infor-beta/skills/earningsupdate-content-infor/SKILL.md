---
name: earningsupdate-content-infor
description: >
  Use this skill as the Phase 3 POC content stage for a quarterly earnings update. It consumes a
  typed SlidePlan plus source inputs and emits a strict EarningsUpdateContent JSON bundle for the
  deck-assembler stage. Activates inside the conductor plan stage `content`.
version: 0.4.3
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

## Content rules carried forward from the monolith

- Company overview bullets: 7–12 bullets, each ≤250 chars, 1,200–1,500 chars total, no terminal periods or semicolons.
  - Use sentence-long or max two-sentence-long bullets that concisely explain what the company does and who they are.
  - Do **not** use bold `Header:` / `Topic:` prefix formatting for general overview bullets. Only use `bold_prefix` for true product / service segment names when the bullet is specifically walking through business segments.
- Business updates: 4–6 bullets, each ≤250 chars, ≤900 chars total, no terminal periods or semicolons.
- KPI rows: exactly 4 rows; rate deltas in `%`, never bps; `delta_sign` is `1`, `0`, or `-1` and controls green/red formatting downstream.
- Broker rows: exactly 5 rows; no `N/A`, `NA`, or `-` cells; `variance_sign` is `1`, `0`, or `-1`.
- Management quotes: exactly 2 quotes; each ≤200 chars and ≤30 words.
- Performance summary: ≤25 words and ≤150 chars.
- Currency footnote convention: full code in footnotes / broker table header (`C$MM`, `US$MM`, etc.); on-slide values may use plain `$` where the footnote scopes the currency.

## Source expectations

Use the company filings, MD&A / press release, earnings call transcript if provided, and the Bloomberg EEO snip path supplied in stage inputs. If a required data point is not recoverable from the provided sources, ask the analyst rather than inventing it.

## Boundary

Do not call `deck-assembler`, do not open PowerPoint, and do not produce the companion cap table. Those are separate conductor stages.
