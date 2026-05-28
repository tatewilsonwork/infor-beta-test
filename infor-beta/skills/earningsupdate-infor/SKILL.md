---
name: earningsupdate-infor
description: >
  Use this skill when the user invokes /earningsupdate-infor or asks to build a quarterly earnings
  update deck for a public company. Builds a branded 5-slide deck cloned from the shared INFOR Slide
  Library (cover, public-company overview, earnings summary, disclaimer, contact) with a company
  overview, an LTM revenue pie placeholder, a Capitalization Summary cap table, four KPI metric
  boxes (actual vs prior-year quarter), a Broker Estimates vs Actuals table sourced from a Bloomberg
  EEO snip, business-update bullets, management quotes, and a short performance summary. Activates on
  "earnings update", "earnings deck", "quarterly earnings", "earnings summary deck", or any request
  to build a branded update deck off a recent 10-Q/10-K and Bloomberg EEO snip.
version: 0.5.1
allowed-tools: [Read, Bash, Write, Glob, WebSearch, WebFetch]
---

# INFOR Earnings Update — Workflow

This skill builds a branded 5-slide earnings update deck for direct (`/earningsupdate-infor`)
invocation. It is a thin orchestrator over the same typed pipeline the conductor's
`earnings-update.yaml` plan runs stage-by-stage — there is no longer a separate monolith
implementation and no standalone earnings template. The deck is cloned from the shared
`INFOR Slide Library.pptx`; the same helpers and content rules apply whether you reach them through
the conductor or this command.

The pipeline, in order:

1. **wireframe** — `build_earnings_update_slide_plan` → typed `SlidePlan` (5 fixed slides).
2. **content** — draft a typed `EarningsUpdateContent` bundle (rules below).
3. **captable** — companion capitalization-table workbook via `captable-infor`.
4. **ltm-revenue** — companion LTM revenue workbook via `ltm-revenue-infor`.
5. **assemble** — `assemble_earnings_update_deck` clones the five library entries, fills them, and
   inserts the cap-table picture into the overview slide.
6. **QA** — render slides to PNG and fix overflow with `enable_normal_autofit`.

Today's date is available from the system context (`currentDate`) — do not shell out to `date`.

Shared helpers live at [`infor-beta/scripts/`](../../scripts/) and import via `CLAUDE_PLUGIN_ROOT`:

```python
import sys, os
sys.path.insert(0, os.environ.get("CLAUDE_PLUGIN_ROOT", "./infor-beta") + "/scripts")

from schemas import Company, EarningsUpdateContent
from earnings_update_wireframe import build_earnings_update_slide_plan, write_slide_plan
from earnings_update_assembler import assemble_earnings_update_deck
from ltm_revenue import build_ltm_revenue_workbook, RevenueSegment
from slide_render import render_deck_to_png
from pptx_helpers import enable_normal_autofit
```

Output lands in cwd for direct invocation. (Under the conductor, the decomposed stage skills —
`earningsupdate-wireframe-infor`, `earningsupdate-content-infor`, `captable-infor`,
`ltm-revenue-infor`, `deck-assembler` — handle the same steps and write to the deal directory.)

---

## Workflow Steps

### Step 1 — Collect Inputs

Required:
- Company name and CapIQ ticker (e.g., `NasdaqGS:VFF`)
- Reporting quarter (e.g., Q4 2025) and prior-year comparison quarter (Q4 2024)
- Most recent 10-Q / 10-K / annual report + MD&A
- Bloomberg EEO snip (image/screenshot showing consensus vs. actual for the quarter)

Optional but highly useful:
- Earnings call transcript
- Company earnings press release URL (if not web-searchable from company name + quarter)

If any required input is missing, ask in a single message:

> "To build the earnings update, I need:
> - **Company name + CapIQ ticker** (e.g., NasdaqGS:VFF)
> - **Reporting quarter** (e.g., Q4 2025) and **prior-year comparison quarter**
> - **10-Q / 10-K / MD&A** attached
> - **Bloomberg EEO snip** attached (image)
>
> Optional: earnings call transcript (attach) for extra detail."

Wait for all required inputs before proceeding.

---

### Step 2 — Determine Reporting Currency

Read the 10-Q/10-K to identify the reporting currency. Village Farms reports in US$ despite being
Canadian-listed — do NOT infer currency from exchange. Read the cover page or the "Basis of
Presentation" footnote in the financial statements.

Output the currency code as one of `US$MM`, `C$MM`, `€MM`, `£MM`, `A$MM`, etc.

**Currency only lives in the footnote / table header.** Use the full code only in:
- The overview and earnings-summary footnote text (`Note: All figures in C$MM, except where indicated otherwise`)
- The Broker Estimates table header cell (`Figures in C$MM`)

**Everywhere else on the slide, values use a plain `$` prefix** — never `C$`, `US$`, `€`. The
footnote scopes the currency; repeating it on every value is redundant and visually noisy. Applies
to KPI metric boxes, performance summary, Business Updates bullets, overview bullets, and broker
cells. For non-dollar currencies, use the symbol once in the footnote and plain `$` (or no prefix)
on values.

---

### Step 3 — Build the SlidePlan

Build the fixed five-slide earnings-update `SlidePlan` and write it to cwd:

```python
company = Company(legal_name="<Company Name>", ticker="<CapIQ ticker>")
plan = build_earnings_update_slide_plan(
    company=company,
    reporting_quarter="<e.g. Q4 2025>",
    comparison_quarter="<e.g. Q4 2024>",
)
slide_plan_path = write_slide_plan(plan, "./earnings_slide_plan.json")
```

The plan is canonical for slide order and titles (`earnings-update-cover`,
`earnings-update-company-overview`, `earnings-update-earnings-summary`, `earnings-update-disclaimer`,
`earnings-update-contact`). Do not draft copy here.

---

### Step 4 — Draft the Content Bundle

Draft a typed `EarningsUpdateContent` from the filings, MD&A / press release, transcript (if
provided), and the Bloomberg EEO snip, and write it to `./earnings_content.json`. The schema is at
`scripts/schemas/json/earnings_update_content.schema.json` (imported from
`schemas.EarningsUpdateContent`).

Content rules (enforce with asserts before writing — they are the primary overflow defense):

- **Company overview bullets:** 6–10 bullets, each ≤250 chars, 650–1,050 chars total, no terminal
  periods or semicolons. The overview slide reserves its lower-left quadrant for the LTM revenue pie
  placeholder, so keep bullets concise or text overflows behind the pie. Use sentence-long (max two
  sentences) bullets describing what the company does and who they are. Do **not** use bold
  `Header:` prefixes for general bullets — only set `bold_prefix` for true product/service segment
  names when a bullet is specifically walking through business segments.
- **Business updates:** 4–6 bullets, each ≤250 chars, ≤900 chars total, no terminal periods/semicolons.
- **KPI rows:** exactly 4. Currency/value metrics are whole numbers in MM with **no decimals**; each
  metric box shows the rounded value plus the metric name. The reporting/comparison period prints in
  the mid-blue bar below the "Financial Highlights" title, **not** in the boxes. Rate deltas in `%`,
  never bps. `delta_sign` is `1`/`0`/`-1` and drives green (`#00B050`) / red (`#C00000`) downstream.
- **Broker rows:** exactly 5; no `N/A`/`NA`/`-` cells; `variance_sign` is `1`/`0`/`-1`. The assembler
  prefixes `$` onto Reported, Bloomberg Estimate, and Variance, so supply plain numerics
  (`1,234`, `(56)`) with no leading currency symbol. Do not repeat the table's MM scope in row labels
  (`Revenue`, `Adj. EBITDA`, `Operating income`, `Free cashflow`); only per-share metrics carry an
  inline unit such as `EPS (US$)`.
- **Management quotes:** exactly 2; each ≤200 chars and ≤30 words; address the key item of the
  quarter. Use abbreviated roles (`CEO`, `CFO`, `Interim CEO`, `Executive VP and CFO`) — never spell
  out "Chief Executive Officer".
- **Performance summary:** one sentence, ≤25 words and ≤150 chars (beat/miss + qualifier).

If a required data point is not recoverable from the provided sources, ask the analyst rather than
inventing it.

---

### Step 5 — Generate the Companion Cap Table

Invoke the **captable-infor** skill's workflow using the same 10-Q/10-K/MD&A attachments and the
CapIQ ticker from Step 1, saving the workbook to cwd as
`./<SANITIZED_TICKER> - Capitalization Table.xlsx`. The assembler inserts a picture of its
`Cap with Links!B15:F31` range into the overview slide's `Rectangle 3` Capitalization Summary
placeholder, so the workbook must populate that range.

> Per the standing constraint, ignore `#NAME?` errors in the workbook — they resolve when the
> analyst refreshes the Capital IQ connector.

---

### Step 6 — Generate the LTM Revenue Workbook

Invoke **ltm-revenue-infor** (or call `build_ltm_revenue_workbook` directly) to produce the
companion LTM revenue breakdown workbook in cwd. Segment by service/product line when disclosed,
else by geography. The overview slide keeps its `[Pie Chart Placeholder]` — the analyst (or a later
stage) builds the pie from this workbook; the assembler does **not** chart it.

---

### Step 7 — Assemble the Deck

Clone the five library entries and fill them, inserting the cap-table picture:

```python
template = os.environ.get("CLAUDE_PLUGIN_ROOT", "./infor-beta") + "/templates/INFOR Slide Library.pptx"
output_path = assemble_earnings_update_deck(
    slide_plan_path="./earnings_slide_plan.json",
    content_path="./earnings_content.json",
    template_path=template,
    output_dir=".",
    captable_workbook_path="./<SANITIZED_TICKER> - Capitalization Table.xlsx",
)
```

The assembler clones library slides 1, 7, 8, 14, 15 into the final five-slide order (cover,
overview, earnings summary, disclaimer, contact), fills the cover date, overview, and earnings
summary, and leaves the disclaimer and contact entries exactly as shipped. It saves
`Earnings Update - <Company>.pptx` to `output_dir`. Cap-table insertion uses Excel COM on Windows
and a LibreOffice fallback on Cowork/Linux; if no workbook is available the placeholder is preserved.

---

### Step 8 — Overflow QA

Claude cannot see PowerPoint overflow directly. The character/word caps in Step 4 are the primary
defense; this step is the visual backstop. Render the populated overview and earnings-summary slides
to PNG and inspect them:

```python
pngs = render_deck_to_png(output_path, "./_qa", slide_indices=[1, 2])
```

`Read` the PNGs. If any shape overflows, call `enable_normal_autofit` on the offending shape (or
tighten the content within the Step 4 caps) and re-render until clean. Rendering uses PowerPoint COM
on Windows and LibreOffice headless on Cowork/Linux; if neither is available, skip and trust the caps.

---

### Step 9 — Report to User

Output a brief summary:

1. **Deck file:** absolute path to saved `.pptx`
2. **Cap table file:** absolute path to saved `.xlsx`
3. **LTM revenue file:** absolute path to saved `.xlsx`
4. **Reporting currency used:** e.g., `US$MM`
5. **KPIs selected:** the four metrics and why (one phrase each)
6. **Sources used for bullets / quotes:** which file or URL each major section came from
7. **Manual steps remaining:**
   - Refresh the Capital IQ connector in the cap table and LTM revenue workbooks
   - Build the LTM revenue pie from the LTM workbook and drop it over `[Pie Chart Placeholder]` on the overview slide
   - Review the Performance Summary box wording before sending
