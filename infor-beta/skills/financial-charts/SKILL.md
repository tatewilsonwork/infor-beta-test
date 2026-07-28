---
name: financial-charts
description: >
  Use this skill as the pitch plan's `financial-charts` stage. It builds the INFOR-formatted
  clustered-column charts for the deck's Financial Summary slide(s) — one per metric row, four
  per slide — on the deal workbook's `financial-summary` tab, then renders them into each FS
  slide's four chart placeholders. It also builds the overview slide's LTM revenue-by-segment pie
  on the deal workbook's `ltm-metrics` tab and drops it into the "[Pie Chart Placeholder]". Runs
  after `deck`, the deck it edits.
allowed-tools: [Read, Write, Bash]
---

# Financial Summary Charts — Workflow

This stage finishes the pitch deck's **Financial Summary** slide(s) (library entry 8 — one slide
by default, two when the deck spec asked for 8 metrics). The `financial-summary` stage already
built the chart-ready data tab and the deck-assembler already filled the metric **name** tiles;
this stage builds the **charts** (one per metric row — 4 or 8) and drops them into each FS slide's
four chart placeholders. The helper discovers the FS slides by scanning the deck for the Metric #1
placeholder and detects the metric rows from the tab, so 4-metric and 8-metric decks run the same
call.

It also fills the **overview** slide's deferred **LTM revenue pie** — see the section below. Both
the FS charts and the pie ride this single stage because the data live on the same
deal workbook (FS charts off the `financial-summary` tab, the pie off the `ltm-metrics` tab).

## Mandatory: build the charts with the helper (no hand-rolling)

This stage **must** build every chart by calling the two `scripts/financial_charts.py`
orchestrators — `render_financial_summary_charts_into_deck(...)` for the Financial Summary
charts and `render_ltm_revenue_pie_into_deck(...)` for the overview pie (see the Reference
command). Do **not** hand-roll charts with `matplotlib`, `plotly`, Pillow, or any other plotting
library, and do not draw the chart images yourself — the analyst needs real, editable Excel
charts on the workbook, not flat pictures.

Each helper does two things that are both required and must not be split apart:

1. **Persists native chart objects on the deal workbook** — the clustered-column charts
   on the `financial-summary` tab and the pie on the `ltm-metrics` tab — so opening
   `pitch-<codename>.xlsx` shows real Excel chart objects.
2. **Inserts the rendered charts into the deck** — the four chart placeholders on each Financial
   Summary slide and the `[Pie Chart Placeholder]` on the overview slide.

The helpers save the native charts to the workbook **first**, then render the deck images. If
LibreOffice (`soffice`/`libreoffice`) is unavailable, the workbook charts are still saved and the
helper returns `None` for the deck step (the slide keeps its placeholders) instead of aborting the
stage — when that happens, **say so explicitly** in the handoff; never substitute a hand-drawn
image.

## ⛔ NEVER re-run `deck-assembler` (or any skill) from this stage

**This stage MUST NOT invoke `deck-assembler` — or any other INFOR skill — via the `Task`/`Agent`
tool. It never re-assembles, re-clones, or re-saves a fresh deck.** It only opens the
**already-assembled** deck at `deck_path`, inserts the chart/pie pictures into the existing
placeholders, and saves it back in place. Its `allowed-tools` are `[Read, Write, Bash]` precisely
so it *cannot* dispatch another skill — do not work around that by shelling out to one.

**Why this is load-bearing:** building charts is the *last* mutation of the deck, and re-running
the assembler here would re-save a clean deck over the top of it. Nothing may re-assemble the deck
afterward.

Before Phase D there was a second, sharper reason: this stage ran after `workbook-aggregation`,
which folded the standalone `captable` / `ownership` workbooks into a combined file and then
**deleted them** — so a re-assembly re-pasted tables from files that no longer existed and
reverted them to empty placeholders (a real regression from a live run). The deal owns one workbook
now, nothing is deleted, and the tabs the assembler reads are still there — but the rule stands,
because a re-assembly still discards this stage's charts.

**If a table ever genuinely needs re-inserting** (it should not — `deck` already did it), read it
from the deal workbook's `captable` / `Ownership` tabs and paste it with
`excel_to_powerpoint.insert_excel_into_placeholder`. Never call the assembler to do it.

## Why it runs after `deck`

Only because it edits the deck `deck` produces. It has no workbook-ordering constraint: each
**flow** metric's LTM cell on the `financial-summary` tab is a label-keyed lookup
`=INDEX('ltm-metrics'!$B:$B, MATCH("(=) <result_label>", 'ltm-metrics'!$A:$A, 0))`, and since both
tabs live in the deal's single workbook that is an ordinary internal reference, live from the moment
`ltm-metrics` is written. (Before Phase D it read `#N/A` until `workbook-aggregation` merged the two
standalone files, which is why this used to be a post-aggregation stage.) The native charts persist
on the tab for the analyst; the rendered copies go into the assembled deck.

## Conductor-mode handoff

Your dispatch envelope carries three paths — plugin root, `inputs.json`, `outputs.json` — and
every command below takes them **as arguments**
(`python <script.py> "<plugin_root>" "<inputs.json>" "<outputs.json>"`, read back by
`stage_io()`). Nothing is exported; nothing is read from the environment.

- Read inputs from `io.inputs`:
  - `deal_workbook` — the deal's single workbook, `pitch-<codename>.xlsx`.
  - `deck_path` — the assembled pitch deck from `deck`.
  - `deal_dir` — deal directory root (used for the QA render output; `io.deal_dir` derives the
    same path from the inputs path when the input is absent).
- If `deal_workbook` has no `financial-summary` tab (the financial-summary stage produced
  nothing), the slide is left with its placeholders — write the handoff and report the skip.
- If a required field is missing, `io.fail("missing input: <field>")` and stop.

## What the charts look like (the only formatting that matters)

One **clustered-column** chart per metric (4 or 8 total): a single data series = that metric's row
(`B{r}:G{r}`, `r` from 6 down — rows 6-9 for four metrics, 6-13 for eight; detected from the tab),
categories = the period header (`B5:G5`). Read the header row **dynamically** so a suppressed-LTM
tab (`B5:F5`) charts five columns instead of six.

- Font: **Palatino Linotype 9 pt, black** (data labels + category-axis labels)
- **No** chart title
- **No** chart border
- **No** major gridlines
- Horizontal (category) axis: a **visible solid black baseline line** beneath the bars — an
  explicit width, never a hairline (a width-less line is dropped by the LibreOffice render, so
  the bars float with no baseline)
- Vertical (value) axis: **hidden** — no line, no label
- Gap width: **50%**
- Data labels on **all** bars, position = **Outside End**, in the **same `$` currency format as
  the tab's value cells** (`$#,##0.0_);($#,##0.0);"--"`) — a bar reads `$102.7` exactly like its
  cell, never a bare `102.7`
- All bars fill **RGB(70, 86, 110)** (hex `46566E`)

## Placeholder ↔ metric mapping (per Financial Summary slide)

The helper finds every FS slide by its `Rectangle 17` / `[Placeholder for Metric #1 Chart]`
marker (slide 8 in the default deck; slides 8-9 when the deck spec asked for two). FS slide *k*
(zero-based) charts tab rows `6+4k .. 9+4k` onto the same four placeholder shapes:

| Metric on the slide | data row (slide k) | chart placeholder | box (L", T") |
|--------|----------|-------------------|--------------|
| #1 | 6+4k | `Rectangle 17` | 0.35, 1.51 |
| #2 | 7+4k | `Rectangle 7`  | 5.12, 1.51 |
| #3 | 8+4k | `Rectangle 19` | 0.35, 4.42 |
| #4 | 9+4k | `Rectangle 18` | 5.12, 4.42 |

Each placeholder is 4.53" × 2.51"; the rendered chart is stretched to its box.

## LTM revenue pie (overview slide = `prs.slides[6]`)

The overview ("Introduction to {company}") slide carries a deferred `[Pie Chart Placeholder]`
(shape `Rectangle 4`, box 4.51" × 1.77") under the "LTM Revenue Breakdown" label. This stage fills
it with a by-segment pie built on the deal workbook's **`ltm-metrics`** tab, off the
**"LTM Revenue Overview"** block (located by its section title — no hardcoded row numbers; the
**Total** row is excluded).

The pie charts **at most 5 slices**: the **4 largest segments** (descending by LTM revenue) plus
an **"Other"** slice grouping the remainder — charting every segment overflowed the legend into
the pie. Both builders write a **"Pie Chart Source"** block in columns **E:G** beside the overview
block (name `=A{r}` / $ amount `=B{r}` / fraction `=F/Btotal`; "Other" = `=Btotal − SUM(top 4)`),
and the chart's categories/series reference that block — Excel charts literal cells, and the
in-cell formulas keep the pie live when the analyst edits a segment $. Five or fewer segments
chart as-is (descending), with no "Other". Charting the fraction column rather than the dollar
amounts gives identical slice geometry but lets the data labels read the slice share.
(Off-Windows, openpyxl can't evaluate the block's formulas, so the PNG render recomputes the
grouped fractions from the literal $ column — the analyst's workbook still charts the live block.)

- Legend docked on the **RIGHT** at **Palatino 8** (one point under the chart text), **pinned to
  the full remaining right side** of the chart box; the pie's plot area is **pinned to the left**
  (manual layouts on both) so pie and legend cannot overlap and Excel cannot silently drop legend
  entries that its undersized auto-layout box wouldn't fit; **no** chart title; **no** chart border
- Slice fills from the **INFOR theme accent palette** (`pptx_helpers.INFOR_ACCENTS`:
  `0E213F, 46566E, ADB9CA, A4844B, 767171, E5E3E3`), in order
- Palatino 9 data labels in **white**, position **Inside End** — every label sits inside its own
  slice (Best Fit floated the small-slice labels outside the pie); **value-only** (percentage /
  category / series / legend-key flags off), number format `#,##0.0%_);(#,##0.0%);"--"` so the
  fraction reads e.g. `45.2%`
- Data labels **only on slices larger than 3%** of the total — smaller slices carry no label
  (their labels overlap each other in the short box); the slice itself still renders

When the deal workbook has no `ltm-metrics` tab or no "LTM Revenue Overview" block, the pie is
skipped and the placeholder is left in place (the null path).

## Workflow

1. Read your resolved inputs (`io.inputs`; also reproduced in the envelope).
2. Call `render_financial_summary_charts_into_deck(...)` (see the reference command). It builds one
   chart per metric row on the deal workbook's `financial-summary` tab and inserts them into
   every FS slide in the deck (four per slide, discovered by scanning).
3. Call `render_ltm_revenue_pie_into_deck(...)` on the **deck written by step 2** (chain the same
   path) to build + insert the overview LTM revenue pie. A `None` return means the pie was skipped
   (no `ltm-metrics` tab / block) — leave the placeholder.
4. **Chart QA — mandatory, do not skip.** Render the overview slide **and** every Financial Summary
   slide to PNG with `render_deck_to_png(deck_path, out, slide_indices=[6, 7])` (append index 8
   when the deck carries a second FS slide), read the PNGs, and confirm:
   - **Overview (slide 6):** the pie filled the placeholder; **at most 5 slices** (top 4 +
     "Other"); legend on the **right** showing **every** slice's entry (including "Other") with
     the pie sitting **clear of it** (no overlap); INFOR accent slice fills; labels read as
     percentages (`45.2%`) in **white, inside their slices** (none floating outside the pie) and
     appear **only on slices above 3%**; no title / no border; reads legibly in the wide/short box.
   - **Each Financial Summary slide (7, and 8 when present):** all four charts landed; INFOR
     formatting (bars `46566E`, data labels outside-end reading **`$102.7`-style currency** exactly
     like the tab's cells, **no border**, a **visible solid black** category-axis baseline line
     under the bars, no value axis / gridlines / title, Palatino 9);
   - the combined-metric and LTM bars use the **same scale** as the fiscal-year bars (no 10⁶
     mismatch — a symptom of a units mismatch between the `financial-summary` and `ltm-metrics`
     tabs; both must be in millions with an `"MM"` suffix).
   If the renderer is unavailable, say so explicitly rather than skipping silently.
5. Write the structured handoff with `io.write(...)`.

On **direct invocation** there is no envelope and no handoff: run the same flow against the
supplied paths.

## Reference command

Run as `python <script.py> "<plugin_root>" "<inputs.json>" "<outputs.json>"` — the three paths
your dispatch envelope prints.

```python
import sys
from pathlib import Path

sys.path.insert(0, str(Path(sys.argv[1]) / "scripts"))   # plugin root — arg 1

from stage_io import stage_io
from financial_charts import (
    render_financial_summary_charts_into_deck,
    render_ltm_revenue_pie_into_deck,
)
from slide_render import render_deck_to_png

io = stage_io()
deal_dir = Path(io.inputs.get("deal_dir") or io.deal_dir)
deck_path = io.inputs["deck_path"]
deal_workbook = io.inputs["deal_workbook"]

# 1) Financial Summary charts — one per metric row, inserted into every FS
#    slide (discovered by scanning the deck); modifies the deck in place.
fs_deck = render_financial_summary_charts_into_deck(
    deck_path=deck_path,
    deal_workbook=deal_workbook,
)
fs_inserted = fs_deck is not None
deck_path = str(fs_deck) if fs_deck is not None else deck_path

# 2) Overview LTM revenue pie (slide 6) — chained onto the FS-charts deck.
pie_deck = render_ltm_revenue_pie_into_deck(
    deck_path=deck_path,
    deal_workbook=deal_workbook,
)
pie_inserted = pie_deck is not None
deck_path = str(pie_deck) if pie_deck is not None else deck_path

# 3) Mandatory QA render — overview (index 6) + every Financial Summary slide
#    (index 7, plus 8 when the deck carries two FS slides).
qa_dir = deal_dir / "runs" / "financial-charts-qa"
render_deck_to_png(deck_path, qa_dir, slide_indices=[6, 7])
# → read qa_dir/slide_7.png (overview pie) and slide_8.png (FS charts) and confirm
#   the formatting described in the Workflow QA step.

io.write(
    {
        "deck_path": deck_path,
        "charts_inserted": fs_inserted,
        "pie_inserted": pie_inserted,
    }
)
```

## Outputs (`outputs.json`)

```json
{
  "deck_path": "/absolute/path/to/the/final/pitch/deck.pptx",
  "charts_inserted": true,
  "pie_inserted": true
}
```

`deck_path` is the final pitch deck (same file as the `deck` stage's output, modified in place).
`charts_inserted` is `false` when there was no `financial-summary` tab to chart; `pie_inserted` is
`false` when there was no `ltm-metrics` "LTM Revenue Overview" block for the overview pie.
