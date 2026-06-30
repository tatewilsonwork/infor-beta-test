---
name: financial-charts
description: >
  Use this skill as the pitch plan's post-aggregation `financial-charts` stage. It builds the four
  INFOR-formatted clustered-column charts for the deck's Financial Summary slide on the combined
  pitch workbook's `financial-summary` tab — where each flow metric's LTM link resolves — then
  renders them into the slide's four chart placeholders. It also builds the overview slide's LTM
  revenue-by-segment pie on the combined workbook's `ltm-metrics` tab and drops it into the
  "[Pie Chart Placeholder]". Runs after `workbook-aggregation`.
version: 0.5.19
allowed-tools: [Read, Write, Bash]
---

# Financial Summary Charts — Workflow

This stage finishes the pitch deck's **Financial Summary** slide (library entry 8). The
`financial-summary` stage already built the chart-ready data tab and the deck-assembler already
filled the four metric **name** tiles; this stage builds the **four charts** (one per metric) and
drops them into the slide's four chart placeholders.

It also fills the **overview** slide's deferred **LTM revenue pie** — see the section below. Both
the FS charts and the pie ride this single post-aggregation stage because the data live on the same
combined workbook (FS charts off the `financial-summary` tab, the pie off the `ltm-metrics` tab).

## Mandatory: build the charts with the helper (no hand-rolling)

This stage **must** build every chart by calling the two `scripts/financial_charts.py`
orchestrators — `render_financial_summary_charts_into_deck(...)` for the four Financial Summary
charts and `render_ltm_revenue_pie_into_deck(...)` for the overview pie (see the Reference
command). Do **not** hand-roll charts with `matplotlib`, `plotly`, Pillow, or any other plotting
library, and do not draw the chart images yourself — the analyst needs real, editable Excel
charts on the workbook, not flat pictures.

Each helper does two things that are both required and must not be split apart:

1. **Persists native chart objects on the combined workbook** — the four clustered-column charts
   on the `financial-summary` tab and the pie on the `ltm-metrics` tab — so opening
   `pitch-<codename>.xlsx` shows real Excel chart objects.
2. **Inserts the rendered charts into the deck** — the four chart placeholders on the Financial
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

**Why this is load-bearing:** this stage runs **after `workbook-aggregation`**, which folds the
standalone `captable` / `ownership` source workbooks into the combined `pitch-<codename>.xlsx` and
then **deletes them**. The `deck-assembler` rebuilds the deck from the slide library and re-pastes
the cap-table and ownership tables from those *standalone* workbooks — which no longer exist. So
re-running it here re-saves a clean deck and **reverts the cap-table and ownership tables to empty
placeholders** (the exact regression seen in a live run). Building charts is the *last* mutation of
the deck; nothing may re-assemble it afterward.

**If a table ever genuinely needs re-inserting** (it should not — `deck` already did it), read it
from the **combined workbook's `captable` / `Ownership` tabs** (the standalone sources are gone
post-aggregation) and paste it with `excel_to_powerpoint.insert_excel_into_placeholder`. Never call
the assembler to do it.

## Why it runs after `workbook-aggregation`

Each **flow** metric's LTM cell on the `financial-summary` tab is a label-keyed lookup
`=INDEX('ltm-metrics'!$B:$B, MATCH("(=) <result_label>", 'ltm-metrics'!$A:$A, 0))`. It is `#N/A`
in the standalone Financial Summary file and resolves **only** in the combined
`pitch-<codename>.xlsx`, where the `ltm-metrics` tab co-exists. So the charts are built on the
**combined workbook's `financial-summary` tab** (after the aggregator folds everything in and a
recalc resolves the links — Excel does the math), and the native charts persist there for the
analyst. The rendered charts are then inserted into the assembled deck.

## Conductor-mode handoff

When invoked as a stage, the environment carries `$STAGE_INPUTS`, `$STAGE_OUTPUTS`, `$DEAL_DIR`:

- Read inputs from `$STAGE_INPUTS`:
  - `combined_workbook_path` — the combined `pitch-<codename>.xlsx` from `workbook-aggregation`.
  - `deck_path` — the assembled pitch deck from `deck`.
  - `deal_dir` — deal directory root (used for the QA render output).
- If `combined_workbook_path` has no `financial-summary` tab (the financial-summary stage produced
  nothing), the slide is left with its placeholders — write the handoff and report the skip.
- If a required field is missing, write `{"error": "missing input: <field>"}` to `$STAGE_OUTPUTS`
  and stop.

## What the charts look like (the only formatting that matters)

One **clustered-column** chart per metric (4 total): a single data series = that metric's row
(`B{r}:G{r}`, `r` ∈ 6,7,8,9), categories = the period header (`B5:G5`). Read the header row
**dynamically** so a suppressed-LTM tab (`B5:F5`) charts five columns instead of six.

- Font: **Palatino Linotype 9 pt, black** (data labels + category-axis labels)
- **No** chart title
- **No** chart border
- **No** major gridlines
- Horizontal (category) axis: a **visible solid black baseline line** beneath the bars — an
  explicit width, never a hairline (a width-less line is dropped by the LibreOffice render, so
  the bars float with no baseline)
- Vertical (value) axis: **hidden** — no line, no label
- Gap width: **50%**
- Data labels on **all** bars, position = **Outside End**
- All bars fill **RGB(70, 86, 110)** (hex `46566E`)

## Placeholder ↔ metric mapping (Financial Summary slide = `prs.slides[7]`)

| Metric | data row | chart placeholder | box (L", T") |
|--------|----------|-------------------|--------------|
| #1 | 6 | `Rectangle 17` | 0.35, 1.51 |
| #2 | 7 | `Rectangle 7`  | 5.12, 1.51 |
| #3 | 8 | `Rectangle 19` | 0.35, 4.42 |
| #4 | 9 | `Rectangle 18` | 5.12, 4.42 |

Each placeholder is 4.53" × 2.51"; the rendered chart is stretched to its box.

## LTM revenue pie (overview slide = `prs.slides[6]`)

The overview ("Introduction to {company}") slide carries a deferred `[Pie Chart Placeholder]`
(shape `Rectangle 4`, box 4.51" × 1.77") under the "LTM Revenue Breakdown" label. This stage fills
it with a by-segment pie built on the combined workbook's **`ltm-metrics`** tab, over the
**"LTM Revenue Overview"** block (categories / legend = the Segment column, series = the
**"% of Total"** column — the `=B/Btotal` fraction Excel computes; the **Total** row is
excluded). The block is located by its section title — no hardcoded row numbers. Charting the
fraction column rather than the dollar amounts gives identical slice geometry but lets the data
labels read the segment share. (Off-Windows, openpyxl can't evaluate the `=B/Btotal` formula, so
the PNG render recomputes the fractions from the literal $ column — the analyst's workbook still
charts the live formula column.)

- Legend at the **TOP**; **no** chart title; **no** chart border
- Slice fills from the **INFOR theme accent palette** (`pptx_helpers.INFOR_ACCENTS`:
  `0E213F, 46566E, ADB9CA, A4844B, 767171, E5E3E3`), in order, cycled past six
- Palatino 9 labels/legend; **value-only** data labels (percentage / category / series / legend-key
  flags off), number format `#,##0.0%_);(#,##0.0%);"--"` so the fraction reads e.g. `45.2%`

When the combined workbook has no `ltm-metrics` tab or no "LTM Revenue Overview" block, the pie is
skipped and the placeholder is left in place (the null path).

## Workflow

1. Read `$STAGE_INPUTS`.
2. Call `render_financial_summary_charts_into_deck(...)` (see the reference command). It builds the
   four charts on the combined workbook's `financial-summary` tab and inserts them into the deck.
3. Call `render_ltm_revenue_pie_into_deck(...)` on the **deck written by step 2** (chain the same
   path) to build + insert the overview LTM revenue pie. A `None` return means the pie was skipped
   (no `ltm-metrics` tab / block) — leave the placeholder.
4. **Chart QA — mandatory, do not skip.** Render the overview slide **and** the Financial Summary
   slide to PNG with `render_deck_to_png(deck_path, out, slide_indices=[6, 7])`, read the PNGs, and
   confirm:
   - **Overview (slide 6):** the pie filled the placeholder; legend at top; INFOR accent slice
     fills; no title / no border; reads legibly in the wide/short box.
   - **Financial Summary (slide 7):** all four charts landed; INFOR formatting (bars `46566E`, data
     labels outside-end, **no border**, a **visible solid black** category-axis baseline line under
     the bars, no value axis / gridlines / title, Palatino 9);
   - the combined-metric and LTM bars use the **same scale** as the fiscal-year bars (no 10⁶
     mismatch — a symptom of a units mismatch between the `financial-summary` and `ltm-metrics`
     tabs; both must be in millions with an `"MM"` suffix).
   If the renderer is unavailable, say so explicitly rather than skipping silently.
5. Write `$STAGE_OUTPUTS`.

When `$STAGE_OUTPUTS` is **unset** (direct invocation), run the same flow against the supplied
paths and skip the JSON handoff.

## Reference command

```python
import json, os, sys
from pathlib import Path

plugin_root = Path(os.environ.get("CLAUDE_PLUGIN_ROOT", "./infor-beta"))
sys.path.insert(0, str(plugin_root / "scripts"))
from financial_charts import (
    render_financial_summary_charts_into_deck,
    render_ltm_revenue_pie_into_deck,
)
from slide_render import render_deck_to_png

inputs = json.loads(Path(os.environ["STAGE_INPUTS"]).read_text())
deal_dir = Path(os.environ.get("DEAL_DIR", "."))
deck_path = inputs["deck_path"]
combined = inputs["combined_workbook_path"]

# 1) Financial Summary charts (slide 7) — modifies the deck in place.
fs_deck = render_financial_summary_charts_into_deck(
    deck_path=deck_path,
    combined_workbook_path=combined,
)
fs_inserted = fs_deck is not None
deck_path = str(fs_deck) if fs_deck is not None else deck_path

# 2) Overview LTM revenue pie (slide 6) — chained onto the FS-charts deck.
pie_deck = render_ltm_revenue_pie_into_deck(
    deck_path=deck_path,
    combined_workbook_path=combined,
)
pie_inserted = pie_deck is not None
deck_path = str(pie_deck) if pie_deck is not None else deck_path

# 3) Mandatory QA render — overview (index 6) + Financial Summary (index 7).
qa_dir = deal_dir / "runs" / "financial-charts-qa"
render_deck_to_png(deck_path, qa_dir, slide_indices=[6, 7])
# → read qa_dir/slide_7.png (overview pie) and slide_8.png (FS charts) and confirm
#   the formatting described in the Workflow QA step.

result = {
    "deck_path": deck_path,
    "charts_inserted": fs_inserted,
    "pie_inserted": pie_inserted,
}

stage_outputs = os.environ.get("STAGE_OUTPUTS")
if stage_outputs:
    Path(stage_outputs).write_text(json.dumps(result, indent=2))
```

## Outputs (`$STAGE_OUTPUTS`)

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
