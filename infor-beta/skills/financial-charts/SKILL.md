---
name: financial-charts
description: >
  Use this skill as the pitch plan's post-aggregation `financial-charts` stage. It builds the four
  INFOR-formatted clustered-column charts for the deck's Financial Summary slide on the combined
  pitch workbook's `financial-summary` tab — where each flow metric's LTM link resolves — then
  renders them into the slide's four chart placeholders. Runs after `workbook-aggregation`.
version: 0.5.16
allowed-tools: [Read, Write, Bash]
---

# Financial Summary Charts — Workflow

This stage finishes the pitch deck's **Financial Summary** slide (library entry 8). The
`financial-summary` stage already built the chart-ready data tab and the deck-assembler already
filled the four metric **name** tiles; this stage builds the **four charts** (one per metric) and
drops them into the slide's four chart placeholders.

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
- **No** major gridlines
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

## Workflow

1. Read `$STAGE_INPUTS`.
2. Call `render_financial_summary_charts_into_deck(...)` (see the reference command). It builds the
   four charts on the combined workbook's `financial-summary` tab and inserts them into the deck.
3. **Overflow / chart QA — mandatory, do not skip.** Render the Financial Summary slide to PNG with
   `render_deck_to_png(deck_path, out, slide_indices=[7])`, read the PNG, and confirm:
   - all four charts landed in their placeholders;
   - the INFOR formatting is applied (bars `46566E`, data labels outside-end, no value axis /
     gridlines / title, Palatino 9);
   - the combined-metric and LTM bars use the **same scale** as the fiscal-year bars (no 10⁶
     mismatch — a symptom of a units mismatch between the `financial-summary` and `ltm-metrics`
     tabs; both must be in millions with an `"MM"` suffix).
   If the renderer is unavailable, say so explicitly rather than skipping silently.
4. Write `$STAGE_OUTPUTS`.

When `$STAGE_OUTPUTS` is **unset** (direct invocation), run the same flow against the supplied
paths and skip the JSON handoff.

## Reference command

```python
import json, os, sys
from pathlib import Path

plugin_root = Path(os.environ.get("CLAUDE_PLUGIN_ROOT", "./infor-beta"))
sys.path.insert(0, str(plugin_root / "scripts"))
from financial_charts import render_financial_summary_charts_into_deck
from slide_render import render_deck_to_png

inputs = json.loads(Path(os.environ["STAGE_INPUTS"]).read_text())
deal_dir = Path(os.environ.get("DEAL_DIR", "."))

out_deck = render_financial_summary_charts_into_deck(
    deck_path=inputs["deck_path"],
    combined_workbook_path=inputs["combined_workbook_path"],
)

if out_deck is None:
    # No financial-summary tab — slide left with placeholders.
    result = {"deck_path": inputs["deck_path"], "charts_inserted": False}
else:
    # Mandatory QA render (zero-based slide index 7 = Financial Summary slide).
    qa_dir = deal_dir / "runs" / "financial-charts-qa"
    render_deck_to_png(out_deck, qa_dir, slide_indices=[7])
    # → read qa_dir/slide_8.png and confirm the four charts + INFOR formatting.
    result = {"deck_path": str(out_deck), "charts_inserted": True}

stage_outputs = os.environ.get("STAGE_OUTPUTS")
if stage_outputs:
    Path(stage_outputs).write_text(json.dumps(result, indent=2))
```

## Outputs (`$STAGE_OUTPUTS`)

```json
{
  "deck_path": "/absolute/path/to/the/final/pitch/deck.pptx",
  "charts_inserted": true
}
```

`deck_path` is the final pitch deck (same file as the `deck` stage's output, modified in place).
`charts_inserted` is `false` when there was no `financial-summary` tab to chart.
