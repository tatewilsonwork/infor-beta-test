---
name: ltm-revenue
description: >
  Use this skill to extract a company's last-twelve-months (LTM) revenue breakdown from filings,
  MD&A, or the 10-K and emit a standalone Excel workbook. Activates as the earnings-update plan
  stage `ltm-revenue`, supplying the companion breakdown behind the overview slide's LTM revenue
  pie placeholder. Segment by service / product line when disclosed, else by geography.
version: 0.5.2
allowed-tools: [Read, Write, Bash, WebSearch, WebFetch]
---

# LTM Revenue Breakdown — Workflow

This stage produces a single Excel workbook with the company's LTM revenue split by segment. It does **not** build a chart and does **not** touch PowerPoint — the overview slide keeps its `[Pie Chart Placeholder]`; the analyst (or a later stage) builds the pie from this workbook.

## Conductor mode

When invoked by the conductor, the environment carries:

- `$STAGE_INPUTS` — JSON with `company`, `ticker`, and `reporting_quarter`
- `$STAGE_OUTPUTS` — path where this stage must write its structured handoff
- `$DEAL_DIR` — deal directory root

If `$STAGE_INPUTS` is missing a field you need, write `{"error": "missing input: <field>"}` to `$STAGE_OUTPUTS` and stop.

## Workflow

1. Read `$STAGE_INPUTS`.
2. From the attached filings (10-K / 10-Q / MD&A / annual report), find the revenue disaggregation note.
   - **Preferred basis:** service / product line (e.g. Cloud Services & Subscriptions, Customer Support, License, Professional Service). This is usually in the MD&A revenue discussion or the revenue-recognition footnote.
   - **Fallback basis:** geography (Americas / EMEA / Asia-Pacific) or any other disclosed segmentation, in that order of preference.
3. Compute **LTM** figures (trailing four quarters). If a clean LTM is not derivable, use the most recent full fiscal year and note the basis in the period label — do not invent quarters.
4. Build the workbook with the shared helper. Segment values are raw numbers in the same currency as the deck; the workbook computes `% of total` and the total via formulas (Excel does the math, not the LLM).
5. Write the workbook to `$DEAL_DIR/artefacts/` (bootstrap the folder if needed), then write `$STAGE_OUTPUTS`:

```json
{
  "workbook_path": "/absolute/path/to/<Company> - LTM Revenue Breakdown.xlsx"
}
```

When `$STAGE_OUTPUTS` is unset (direct invocation), write the workbook to cwd and skip the JSON handoff.

## Reference command

```python
import json, os, sys
from pathlib import Path

plugin_root = Path(os.environ.get("CLAUDE_PLUGIN_ROOT", "./infor-beta"))
sys.path.insert(0, str(plugin_root / "scripts"))
from ltm_revenue import build_ltm_revenue_workbook, RevenueSegment

inputs = json.loads(Path(os.environ["STAGE_INPUTS"]).read_text())
deal_dir = Path(os.environ.get("DEAL_DIR", "."))
out_dir = deal_dir / "artefacts"

workbook_path = build_ltm_revenue_workbook(
    company_name=inputs["company"]["legal_name"],
    period_label="LTM ended March 31, 2026",      # set from the actual reporting period
    currency="US$MM",                              # match the deck currency footnote
    segmentation_basis="Service line",             # or "Geography" on fallback
    segments=[
        RevenueSegment("Cloud Services & Subscriptions", 1932.0),
        RevenueSegment("Customer Support", 1480.0),
        RevenueSegment("License", 360.0),
        RevenueSegment("Professional Service & Other", 290.0),
    ],
    output_dir=out_dir,
)
Path(os.environ["STAGE_OUTPUTS"]).write_text(json.dumps({"workbook_path": str(workbook_path)}, indent=2) + "\n")
```

## Boundary

Do not build the pie chart, do not edit the deck, and do not produce the cap table. Those are separate stages. If revenue disaggregation is not disclosed in the provided sources, ask the analyst rather than estimating segment splits.
