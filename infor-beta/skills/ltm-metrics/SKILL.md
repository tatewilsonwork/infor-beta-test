---
name: ltm-metrics
description: >
  Use this skill to compute a company's last-twelve-months (LTM) metrics from filings, MD&A,
  or the 10-K/10-Q and emit a standalone Excel workbook. Produces, on one tab: an LTM revenue
  breakdown by segment, an LTM revenue bridge (FY + YTD − prior-year YTD), and an LTM Adj. EBITDA
  (or EBITDA) bridge. Activates as the earnings-update plan stage `ltm-metrics`, supplying the
  companion data behind the overview slide's LTM revenue pie placeholder. Segment by service /
  product line when disclosed, else by geography.
version: 0.5.14
allowed-tools: [Read, Write, Bash, WebSearch, WebFetch]
---

# LTM Metrics — Workflow

This stage produces a single Excel workbook (one **"LTM Metrics"** tab) with three stacked blocks:

1. **LTM Revenue overview** — revenue split by segment, with `% of total` and a total row.
2. **LTM Revenue bridge** — how the LTM revenue total is derived: `LTM = FY + current-year YTD − prior-year YTD`.
3. **LTM Adj. EBITDA bridge** — the same FY + YTD − prior-YTD bridge for Adjusted EBITDA (or unadjusted EBITDA when no Adj. figure is disclosed). **Bridge only — no segment overview.**

It does **not** build a chart and does **not** touch PowerPoint — the overview slide keeps its `[Pie Chart Placeholder]`; the analyst (or a later stage) builds the pie from this workbook.

## Inputs you must have before computing (read first)

LTM is a trailing-four-quarters figure, so a single statement is **not** enough. To bridge `FY + YTD − prior YTD` you need **three** sets of figures:

- the **most-recent full fiscal year** (the FY base), and
- the **current-year interim (YTD) stub** through the latest reported quarter, and
- the **prior-year same-period interim (YTD) stub**.

Example: if the latest report is **Q3 2026** (nine months YTD), the LTM = **FY2025 + Q3 2026 YTD − Q3 2025 YTD**. You therefore need the FY2025 statements/MD&A in addition to the Q3 2026 filing.

If any of these three are missing from the provided sources, **ask the analyst for the specific additional statement** before proceeding (name the period, e.g. "I also need the FY2025 income statement / MD&A to bridge to LTM"). Do not invent quarters or estimate the stub.

> **Do not confuse this with the cap table.** The extra prior-period statement is needed *only* for the LTM bridge math. The cap table is always built off the **most recent** reported statement — never off the older FY statement attached for the LTM calculation.

## Conductor mode

When invoked by the conductor, the environment carries:

- `$STAGE_INPUTS` — JSON with `company`, `ticker`, `reporting_quarter`, and `comparison_quarter`
- `$STAGE_OUTPUTS` — path where this stage must write its structured handoff
- `$DEAL_DIR` — deal directory root

If `$STAGE_INPUTS` is missing a field you need, write `{"error": "missing input: <field>"}` to `$STAGE_OUTPUTS` and stop.

## Workflow

1. Read `$STAGE_INPUTS`.
2. Confirm you have the three input sets above (FY base + current YTD stub + prior-year YTD stub). If not, surface the gap (see "Inputs you must have").
3. From the attached filings (10-K / 10-Q / MD&A / annual report), find the revenue disaggregation note for the **revenue overview**.
   - **Preferred basis:** service / product line (e.g. Cloud Services & Subscriptions, Customer Support, License, Professional Service). Usually in the MD&A revenue discussion or the revenue-recognition footnote.
   - **Fallback basis:** geography (Americas / EMEA / Asia-Pacific) or any other disclosed segmentation, in that order of preference.
4. Build the **revenue bridge** components: FY revenue (additive), current-year YTD revenue (additive), prior-year YTD revenue (subtractive). The workbook computes the LTM total via a formula.
5. Build the **EBITDA bridge** the same way. Prefer **Adjusted EBITDA** if the company discloses it (pass `ebitda_label="LTM Adj. EBITDA"`); fall back to unadjusted EBITDA (`ebitda_label="LTM EBITDA"`) if no Adj. figure is available. If EBITDA is not directly disclosed, derive it from operating income + D&A (+ disclosed adjustments) for each period, and note the basis.
6. Build the workbook with the shared helper. All arithmetic (% of total, totals, the bridge sums) lives in cell formulas — Excel does the math, not the LLM.
7. Write the workbook to `$DEAL_DIR/artefacts/` (bootstrap the folder if needed), then write `$STAGE_OUTPUTS`:

```json
{
  "workbook_path": "/absolute/path/to/<Company> - LTM Metrics.xlsx",
  "ltm_revenue": 4062.0,
  "ltm_adj_ebitda": 2045.0
}
```

`ltm_revenue` and `ltm_adj_ebitda` are the bridge totals **in millions, in the filing's reporting currency** — the same currency as the bridge components. The downstream `captable` stage reads them and writes them to the cap table's LTM column (D47 / D48), applying the cap table's FX rate F7 to convert into the output currency, so emit them unconverted here. Use `bridge_total(...)` to compute each from the same component list you pass to the workbook (omit a key when its bridge is absent).

When `$STAGE_OUTPUTS` is unset (direct invocation), write the workbook to cwd and skip the JSON handoff.

## Reference command

```python
import json, os, sys
from pathlib import Path

plugin_root = Path(os.environ.get("CLAUDE_PLUGIN_ROOT", "./infor-beta"))
sys.path.insert(0, str(plugin_root / "scripts"))
from ltm_metrics import build_ltm_metrics_workbook, bridge_total, RevenueSegment, BridgeComponent

inputs = json.loads(Path(os.environ["STAGE_INPUTS"]).read_text())
deal_dir = Path(os.environ.get("DEAL_DIR", "."))
out_dir = deal_dir / "artefacts"

revenue_bridge = [
    BridgeComponent("FY2025 Revenue", 5400.0),
    BridgeComponent("Q3 2026 YTD Revenue", 3050.0),
    BridgeComponent("Q3 2025 YTD Revenue", 2388.0, subtract=True),
]
ebitda_bridge = [
    BridgeComponent("FY2025 Adj. EBITDA", 1820.0),
    BridgeComponent("Q3 2026 YTD Adj. EBITDA", 1040.0),
    BridgeComponent("Q3 2025 YTD Adj. EBITDA", 815.0, subtract=True),
]

workbook_path = build_ltm_metrics_workbook(
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
    revenue_bridge=revenue_bridge,
    ebitda_bridge=ebitda_bridge,
    ebitda_label="LTM Adj. EBITDA",                # or "LTM EBITDA" if no Adj. disclosed
    output_dir=out_dir,
)

# Mirror the workbook's bridge formulas for the typed handoff (filing-currency millions).
handoff = {"workbook_path": str(workbook_path)}
ltm_revenue = bridge_total(revenue_bridge)
ltm_adj_ebitda = bridge_total(ebitda_bridge)
if ltm_revenue is not None:
    handoff["ltm_revenue"] = ltm_revenue
if ltm_adj_ebitda is not None:
    handoff["ltm_adj_ebitda"] = ltm_adj_ebitda
Path(os.environ["STAGE_OUTPUTS"]).write_text(json.dumps(handoff, indent=2) + "\n")
```

## Boundary

Do not build the pie chart, do not edit the deck, and do not produce the cap table. Those are separate stages. If revenue disaggregation is not disclosed in the provided sources, ask the analyst rather than estimating segment splits. If the prior-period statement needed for the LTM bridge is missing, ask for it rather than estimating the stub.
