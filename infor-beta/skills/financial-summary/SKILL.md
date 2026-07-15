---
name: financial-summary
description: >
  Use this skill to build the INFOR Financial Summary data tab for a target — the chart-ready
  companion behind the pitch deck's Financial Summary slide(s). Activates on /financial-summary and
  as the pitch plan `financial-summary` stage. Selects the most relevant metrics for the target
  (four per Financial Summary slide: 4 by default, 8 when the deck spec asks for two slides;
  industry-aware: operating company vs. financial institution), gathers their last five fiscal
  years from the latest 10-Ks plus an LTM column, and emits the metric labels the deck tiles use.
  The single source of truth for the deck's financial metrics.
version: 0.5.30
allowed-tools: [Read, Write, Bash, WebSearch, WebFetch]
---

# Financial Summary — Workflow & Domain Knowledge

This skill produces one **"Financial Summary"** Excel tab: the **metrics** it selects as most
relevant to the target — **four per deck Financial Summary slide**, so 4 (the default single
slide) or 8 (the deck spec's two-slide option, passed in as `financial_metric_count`) — each with
its **last five fiscal years** and an **LTM** column. The tab is laid out *chart-ready* so a later
step can drop native Excel charts on it with no reshaping. This stage is the **single source of
truth** for those metrics — it emits their labels, and the deck's Financial Summary tiles read
them from here (no longer from `pitch-content`).

> **It does not build charts and does not touch PowerPoint.** The FS slides' chart regions stay
> placeholders; a later task builds the charts off this tab. This stage produces only the data tab
> plus the labels.

> **Where the LTM value comes from.** LTM is computed on the **`ltm-metrics`** tab (the LTM bridge:
> `FY + current-YTD − prior-YTD`) and **linked** here by a label-keyed formula — not re-derived.
> Because this stage runs *before* `ltm-metrics` (it tells `ltm-metrics` which metrics to bridge),
> the LTM link is a lookup that stays `#N/A` in the standalone file and resolves only in the
> combined `pitch-<codename>.xlsx` — exactly like the cap table's CapIQ formulas. The builder
> writes this for you.

Today's date is available from the system context (`currentDate`) — do not shell out to `date`.

## Conductor-mode handoff (read first when running under the conductor)

When invoked as a stage of a conductor plan, the environment carries `$STAGE_INPUTS`,
`$STAGE_OUTPUTS`, and `$DEAL_DIR`:

- Read inputs from `$STAGE_INPUTS`: `company` (the subject-company facts), `ticker` (used only to
  name the file), `reporting_quarter` / `comparison_quarter` (the latest reported period and
  its prior-year comparison — used to decide the LTM-suppression rule below), and
  `financial_metric_count` (4 or 8; `null`/absent → 4 — sets how many metrics you select and pass
  to the builder's `metric_count`).
- Write the workbook to `$DEAL_DIR/artefacts/<SANITIZED_NAME> - Financial Summary.xlsx` (bootstrap
  `artefacts/` if absent). At the end, write the structured handoff (see **Outputs**).
- If `$STAGE_INPUTS` is missing a field you need, write `{"error": "missing input: <field>"}` to
  `$STAGE_OUTPUTS` and stop.

When `$STAGE_OUTPUTS` is **unset** (direct `/financial-summary` invocation), follow the workflow
as-is — output lands in cwd, no JSON handoff needed.

---

## Workflow Steps

### Step 1 — Understand the target and pick the metric family

From the company facts (and analyst notes / filings), establish what the target does and which
metric **family** fits, following the `comps` / `precedents` precedent:

- **Operating company** → income-statement flow metrics: Revenue, Gross Profit, Adjusted EBITDA,
  EBIT / Operating Income, Net Income (and the like).
- **Financial institution** (bank / lender / insurer / asset manager) → Total Revenue (or Net
  Interest Income + Non-Interest Income), Pre-Provision Net Revenue, Net Income, and balance /
  ratio metrics such as Combined Loan Balances, Book Value, or Return on Equity.

Verify with WebSearch if the facts are thin. If invoked directly with no company named, ask for the
target first.

### Step 2 — Select exactly the requested metric count (NAMES only)

Choose exactly **`financial_metric_count`** metrics (4 when unset) a banker would put on this
target's Financial Summary slide(s), in tile order — with 8, the first four land on FS slide
`(1 of 2)` and the next four on `(2 of 2)`, so lead with the headline metrics (Revenue, EBITDA)
and follow with the secondary ones (margins, income, balances). Use metric **NAMES only** — no
amounts, currency, units, or YoY deltas (e.g. `Revenue`, `Adjusted EBITDA`, `Net Income`,
`Combined Loan Balances`). These labels are the deck tiles' single source of truth. Keep each
≤ 40 characters.

Note for each metric whether it is a **flow** metric (an income-statement figure with a clean
trailing-twelve-month bridge — Revenue, Gross Profit, EBITDA, Operating Income, Net Income) or a
**non-flow** metric (a point-in-time balance or a ratio — loan balances, book value, ROE). This
drives the LTM column below.

### Step 3 — Gather five fiscal years from the latest 10-Ks

Pull each metric's value for the **last five fiscal years** from the **latest four fiscal-year
10-Ks** (each 10-K reports the current + prior fiscal year, so four filings cover five distinct
years; use the one-year overlap between consecutive filings to cross-check the figures). Record the
five values **chronologically (oldest → newest)** with their fiscal-year labels (e.g. `FY2021` …
`FY2025`). Do not estimate a missing year — ask the analyst for the specific 10-K.

**Units — millions, matching the LTM tab (required).** A dollar metric's values **and** its `units`
string are in **millions** of the filing's reporting currency, labelled with an `"MM"` suffix
(`US$MM`, `C$MM`) — never `US$` / full dollars. This must match the `ltm-metrics` tab exactly:
each flow metric's LTM cell links the `ltm-metrics` bridge total **value-for-value** (`=INDEX/MATCH`),
so a scale mismatch (e.g. this tab in `US$MM` while `ltm-metrics` used `US$`) makes the LTM bar 10⁶×
off. Non-dollar metrics use their natural unit (`%`, `x`). Keep the unit constant down the row.

**Combined metrics — pass an Excel formula, never pre-summed.** If a metric combines two or more
reported figures (e.g. "Ending Combined Loan & Advance Bal." = loans + advances), pass each fiscal
value (and, for a non-flow combined metric, the `ltm_value`) as an Excel **formula string** of the
components — `"=9000+800"`, not `9800` — so the arithmetic lives in the cell and stays auditable
(Excel does the math, not you). The builder writes a `"="` string straight through as a formula.

### Step 4 — Decide the LTM column (suppression rule)

LTM = trailing twelve months as of the **most recent reported quarter**.

- If a later interim 10-Q stub exists after the latest fiscal year-end, **show the LTM column**:
  each **flow** metric links to its `ltm-metrics` bridge, and each **non-flow** metric shows its
  latest reported value (`ltm_value`).
- If the most recent filing is a **fiscal year-end 10-K with no later 10-Q stub**, then LTM == the
  latest fiscal year. **Suppress the LTM column** (`show_ltm=False`): show only the five FY columns,
  and for **flow** metrics the most-recent FY cell carries the `ltm-metrics` link instead.

### Step 5 — Build the workbook

Build with the shared helper `build_financial_summary_workbook` (see the reference command). The
builder lays out the chart-ready grid, writes the FY values (numbers, or `"="` formulas for combined
metrics — see Step 3), and writes the LTM links — you supply the metric series. For each **flow**
metric set `result_label` to the exact `ltm-metrics` bridge label its LTM links to (see the
coordination note below). For each **non-flow** metric set `ltm_value` to the latest reported figure
(a number, or a `"="` formula when that figure is itself combined) and leave `result_label = None`.

**Coordinating with `ltm-metrics` (read carefully).** `ltm-metrics` always builds a **Revenue**
bridge (result row `(=) LTM Revenue`) and an **EBITDA** bridge (`(=) LTM Adj. EBITDA`, or
`(=) LTM EBITDA` when no Adjusted figure is disclosed). So:

- If your selection includes **Revenue**, set its `result_label = "LTM Revenue"` and **do not** list it in
  `ltm_bridge_specs`.
- If your selection includes **Adjusted EBITDA**, set its `result_label = "LTM Adj. EBITDA"` (or
  `"LTM EBITDA"` if the company discloses no Adjusted figure) and **do not** list it in
  `ltm_bridge_specs`.
- For every **other flow** metric (Gross Profit, Net Income, Operating Income, …), choose a
  `result_label` of the form `"LTM <Metric>"` (e.g. `"LTM Net Income"`), set it on the metric, **and**
  add `{ "tile_label": <metric name>, "result_label": <same label> }` to `ltm_bridge_specs`.
  `ltm-metrics` will build a matching `(=) <result_label>` bridge for it.

`ltm_bridge_specs` is what makes `ltm-metrics` bridge your selected metrics, so it must list every
non-Revenue/EBITDA **flow** metric you link.

### Step 6 — Summary

Report: the output path; the selected metrics with their family and the five fiscal years used; whether
the LTM column is shown or suppressed; and a reminder that the LTM cells resolve once the workbook
aggregator folds this tab and the `ltm-metrics` tab into the combined `pitch-<codename>.xlsx`.

---

## Outputs (`$STAGE_OUTPUTS`)

```json
{
  "workbook_path": "/absolute/path/to/<Company> - Financial Summary.xlsx",
  "financial_metric_labels": ["Revenue", "Gross Profit", "Adjusted EBITDA", "Net Income"],
  "ltm_bridge_specs": [
    {"tile_label": "Gross Profit", "result_label": "LTM Gross Profit"},
    {"tile_label": "Net Income",   "result_label": "LTM Net Income"}
  ]
}
```

- `financial_metric_labels` — exactly `financial_metric_count` metric names (4 when unset), in
  slide-tile order; the `deck` stage reads these for the FS slide tiles (four per slide).
- `ltm_bridge_specs` — the flow metrics `ltm-metrics` must additionally bridge (excludes Revenue and
  Adj./unadj. EBITDA, which `ltm-metrics` always bridges; excludes non-flow metrics, which have no
  bridge). May be an empty list.

## Chart-ready layout contract

Future chart steps rely on this exact shape on the `Financial Summary` tab (renamed
`financial-summary` in the combined workbook):

| Cell(s) | Contents |
|---------|----------|
| `A1` | `<company> — Financial Summary` (title) |
| `A2` | currency / units note |
| `A3` | period note (`FY = fiscal year; LTM = trailing twelve months as of <period>`) |
| row 5 | header: `A5="Metric"`, `B5..F5` = the five FY labels oldest→newest, `G5="LTM"` (omitted when suppressed), last col `"Units"` |
| rows 6..5+N | one metric per row (N = `financial_metric_count`: rows 6–9 for 4, 6–13 for 8): `A` = metric label, `B..F` = five numeric FY values, LTM cell (flow → `=INDEX('ltm-metrics'!$B:$B, MATCH("(=) <result_label>", 'ltm-metrics'!$A:$A, 0))`; non-flow → latest value), last col = units |

The data block (period columns × metric rows) carries **no merged cells** and **numeric** value
cells, with a single contiguous period header row — so the chart step selects the header row + a
metric row and charts it directly.

---

## Reference command

```python
import json, os, subprocess, sys
from pathlib import Path

plugin_root = Path(os.environ.get("CLAUDE_PLUGIN_ROOT", "./infor-beta"))
sys.path.insert(0, str(plugin_root / "scripts"))
from financial_summary_workbook import build_financial_summary_workbook, MetricSeries

inputs = json.loads(Path(os.environ["STAGE_INPUTS"]).read_text())   # conductor mode
company_name = inputs["company"]["legal_name"]
sanitized = subprocess.run(
    ["bash", str(plugin_root / "scripts" / "sanitize_name.sh"), company_name],
    capture_output=True, text=True,
).stdout.strip() or company_name
out_dir = Path(os.environ.get("DEAL_DIR", ".")) / "artefacts"

fiscal_labels = ["FY2021", "FY2022", "FY2023", "FY2024", "FY2025"]   # oldest -> newest
metrics = [
    MetricSeries("Revenue", "US$MM", [3100.0, 3450.0, 3820.0, 4180.0, 4520.0],
                 result_label="LTM Revenue"),
    MetricSeries("Gross Profit", "US$MM", [1240.0, 1410.0, 1570.0, 1740.0, 1900.0],
                 result_label="LTM Gross Profit"),
    MetricSeries("Adjusted EBITDA", "US$MM", [820.0, 940.0, 1080.0, 1210.0, 1330.0],
                 result_label="LTM Adj. EBITDA"),
    MetricSeries("Net Income", "US$MM", [410.0, 470.0, 540.0, 600.0, 660.0],
                 result_label="LTM Net Income"),
]

workbook_path = build_financial_summary_workbook(
    company_name=company_name,
    currency_note="Figures in US$MM unless noted",
    period_note="FY = fiscal year; LTM = trailing twelve months as of Q3 2026",
    fiscal_labels=fiscal_labels,
    metrics=metrics,
    metric_count=inputs.get("financial_metric_count") or 4,  # 4 or 8; must equal len(metrics)
    show_ltm=True,                       # False when the latest filing is a 10-K with no later stub
    output_dir=out_dir,
    file_stem=f"{sanitized} - Financial Summary",
)

# ltm_bridge_specs: the flow metrics ltm-metrics must additionally bridge (Revenue
# and Adj. EBITDA are always bridged by ltm-metrics, so they are excluded here).
handoff = {
    "workbook_path": str(workbook_path),
    "financial_metric_labels": [m.label for m in metrics],
    "ltm_bridge_specs": [
        {"tile_label": "Gross Profit", "result_label": "LTM Gross Profit"},
        {"tile_label": "Net Income", "result_label": "LTM Net Income"},
    ],
}
Path(os.environ["STAGE_OUTPUTS"]).write_text(json.dumps(handoff, indent=2) + "\n")
```

## Boundary

- Select **exactly `financial_metric_count`** metrics (4 when unset; 8 = the two-slide deck spec);
  emit their names as `financial_metric_labels`. Use NAMES only.
- Gather only the **annual** five-year history here. The LTM bridge math is `ltm-metrics`' job — you
  only *link* to it. Do not hardcode an LTM value for a flow metric.
- Do **not** build charts or edit the deck — the Financial Summary slide's chart regions stay
  placeholders; a later task builds them off this tab.
- The workbook is a companion artefact: the conductor's `workbook-aggregation` stage folds it into
  the combined `pitch-<codename>.xlsx` as the `financial-summary` tab, where the LTM links resolve
  against the `ltm-metrics` tab.
