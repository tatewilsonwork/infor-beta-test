---
name: precedents
description: >
  Use this skill to build the INFOR precedent-transactions table for a target. Activates on
  /precedents and as the pitch plan `precedents` stage. Researches up to 12 relevant M&A deals
  (two peer groups of six), picks the metric family that fits the target's industry — EV/Revenue
  + EV/EBITDA for operating companies, or P/E + P/B + P/TBV for financial institutions — and
  fills the INFOR Precedents Template, hyperlinking each figure's source. The companion workbook
  behind the deck's precedent-transactions slide.
version: 0.5.16
allowed-tools: [Read, Bash, Write, Glob, WebSearch, WebFetch]
---

# INFOR Precedent Transactions — Workflow & Domain Knowledge

This skill populates the **INFOR Precedents Template** with up to **12 precedent M&A
transactions** (two peer groups of six), each carrying the deal identity, the source-FX $
metric inputs for the **one metric family** that fits the target's industry, source
hyperlinks, and — when a multiple is disclosed in the deal PR — the multiple written straight
over the template's ratio formula. The FX conversion, the per-row ratio formulas, and the
group / global statistic rows are Capital IQ / template formulas that resolve in Excel — the
skill only writes the inputs.

> **Capital IQ cannot be refreshed here.** The column-C FX rate is a CapIQ array formula, so
> the converted TEV (column J), the ratios (S–Z), and the statistic rows stay un-evaluated
> (blank / `#NAME?`) until the analyst opens the workbook in Excel with the Capital IQ add-in
> active and refreshes. A **disclosed multiple** you write into S–Z is a literal and shows
> immediately; a **$ metric** you write into K–R only resolves into a ratio after refresh.
> The deck's precedent-transactions slide therefore stays a placeholder for now.

Today's date is available from the system context (`currentDate`) — do not shell out to `date`.

**Detailed sourcing rules** (load on demand): [`references/sourcing-criteria.md`](references/sourcing-criteria.md)
— the source ladder, the "disclosed" definition, the pro-forma rule, currency handling, the
metric-family decision, transaction-selection criteria, and the reputable-source domains.

## Conductor-mode handoff (read first when running under the conductor)

When invoked as a stage of a conductor plan, the environment carries `$STAGE_INPUTS`,
`$STAGE_OUTPUTS`, and `$DEAL_DIR`:

- Read inputs from `$STAGE_INPUTS`: `company` (the subject-company facts, to anchor the metric
  family and the comparable-deal set) and `ticker` (the target's own symbol, if public — used
  only to name the file).
- Write the workbook to `$DEAL_DIR/artefacts/<SANITIZED_NAME> - Precedent Transactions.xlsx`
  (bootstrap `artefacts/` if absent). At the end, write the structured handoff:
  ```bash
  python -c "import json,os; json.dump({'workbook_path': os.environ['OUTPUT']}, open(os.environ['STAGE_OUTPUTS'], 'w'))"
  ```
- If `$STAGE_INPUTS` is missing a field you need, write `{"error": "missing input: <field>"}`
  to `$STAGE_OUTPUTS` and stop.

When `$STAGE_OUTPUTS` is **unset** (direct `/precedents` invocation), follow the workflow below
as-is — output lands in cwd, no JSON handoff needed.

---

## Workflow Steps

### Step 1 — Understand the target and choose the metric family

From the company facts (and analyst notes / `cim_path` if attached), establish what the target
does, its sector, and its business model. Verify with **WebSearch** if the facts are thin. If
invoked directly with no company named, ask for the target before proceeding.

Then choose **one** metric family for the whole table, by industry:

| Family | Use for | Fill (source-FX $) | Multiples shown |
|--------|---------|--------------------|-----------------|
| **Operating** | most companies — software, industrials, consumer, healthcare, services | Revenue (LTM/NTM), Adj. EBITDA (LTM/NTM) | EV/Revenue, EV/EBITDA |
| **Financial** | banks, insurers, asset/wealth managers, other balance-sheet-driven financials | Net Income (LTM/NTM), Book Value, Tangible Book Value | P/E, P/B, P/TBV |

Fill **only** the chosen family's input columns; leave the other family's columns blank. See
[`references/sourcing-criteria.md`](references/sourcing-criteria.md) for edge cases.

### Step 2 — Research the precedent transactions

Search for relevant M&A deals where the **target** is comparable to the input company (sector,
business model, client segment, scale). Two hard requirements for every deal:

1. A **disclosed TEV** — never include a deal with no deal value.
2. **At least one multiple you can show** from the chosen family — either a multiple disclosed
   in the deal PR, or a disclosed $ metric (Revenue / EBITDA, or Net Income / Book Value for
   the financial family) the template's ratio formula turns into one. **Do not include a deal
   you can only give a TEV for** — it just adds an empty row. The builder enforces this and will
   reject a metric-less deal.

Prefer deals announced in the last ~6–8 years. Organise them into **two peer groups, and aim to
fill all six rows in each** (12 deals total, e.g. by sub-sector or deal type); each group gets a
short label (`E7` / `E16`). Cast a wide enough net that you can find six **valuable** deals per
group — don't stop at the first handful. A single group is acceptable only when a credible
second peer group genuinely doesn't exist; otherwise fill both. Leave an unused group as its
`[Group #2]` placeholder.

See [`references/sourcing-criteria.md`](references/sourcing-criteria.md) for the full
selection criteria.

### Step 3 — Source each value (multiple first, then $ metric)

For each deal, source the family's figures top-down — stop at the first that's disclosed:

1. **Disclosed multiple** in the deal PR / 8-K / transcript / major news — *preferred*. Write
   it as a literal into the multiple column (S–Z); it overwrites the ratio formula. A directly
   quoted "~12.5x LTM EBITDA" beats working a ratio back from a $ figure and TEV.
2. **Disclosed $ metric** — write the source-currency $MM value into the input column (K–R);
   the template's ratio formula computes the multiple on refresh. When a clean LTM/NTM figure
   isn't disclosed, **use the most recent reported figure as the LTM/NTM proxy** — do *not*
   reconstruct LTM by stitching multiple filings together.
3. Otherwise leave that individual figure's cell blank — but a deal must end up with **at least
   one** multiple or $ metric across the family (see Step 2); if you can't source any, drop the
   deal and find another, rather than shipping a TEV-only row.

**Hyperlink every figure's source** onto its link cell (AB–AG), one per metric concept. Use
reputable primary sources (deal PR, filings, major financial news) — see the reputable-source
list in the reference. The HQ country of the target goes in **column AI as a 3-letter code**
(`USA`, `CAN`, `GBR`). Column H is left empty.

### Step 4 — Build the workbook

Resolve the template at `$CLAUDE_PLUGIN_ROOT/templates/INFOR Precedents Template.xlsx` **in
Python** (resolving it in Python avoids the Git-Bash `/c/…` path that breaks `pathlib` on
Windows) and build with the shared helper — see the reference command below. **Do not build
the workbook by hand or in any other format.** If the template can't be found, stop and tell
the analyst to confirm `INFOR Precedents Template.xlsx` exists in the plugin `templates/`.

### Step 5 — Summary

Report: the output path; the group labels with each deal (target → acquiror, announce date,
TEV); which metric family was used; for each figure whether it came from a disclosed multiple
or a disclosed $ metric (note proxies); any figure left blank for non-disclosure; and the
reminder that the analyst must **open the file in Excel with the Capital IQ add-in active and
refresh** to populate the FX rate, the converted TEV, the $-metric ratios, and the statistic
rows.

---

## Reference command

```python
import json, os, subprocess, sys
from datetime import date
from pathlib import Path

plugin_root = Path(os.environ.get("CLAUDE_PLUGIN_ROOT", "./infor-beta"))
sys.path.insert(0, str(plugin_root / "scripts"))
from precedents_workbook import build_precedents_workbook, PrecedentGroup, PrecedentTransaction

inputs = json.loads(Path(os.environ["STAGE_INPUTS"]).read_text())   # conductor mode
company_name = inputs["company"]["legal_name"]
template = plugin_root / "templates" / "INFOR Precedents Template.xlsx"  # native path (no /c/…)
sanitized = subprocess.run(
    ["bash", str(plugin_root / "scripts" / "sanitize_name.sh"), company_name],
    capture_output=True, text=True,
).stdout.strip() or company_name
out_dir = Path(os.environ.get("DEAL_DIR", ".")) / "artefacts"       # cwd for direct /precedents

# Operating-family example (Revenue + EBITDA). Pick ONE family for the whole table.
groups = [
    PrecedentGroup("Vertical Software", [
        # Row sourced from a disclosed multiple -> literal into the ratio column.
        PrecedentTransaction(
            input_currency="USD", announce_date=date(2022, 8, 25),
            target="Micro Focus International", acquiror="OpenText", tev=6000.0, hq_country="GBR",
            ev_ebitda_ltm=6.3, ev_revenue_ltm=2.3,            # disclosed multiples (S/U)
            tev_link="https://investors.opentext.com/press-releases/.../default.aspx",
            ebitda_link="https://investors.opentext.com/press-releases/.../default.aspx",
            revenue_link="https://investors.opentext.com/press-releases/.../default.aspx",
        ),
        # Row sourced from disclosed $ figures -> ratio formulas compute on refresh.
        PrecedentTransaction(
            input_currency="USD", announce_date=date(2025, 5, 6),
            target="AvidXchange Holdings", acquiror="TPG / Corpay", tev=2200.0, hq_country="USA",
            revenue_ltm=440.0, ebitda_ltm=85.0,               # source-FX $MM (K/O)
            tev_link="https://www.tpg.com/news-and-insights/...",
            revenue_link="https://www.sec.gov/...",
            ebitda_link="https://www.sec.gov/...",
        ),
        # ... up to six per group ...
    ]),
    # Optional second group; omit it to leave the [Group #2] placeholder.
]

workbook_path = build_precedents_workbook(
    template_path=template,
    groups=groups,
    output_currency="USD",                                   # -> C2 (the FX formula keys off it)
    output_path=out_dir / f"{sanitized} - Precedent Transactions.xlsx",
)

Path(os.environ["STAGE_OUTPUTS"]).write_text(
    json.dumps({"workbook_path": str(workbook_path)}, indent=2) + "\n"
)
```

For a **financial-institution** target, swap the metric inputs: use `net_income_ltm` /
`net_income_ntm` (P/E), `book_value` (P/B), `tangible_book_value` (P/TBV), and the matching
`pe_ltm` / `pe_ntm`, `pb`, `ptbv` multiple-overwrites, with `net_income_link` / `book_value_link`
/ `tangible_book_value_link`.

The builder validates the shape (≤2 groups, ≤6 transactions each, a positive TEV, 3-letter
currency / HQ codes, numeric metrics, http(s) links), writes only the inputs, and leaves the FX
/ ratio / statistic formulas untouched.

---

## Domain Reference

### Template cell map (`Precedents` sheet)

Two groups: rows **8–13** (label `E7`) and **17–22** (label `E16`). Write per transaction row:

| Cell | Field | Notes |
|------|-------|-------|
| `B` | Input currency | ISO-3 (drives the column-C FX formula) |
| `E` | Announce date | a real date; FX is taken as of the workday before it |
| `F` / `G` | Target / Acquiror | legal names |
| `I` | TEV | source-FX $MM (column J `=+I*C` converts it) |
| `AI` | HQ | **3-letter** country code (`USA`, `CAN`, `GBR`); column H stays empty |
| `K` / `L` | Revenue LTM / NTM | *operating family* — source-FX $MM |
| `O` / `P` | Adj. EBITDA LTM / NTM | *operating family* — source-FX $MM |
| `M` / `N` | Net Income LTM / NTM | *financial family* — source-FX $MM |
| `Q` / `R` | Book Value / Tangible BV | *financial family* — source-FX $MM |
| `S` / `T` | EV/Revenue LTM / NTM | **formula** — overwrite with a literal only if the multiple is disclosed |
| `U` / `V` | EV/EBITDA LTM / NTM | same |
| `W` / `X` | P/E LTM / NTM | same |
| `Y` / `Z` | P/B / P/TBV | same |
| `AB`–`AG` | Source links | hyperlink per metric: AB TEV, AC Revenue, AD EBITDA, AE Net Income, AF Book Value, AG Tangible BV |
| `C2` | Output currency | ISO-3 the FX formula keys off (builder writes it; default `USD`) |

### Cells the skill must NOT touch

- `C` (CapIQ FX array formula), `J` (`=+I*C`), and the `S–Z` ratio formulas **except** when
  overwriting a ratio cell with a *disclosed* multiple.
- The group-average (rows 14 / 23), global-average / median (25 / 26), and percentile (29 /
  30) rows — all template formulas.
- The two-row header (rows 4–5).

### HQ country codes — common 3-letter examples

| Country | Code | Country | Code |
|---------|------|---------|------|
| United States | USA | Australia | AUS |
| Canada | CAN | France | FRA |
| United Kingdom | GBR | Germany | DEU |
| Ireland | IRL | Netherlands | NLD |

### Boundary

- Write **only** the inputs above. Never fabricate a figure — leave an individual unsourced
  figure's cell blank. But every **deal** must carry at least one multiple or $ metric (Step 2);
  the builder rejects a deal with only a TEV, so drop and replace any deal you can't value.
- Pick **one** metric family per table; do not fill both families' columns.
- Do **not** build the precedents chart or edit the deck. The slide stays a placeholder until
  Capital IQ refresh is available.
- The workbook is a companion artefact: the conductor's `workbook-aggregation` stage folds it
  into the combined `pitch-<deal>.xlsx` as the `precedents` tab.
