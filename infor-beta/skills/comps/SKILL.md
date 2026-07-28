---
name: comps
description: >
  Use this skill to build the INFOR public comparable-companies ("public comps" / trading
  comps) table for a target. Activates on /comps and as the pitch plan `comps` stage. Finds
  three verticals (peer groups) relevant to the target, selects six public companies per
  vertical with their Capital IQ tickers, writes a short description for each, and fills the
  INFOR Comps Template — the companion workbook behind the deck's comps slide.
allowed-tools: [Read, Bash, Write, Glob, WebSearch, WebFetch]
---

# INFOR Public Comparables — Workflow & Domain Knowledge

This skill populates the **INFOR Comps Template** with three verticals (peer groups), six
public comparables per vertical, their Capital IQ tickers, and a one-line description for
each. Every market-data / multiple / statistic column in the template is a **Capital IQ
array formula** that resolves off the ticker — so the skill writes only the inputs the
analyst can't compute from CapIQ: the vertical labels, the tickers, and the descriptions.

> **Capital IQ cannot be refreshed here.** This environment has no CapIQ connector, so the
> workbook ships with its formulas **un-evaluated** — every metric cell stays blank / `#NAME?`
> until the analyst opens the file in Excel with the Capital IQ add-in active and refreshes.
> Do **not** try to fetch or hand-fill market data, multiples, or financials; CapIQ owns those.
> The deck's comps slide therefore stays a placeholder for now (no Excel→PowerPoint step).

Today's date is available from the system context (`currentDate`) — do not shell out to `date`.

## Conductor-mode handoff (read first when running under the conductor)

Your dispatch envelope carries three paths — plugin root, `inputs.json`, `outputs.json` — and
every command below takes them **as arguments**
(`python <script.py> "<plugin_root>" "<inputs.json>" "<outputs.json>"`, read back by
`stage_io()`). Nothing is exported; nothing is read from the environment.

- Read inputs from `io.inputs`: `company` (the subject-company facts, to anchor the
  vertical selection and peer set), `ticker` (the target's own symbol, if public — used
  only to avoid listing the target as its own comparable), and `deal_workbook` (the deal's
  ONE workbook, whose `comps` tab you write — do NOT create a standalone comps file).
- At the end, write the structured handoff: `io.write({"workbook_path": str(deal_workbook)})`.
- If an input you need is missing, `io.fail("missing input: <field>")` and stop.

On **direct `/comps` invocation** there is no envelope and no handoff: the analyst supplies
the deal workbook path.

---

## Workflow Steps

### Step 1 — Understand the target

From the company facts (and analyst notes / `cim_path` if attached), establish what the
target actually does: its products / services, end markets, business model, and rough
scale. Verify with **WebSearch** if the facts are thin. If invoked directly with no company
named, ask for the target company before proceeding.

### Step 2 — Choose three verticals (peer groups)

Pick **exactly three** verticals relevant to the target — distinct sub-sectors, business
lines, or peer archetypes a banker would benchmark it against. Each becomes a labelled block
in the workbook (`D9` / `D19` / `D29`). Keep labels concise (≈3–6 words), e.g. for a
diversified financial-software target: `Cloud ERP & Back-Office`, `Payments & Merchant
Acquiring`, `Capital Markets Technology`. Verticals are **business-line / sub-sector** groups,
not geographies — the exchange prefix in each ticker already signals geography.

### Step 3 — Select six public comparables per vertical

For each vertical, choose **exactly six publicly traded** comparables and write their Capital
IQ tickers in **`Exchange:Ticker`** format (see the table below). Use WebSearch to confirm a
current ticker / primary listing when unsure.

Selection criteria (match on as many as possible):
- Business model (revenue type, margin profile) and end market / vertical
- Size — market cap within roughly 0.3×–3× of the target where a peer set allows
- Growth profile (high-growth vs. mature)

Prefer listed companies with **active Capital IQ coverage** and liquid trading. Avoid shell
companies, SPACs, and names mid-M&A (a pending take-private distorts the multiples). Do not
list the **target itself** as one of its comparables. A company may legitimately appear in
only one vertical — don't repeat the same name across blocks.

### Step 4 — Write a short description for each company

One line per company in column AA — **what the company does or sells** (product, service,
asset class, client segment, or business model). Rules:
- Target **30–50 characters**; **never exceed 50** (column AA is ~50 wide — longer overflows,
  and the builder rejects >50).
- **No geography** — the exchange prefix already signals it (`TSX:RY` → Canada).
- No trailing punctuation; title case preferred.
- Examples: `Enterprise workflow & IT service mgmt` (38), `Global payments & merchant acquiring`
  (37), `Diversified multi-asset & alternatives mgr` (43).

### Step 5 — Write the `comps` tab

The deal owns ONE workbook, created at deal-init, and the `comps` tab is already in it with
its CapIQ array formulas and `infor_comps_*` defined names. Pass the workbook path from
`io.inputs["deal_workbook"]` to the shared helper — see the reference command below.

**Do not create a standalone comps workbook, and do not copy a template.** There is no
`output_path` and no `template_path` any more; `write_tab` serializes the write so concurrent
stages in the same wave cannot clobber one another. If the tab is missing or has lost its
defined names, the helper raises `TemplateLayoutError` — stop and report it verbatim.

### Step 6 — Summary

Report: the deal workbook path and the tab written; the three verticals, each with its six companies (ticker — description);
and the reminder that the analyst must **open the file in Excel with the Capital IQ add-in
active and refresh** to populate every market-data / multiple / statistic column (this
environment can't refresh CapIQ). Flag any ticker whose current listing you could not verify.

---

## Reference command

Run as `python <script.py> "<plugin_root>" "<inputs.json>" "<outputs.json>"` — the three paths
your dispatch envelope prints.

```python
import sys
from pathlib import Path

sys.path.insert(0, str(Path(sys.argv[1]) / "scripts"))   # plugin root — arg 1

from stage_io import stage_io
from comps_workbook import build_comps_workbook, Vertical, CompCompany

io = stage_io()
deal_workbook = io.inputs["deal_workbook"]   # the deal's ONE workbook; the `comps` tab is in it

# FORMAT ILLUSTRATION ONLY — the tickers/labels below are obviously-synthetic
# placeholders showing the call shape; NEVER reuse them as data. Every real
# ticker comes from your Step 1-2 peer research, listing-verified.
verticals = [
    Vertical("[Vertical #1 label]", [
        CompCompany("NYSE:AAAA", "Placeholder peer description (<=50 chars)"),
        CompCompany("NasdaqGS:BBBB", "Placeholder peer description (<=50 chars)"),
        # ... six per vertical ...
    ]),
    Vertical("[Vertical #2 label]", [
        CompCompany("TSX:CCCC", "Placeholder peer description (<=50 chars)"),
        # ... six per vertical ...
    ]),
    Vertical("[Vertical #3 label]", [
        CompCompany("NYSE:DDDD", "Placeholder peer description (<=50 chars)"),
        # ... six per vertical ...
    ]),
]

workbook_path = build_comps_workbook(
    verticals=verticals,
    deal_workbook=deal_workbook,
)

io.write({"workbook_path": str(workbook_path)})
```

The builder validates the shape (≤3 verticals, ≤6 companies each, non-empty tickers,
descriptions ≤50 chars), verifies the template's sentinel labels around the hardcoded block
addresses before writing (shared `template_layout` map — a re-saved template with shifted rows
raises `TemplateLayoutError` instead of writing blind), writes only the labels / tickers /
descriptions, and leaves the CapIQ array formulas and statistic rows untouched.

---

## Domain Reference

### Capital IQ ticker format

Use `Exchange:Ticker`. Common prefixes:

| Exchange | Format | Example |
|----------|--------|---------|
| Nasdaq Global Select | `NasdaqGS:TICK` | `NasdaqGS:MSFT` |
| NYSE | `NYSE:TICK` | `NYSE:JPM` |
| TSX | `TSX:TICK` | `TSX:RY` |
| TSX Venture | `TSXV:TICK` | `TSXV:XYZ` |
| London Stock Exchange | `LSE:TICK` | `LSE:HSBA` |
| ASX | `ASX:TICK` | `ASX:CBA` |

### Template cell map (`Comps` sheet)

| Cell(s) | Written by | Meaning |
|---------|-----------|---------|
| `D9` / `D19` / `D29` | skill | Vertical (peer-group) labels — overwrite the `[Group #N]` placeholders |
| `B10:B15` / `B20:B25` / `B30:B35` | skill | Six CapIQ `Exchange:Ticker` identifiers per vertical |
| `AA10:AA15` / `AA20:AA25` / `AA30:AA35` | skill | One-line company descriptions (≤50 chars) |
| `D10:D15`, `G:`–`BH:` | template | Capital IQ array formulas (`SPG($B<row>, …)`) — resolve on refresh — **never overwrite** |
| `D17` / `D27` / `D37`, `D39:D44` | template | Group-average / global average / median / percentile rows — **never overwrite** |
| `F3` / `F4` | template | Output currency / period — left at the template defaults; the analyst adjusts these in Excel |

### Boundary

- Write **only** the vertical labels, tickers, and descriptions. Never touch the CapIQ
  formulas, the statistic rows, or `F3` / `F4`.
- Do **not** fabricate or fetch market data, multiples, growth, or any financial figure —
  Capital IQ populates every metric column on refresh; this environment cannot refresh it.
- Do **not** build the comps chart or edit the deck. The comps slide stays a placeholder; a
  later stage handles the Excel→PowerPoint step once CapIQ refresh is available.
- You are writing the `comps` tab of the deal's single workbook, `pitch-<codename>.xlsx`.
  Since Phase D there is no companion file and no aggregation stage to fold one in; the CapIQ
  `__snloffice` helper sheet was already dropped when the deal-workbook template was built.
