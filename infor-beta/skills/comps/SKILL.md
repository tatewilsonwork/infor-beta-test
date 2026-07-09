---
name: comps
description: >
  Use this skill to build the INFOR public comparable-companies ("public comps" / trading
  comps) table for a target. Activates on /comps and as the pitch plan `comps` stage. Finds
  three verticals (peer groups) relevant to the target, selects six public companies per
  vertical with their Capital IQ tickers, writes a short description for each, and fills the
  INFOR Comps Template — the companion workbook behind the deck's comps slide.
version: 0.5.23
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

When invoked as a stage of a conductor plan, the environment carries `$STAGE_INPUTS`,
`$STAGE_OUTPUTS`, and `$DEAL_DIR`:

- Read inputs from `$STAGE_INPUTS`: `company` (the subject-company facts, to anchor the
  vertical selection and peer set) and `ticker` (the target's own symbol, if public — used
  only to name the file and to avoid listing the target as its own comparable).
- Write the workbook to `$DEAL_DIR/artefacts/<SANITIZED_NAME> - Comparable Companies.xlsx`
  (bootstrap `artefacts/` if absent). At the end, write the structured handoff:
  ```bash
  python -c "import json,os; json.dump({'workbook_path': os.environ['OUTPUT']}, open(os.environ['STAGE_OUTPUTS'], 'w'))"
  ```
- If `$STAGE_INPUTS` is missing a field you need, write `{"error": "missing input: <field>"}`
  to `$STAGE_OUTPUTS` and stop.

When `$STAGE_OUTPUTS` is **unset** (direct `/comps` invocation), follow the workflow below
as-is — output lands in cwd, no JSON handoff needed.

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

### Step 5 — Build the workbook

Resolve the template at `$CLAUDE_PLUGIN_ROOT/templates/INFOR Comps Template.xlsx` **in Python**
(the same primary location `find_template.sh` searches; resolving it in Python avoids the
Git-Bash `/c/…` path that breaks `pathlib` on Windows) and build with the shared helper — see
the reference command below.

**Do not build the workbook by hand or in any other format.** If the template can't be found,
stop and tell the analyst to confirm `INFOR Comps Template.xlsx` exists in the plugin
`templates/`.

### Step 6 — Summary

Report: the output path; the three verticals, each with its six companies (ticker — description);
and the reminder that the analyst must **open the file in Excel with the Capital IQ add-in
active and refresh** to populate every market-data / multiple / statistic column (this
environment can't refresh CapIQ). Flag any ticker whose current listing you could not verify.

---

## Reference command

```python
import json, os, subprocess, sys
from pathlib import Path

plugin_root = Path(os.environ.get("CLAUDE_PLUGIN_ROOT", "./infor-beta"))
sys.path.insert(0, str(plugin_root / "scripts"))
from comps_workbook import build_comps_workbook, Vertical, CompCompany

inputs = json.loads(Path(os.environ["STAGE_INPUTS"]).read_text())   # conductor mode
company_name = inputs["company"]["legal_name"]
template = plugin_root / "templates" / "INFOR Comps Template.xlsx"   # native path (no /c/… )
sanitized = subprocess.run(
    ["bash", str(plugin_root / "scripts" / "sanitize_name.sh"), company_name],
    capture_output=True, text=True,
).stdout.strip() or company_name
out_dir = Path(os.environ.get("DEAL_DIR", ".")) / "artefacts"       # cwd for direct /comps

verticals = [
    Vertical("Cloud ERP & Back-Office", [
        CompCompany("NYSE:NOW", "Enterprise workflow & IT service mgmt"),
        CompCompany("NasdaqGS:WDAY", "Cloud HR & finance suite"),
        # ... six per vertical ...
    ]),
    Vertical("Payments & Merchant Acquiring", [
        CompCompany("NYSE:FIS", "Banking & payments technology"),
        # ... six per vertical ...
    ]),
    Vertical("Capital Markets Technology", [
        CompCompany("NasdaqGS:SSNC", "Investment & fund administration software"),
        # ... six per vertical ...
    ]),
]

workbook_path = build_comps_workbook(
    template_path=template,
    verticals=verticals,
    output_path=out_dir / f"{sanitized} - Comparable Companies.xlsx",
)

Path(os.environ["STAGE_OUTPUTS"]).write_text(
    json.dumps({"workbook_path": str(workbook_path)}, indent=2) + "\n"
)
```

The builder validates the shape (≤3 verticals, ≤6 companies each, non-empty tickers,
descriptions ≤50 chars), writes only the labels / tickers / descriptions, and leaves the
CapIQ array formulas and statistic rows untouched.

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
- The workbook is a companion artefact: the conductor's `workbook-aggregation` stage folds it
  into the combined `pitch-<deal>.xlsx` as the `comps` tab (the CapIQ `__snloffice` helper
  sheet is dropped automatically).
