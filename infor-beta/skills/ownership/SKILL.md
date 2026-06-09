---
name: ownership
description: >
  Use this skill to build the insider-ownership table for a Canadian public company from a SEDI
  "Insider Information by Issuer" report. Activates on /ownership and as the pitch plan `ownership`
  stage. Parses the analyst-attached SEDI PDF, keeps only current insiders, sums each one's common
  shares, looks up roles, and writes the INFOR ownership workbook (companion to the ownership slide).
version: 0.5.12
allowed-tools: [Read, Bash, Write, Glob, WebSearch, WebFetch]
---

# INFOR Insider Ownership — Workflow & Domain Knowledge

This skill populates the INFOR ownership template's **Select Insiders** block from a SEDI
"Insider Information by Issuer" report, producing the companion workbook whose `B4:G17`
range the deck-assembler pastes into the pitch deck's ownership slide (left "Insiders" side).
The right "Institutions" side is Bloomberg-sourced and out of scope for this skill.

Today's date is available from the system context (`currentDate`) — do not shell out to `date`.

> **SEDI is Canadian-only and cannot be auto-fetched.** SEDI (`sedi.ca`) sits behind a
> Radware/ShieldSquare bot wall, so the report cannot be scraped — the **analyst downloads it
> manually** (English → Access Public Filings → View Summary Reports → Insider information by
> issuer → [company] → View → **Download PDF Report**) and attaches the PDF. See
> `references/sedi-extraction.md`. There is no US equivalent here (US insiders are SEC Forms 3/4/5
> on EDGAR), so this slide applies **only to Canadian reporting issuers**.

## Conductor-mode handoff (read first when running under the conductor)

When invoked as a stage of a conductor plan, the environment carries `$STAGE_INPUTS`,
`$STAGE_OUTPUTS`, and `$DEAL_DIR`:

- Read inputs from `$STAGE_INPUTS`: `company` (facts, to confirm Canadian-public + name), `ticker`,
  and `captable_workbook_path` (the companion cap table, for total shares outstanding). The SEDI PDF
  is an analyst attachment in the chat / deal directory.
- The ownership slide is **optional** — when it can't be produced, emit a **null** workbook so the
  rest of the deck still assembles (the deck-assembler then leaves the slide's placeholders in place).
  Always write a `workbook_path` key (the deck stage references it), null when skipping:
  - Target **not a Canadian reporting issuer** (SEDI is Canadian-only): write
    `{"workbook_path": null, "note": "target is not a Canadian reporting issuer; SEDI ownership N/A"}`
    and stop.
  - **No SEDI PDF attached**: write
    `{"workbook_path": null, "note": "no SEDI PDF attached; ownership slide left as placeholder"}`
    and stop — surface the note at the checkpoint so the analyst can attach it and re-run.
- Otherwise write the workbook to `$DEAL_DIR/artefacts/<SANITIZED_TICKER> - Ownership.xlsx` (bootstrap
  `artefacts/` if absent). At the end, write the structured handoff:
  ```bash
  python -c "import json,os; json.dump({'workbook_path': os.environ['OUTPUT']}, open(os.environ['STAGE_OUTPUTS'], 'w'))"
  ```

When `$STAGE_OUTPUTS` is **unset** (direct `/ownership` invocation), follow the workflow below as-is —
output lands in cwd, no JSON handoff needed.

---

## Workflow Steps

### Step 1 — Confirm the target is a Canadian public company

SEDI only covers Canadian reporting issuers. Confirm from company facts / the ticker exchange
(TSX, TSXV, CSE, or a Canadian-listed interlisted name) — verify with WebSearch if unsure. If the
target is not a Canadian reporting issuer, **stop** (skip handoff above for conductor mode; tell the
analyst the ownership slide does not apply for direct invocation).

### Step 2 — Confirm the SEDI PDF is attached

You need the analyst-downloaded **"Insider Information by Issuer – View Results"** PDF for this
issuer. If it is not attached, ask for it (cite the manual download path in the note above) and wait.

### Step 3 — Locate the template and set the output path

The template is resolved **in Python** in Step 9 as
`Path(os.environ.get("CLAUDE_PLUGIN_ROOT", "./infor-beta")) / "templates" / "INFOR Ownership Template.xlsx"`
— the same primary location `find_template.sh` searches. (Resolving it in Python, rather than via
`TEMPLATE=$(bash find_template.sh …)`, avoids the Git-Bash `/c/…` path that `pathlib` mis-reads on
Windows; this matches the `deck-assembler` / `comps` skills.) Set the sanitized ticker for the
output filename:

```bash
SANITIZED_TICKER=$(bash "${CLAUDE_PLUGIN_ROOT:-./infor-beta}/scripts/sanitize_name.sh" "$TICKER")
OUTPUT="./$SANITIZED_TICKER - Ownership.xlsx"   # or $DEAL_DIR/artefacts/... under the conductor
```

**Do not build the workbook by hand or in any other format.** If the template can't be found, stop
and tell the analyst to confirm `INFOR Ownership Template.xlsx` exists in the plugin `templates/`.

### Step 4 — Parse the SEDI report and keep only CURRENT insiders

Read the attached PDF. It lists every insider the issuer has ever had, each as a block:
`Insider Name:` … `Insider Relationship:` (codes 1–8) … `Ceased to be Insider:` … then security rows.

**Keep only insiders whose `Ceased to be Insider:` is `Not Applicable`.** Drop every insider with a
ceased date — they are former insiders. (See `references/sedi-extraction.md` for the relationship-code
legend and worked parsing examples.)

### Step 5 — Extract common shares per current insider (common shares ONLY)

For each current insider, sum **only the "Common Shares" security rows** — never Options, Warrants,
DSUs, PSUs, or RSUs. An insider may hold common shares under several registered holders (direct, RRSP,
Holdco, spouse, …); collect each tranche's closing balance into a **list** so the workbook writes them
as an in-cell sum formula (`=193000+0+0`) rather than a pre-computed total — auditable, and it never
hand-calculates. A current insider holding **no** common shares is still included (common shares = 0).

### Step 6 — Most recent common-share date per insider

Record the latest transaction date among that insider's common-share rows (ISO `YYYY-MM-DD`).

### Step 7 — Build the adjusted name + role

Convert the SEDI name (usually `Last, First Middle`) to **`First Last`**, then append the person's
role in brackets, e.g. `Mark Barrenechea (CEO & Director)`, `Randy Fowlie (Director)`. Derive the
role from the relationship code (4 = Director, 5 = Senior Officer, 3 = 10% Holder) **plus** the
specific title from the company website / LinkedIn (WebSearch/WebFetch) — e.g. CEO, CFO, Chairman.
The slide shows only the **top 12 insiders by common shares**, so prioritise role accuracy for the
largest holders; minor holders' titles can fall back to the relationship-code role. Flag any role you
could not verify in the Step 10 summary.

### Step 8 — Total basic shares outstanding (F35)

`F35` is the % denominator. Source it from the companion cap table (`captable_workbook_path`):

```python
from ownership_workbook import read_basic_shares_from_cap_table
total = read_basic_shares_from_cap_table(captable_workbook_path)  # full units, or None
```

This sums the cap table's Section VII basic-share input rows (full units). If it returns `None`
(no cap table / unreadable), leave `F35` blank and flag it as a manual step in the summary. Because
F35 comes from the cap table, this stage must run **after** `captable`.

### Step 9 — Write the workbook

```python
import os, sys
from pathlib import Path

plugin_root = Path(os.environ.get("CLAUDE_PLUGIN_ROOT", "./infor-beta"))
sys.path.insert(0, str(plugin_root / "scripts"))
from ownership_workbook import build_ownership_workbook, InsiderHolding

template = plugin_root / "templates" / "INFOR Ownership Template.xlsx"   # native path (no /c/… )

insiders = [
    InsiderHolding("Barrenechea, Mark James", "Mark Barrenechea (CEO & Director)", 1219092, "2025-03-31"),
    InsiderHolding("Fowlie, Randy", "Randy Fowlie (Director)", [193000, 0, 0, 0, 0], "2025-12-01"),
    # ... one InsiderHolding per current insider ...
]
build_ownership_workbook(
    template_path=template,
    insiders=insiders,
    total_shares_outstanding=total,   # full units, or None
    output_path=OUTPUT,
)
```

The builder writes the SEDI name (B), common shares (F, plain or sum formula), date (G), and adjusted
name (J) into rows 39-65, and sets `F35`. It does **not** touch the display block (`B4:G17`), the
include flags (H), the `=H*F` formulas (I), or the institutional side. (The shipped template is
pre-cleaned of the vestigial external links / legacy defined names that would otherwise make the
openpyxl output unopenable in Excel for the render — a regression test guards that it stays clean.)

### Step 10 — Summary

Report: output path; number of current insiders written (and how many ceased insiders were dropped);
the total basic shares used for `F35` and its source (cap table, or "left blank — fill manually");
any insider roles you could not verify; and the reminder that the **institutional / Bloomberg side of
the slide is filled separately** (this skill fills only the insider side).

---

## Domain Reference

- **`references/sedi-extraction.md`** — SEDI access path + bot-wall note, the relationship-code legend
  (1–8), how to read an insider block, common-share vs. equity-award disambiguation, and worked
  examples from a real report.

### Template cell map (`Ownership` sheet)

| Cell(s) | Written by | Meaning |
|---------|-----------|---------|
| `B39:B65` | skill | Raw SEDI name (reference / matching) |
| `F39:F65` | skill | Common shares — plain int, or `=a+b+c` for multiple tranches |
| `G39:G65` | skill | Most recent common-share date |
| `J39:J65` | skill | **Adjusted name + role** — what the slide shows (via `XLOOKUP`) |
| `F35` | skill | Total basic shares outstanding (full units) — % denominator, from the cap table |
| `H39:H65` | template | Include flag (1); analyst toggles to 0 to exclude a row — **do not change** |
| `I39:I65` | template | `=H*F` helper feeding the top-12 `LARGE`/`XLOOKUP` — **do not change** |
| `B4:G17` | template | Display block (top-12 insiders by common shares) — the slide picture range |
| rows 67+ / `Bloomberg Output` | — | Institutional side (Bloomberg) — **out of scope** for this skill |

### Known limitations

- The display block shows the **top 12 insiders by common shares**. If fewer than 12 current insiders
  hold a positive common-share balance, the template's `XLOOKUP(0, …)` repeats the lowest-ranked name
  in the empty slots. With a full current-insider list this does not occur; if it does, the analyst can
  blank the surplus display rows.
- **Render fidelity:** the picture render uses Excel COM on Windows (full fidelity, native `XLOOKUP`)
  and LibreOffice headless elsewhere (Cowork). LibreOffice needs a recent build (24.8+) for `XLOOKUP`;
  on older builds the insider **name** column may render blank — run the ownership stage on Windows +
  Excel for production decks.
