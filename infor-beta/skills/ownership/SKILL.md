---
name: ownership
description: >
  Use this skill to build the ownership table for a Canadian public company from a SEDI
  "Insider Information by Issuer" report, plus — when attached — a Bloomberg ownership export for
  the institutional side. Activates on /ownership and as the pitch plan `ownership` stage. Parses
  the analyst-attached SEDI PDF, keeps only current insiders, sums each one's common shares, looks
  up roles, ingests the Bloomberg holders (excluding SEDI duplicates — the SEDI figure always
  wins), and writes the INFOR ownership workbook (companion to the ownership slide).
allowed-tools: [Read, Bash, Write, Glob, WebSearch, WebFetch]
---

# INFOR Insider Ownership — Workflow & Domain Knowledge

This skill populates the INFOR ownership template's **Select Insiders** block from a SEDI
"Insider Information by Issuer" report, producing the companion workbook whose `B4:G17`
range the deck-assembler pastes into the pitch deck's ownership slide (left "Insiders" side).
When the analyst also attaches a **Bloomberg ownership export**, the skill fills the
**Select Institutions** side too (Step 9): the export replaces the workbook's `Bloomberg
Output` tab, the Ownership tab's pre-wired link rows (68-185) pick it up, Bloomberg rows
that duplicate a SEDI insider are excluded (**the SEDI figure always wins — only the
Bloomberg duplicate is flagged out, never the SEDI row**), and the deck-assembler then
pastes `B19:G35` into the slide's right "Institutions" side. Without a Bloomberg export
the institutional side stays a placeholder, as before.

Today's date is available from the system context (`currentDate`) — do not shell out to `date`.

> **SEDI is Canadian-only and cannot be auto-fetched.** SEDI (`sedi.ca`) sits behind a
> Radware/ShieldSquare bot wall, so the report cannot be scraped — the **analyst downloads it
> manually** (English → Access Public Filings → View Summary Reports → Insider information by
> issuer → [company] → View → **Download PDF Report**) and attaches the PDF. See
> `references/sedi-extraction.md`. There is no US equivalent here (US insiders are SEC Forms 3/4/5
> on EDGAR), so this slide applies **only to Canadian reporting issuers**.

## Conductor-mode handoff (read first when running under the conductor)

Your dispatch envelope carries three paths — plugin root, `inputs.json`, `outputs.json` — and
every command below takes them **as arguments**
(`python <script.py> "<plugin_root>" "<inputs.json>" "<outputs.json>"`, read back by
`stage_io()`). Nothing is exported; nothing is read from the environment.

- Read inputs from `io.inputs`: `company` (facts, to confirm Canadian-public + name), `ticker`,
  `deal_workbook` (the deal's ONE workbook, whose `Ownership` + `Bloomberg Output` tabs you write),
  and `captable_workbook_path` (the companion cap table, for total shares outstanding). The SEDI PDF
  — and, optionally, the Bloomberg ownership export (`.xlsm`/`.xlsx`) — are analyst attachments in
  the chat / deal directory. A missing Bloomberg export is **not** an error: build insider-only and
  note that the institutions side was left as a placeholder.
- The ownership slide is **optional** — when it can't be produced, emit a **null** workbook so the
  rest of the deck still assembles (the deck-assembler then leaves the slide's placeholders in place).
  Always write a `workbook_path` key (the deck stage references it), null when skipping:
  - Target **not a Canadian reporting issuer** (SEDI is Canadian-only): write
    `{"workbook_path": null, "note": "target is not a Canadian reporting issuer; SEDI ownership N/A"}`
    and stop.
  - **No SEDI PDF attached**: write
    `{"workbook_path": null, "note": "no SEDI PDF attached; ownership slide left as placeholder"}`
    and stop — surface the note at the checkpoint so the analyst can attach it and re-run.
- Otherwise write the `Ownership` + `Bloomberg Output` tabs of the deal workbook, then write the
  structured handoff: `io.write({"workbook_path": str(deal_workbook)})`.

On **direct `/ownership` invocation** there is no envelope and no handoff: the analyst supplies
the deal workbook path.

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

The template is resolved **in Python** in Step 10 as
`Path(sys.argv[1]) / "templates" / "INFOR Ownership Template.xlsx"` — the plugin root your
dispatch envelope passes as the first argument, and the same primary location
`find_template.sh` searches. (Resolving it in Python, rather than via
`TEMPLATE=$(bash find_template.sh …)`, avoids the Git-Bash `/c/…` path that `pathlib` mis-reads on
Windows; this matches the `comps` skill and the in-process deck assembler.) The deal workbook path is
`io.inputs["deal_workbook"]`.

**Do not build the workbook by hand or in any other format.** If the template can't be found, stop
and tell the analyst to confirm `INFOR Ownership Template.xlsx` exists in the plugin `templates/`.

### Step 4 — Parse the SEDI report and keep only CURRENT insiders

Read the attached PDF. It lists every insider the issuer has ever had, each as a block:
`Insider Name:` … `Insider Relationship:` (codes 1–8) … `Ceased to be Insider:` … then security rows.

SEDI exports are normally clean text (see `references/sedi-extraction.md`), but if a particular report comes through garbled (CID-font scramble, U+FFFD characters, or blank pages), read it via the shared `<plugin_root>/scripts/pdf_extract.py` helper (`from pdf_extract import extract_pdf_text`), which detects garble and falls back to rendering + tesseract OCR — never transcribe scrambled glyphs.

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
could not verify in the Step 11 summary.

### Step 8 — Total basic shares outstanding (F35)

`F35` is the % denominator. Source it from the companion cap table (`captable_workbook_path`):

```python
from ownership_workbook import read_basic_shares_from_cap_table
total = read_basic_shares_from_cap_table(captable_workbook_path)  # full units, or None
```

This sums the cap table's Section VII basic-share input rows (full units). If it returns `None`
(no cap table / unreadable), leave `F35` blank and flag it as a manual step in the summary. Because
F35 comes from the cap table, this stage must run **after** `captable`. The summing window comes from
the cap table's `infor_cap_share_inputs` name, so an inserted Section VII row is summed too; a
*readable* cap table that has lost that name raises `TemplateLayoutError` instead of silently summing
a hardcoded window — report that to the analyst rather than working around it.

### Step 9 — Bloomberg institutional side (optional)

Look for an analyst-attached **Bloomberg ownership export** — the BBG Excel add-in "Ownership"
template (`.xlsm`, sometimes `.xlsx`) whose **Summary View** sheet lists holders from row 14
(`Holder Name` / `Position` / `Filing Date` / `Insider Status` / …). Like the SEDI PDF, it cannot
be auto-fetched — the analyst runs the Bloomberg pull and attaches the file. **If none is
attached, skip this step** (insider-only workbook; the slide's Institutions side stays a
placeholder) and say so in the summary — do not stop the stage.

When attached:

1. **Parse + review the duplicate matches.** `read_bloomberg_export` parses the file;
   `match_bloomberg_to_sedi` deterministically maps each Bloomberg holder to the SEDI insider it
   duplicates (SEDI `Last, First Middle` ↔ BBG `Last First Middle`, initials tolerated; corporate
   holders match after legal-suffix stripping). Every matched Bloomberg row is excluded (`H=0`)
   by the builder — **the SEDI figure always wins; never exclude the SEDI row.** Review the
   report: a Bloomberg row flagged `Insider Status = "Y"` that did **not** match stays *included*
   (only confirmed SEDI duplicates are excluded) — list those unmatched insiders in the summary
   for the analyst, and use `bloomberg_include_overrides={"<Holder Name>": 0}` only when the
   analyst confirms one is a missed duplicate.
2. **Adjusted names.** The builder defaults each Bloomberg row's display name (col J) to
   `strip_legal_suffixes` — trailing Inc / Corp / Group / Partners / LP / Ltd / … tokens dropped
   (`"T Rowe Price Group Inc"` → `"T Rowe Price"`, `"Kelso & Co LP"` → `"Kelso & Co"`; an `"& Co"`
   brand is kept). Where house style clearly differs, pass
   `bloomberg_adjusted_names={"<Holder Name>": "<display name>"}` (e.g. restore `"T. Rowe Price"`
   punctuation). Do not shorten person names.

### Step 9b — Provenance — REQUIRED, and structured

**Every share count carries a record.** Each insider's holding is a figure you read off the
analyst-attached SEDI report, and each institutional position is a figure you read off the Bloomberg
export — but until now this stage recorded neither, so a run's provenance ledger held 70 records with
`ownership` contributing zero and every percentage on the ownership slide untraceable.

Give each `InsiderHolding` a `source`: a `provenance.FigureSource` naming the **SEDI report** as the
filing, the insider as the statement/section, and the **page** their holdings are on —
`FigureSource(filing="SEDI Insider Information by Issuer report", statement="<insider name>",
page=99)`. A **citation string is rejected**: it would build a record with the whole sentence in
`filing` and no statement or page, which reads like provenance and cannot be followed. The builder
records each one and renders the share cell's `Source: …` comment **from** the record, so the
artefact and the machine-readable record cannot disagree. The Bloomberg side is recorded for you from
the attached export's filename — you supply nothing extra.

**`F35` is recorded as a DERIVED figure.** Every percentage on the slide divides by it, and its
provenance is the cap table's: the builder records it with a `derived_from` ref to the cap table's
`Total Basic Shares Outstanding`, which the `captable` stage recorded against the filing it read the
share count from. That cross-stage ref resolves in the run's merged record — which is what lets
`deckcheck` walk an ownership percentage back to a capital-stock note instead of stopping at "from
the companion cap table".

Write the fragment with `ledger.write(io.stage_dir)` — your **own** stage directory, never a shared
file, because wave-mates run concurrently and a shared ledger would be a read-modify-write race
between sub-agents. `deckcheck` merges every stage's fragment into the run's record.

### Step 10 — Write the workbook

Run as `python <script.py> "<plugin_root>" "<inputs.json>" "<outputs.json>"` — the three paths
your dispatch envelope prints.

```python
import sys
from pathlib import Path

sys.path.insert(0, str(Path(sys.argv[1]) / "scripts"))   # plugin root — arg 1

from stage_io import stage_io
from ownership_workbook import build_ownership_workbook, InsiderHolding
from provenance import FigureSource, ProvenanceLedger

io = stage_io()
deal_workbook = io.inputs["deal_workbook"]   # the Ownership + Bloomberg Output tabs are in it
ledger = ProvenanceLedger(stage=io.stage_id)   # one record per holding; written at the end

SEDI = "SEDI Insider Information by Issuer report"   # the attached PDF

insiders = [
    InsiderHolding("Barrenechea, Mark James", "Mark Barrenechea (CEO & Director)", 1219092, "2025-03-31",
                   source=FigureSource(filing=SEDI, statement="Barrenechea, Mark James", page=99)),
    InsiderHolding("Fowlie, Randy", "Randy Fowlie (Director)", [193000, 0, 0, 0, 0], "2025-12-01",
                   source=FigureSource(filing=SEDI, statement="Fowlie, Randy", page=99)),
    # ... one InsiderHolding per current insider, each with the page you read it on ...
]
build_ownership_workbook(

    insiders=insiders,
    total_shares_outstanding=total,   # full units, or None
    deal_workbook=deal_workbook,
    bloomberg_export_path=BBG_PATH,   # or None / omit when no export is attached
    provenance=ledger,                # filled in place: insiders, institutions, and F35
    # bloomberg_adjusted_names={"T Rowe Price Group Inc": "T. Rowe Price"},   # optional
    # bloomberg_include_overrides={"Some Holder": 0},                         # optional
)

# The stage's provenance fragment, beside its inputs/outputs. `deckcheck` merges
# every stage's fragment into the run's <run_dir>/provenance.json.
ledger.write(io.stage_dir)
```

The builder writes the SEDI name (B), common shares (F, plain or sum formula), date (G), and adjusted
name (J) into rows 39-65, and sets `F35`. Before writing it verifies the **defined names** those
writes resolve through (`infor_own_insider_block`, `infor_own_total_shares` — shared
`template_layout` map) and raises `TemplateLayoutError` if the tab has lost them.
It does **not** touch the display blocks (`B4:G17`,
`B19:G35`), the insider include flags (H39:H65), or the `=H*F` formulas (I). With a Bloomberg export
it additionally copies the holder rows into the `Bloomberg Output` tab (values + number formats;
capped at the template's 118 rows — the Summary View is position-sorted, so any truncated tail is the
smallest holders), sets `H=0` with an audit comment on each SEDI-duplicate row, writes the adjusted
names (J68+), and neutralises the unused link rows so the Select-Institutions `LARGE` block computes.
(The shipped template is pre-cleaned of the vestigial external links / legacy defined names that would
otherwise make the openpyxl output unopenable in Excel for the render — a regression test guards that
it stays clean.)

### Step 11 — Summary

Report: output path; number of current insiders written (and how many ceased insiders were dropped);
the total basic shares used for `F35` and its source (cap table, or "left blank — fill manually");
any insider roles you could not verify; the Bloomberg side — export ingested or not, holders written,
which rows were excluded as SEDI duplicates, any Bloomberg-flagged insiders that did **not** match a
SEDI insider (kept included — analyst should confirm), any truncation past 118 holders, and any
adjusted-name overrides applied. Without an export, remind the analyst the slide's Institutions side
was left as a placeholder and can be filled by re-running with the Bloomberg attachment.

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
| `H39:H65` | template | Insider include flag (1); analyst toggles to 0 to exclude a row — **do not change** |
| `I39:I185` | template | `=H*F` helper feeding the top-12 `LARGE`/`XLOOKUP` — **do not change** |
| `B4:G17` | template | Insiders display block (top 12 by common shares) — the slide's left picture range |
| `B68:G185` | template | Bloomberg link rows — `B` = `='Bloomberg Output'!C14…`, `F`/`G` = `XLOOKUP` into Position / Filing Date — **do not change** (the builder blanks only the unused tail) |
| `H68:H185` | skill (BBG) | Institution include flag — builder writes **0 + audit comment** on SEDI duplicates and on unused rows; analyst-directed overrides via `bloomberg_include_overrides` |
| `J68:J185` | skill (BBG) | **Adjusted institution name** — suffix-stripped, or `bloomberg_adjusted_names` override |
| `B19:G35` | template | Institutions display block (top 12 + subtotal + Other Shareholders + Total) — the slide's right picture range |
| `Bloomberg Output` `C14:AC131` + info cells | skill (BBG) | The attached Bloomberg export's Summary View rows, copied verbatim (values + number formats) |

### Known limitations

- The display blocks show the **top 12** insiders / institutions. If fewer than 12 rows carry a
  positive included balance, the template's `XLOOKUP(0, …)` renders `0` / `--` in the surplus slots
  (both sides). With a realistic SEDI list and Bloomberg pull this does not occur; if it does, the
  analyst can blank the surplus display rows.
- The `Bloomberg Output` tab holds **118 holders** (`C14:C131`); a longer export is truncated to the
  118 largest (the Summary View is position-sorted) and the truncation is reported in the summary.
- **Render fidelity:** the picture render is **LibreOffice headless on every platform** (Phase D
  deleted the Excel-COM path, which only ever ran on a Windows dev box — production is Cowork/Linux
  with no Excel). LibreOffice needs a recent build (24.8+) for `XLOOKUP`; on older builds the
  insider/institution **name** columns may render blank, so check the build before a production run
  rather than reaching for Excel.
