---
name: excel-to-powerpoint
description: Reusable POC skill for moving Excel chart/table outputs into PowerPoint placeholders, initially cap table and comps artefacts for the slide-library POC.
version: 0.5.7
allowed-tools:
  - Read
  - Write
  - Bash
inputs:
  - name: workbook_path
    type: Path
    required: true
  - name: deck_path
    type: Path
    required: true
  - name: placeholder_id
    type: str
    required: true
outputs:
  - name: deck_path
    type: Path
side_effects:
  - kind: file_write
    description: Writes or updates a PowerPoint deck with Excel content placed into a placeholder.
---

# excel-to-powerpoint

Reusable Excel-to-PowerPoint insertion stage for INFOR decks.

Current implemented use:
- earnings-update cap table workbook from `captable` → overview slide `Rectangle 3` Capitalization Summary placeholder via `scripts/excel_to_powerpoint.py::insert_cap_table_into_placeholder`
- pitch slide-library cap table workbook from `captable` → Slide 7 `Rectangle 3` cap table placeholder. Wired inside `pitch_deck_assembler.assemble_pitch_deck` (same `insert_cap_table_into_placeholder` call, `slide_index=6`, range `B15:F40`) and invoked whenever the deck-assembler stage receives `captable_workbook_path` — so no separate plan stage is needed.

POC/future uses:
- comps workbook/chart → Slide 10 comps placeholder

Current boundary:
- Earnings-update cap-table insertion renders `Cap with Links!B15:F40` (capitalization summary + Financial Metrics + Valuation Metrics) as a picture via Excel COM (`Range.CopyPicture` → temporary `ChartObject` → `Chart.Export` to PNG) and pastes the PNG into the placeholder at the placeholder's exact width and height. The chart-export round-trip lets Excel stay invisible (no window flash). The picture is stretched to fit; tune the workbook's column widths and row heights so the source range's natural aspect ratio matches the placeholder (currently ~0.84 w/h for the overview slide's `Rectangle 3` Capitalization Summary placeholder).
- Renders on two backends by platform: the **Excel COM** path (Windows + Microsoft Excel; `pywin32`, Windows-only env marker in `pyproject.toml`) and the **headless LibreOffice** fallback for Cowork/Linux (`soffice`/`libreoffice` on PATH + `pypdfium2`).
- **Formula recalculation (both backends).** The cap-table workbook is authored by openpyxl, which writes formula strings with **no cached values**, and the template is manual-calc — so the rendered range would otherwise print 0/blank from Basic Market Cap down through Enterprise Value and the Financial/Valuation block. Each backend forces a recalc before the snapshot: the Windows COM path calls `excel.CalculateFull()` (running visible-but-parked-off-screen, because a recalc blanks an invisible instance's render buffer); the LibreOffice path injects a self-contained throwaway user profile (`-env:UserInstallation` + a `registrymodifications.xcu` that sets `OOXMLRecalcMode=0`, "Always recalculate") so headless LibreOffice recomputes on load. 0 = Always (not 2 = Prompt) matters — a prompt is silently skipped headless.
- **CapIQ caveat.** Cells backed by Capital IQ add-in functions — the forward consensus estimates `E47/F47` and `E48/F48` — resolve to `#NAME?` under LibreOffice (no add-in) and degrade to `n/a`/blank; they only populate when the deck is built on **Windows with Excel + the CapIQ add-in**. The LTM column (`D47/D48`, fed as hardcoded numbers by the `ltm-metrics` stage) and all pure-arithmetic cells still compute under both backends.
- Other chart/table insertions still preserve placeholders when no inserted artefact is available.
- This skill owns future chart/table placement from Excel into PowerPoint so `deck-assembler` remains focused on typed content assembly.

Future uses:
- LTM revenue breakdown pie chart
- financial summary metric charts
