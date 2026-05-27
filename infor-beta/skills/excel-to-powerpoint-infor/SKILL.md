---
name: excel-to-powerpoint-infor
description: Reusable POC skill for moving Excel chart/table outputs into PowerPoint placeholders, initially cap table and comps artefacts for the slide-library POC.
version: 0.4.3
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

# excel-to-powerpoint-infor

Reusable Excel-to-PowerPoint insertion stage for INFOR decks.

Current implemented use:
- earnings-update cap table workbook from `captable-infor` → slide 2 `Rectangle 4` cap-table placeholder via `scripts/excel_to_powerpoint.py::insert_cap_table_into_placeholder`

POC/future uses:
- pitch slide-library cap table workbook from `captable-infor` → Slide 7 cap table placeholder
- comps workbook/chart → Slide 10 comps placeholder

Current boundary:
- Earnings-update cap-table insertion renders `Cap with Links!B13:F31` as a picture via Excel COM (`Range.CopyPicture` → temporary `ChartObject` → `Chart.Export` to PNG) and pastes the PNG into the placeholder at the placeholder's exact width and height. The chart-export round-trip lets Excel stay invisible (no window flash). The picture is stretched to fit; tune the workbook's column widths and row heights so the source range's natural aspect ratio matches the placeholder (currently ~0.84 w/h for slide 2 `Rectangle 4`).
- Requires Windows + Microsoft Excel installed; `pywin32` is a runtime dependency (Windows-only env marker in `pyproject.toml`).
- Other chart/table insertions still preserve placeholders when no inserted artefact is available.
- This skill owns future chart/table placement from Excel into PowerPoint so `deck-assembler` remains focused on typed content assembly.

Future uses:
- LTM revenue breakdown pie chart
- financial summary metric charts
