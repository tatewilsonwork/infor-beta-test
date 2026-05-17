---
name: excel-to-powerpoint-infor
description: Reusable POC skill for moving Excel chart/table outputs into PowerPoint placeholders, initially cap table and comps artefacts for the slide-library POC.
version: 0.4.1
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

Reusable Excel-to-PowerPoint insertion stage for the slide-library POC.

Initial POC uses:
- cap table workbook from `captable-infor` → Slide 7 cap table placeholder
- comps workbook/chart → Slide 10 comps placeholder

Current POC boundary:
- The deck assembler preserves placeholders when no inserted artefact is available.
- This skill owns future chart/table placement from Excel into PowerPoint so `deck-assembler` remains focused on typed content assembly.

Future uses:
- LTM revenue breakdown pie chart
- financial summary metric charts
