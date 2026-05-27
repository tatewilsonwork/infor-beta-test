---
name: pitch-content-infor
description: Drafts the broad typed content bundle for the 12-slide INFOR slide-library POC deck from analyst notes and optional supporting sources.
version: 0.4.3
allowed-tools:
  - Read
  - Write
  - Bash
  - WebSearch
inputs:
  - name: analyst_notes
    type: str
    required: true
  - name: company
    type: Company
    required: true
  - name: presentation_date
    type: str
    required: true
  - name: slide_plan_path
    type: Path
    required: true
outputs:
  - name: content_bundle_path
    type: Path
side_effects:
  - kind: file_write
    description: Writes a typed PitchDeckContent JSON artefact to the stage output directory.
---

# pitch-content-infor

Drafts the single broad `PitchDeckContent` handoff for the Phase 3 slide-library POC.

Scope:
- Executive Summary bullets: flexible count; choose main/sub-bullets.
- Public-company overview bullets: concise description of who the company is and what it does.
- Financial Summary metric labels only; charts remain placeholders.
- Acquirer risks/mitigants: concise rows, exactly three short mitigants each.
- Comps takeaway: one sentence.

Required source:
- analyst notes

Optional sources:
- CIM / management deck
- valuation range
- public filings
- company website
- S&P Capital IQ snippets
- analyst risk notes

Do not write PowerPoint. Write `content_bundle_path` to `$STAGE_OUTPUTS`.
