---
name: pitch-wireframe-infor
description: Builds the typed SlidePlan for the 12-slide INFOR slide-library POC deck, using the blank INFOR Slide Library order as canonical.
version: 0.4.3
allowed-tools:
  - Read
  - Write
  - Bash
inputs:
  - name: company
    type: Company
    required: true
  - name: section_labels
    type: list[str]
    required: false
  - name: current_section
    type: str
    required: false
outputs:
  - name: slide_plan_path
    type: Path
side_effects:
  - kind: file_write
    description: Writes a typed SlidePlan JSON artefact to the stage output directory.
---

# pitch-wireframe-infor

Builds the fixed 12-slide `SlidePlan` for the Phase 3 slide-library POC.

Rules:
- Use `/templates/INFOR Slide Library.pptx` names/order as canonical.
- Do not draft copy.
- Do not touch PowerPoint.
- Infer section labels from the selected slide sequence unless analyst overrides them.
- Write `slide_plan_path` to `$STAGE_OUTPUTS` when running under the conductor.

Implementation helper:

```python
from pitch_deck_wireframe import build_pitch_deck_slide_plan, write_slide_plan
```
