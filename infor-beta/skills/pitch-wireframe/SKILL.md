---
name: pitch-wireframe
description: Builds the typed SlidePlan for the INFOR slide-library pitch deck, using the blank INFOR Slide Library order as canonical; the market-entry section expands to two targets per slide.
version: 0.5.9
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
  - name: market_entry_target_count
    type: int
    required: false
outputs:
  - name: slide_plan_path
    type: Path
side_effects:
  - kind: file_write
    description: Writes a typed SlidePlan JSON artefact to the stage output directory.
---

# pitch-wireframe

Builds the `SlidePlan` for the slide-library pitch deck. The blank library is 15
slides (including the insider-ownership slide); the market-entry section grows to
`ceil(market_entry_target_count / 2)` slides (two targets per slide), so a deck
with 8 targets has 4 market-entry slides (18 slides total).

Rules:
- Use `/templates/INFOR Slide Library.pptx` names/order as canonical.
- Do not draft copy.
- Do not touch PowerPoint.
- Infer section labels from the selected slide sequence unless analyst overrides them.
- Pass `market_entry_target_count` when the analyst has specified how many
  market-entry (acquisition) targets they want, so the plan reflects the real
  market-entry slide count and titles them `(N of M)`. When it is omitted (the
  analyst didn't ask for a specific number), the plan **defaults to 8 targets
  (4 market-entry slides)** — the standard layout — and the **deck-assembler
  still clones to the true count from the content bundle**, so the deck is
  always correct regardless.
- Write `slide_plan_path` to `$STAGE_OUTPUTS` when running under the conductor.

Implementation helper:

```python
from pitch_deck_wireframe import build_pitch_deck_slide_plan, write_slide_plan
```
