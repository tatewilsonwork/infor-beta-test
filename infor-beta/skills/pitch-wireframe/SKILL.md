---
name: pitch-wireframe
description: Builds the typed SlidePlan for the INFOR slide-library pitch deck, using the blank INFOR Slide Library order as canonical; the market-entry section expands to two targets per slide, the Financial Summary section to four metrics per slide, and the Key Investment Highlights slide can be omitted.
version: 0.5.34
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
  - name: financial_metric_count
    type: int
    required: false
  - name: include_investment_highlights
    type: bool
    required: false
outputs:
  - name: slide_plan_path
    type: Path
side_effects:
  - kind: file_write
    description: Writes a typed SlidePlan JSON artefact to the stage output directory.
---

# pitch-wireframe

Builds the `SlidePlan` for the slide-library pitch deck. The blank library is 16
slides (including the insider-ownership and precedent-transactions slides — the
canonical order lives in `scripts/slide_library_registry.py`); three deck-spec
inputs adjust the slide mix:

- the market-entry section grows to `ceil(market_entry_target_count / 2)` slides
  (two targets per slide), so a deck with 8 targets has 4 market-entry slides
  (19 slides total);
- the Financial Summary section grows to `financial_metric_count / 4` slides
  (four metric tiles per slide; 4 = one slide, 8 = two slides titled
  `(1 of 2)` / `(2 of 2)`);
- `include_investment_highlights: false` drops the Key Investment Highlights
  entry from the plan entirely (the deck-assembler then deletes the slide).

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
- Pass `financial_metric_count` (a positive multiple of 4; the deck spec offers
  4 or 8) when supplied; omitted → **one Financial Summary slide (4 metrics)**.
  Unlike market-entry, the assembler follows the PLAN's Financial Summary count,
  so this input is authoritative.
- Pass `include_investment_highlights` when supplied; omitted → the slide is
  **included**. The assembler follows the plan here too — no
  `key-investment-highlights` entry means no slide in the deck.
- Write `slide_plan_path` to `$STAGE_OUTPUTS` when running under the conductor.

Implementation helper:

```python
from pitch_deck_wireframe import build_pitch_deck_slide_plan, write_slide_plan
```
