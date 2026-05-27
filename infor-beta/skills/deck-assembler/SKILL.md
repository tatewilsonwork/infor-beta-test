---
name: deck-assembler
description: >
  Use this skill as the deck assembly stage. It consumes a typed SlidePlan and typed content bundle
  and writes either the decomposed earnings-update POC deck or the 12-slide INFOR slide-library POC deck.
version: 0.4.3
allowed-tools: [Read, Write, Bash]
---

# Deck Assembler

This stage assembles typed slide/content handoffs into PowerPoint decks.

## Supported POC paths

1. **Earnings update POC**
   - `SlidePlan.deliverable_type = "earnings-update"`
   - content bundle schema: `EarningsUpdateContent`
   - template: `INFOR Earnings Update Template.pptx`
   - helper: `scripts/earnings_update_assembler.py`

2. **Slide-library pitch POC**
   - `SlidePlan.deliverable_type = "pitch"`
   - content bundle schema: `PitchDeckContent`
   - template: `INFOR Slide Library.pptx`
   - helper: `scripts/pitch_deck_assembler.py`

## Conductor mode

When invoked by the conductor, read:

- `$STAGE_INPUTS` — JSON with `slide_plan_path`, `content_bundle_path`, `template_name`, and `output_dir`; may also include `captable_workbook_path` and `comps_workbook_path`
- `$STAGE_OUTPUTS` — path where this stage must write its structured handoff
- `$DEAL_DIR` — deal directory root

## Workflow

1. Read `$STAGE_INPUTS`.
2. Resolve `template_name` under `${CLAUDE_PLUGIN_ROOT:-./infor-beta}/templates/`.
3. Inspect the `SlidePlan.deliverable_type`.
4. If `earnings-update`, call `assemble_earnings_update_deck(...)` and pass `captable_workbook_path` when supplied so slide 2's cap-table placeholder is replaced.
5. If `pitch`, call `assemble_pitch_deck(...)`.
6. Write the deck under `$DEAL_DIR/artefacts/` when `$DEAL_DIR` is set; otherwise use the supplied `output_dir`.
7. Write `$STAGE_OUTPUTS` as:

```json
{
  "deck_path": "/absolute/path/to/Pitch Deck - Company.pptx"
}
```

## Pitch POC boundaries

- Slide 1: client name/date only.
- Slide 2: executive summary bullets.
- Slides 3–5: static, do not touch.
- Slide 6: section divider labels.
- Slide 7: company overview bullets; cap table placeholder stays unless `excel-to-powerpoint-infor` replaces it later.
- Slide 8: metric labels only; charts stay placeholders.
- Slide 9: concise risks/mitigants + tagline.
- Slide 10: comps takeaway; chart placeholder stays unless insertion replaces it later.
- Slides 11–12: static, do not touch.

## Reference command

```python
import json, os, sys
from pathlib import Path

plugin_root = Path(os.environ.get("CLAUDE_PLUGIN_ROOT", "./infor-beta"))
sys.path.insert(0, str(plugin_root / "scripts"))

from schemas import SlidePlan
from earnings_update_assembler import assemble_earnings_update_deck
from pitch_deck_assembler import assemble_pitch_deck

inputs = json.loads(Path(os.environ["STAGE_INPUTS"]).read_text())
slide_plan = SlidePlan.model_validate_json(Path(inputs["slide_plan_path"]).read_text())
template_path = plugin_root / "templates" / inputs["template_name"]
output_dir = Path(os.environ.get("DEAL_DIR", inputs["output_dir"])) / "artefacts"

if slide_plan.deliverable_type == "earnings-update":
    deck_path = assemble_earnings_update_deck(
        slide_plan_path=inputs["slide_plan_path"],
        content_path=inputs["content_bundle_path"],
        template_path=template_path,
        output_dir=output_dir,
        captable_workbook_path=inputs.get("captable_workbook_path"),
    )
elif slide_plan.deliverable_type == "pitch":
    deck_path = assemble_pitch_deck(
        slide_plan_path=inputs["slide_plan_path"],
        content_path=inputs["content_bundle_path"],
        template_path=template_path,
        output_dir=output_dir,
        captable_workbook_path=inputs.get("captable_workbook_path"),
        comps_workbook_path=inputs.get("comps_workbook_path"),
    )
else:
    raise ValueError(f"unsupported deliverable_type: {slide_plan.deliverable_type}")

Path(os.environ["STAGE_OUTPUTS"]).write_text(json.dumps({"deck_path": str(deck_path)}, indent=2) + "\n")
```
