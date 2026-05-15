---
name: deck-assembler
description: >
  Use this skill as the Phase 3 POC deck assembly stage. It consumes a typed SlidePlan and typed
  EarningsUpdateContent bundle and writes an INFOR Earnings Update .pptx from the existing
  INFOR Earnings Update Template. POC scope only; not the generalized slide-library assembler yet.
version: 0.3.1
allowed-tools: [Read, Write, Bash]
---

# Deck Assembler — Earnings Update POC

This is the minimum deck assembler needed to prove the decomposed earnings-update conductor path. It is intentionally template-specific.

## Conductor mode

When invoked by the conductor, read:

- `$STAGE_INPUTS` — JSON with `slide_plan_path`, `content_bundle_path`, `template_name`, and `output_dir`
- `$STAGE_OUTPUTS` — path where this stage must write its structured handoff
- `$DEAL_DIR` — deal directory root

## Workflow

1. Read `$STAGE_INPUTS`.
2. Resolve `template_name` under `${CLAUDE_PLUGIN_ROOT:-./infor-beta}/templates/`.
3. Use `scripts/earnings_update_assembler.py` to assemble the deck from:
   - the typed `SlidePlan`
   - the typed `EarningsUpdateContent`
   - `INFOR Earnings Update Template.pptx`
4. Write the deck under `$DEAL_DIR/artefacts/` when `$DEAL_DIR` is set; otherwise use the supplied `output_dir`.
5. Verify that slides 1–3 no longer contain template placeholders and that slide 2's `[Macabacus Placeholder]` remains intact for manual cap-table paste.
6. Write `$STAGE_OUTPUTS` as:

```json
{
  "deck_path": "/absolute/path/to/Earnings Update - Company.pptx"
}
```

## Reference command

```python
import json, os
from pathlib import Path
from earnings_update_assembler import assemble_earnings_update_deck

inputs = json.loads(Path(os.environ["STAGE_INPUTS"]).read_text())
plugin_root = Path(os.environ.get("CLAUDE_PLUGIN_ROOT", "./infor-beta"))
template_path = plugin_root / "templates" / inputs.get("template_name", "INFOR Earnings Update Template.pptx")
output_dir = Path(os.environ.get("DEAL_DIR", inputs["output_dir"])) / "artefacts"
deck_path = assemble_earnings_update_deck(
    slide_plan_path=inputs["slide_plan_path"],
    content_path=inputs["content_bundle_path"],
    template_path=template_path,
    output_dir=output_dir,
)
Path(os.environ["STAGE_OUTPUTS"]).write_text(json.dumps({"deck_path": str(deck_path)}, indent=2) + "\n")
```

## Boundary

Do not generalize this into the full slide-library assembler during the POC. No registry, no CIM/pitch/teaser slide entries, no variant logic.
