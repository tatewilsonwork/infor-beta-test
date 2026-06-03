---
name: earningsupdate-wireframe
description: >
  Use this skill as the Phase 3 POC wireframe stage for a quarterly earnings update. It consumes
  conductor stage inputs and emits a typed SlidePlan JSON artefact for the fixed 5-slide
  earnings-update structure cloned from the shared INFOR Slide Library. Activates inside the
  conductor plan stage `wireframe`.
version: 0.5.8
allowed-tools: [Read, Write, Bash]
---

# Earnings Update Wireframe — Phase 3 POC

This is a narrow conductor-stage skill. It does not write copy, extract KPIs, parse the Bloomberg EEO snip, or touch PowerPoint.

## Conductor mode

When invoked by the conductor, the environment carries:

- `$STAGE_INPUTS` — JSON with `company`, `ticker`, `reporting_quarter`, `comparison_quarter`
- `$STAGE_OUTPUTS` — path where this stage must write its structured handoff
- `$DEAL_DIR` — deal directory root

## Workflow

1. Read `$STAGE_INPUTS`.
2. Build the fixed five-slide earnings-update `SlidePlan` using `scripts/earnings_update_wireframe.py`:
   - `earnings-update-cover`
   - `earnings-update-company-overview`
   - `earnings-update-earnings-summary`
   - `earnings-update-disclaimer`
   - `earnings-update-contact`
3. Write the plan artefact to the current stage directory as `slide_plan.json`.
4. Write `$STAGE_OUTPUTS` as:

```json
{
  "slide_plan_path": "/absolute/path/to/slide_plan.json"
}
```

## Reference command

From the plugin root with `infor-beta/scripts` on `PYTHONPATH`, the core operation is:

```python
import json, os
from pathlib import Path
from schemas import Company
from earnings_update_wireframe import build_earnings_update_slide_plan, write_slide_plan

inputs_path = Path(os.environ["STAGE_INPUTS"])
outputs_path = Path(os.environ["STAGE_OUTPUTS"])
inputs = json.loads(inputs_path.read_text())
company = Company.model_validate(inputs["company"])
plan = build_earnings_update_slide_plan(
    company=company,
    reporting_quarter=inputs["reporting_quarter"],
    comparison_quarter=inputs["comparison_quarter"],
)
slide_plan_path = write_slide_plan(plan, outputs_path.parent / "slide_plan.json")
outputs_path.write_text(json.dumps({"slide_plan_path": str(slide_plan_path)}, indent=2) + "\n")
```

## Boundary

If asked to draft bullets, select metrics, or build PowerPoint, stop. Those belong to the content and deck stages.
