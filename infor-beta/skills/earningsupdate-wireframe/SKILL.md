---
name: earningsupdate-wireframe
description: >
  Use this skill as the Phase 3 POC wireframe stage for a quarterly earnings update. It consumes
  conductor stage inputs and emits a typed SlidePlan JSON artefact for the fixed 5-slide
  earnings-update structure cloned from the shared INFOR Slide Library. Activates inside the
  conductor plan stage `wireframe`.
allowed-tools: [Read, Write, Bash]
---

# Earnings Update Wireframe — Phase 3 POC

This is a narrow conductor-stage skill. It does not write copy, extract KPIs, parse the Bloomberg EEO snip, or touch PowerPoint.

## Conductor mode

Your dispatch envelope carries three paths — plugin root, `inputs.json`, `outputs.json` — passed **as arguments** to every command. Nothing is exported; nothing is read from the environment.

Your resolved inputs carry `company`, `ticker`, `reporting_quarter`, `comparison_quarter`.

## Workflow

1. Read your resolved inputs (`io.inputs`; also reproduced in the envelope).
2. Build the fixed five-slide earnings-update `SlidePlan` using `scripts/earnings_update_wireframe.py`:
   - `earnings-update-cover`
   - `earnings-update-company-overview`
   - `earnings-update-earnings-summary`
   - `earnings-update-disclaimer`
   - `earnings-update-contact`
3. Write the plan artefact to the stage directory (`io.stage_dir`) as `slide_plan.json`.
4. Write the structured handoff — `io.write({"slide_plan_path": ...})`, i.e. `outputs.json` holding:

```json
{
  "slide_plan_path": "/absolute/path/to/slide_plan.json"
}
```

## Reference command

Run as `python <script.py> "<plugin_root>" "<inputs.json>" "<outputs.json>"` — the three paths
your dispatch envelope prints.

```python
import sys
from pathlib import Path

sys.path.insert(0, str(Path(sys.argv[1]) / "scripts"))   # plugin root — arg 1

from stage_io import stage_io
from schemas import Company
from earnings_update_wireframe import build_earnings_update_slide_plan, write_slide_plan

io = stage_io()
company = Company.model_validate(io.inputs["company"])
plan = build_earnings_update_slide_plan(
    company=company,
    reporting_quarter=io.inputs["reporting_quarter"],
    comparison_quarter=io.inputs["comparison_quarter"],
)
slide_plan_path = write_slide_plan(plan, io.stage_dir / "slide_plan.json")
io.write({"slide_plan_path": str(slide_plan_path)})
```

## Boundary

If asked to draft bullets, select metrics, or build PowerPoint, stop. Those belong to the content and deck stages.
