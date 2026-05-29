---
name: workbook-aggregator
description: >
  Use this skill as the final consolidation stage of a conductor plan to merge every Excel (.xlsx)
  workbook produced during the deliverable into a single combined workbook. Each producing skill's
  workbook becomes a tab named after that skill (captable, ltm-revenue, comps, ...). The combined
  file is named <deliverable>-<deal name>.xlsx — e.g. earningsupdate-Project Atlas.xlsx,
  pitch-Project Atlas.xlsx. Activates as the plan stage `workbook-aggregation`. Preserves formulas,
  CapIQ links, charts, and formatting via Excel COM on Windows; falls back to a best-effort openpyxl
  merge off-Windows. The individual source workbooks are replaced by the combined file.
version: 0.5.3
allowed-tools: [Read, Write, Bash]
---

# Workbook Aggregator — Workflow

This is the last stage of a deliverable run. By the time it executes, every workbook-producing stage (cap table, LTM revenue, comps, …) has already emitted a standalone `.xlsx` under the deal's `artefacts/`, and the deck stage has already consumed whatever it needed from them. This stage merges those workbooks into one combined file and removes the individual sources, so the analyst is left with a single workbook per deal.

This skill does **not** read filings, compute figures, or touch PowerPoint. It only consolidates files that earlier stages produced.

## Conductor mode

When invoked by the conductor, the environment carries:

- `$STAGE_INPUTS` — JSON with `workbooks` (a mapping of producing-skill name → workbook path), `deliverable_type`, `deal_name`, and `output_dir`
- `$STAGE_OUTPUTS` — path where this stage must write its structured handoff
- `$DEAL_DIR` — deal directory root

If `$STAGE_INPUTS` is missing a field you need, write `{"error": "missing input: <field>"}` to `$STAGE_OUTPUTS` and stop.

## Tab naming

- A **single-sheet** source workbook becomes one tab named exactly after the producing skill (`captable`, `ltm-revenue`).
- A **multi-sheet** source contributes one tab per sheet, each named `<skill>-<sheet>` (e.g. `captable-Cap with Links`, `captable-Inputs`).
- Excel's constraints are enforced automatically: tab names are truncated to 31 chars, forbidden characters (`[]:*?/\`) are replaced, and collisions are disambiguated with a ` (2)`, ` (3)`, … suffix.

## Output naming

The combined workbook is `<deliverable>-<deal name>.xlsx`, where the deliverable prefix drops hyphens — so `earnings-update` reads as `earningsupdate` and `pitch` stays `pitch`:

- `earningsupdate-Project Atlas.xlsx`
- `pitch-Project Atlas.xlsx`

## Workflow

1. Read `$STAGE_INPUTS`. Optional / not-yet-produced workbooks may arrive as `null` or point at a non-existent path — the helper skips those, so pass them through unconditionally.
2. Call the shared helper to merge the workbooks into `$DEAL_DIR/artefacts/` (bootstrap the folder if needed). On Windows with Excel the merge uses COM and preserves formulas, CapIQ links, charts, and formatting; otherwise it uses a best-effort openpyxl copy (external CapIQ data connections and charts will **not** survive — note this in the summary if that path is taken).
3. The helper deletes the individual source workbooks once the merge succeeds (the combined file replaces them). Pass `delete_sources=False` only if the analyst explicitly wants the standalone files kept.
4. Write `$STAGE_OUTPUTS`:

```json
{
  "combined_workbook_path": "/absolute/path/to/<deliverable>-<deal name>.xlsx"
}
```

When `$STAGE_OUTPUTS` is unset (direct invocation), write the combined workbook to cwd and skip the JSON handoff.

## Reference command

```python
import json, os, sys
from pathlib import Path

plugin_root = Path(os.environ.get("CLAUDE_PLUGIN_ROOT", "./infor-beta"))
sys.path.insert(0, str(plugin_root / "scripts"))
from workbook_aggregator import combine_workbooks

inputs = json.loads(Path(os.environ["STAGE_INPUTS"]).read_text())
deal_dir = Path(os.environ.get("DEAL_DIR", "."))
out_dir = deal_dir / "artefacts"

combined_path = combine_workbooks(
    sources=inputs["workbooks"],                 # {"captable": "...", "ltm-revenue": "..."}
    output_dir=out_dir,
    deliverable_type=inputs["deliverable_type"], # e.g. "earnings-update" or "pitch"
    deal_name=inputs["deal_name"],               # the deal codename
)
Path(os.environ["STAGE_OUTPUTS"]).write_text(
    json.dumps({"combined_workbook_path": str(combined_path)}, indent=2) + "\n"
)
```

## Boundary

Do not produce, edit, or recompute any workbook content — earlier stages own that. Do not run before the deck stage: the deck-assembler reads the standalone cap-table workbook to render its snapshot, and this stage deletes that file. If `workbooks` resolves to nothing (every entry null or missing), surface the gap to the conductor via `{"error": "no workbooks to combine"}` rather than writing an empty file.
