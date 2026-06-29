---
name: workbook-aggregator
description: >
  Use this skill as the final consolidation stage of a conductor plan to merge every Excel (.xlsx)
  workbook produced during the deliverable into a single combined workbook. Each producing skill's
  workbook becomes a tab named after that skill (captable, ltm-metrics, comps, ...). The combined
  file is named <deliverable>-<deal name>.xlsx — e.g. earningsupdate-Project Atlas.xlsx,
  pitch-Project Atlas.xlsx. Activates as the plan stage `workbook-aggregation`. Preserves formulas,
  CapIQ links, charts, and formatting via Excel COM on Windows; falls back to a best-effort openpyxl
  merge off-Windows. The individual source workbooks are replaced by the combined file.
version: 0.5.17
allowed-tools: [Read, Write, Bash]
---

# Workbook Aggregator — Workflow

This is the last stage of a deliverable run. By the time it executes, every workbook-producing stage (cap table, LTM metrics, comps, …) has already emitted a standalone `.xlsx` under the deal's `artefacts/`, and the deck stage has already consumed whatever it needed from them. This stage merges those workbooks into one combined file and removes the individual sources, so the analyst is left with a single workbook per deal.

This skill does **not** read filings, compute figures, or touch PowerPoint. It only consolidates files that earlier stages produced.

## Conductor mode

When invoked by the conductor, the environment carries:

- `$STAGE_INPUTS` — JSON with `workbooks` (a mapping of producing-skill name → workbook path), `deliverable_type`, `deal_name`, and `output_dir`
- `$STAGE_OUTPUTS` — path where this stage must write its structured handoff
- `$DEAL_DIR` — deal directory root

If `$STAGE_INPUTS` is missing a field you need, write `{"error": "missing input: <field>"}` to `$STAGE_OUTPUTS` and stop.

## Tab naming

- A **single-sheet** source workbook becomes one tab named exactly after the producing skill (`captable`, `ltm-metrics`).
- A **multi-sheet** source keeps its **original sheet names**, unprefixed — e.g. the `ownership` workbook contributes `Ownership` and `Bloomberg Output` (not `ownership-Ownership` / `ownership-Bloomberg Output`). The sheet names are self-describing, and prefixing them would force a rename that breaks the source's intra-workbook cross-sheet references (the ownership `Ownership` sheet's hundreds of `='Bloomberg Output'!…` lookups → `#REF`).
- Excel's constraints are enforced automatically: tab names are truncated to 31 chars, forbidden characters (`[]:*?/\`) are replaced, and collisions are disambiguated with a ` (2)`, ` (3)`, … suffix.

## Output naming

The combined workbook is `<deliverable>-<deal name>.xlsx`, where the deliverable prefix drops hyphens — so `earnings-update` reads as `earningsupdate` and `pitch` stays `pitch`:

- `earningsupdate-Project Atlas.xlsx`
- `pitch-Project Atlas.xlsx`

## Merge base + cross-tab links

When a `captable` workbook is among the sources, the COM backend opens **it** as the base workbook (and saves it as the combined file), then copies the other skills' sheets in after it. This preserves the cap table's theme, CapIQ links, and formatting — the previous blank-base merge shifted colours and could drop CapIQ links, which made the combined file hard to format and link. (Without a cap table, a blank base is used, as before.)

Each source's content sheets are copied **as a group in one operation** so a source's intra-workbook cross-sheet references survive as internal references (copying sheet-by-sheet turns them into external links to the soon-deleted source → `#REF`; this is the ownership `Ownership` → `Bloomberg Output` case), and the copy destination is passed **positionally** — the named `After=` argument is silently dropped by some Excel builds, which then copy into a stray new workbook (a no-op append that, before, forced the openpyxl fallback and lost the theme).

## Brand theme

The combined workbook is stamped with the INFOR brand theme (`templates/INFORFG.thmx`) — via `ApplyTheme` on the COM backend, `loaded_theme` on the openpyxl backend — so it carries the INFOR colour scheme / fonts even when the merge base is a blank workbook (no cap table) or the openpyxl fallback runs. It is best-effort: a missing theme file or a theme-apply error never loses the merged workbook.

Once every workbook is one file, a **relink** pass rewrites the skills' standalone scalar handoffs into live cross-tab formulas, so the analyst's combined workbook stays internally linked:

- Cap table `D47` / `D48` (LTM Revenue / Adj. EBITDA) → `='ltm-metrics'!B<row>*F7`, where `<row>` is found by the `(=) LTM Revenue` / `(=) LTM Adj. EBITDA` bridge-total label (the bridge rows are dynamic). Units already match (millions); `*F7` keeps the existing FX conversion.
- Ownership `F35` (% denominator) → `='captable'!F17*1000000` (the cap table is in millions; ownership share counts are full units).

The relink is best-effort and only fires when the relevant tabs are present; it is a no-op otherwise. On the openpyxl fallback the formulas are still written, but CapIQ links are already lost on that path.

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
    sources=inputs["workbooks"],                 # {"captable": "...", "ltm-metrics": "..."}
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
