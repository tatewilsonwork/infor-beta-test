---
name: workbook-aggregator
description: >
  Use this skill as the workbook-consolidation stage of a conductor plan to merge every Excel (.xlsx)
  workbook produced during the deliverable into a single combined workbook. Each producing skill's
  workbook becomes a tab named after that skill (captable, ltm-metrics, comps, ...). The combined
  file is named <deliverable>-<deal name>.xlsx — e.g. earningsupdate-Project Atlas.xlsx,
  pitch-Project Atlas.xlsx. Activates as the plan stage `workbook-aggregation`. Preserves formulas,
  CapIQ links, charts, and formatting via Excel COM on Windows; falls back to a best-effort openpyxl
  merge off-Windows. The individual source workbooks are replaced by the combined file once the
  merge is verified; a degraded or partially-linked merge preserves them for a retry.
version: 0.5.33
allowed-tools: [Read, Write, Bash]
---

# Workbook Aggregator — Workflow

This stage consolidates the run's workbooks. By the time it executes, every workbook-producing stage (cap table, LTM metrics, comps, …) has already emitted a standalone `.xlsx` under the deal's `artefacts/`, and the deck stage has already consumed whatever it needed from them. This stage merges those workbooks into one combined file and removes the individual sources — removal is gated on verified success (see **Source-deletion gate** below) — so on a clean run the analyst is left with a single workbook per deal. It is the last stage of the earnings-update plan; in the pitch plan the post-aggregation `financial-charts` stage runs after it, drawing the deck's charts on the combined workbook (the only place the financial-summary tab's LTM links resolve).

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
- Comps `F3` and precedents `C2` (output currency) → `='captable'!F5`, so the combined workbook shows one consistent currency and updating `F5` flows through (restyled to Palatino 9 blue to match).
- Financial-summary LTM link cells — each `=INDEX('ltm-metrics'!…)` formula is re-assigned so it re-binds to the sibling `ltm-metrics` tab (the COM sheet-copy leaves it bound to an *external* workbook relationship that Excel's `.Formula` getter masks as internal, resolving `#N/A`; see v0.5.16).

The relink is best-effort and only fires when the relevant tabs are present; it is a no-op otherwise. On the openpyxl fallback the formulas are still written, but CapIQ links are already lost on that path. A relink failure never loses the merged workbook — but it is reported (stderr + `CombineResult.relink_ok`) and blocks source deletion, because a combined workbook whose links are broken plus deleted sources would be unrecoverable.

**Layout pre-flight (v0.5.32).** Because the relink writes those fixed addresses blind, `combine_workbooks` first verifies each source's sentinel anchors from the shared `template_layout` map (cap table `B47`/`B48` `Revenue`/`Adj. EBITDA` labels + the `D33` `LTM` header for `D47`/`D48`, `B17` for `F17`, `B5` for `F5`, `B7` for `F7`; ownership `B35`; comps `E3`; precedents `B2`) — but only for the relink pairs actually present. The check runs before either merge backend, on the still-on-disk sources, so a re-saved template with shifted rows raises `TemplateLayoutError` with nothing merged and no source deleted, instead of relinking the wrong cells.

## LibreOffice recalc (openpyxl path only)

The openpyxl copy writes formula **strings** with no cached values, so the cross-tab links above — and the `financial-summary` tab's `=INDEX('ltm-metrics'!…)` LTM lookups — sit **un-evaluated** in the merged file until something recalculates them. A downstream stage that reads those cells programmatically (e.g. `financial-charts`) would otherwise get `None`. So after the openpyxl merge the helper re-saves the combined workbook through **headless LibreOffice** (recalc-on-load), caching the evaluated values **while preserving the formulas** (Excel does the math; the analyst still sees live formulas). This is automatic inside `combine_workbooks` — no extra step in this skill. It is best-effort: when `soffice`/`libreoffice` is absent the workbook simply keeps its un-evaluated formulas (the analyst's Excel recalcs on open) rather than the stage failing. The COM path needs no recalc — Excel evaluates natively on save.

One LibreOffice export artefact is cleaned up right after the recalc: LibreOffice rewrites a parenthesized multi-range union argument with its own **`~` union operator** (the comps / precedents quartile rows come back as `PERCENTILE.INC((L10:L15~L20:L25~L30:L35),0.25)`), which Excel cannot parse — the combined workbook then opens with a **"Removed Records: Formula"** repair that strips those formulas. `_strip_lo_union_operators` rewrites `~` back to `,` inside every sheet formula at the XML level (string literals excluded — `~` is Excel's wildcard escape there), leaving the recalc's cached values untouched. Automatic; no extra step in this skill. The whole recalc step is genuinely never-fatal: any failure inside it — including a bad LibreOffice export that breaks the `~` fix — rolls back to the pre-recalc merged file with its formulas left un-evaluated.

## Source-deletion gate

Deleting the standalone sources is the one irreversible act in this stage, so `combine_workbooks` deletes them (default `delete_sources=True`) only when **all** of the following verified:

1. **The merge ran on the platform's intended backend.** Off-Windows, openpyxl *is* the intended backend — deletion stays allowed. On Windows, if Excel COM failed and the lossy openpyxl fallback ran, the merge is **degraded** (live CapIQ connections, charts, and full formatting did not survive) and the full-fidelity sources are the only way to retry — they are preserved.
2. **The cross-tab relink pass succeeded** (both backends report it; a failure is traced to stderr, never raised — the merged workbook is still saved).
3. **The combined workbook carries no external-workbook references** — no `xl/externalLinks` part and no `[n]`-indexed sheet formula. A failed relink can leave the financial-summary LTM links bound to an *external* `'[1]ltm-metrics'` workbook reference; deleting the sources would turn that into a permanent `#REF!`/`#N/A`.

When any check fails, the sources are **preserved**, each problem is printed to stderr, and the same warnings come back on the returned `CombineResult` (`warnings`, plus `backend` / `degraded` / `relink_ok` / `external_refs` / `sources_deleted` / `kept_sources`) so this stage can surface them in its outputs. `delete_sources=False` keeps the sources unconditionally, exactly as before.

## Workflow

1. Read `$STAGE_INPUTS`. Optional / not-yet-produced workbooks may arrive as `null` or point at a non-existent path — the helper skips those, so pass them through unconditionally.
2. Call the shared helper to merge the workbooks into `$DEAL_DIR/artefacts/` (bootstrap the folder if needed). On Windows with Excel the merge uses COM and preserves formulas, CapIQ links, charts, and formatting; otherwise it uses a best-effort openpyxl copy (external CapIQ data connections and charts will **not** survive — note this in the summary if that path is taken).
3. The helper deletes the individual source workbooks only once the merge is **verified** (see the Source-deletion gate above); otherwise it preserves them and returns the warnings on the `CombineResult`. Pass `delete_sources=False` only if the analyst explicitly wants the standalone files kept.
4. Write `$STAGE_OUTPUTS`, carrying the verification facts alongside the declared path (extra keys beyond the plan's declared outputs are allowed):

```json
{
  "combined_workbook_path": "/absolute/path/to/<deliverable>-<deal name>.xlsx",
  "sources_deleted": true,
  "warnings": []
}
```

If `warnings` is non-empty (degraded merge, failed relink, or leftover external references), say so plainly in the stage summary: the combined workbook was produced but the standalone sources were **preserved** so the merge can be retried.

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

result = combine_workbooks(
    sources=inputs["workbooks"],                 # {"captable": "...", "ltm-metrics": "..."}
    output_dir=out_dir,
    deliverable_type=inputs["deliverable_type"], # e.g. "earnings-update" or "pitch"
    deal_name=inputs["deal_name"],               # the deal codename
)
Path(os.environ["STAGE_OUTPUTS"]).write_text(
    json.dumps(
        {
            "combined_workbook_path": str(result.output_path),
            "sources_deleted": result.sources_deleted,
            "warnings": list(result.warnings),
        },
        indent=2,
    )
    + "\n"
)
```

## Boundary

Do not produce, edit, or recompute any workbook content — earlier stages own that. Do not run before the deck stage: the deck-assembler reads the standalone cap-table workbook to render its snapshot, and this stage deletes that file. If `workbooks` resolves to nothing (every entry null or missing), surface the gap to the conductor via `{"error": "no workbooks to combine"}` rather than writing an empty file.
