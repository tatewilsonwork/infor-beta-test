# infor-beta architecture migration plan

> **This is a temporary working document. Delete it — and this `docs/` directory
> if it is otherwise empty — once Phase G ships.** Record each completed phase in
> `CHANGELOG.md` and the CLAUDE.md phase bullets as usual; this file is the
> sequencing scratchpad, not the permanent record.

Written 2026-07-27, against v0.5.34.

---

## Why

Of the ~34 shipped versions, roughly **13 are "the analyst opened the file and
something looked wrong"**: PRL1 (v0.5.6), PRL2 (v0.5.7), Test #5 / OTEX, PRL17
(v0.5.23), PRL18 (v0.5.24), the market-entry table clamp (v0.5.10), the pitch
companion-workbook fixes (v0.5.14), and four separate chart-formatting rounds
(v0.5.17 / 18 / 19 / 22).

Almost none are logic bugs or orchestration bugs. **The dominant failure mode of
this plugin is visual fidelity that nothing in the system can see until a human
opens the artefact.** The plan below is ordered around that fact.

Three structural consequences of building blind, all of which this plan removes:

- A text layout engine reimplemented in Python — `_PALATINO_CHAR_WIDTH_PER_PT`,
  `palatino_text_width_in`, `fit_overview_textbox`'s hand-calibrated em constants
  (recalibrated ~15% in v0.5.23), `_set_table_height`'s growth-aware row floors.
- A 1,123-line workbook merger (+820 lines of tests) reconciling six standalone
  workbooks after the fact, and the bug class that comes with it: LibreOffice `~`
  union operators triggering Excel repair, external-ref rebinding, order-dependent
  sheet renames, the v0.5.33 source-deletion gate.
- 441 lines of sentinel tables (`template_layout.py`) guarding dozens of
  hardcoded cell addresses and slide indices — `_KEEP_LIBRARY_INDICES` alone has
  needed three manual migrations.

## Ordering principle

**Build the oracle before the refactors.** Every phase here is a high-risk change
to a system whose failure mode is invisible. Once a machine can look at the deck,
the PRL-class bug becomes a caught test failure instead of a v0.5.35.

| Phase | Change | Depends on | Net lines |
|---|---|---|---|
| A | Clear the decks — versioning, render parity, golden fixtures | — | ~0 |
| B | Visual oracle → converge loop | A | +400 / −300 |
| C | Name-based template addressing | B | −350 |
| D | One workbook, one backend | C | **−2,000** |
| E | Conductor as code | D | −150 |
| F | Stage granularity | E | −400 |
| G | Falsification pass | B | +500 |

A gates everything. After A, **G can run in parallel with C/D/E** — different
files, no overlap. B → C → D → E → F is a hard chain.

---

## Phase A — Clear the decks

No behaviour change. One release.

1. **Drop lock-step versioning.** Single version in
   `infor-beta/.claude-plugin/plugin.json`, `.claude-plugin/marketplace.json`, and
   `pyproject.toml`; remove the `version:` frontmatter key from all 14 skills and
   the bump checklist entry from CLAUDE.md. Do this first — every later phase
   produces a large diff and you do not want 14 one-digit edits inside it.
   (`pyproject.toml` sat 12 versions behind until v0.5.34; this removes the drift
   surface entirely.)
2. **Render parity between dev and prod.** `pypdfium2` is currently pinned
   `sys_platform != 'win32'`, so Windows dev renders through PowerPoint COM and
   Cowork/Linux prod renders through LibreOffice — **a production rendering bug
   cannot currently be reproduced locally.** Make the dependency unconditional and
   flip `slide_render.render_deck_to_png` to prefer the LibreOffice path on all
   platforms.
3. **Freeze golden fixtures.** Take the last PRL run's `.pptx` and combined
   `.xlsx` (or run both shipped plans once) and commit them as reference
   artefacts. Everything after this is validated against them.

**Exit:** `pytest` green; a fixture deck renders to comparable PNGs on both
render paths. "Byte-comparable" was the original wording and is unachievable
across two rasterizers — the criterion is **same slide count, same page
dimensions, and identical findings from the geometric checks**; visual
comparison is an eyeball pass, not an assertion. Encoded as
`scripts/tests/test_render_parity.py`.

✅ **Shipped 2026-07-27 (v0.5.35),** exit criterion actually met in **v0.5.36**:
v0.5.35 wired the new `find_soffice()` into `slide_render` only and left five
other `shutil.which("soffice")` sites for Phase D, which kept `pytest` red (the
cap-table renderer's COM→LibreOffice fallback could not find a Windows MSI
install) and silently degraded the aggregator recalc and the chart renders on
dev. All five now route through the locator, with a drift lock against a
reappearing bare probe.

See the CHANGELOG entry for the five
findings recorded and deliberately left for later phases — in particular that
LibreOffice wraps Palatino ~one line taller than PowerPoint even with the same
font file, making LibreOffice the *conservative* oracle for Phase B, and that
LibreOffice has no `palatinolinotype` substitution entry (prod font resolution
still needs one `fc-match` on a Cowork shell).

> **Finding #2 reclassified (2026-07-27, during Phase B step 1).** Phase A
> recorded "the overview cap-table *picture* renders `#VALUE!` and is invisible to
> every string scan" as a defect, and used it to argue that Phase B's vision
> checks are load-bearing rather than optional. **The `#VALUE!` is expected, not a
> defect.** The cap table's forward-estimate columns are CapIQ UDF calls
> (`SP_REV_EST` / `SP_EBITDA_EST` at `E47:F48`) wrapped by rows 34/35 in
> `IFERROR(..., "n/a ")`; the EV/metric rows divide `$F$31` by that text, so an
> un-refreshed CapIQ estimate propagates `#VALUE!` into `E39:F40` by design — and
> CapIQ cannot be refreshed in this environment. Same for comps/precedents (array
> formulas ship un-evaluated) and the financial-summary LTM link (`#N/A` until the
> combined workbook exists).
>
> Consequences, all carried into `deck_contract.py`: the error scan is scoped to
> text shapes and table cells and never reaches into rasterised range pictures; a
> workbook-side scan would also not help, because **zero of the 121 formula cells**
> on the fixture's `captable` tab carry a cached value under `data_only=True`
> (`=TODAY()-1` included), so it would see `None` everywhere and look green.
> The general claim — a string scan cannot see inside a raster — still holds and
> is still why the vision tier is agent-inspected; `#VALUE!` is simply the wrong
> example of it. Detail in `scripts/tests/fixtures/README.md`.

## Phase B — Visual oracle, then converge loop

Three steps. Do not put the checker inside the assembler until it is trusted.

1. **`scripts/deck_contract.py`** — a declarative per-slide contract (expected
   shape inventory, box bounds, expected table height, forbidden strings
   `[Placeholder]` / `#N/A` / `#REF!` / `{{`, expected currency token) plus
   `verify_deck(path) -> list[Finding]`.
   - **Deterministic checks** — python-pptx shape extents vs. box bounds, table
     height vs. the library height, string scan. These are **blocking**.
   - **Vision checks** over the rendered PNGs — overlap, collision, unreadable
     contrast, chart-label pileup. These are **surfaced at the `deck` checkpoint,
     not blocking**, to keep non-determinism out of the gate.
2. **Wire as a regression test** against the Phase A fixtures. Confirm it
   independently catches three known-historical bugs by replaying their artefacts:
   the PRL17 5.91" market-entry table, the PRL18 5.36" risk table, and the v0.5.23
   overview-bullet overflow into the LTM section. **If it does not catch all
   three, the contract is wrong — fix it before step 3.**

✅ **Steps 1–2 shipped 2026-07-27 (v0.5.37).** All three bugs caught, each from a
real artefact (see `scripts/tests/fixtures/regressions/README.md` — the pre-fix
PRL17 overview deck was overwritten by a post-fix re-run, so the overview
overflow replays from PRL14, a genuine pre-fix artefact of the identical defect
one release earlier). Two design decisions worth carrying into step 3:

- **The blank library is the geometric baseline, not zero.** The library's own
  footnote placeholders render ~0.07" past their boxes and its tombstone slides
  park shapes fully off-canvas. Measured on the pitch fixture: against zero, 64
  blocking findings, 52 of them template-inherent geometry (36
  `shape-outside-slide`, 16 `rendered-overflow`); the baseline removes all 52 and
  leaves the 12 genuine string findings. Every geometric check now measures the
  deck, measures the matching library slide, and reports the excess — which also
  makes the reference self-maintaining in the Phase C sense. Built slides are
  matched to library slides by **shape-name signature**, since titles are filled
  and the pitch flow deletes and clones.
- **The two table checks are not redundant.** PRL18 declares 5.360" against a
  5.1715" library height and is caught in the XML; PRL17 declares 5.710" — *under*
  the library's 5.7197" — and renders 5.91", so it is only visible on the raster.
  A contract with either check alone catches one of the three bugs, not two.

Step 3 has **not** started (it needs analyst sign-off), so no assembler was
touched and no estimation code was deleted.
3. **Promote into both assemblers** as write → verify → repair → re-verify,
   bounded to ~3 iterations, stage fails if it cannot converge.

**Then delete** — only once step 3 is trusted:

- `_PALATINO_CHAR_WIDTH_PER_PT` and `palatino_text_width_in` (`pptx_helpers.py:81`)
- the `fit_overview_textbox` em constants (`pptx_helpers.py:458`)
- `_set_table_height`'s `min_heights` estimation (`pitch_deck_assembler.py:151`)
- `_fill_risk_table`'s font-stepping ladder

Each becomes measure-then-adjust against the real render.

**Exit:** all three historical bugs caught by CI; the estimation code is gone.

## Phase C — Name-based template addressing

Safe now — Phase B catches breakage.

1. **Named ranges into the four workbook templates** (one manual Excel pass):
   `fx_rate`, `share_price`, `ltm_revenue_valuation`, `ltm_ebitda_valuation`,
   `cap_picture_range`, `basic_shares`, `comps_output_ccy`,
   `precedents_input_ccy`, plus the block anchors. Named ranges are metadata and
   survive a re-save — that is the entire point.
2. **Writers resolve through names.** Keep the `template_layout.py` sentinels for
   one release as a *cross-check* (assert the name and the sentinel agree), then
   delete the sentinel tables.
3. **PowerPoint side:** generalize the existing `Rectangle 17` FS-slide
   self-discovery into `find_slide_by_marker`, replacing `prs.slides[6]`,
   `OVERVIEW_SLIDE_INDEX`, and `_KEEP_LIBRARY_INDICES`.

**Payoff worth naming:** after C, an analyst can re-save a template or insert a
library slide without a code migration.

## Phase D — One workbook, one backend

The large one. Single workstream, because deleting the merge deletes most of the
COM surface with it. (COM appears in exactly four files today: `financial_charts`
12 hits, `excel_to_powerpoint` 7, `slide_render` 6, `workbook_aggregator` 6.)

1. **`scripts/deal_workbook.py`** — the deal owns `pitch-<codename>.xlsx` from
   stage one. Single serialized `write_tab(path, tab_name, TabSpec)`. Templates
   are copied in as tabs at init rather than produced as standalone files.
2. **Convert producers one at a time, cheapest first:** `financial-summary` and
   `ltm-metrics` (anchor-free by design) → `comps` / `precedents` (clean openpyxl
   round-trip) → `captable` → `ownership` last (the cruft-stripped template is the
   hardest).
3. **Delete the aggregator:** `workbook_aggregator.py`, its 820-line test, the
   `workbook-aggregation` stage in both plans, and the hardcoded barrier in
   `plan_schedule.py:105`. Cross-tab links now resolve immediately, so
   `financial-charts` loses its post-aggregation constraint and becomes an
   ordinary data-edge stage — or folds into `deck`.
4. **Then delete COM** from `financial_charts.py` and `excel_to_powerpoint.py`.
   Order matters: step 3 removes `_combine_via_com` for free.

**Permanently killed by this phase:** the LibreOffice `~` union-operator Excel
repair bug, external-ref rebinding, order-dependent sheet renames, the v0.5.33
deletion gate, comment/hyperlink carry-over, and the
`=INDEX('ltm-metrics'!…)`-resolves-only-in-the-combined-workbook constraint.

> **Open decision — resolve before starting D.** Does anything actually depend on
> Excel COM preserving **live CapIQ links** through the merge? Per CLAUDE.md the
> comps and precedents workbooks ship with their CapIQ formulas un-evaluated and
> the analyst refreshes in Excel anyway, which would make live-link preservation
> already moot and D a clean deletion. If some artefact does depend on it, D needs
> a narrow COM-only export step instead. Does not block A–C.

## Phase E — Conductor as code

Much smaller now — no barrier, fewer stages.

1. Extend `conductor_cli` into `conductor.py` with `run_wave(run_dir, n)`. The
   model's loop becomes: call it, issue the returned `Task` calls, collect,
   surface the checkpoint. (`load_plan` / `prep_wave` / `collect_wave` already
   exist — this finishes the trip they started.)
2. **Remove the env-var export block** from
   `skills/conductor/references/stage-envelope.md`. Render resolved inputs inline,
   or have SKILL.md commands take `sys.argv[1]` instead of `os.environ`. This
   removes the "did the export survive across the sub-agent's tool calls" failure
   mode entirely.
3. `skills/conductor/SKILL.md`: 216 lines → ~80 (intake, dispatch, checkpoints,
   summary).

## Phase F — Stage granularity

Requires E — "call it in-process" needs the conductor to be a process.

Reclassify every stage as **transform** (call the function directly, no subagent)
or **judgment** (subagent with a real tool allow-list):

- Transforms: `pitch-wireframe`, `earningsupdate-wireframe`, `deck-assembler`,
  `financial-charts`.
- Judgment: `captable`, `comps`, `precedents`, `ownership`, `financial-summary`,
  `ltm-metrics`, `pitch-content`, `earningsupdate-content`.

Pitch drops from 11 dispatches to ~7 and the wave graph shortens. Keep a SKILL.md
for a transform only if it stays separately invocable (`/captable`-style).

## Phase G — Falsification

Additive; may start any time after B, in parallel with C/D/E.

1. Promote the v0.5.31 / v0.5.34 citations from a cell-comment convention to a
   structured `provenance.json` per run (figure → filing → statement → page), with
   the cell comment *rendered from* it rather than being the record.
2. A `deckcheck` stage after `deck` that reads the rendered PNGs, the provenance
   records, and the source filings, and attempts to **disprove** every figure on
   the deck.

---

## Not changing

Deliberately kept, and would be chosen again:

- The typed I/O contract and plan YAML.
- The deal directory and run-log layout.
- Excel-does-the-math (formulas, not Python scalars).
- Progressive disclosure in SKILL.md.
- A slide library of real `.pptx` entries rather than slides generated from
  scratch — essential for brand fidelity.
- The `required` checkpoint on `deck` before the final artefacts.
- The prompt-injection clause in the stage envelope.
