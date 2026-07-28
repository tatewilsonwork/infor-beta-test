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

- ~~A text layout engine reimplemented in Python — `_PALATINO_CHAR_WIDTH_PER_PT`,
  `palatino_text_width_in`, `fit_overview_textbox`'s hand-calibrated em constants
  (recalibrated ~15% in v0.5.23), `_set_table_height`'s growth-aware row floors.~~
  **Gone in v0.5.39** — 224 lines, replaced by measure-then-adjust against a real
  render (Phase B step 3).
- A 1,123-line workbook merger (+820 lines of tests) reconciling six standalone
  workbooks after the fact, and the bug class that comes with it: LibreOffice `~`
  union operators triggering Excel repair, external-ref rebinding, order-dependent
  sheet renames, the v0.5.33 source-deletion gate.
- 441 lines of sentinel tables (`template_layout.py`) guarding dozens of
  hardcoded cell addresses and slide indices — `_KEEP_LIBRARY_INDICES` alone has
  needed three manual migrations. *(Addressed in Phase C: the addresses and
  indices are gone, replaced by defined names and marker shapes. The sentinel
  tables remain for one release as the cross-check that proves the migration was
  faithful, then get deleted.)*

## Ordering principle

**Build the oracle before the refactors.** Every phase here is a high-risk change
to a system whose failure mode is invisible. Once a machine can look at the deck,
the PRL-class bug becomes a caught test failure instead of a v0.5.35.

| Phase | Change | Depends on | Net lines |
|---|---|---|---|
| A | Clear the decks — versioning, render parity, golden fixtures | — | ~0 |
| B | ✅ Visual oracle → converge loop | A | +1,667 / −389 |
| I | Windows COM dev-path hygiene (runs between B1 and B2) | B1 | ~+60 |
| C | Name-based template addressing | B | −350 |
| D | One workbook, one backend | C | **−2,000** |
| E | ✅ Conductor as code | D | +536 code / +453 tests / −85 docs (est. −150) |
| F | Stage granularity | E | −400 |
| G | Falsification pass | B | +500 |
| H | Single-surface analyst intake | E (H1: none) | +300 / −200 |

A gates everything. After A, **G and H1 can run in parallel with C/D/E** —
different files, no overlap. B → C → D → E → F is a hard chain; H2 wants E done
first.

Phases A–G are structural: they attack the failure mode the release history
exposes. Two phases are not. **H is driven by analyst experience rather than
defect data** — the intake is unpleasant to use, not broken. **I is dev-machine
hygiene** on a path production never touches and Phase D deletes; it is scheduled
only because a red test, a modal dialog, and a memory leak obstruct the migration
work itself.

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

3. **Promote into both assemblers** as write → verify → repair → re-verify,
   bounded to ~3 iterations, stage fails if it cannot converge.

**Then delete** — only once step 3 is trusted:

- `_PALATINO_CHAR_WIDTH_PER_PT` and `palatino_text_width_in`
- the `fit_overview_textbox` em constants
- `_set_table_height`'s `min_heights` estimation
- `_fill_risk_table`'s font-stepping ladder

Each becomes measure-then-adjust against the real render.

**Exit:** all three historical bugs caught by CI; the estimation code is gone.

✅ **Step 3 shipped 2026-07-27 (v0.5.39). PHASE B COMPLETE.** `scripts/deck_repair.py`
is the loop; **224 lines** of estimation code deleted across four commits (3a alone
first, 3d last), with the three fixtures re-verified between each. Exit criteria met:
all three still caught, and PRL17 / PRL18 are now also *repaired* in one measured
pass each. Detail in the CHANGELOG; four things worth carrying forward:

- **Attribution came first, as a prerequisite rather than a refinement.** A repair
  step that cannot attribute overflowing ink to a shape does not know which shape
  to shrink. `masked-overflow` renders a shape *alone* (its own layout, minus an
  empty slide of that layout) and measures its own ink; every probe rides in one
  deck, so the pass costs one extra conversion. It closes step 2's documented blind
  spot — the EU broker table, 0.037" unclaimed but 0.153" attributed.
- **LibreOffice is OPTIMISTIC about autofit, not conservative.** It treats any
  `<a:normAutofit>` as shrink-to-fit and recomputes its own scale, explicit
  `fontScale` included; PowerPoint applies the stored scale and nothing more.
  Measured: identical over-long copy renders 0.100" over with no autofit element
  and **clean at fontScale 100/90/80/70**. So v0.5.37's contract was blind to the
  PRL14 defect class on any freshly built deck — every INFOR overview block carries
  such an autofit, and PRL14 was caught only because that artefact has none. Probes
  now bake the scale in and strip the autofit; `rendered-overflow` skips autofit
  shapes so it cannot outrank attribution (the two disagree 0.14" vs 1.51" on the
  EU fixture). **The Phase A "LibreOffice is the conservative oracle" finding holds
  only for shapes without autofit** — worth remembering in Phase G.
- **The oracle found two defects nobody had recorded, both fixed at source:** the EU
  assembler never sized the Business Updates box (0.28" tall, 2.4" of empty band
  beneath, relying on autofit — the v0.5.23 defect on an unrendered shape), and
  `set_cell_text("")` wrote a zero-length run with no declared size, so an empty
  cell reserved a line at the table style's default and **every pitch with an odd
  market-entry target count rendered 0.587" off the slide edge**.
- **Deleting an estimate can cost quality if the replacement is naive.** Dropping
  the Palatino width table with a per-run "subtract k points" repair converged, but
  took market-entry values from 9 pt to 8 pt for a defect the 11 pt labels caused.
  The repair now **caps** every body run at `max_size - k` instead, so the shrink
  lands where there is most to give. Measure the replacement, not just the deletion.

The frozen fixtures deliberately do **not** converge: `earnings-update-deck.pptx` is
a v0.5.5 artefact whose text blocks are over budget for boxes that were never
sized, and `pitch-deck.pptx` converges in zero iterations because all 12 of its
blocking findings are string-tier. Freshly assembled decks of both deliverables do
converge.

## Phase I — Windows COM dev-path hygiene

**Scheduled: between B1 and B2.** Dev-environment friction only — nothing here
touches production (Cowork has no Excel and never enters these paths) and Phase D
deletes the COM path outright. Keep it minimal; do not gold-plate a path with a
scheduled execution date.

It earns a slot anyway because all three items actively obstruct migration work:
a red test, a modal dialog that stalls every run, and a memory leak that makes
everything else flakier.

1. ✅ **shipped v0.5.38** (with items 2–3 — it is what made the new opt-in gate
   trustworthy). **Degrade an exhausted clipboard retry.** `_ClipboardPasteError`
   (`excel_to_powerpoint.py:45`) is deliberately *not* a `RuntimeError` so a
   transient clipboard race is not mislabeled "Excel unavailable" — that
   reasoning is kept. But the fallback site (`:107`) caught only `RuntimeError`,
   so when the retries **exhausted**, the error escaped and aborted instead of
   degrading to LibreOffice, which is installed and renders it fine. Now caught
   explicitly alongside `RuntimeError`, with a distinct stderr note that retries
   were spent. This is why `test_pitch_deck_inserts_cap_table_into_slide7` was
   red at v0.5.36 — and why the opt-in gate could not otherwise tell you whether
   the COM path actually works.

2. ✅ **shipped v0.5.38.** **Gate the four Excel-COM tests behind an opt-in env
   var (default off).** `test_slide_library_poc.py:724`, `:744`, `:778` and
   `test_earnings_update_assembler.py:302` each spawn a real Excel via
   `DispatchEx` (`excel_to_powerpoint.py:130`), and each **must** set
   `Visible = True` (`:142`) because `CopyPicture(xlScreen)` renders nothing in an
   invisible instance. The S&P Cap IQ Office Tools add-in loads into that
   throwaway instance, finds its Cap IQ Pro sibling absent, and raises its own
   `MessageBox` — which `excel.DisplayAlerts = False` cannot suppress, because it
   gates Excel's alerts, not a third-party add-in's. Up to four modal dialogs per
   run, each stalling the suite until dismissed. Gating matches where Phase A
   already went (LibreOffice default everywhere; PowerPoint COM opt-in via
   `INFOR_SLIDE_RENDER_BACKEND`).
   *Accepted consequence:* the Windows COM path loses routine coverage. That is
   fine — Phase D deletes it, and nothing else depends on those tests.
   Delivered as an `excel_com` marker (registered in `pyproject.toml`, enforced by
   one `pytest_collection_modifyitems` hook in `tests/conftest.py`) gated on
   `INFOR_EXCEL_COM_TESTS`. The LibreOffice cap-table variant,
   `test_render_parity.py` and `test_deck_contract.py` are deliberately ungated.

3. ✅ **shipped v0.5.38 — but the diagnosis above was wrong, and the leak is only
   partly preventable.** All measurements are `EXCEL.EXE` deltas around a real
   render, 2026-07-27:
   - `.Quit()` was **already called at all four sites**.
   - **Refs outstanding at `CoUninitialize()` do not orphan the process.**
     `CoUninitialize()` tears the apartment down and releases its proxies, so the
     server exits regardless — **delta 0 with and without** an explicit release,
     including with the raising frame's traceback retained the way a pytest report
     retains it. Skipping `CoUninitialize()` outright also measured **delta 0**.
     So "release the instance per render" was not the fix; it shipped as ordering
     hygiene.
   - **The real source is aborted COM *startup*.** `CoCreateInstanceEx` failing
     with `-2146959355 "Server execution failed"` happens *after* Excel has
     launched: it is up with ~60 threads and a full add-in set, but the interface
     handoff timed out, so no pointer comes back and there is nothing to `Quit()`.
     Confirmed **one orphan per failure** — a full opted-in run logged exactly two
     startup failures and ended delta 2. All orphans share one signature
     (0.22–0.28 GB, ~60–65 threads, 19 windows, all invisible, **no dialog**),
     matching the 13 above (3.18 GB ÷ 13 ≈ 0.245 GB each).
   - **Not fixable in-process:** a graceful `Quit()` needs a pointer that was
     never handed back, force-killing is barred, and a ROT attach would grab the
     analyst's own Excel. So it is made **loud** on stderr instead of silent —
     previously it normalized to `RuntimeError`, the caller degraded to
     LibreOffice, the test passed, and the orphan was invisible, which is how 13
     accumulate unnoticed. **Item 2's gating is the prevention that works:** a
     default `pytest` run launches zero Excels.
   - Also fixed here: the **`Quit()` guard divergence** (all four Excel sites
     called `excel.Quit()` bare in a `finally`, so a raising Quit skipped
     `CoUninitialize()`; `slide_render` had always guarded its own), consolidated
     into one `excel_to_powerpoint.excel_com_app()` instance owner — the same
     consolidation `find_soffice()` did, with a matching drift lock. Excel gets
     **no** `Workbooks.Count` guard: its `DispatchEx` makes a separate process so
     `Quit` is always safe, and a count guard would skip the Quit exactly when a
     workbook failed to close.

> **Do NOT disable Office Tools with an HKCU `SPGMI.ExcelShell` key.** A
> same-named HKCU key overrides HKLM wholesale, and a manifest-less stub is
> exactly what disabled the analyst's CapIQ add-ins on 2026-07-13. The registry is
> currently healthy (HKLM `LoadBehavior=3`, intact `.vsto` manifest, no HKCU stub,
> no `Resiliency\DisabledItems`) — this is **not** a recurrence of that incident
> and needs no repair. Likewise never force-kill `EXCEL.EXE` / `POWERPNT.EXE`;
> close orphans with a graceful `Quit()`.

**Exit:** `pytest` green with no Excel spawned by default — ✅ met (597 passed, 6
skipped, delta 0, twice). Orphan count zero after an opted-in COM run — ✅ met for
the four gated tests (delta 0, and delta 0 pre-fix on the identical invocation, so
a no-regression rather than a cure), ❌ **not met for a full opted-in suite run**
(delta 2). Across eight measured runs the **delta equals the number of logged
startup aborts every time, and zero aborts is always delta 0** — no run leaked on
the render or Quit path. The abort rate is load-dependent and the orphans are
**self-amplifying**: the first run, with one resident orphan, aborted zero times;
by the end of the session, with 11 resident, aborts were frequent, since each
resident instance slows the next boot into the handoff-timeout window. That
criterion is unreachable under the stated constraints; it needs the startup-abort
window itself narrowed.

**Open follow-up (not scheduled — Phase D deletes this path):** the exposure is
per-`DispatchEx`, and a single pitch deck assembly launches a fresh Excel for
*each* range render (cap table, ownership insiders, ownership institutions).
Reusing one instance across the renders in one assembly would cut the
startup-abort exposure roughly threefold. Deliberately not done here: it changes
`insert_excel_into_placeholder`'s contract and both assemblers, which is
gold-plating a path with a scheduled deletion date.

## Phase C — Name-based template addressing ✅ shipped 2026-07-28 (v0.5.40)

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

### What shipped

27 `infor_`-prefixed, **worksheet-scoped** names across the four templates
(cap table 9, ownership 6, comps 7, precedents 5), registered in
`template_layout.TEMPLATE_NAMED_RANGES` — derived from the same `CellAnchor`
declarations the writers and the sentinel checks read, so the registry and the
stamped file cannot drift. `precedents_input_ccy` ships as
`infor_prec_output_ccy`: the cell is labelled "Output:" and the aggregator
relinks it to the cap table's *output* currency, so the name follows the
artefact rather than the plan's shorthand.

Added by **`tools/add_template_named_ranges.py`** — re-runnable prep tooling,
COM-free, exempt from Phase D. It rewrites only `xl/workbook.xml` and copies
every other zip entry's payload through byte-for-byte; verified after stamping
that exactly one entry differs per template. Neither of the two obvious routes
was used, and the reasons are measured, not assumed: **openpyxl** keeps the
defined names but silently drops `xl/printerSettings1.bin`, and **Excel COM**
would rewrite the whole file through Excel with the Cap IQ add-in loaded. Excel
is still the *oracle* (`--verify-excel`), and it was self-tested against two
deliberately damaged copies to prove it can fire.

The sentinel tables are **kept for this release** as the cross-check, per step 2.
`verify_anchors` now fails when a name and its sentinel disagree about where a
cell is. **Follow-up release: delete the sentinel tables** — that is the debt
this cross-check exists to retire, and nothing else in the phase is outstanding.

Step 3 went further than replacing three call sites: `_PitchLayout`'s index
arithmetic is gone entirely, replaced by marker lookups over the finished slide
mix, so constructing the layout *is* the verification pass it used to precede.
`MARKER_BUILT_*` markers were needed because the library markers key off titles
that assembly overwrites — a post-fill lookup keys off the deferred placeholders
instead.

### Suite runtime, added to the phase after measuring Phase B

712s → 137–247s across six green runs (median 168s; four of the six inside
the ~3-minute target). Converted PDFs are cached by deck
*content* (hashing zip members, not file bytes — python-pptx stamps save times,
so byte-hashing would miss every generated deck); each process gets a private
LibreOffice profile (which also fixes concurrent conversions — 2 of 4 failed
sharing the default profile); and the suite runs `-n 6 --dist loadgroup` with one
render cache shared across workers. `-n 6` beat `-n auto`/20 (159s vs 232s):
every worker carries a fixed library-baseline render, so more workers thrash.

Two defects fell out of making it concurrent, both fixed at source: both
renderers were holding the **caller's** file open (PowerPoint denies other
readers; LibreOffice drops a `.~lock`), which surfaced as a bogus
`PackageNotFoundError` because `zipfile.is_zipfile` swallows the `OSError` — they
now render a private copy, which also cleared the PowerPoint-COM flake this phase
had accepted. And one fill-logic test was running the converge loop because it
called the assembler directly instead of the module's `converge=False` helper.

**Remaining variance is the box, not the change:** the same single conversion of
the same fixture measured 3.4s in one session and 8.7s in another. Squeezing
further would mean restructuring `verify_deck`'s two-conversion flow — the deck
render, then the attribution probe deck that depends on its results — which is
Phase B's oracle and was deliberately left alone.

## Phase D — One workbook, one backend — ✅ SHIPPED (v0.5.41, 2026-07-28)

The large one, and the largest single diff of the migration: **net −2,685 lines**
(2,324 added, 5,009 deleted) against the ≈−2,000 estimate. Single workstream, as
planned, because deleting the merge deleted most of the COM surface with it.

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

> **Open decision — RESOLVED, empirically, before any code changed.** Nothing
> depended on it; D was a clean deletion.
>
> The frozen production fixture carries **zero** defined names while the cap-table
> template ships 33 CapIQ ones including `CIQWBGuid`/`CIQWBInfo`, and the
> aggregator had no defined-name handling — so the merge already stripped CapIQ's
> workbook identity in production. The refresh was then measured on copies with
> the add-in active, over `comps` (where the live `_xll.…SPG` formulas actually
> are — the cap table's live in *cell comments* by the v0.5.3 design): the merged
> workbook, a pristine template with `CIQWBGuid` intact, and the merged workbook
> with the namespace case restored all read **`#PEND`** — Cap IQ's own async-fetch
> marker, emitted by `SPG` itself. The add-in recognises the merged workbook, and
> `CIQWBGuid` makes no difference.
>
> Two notes for the record. A first probe through `DispatchEx` was discarded as
> worthless — an automation-started Excel loads no `OPEN=`-registered SNL add-in,
> so every subject read `#NAME?`, control included; the probe had to relaunch
> `excel.exe` normally. And values never populated in *any* subject because a Cap
> IQ sign-in window was waiting — a gate on the environment, identical across all
> three, so the differential result stands. The cautious branch's requirement was
> adopted anyway at zero cost: the deal workbook preserves CapIQ's defined names
> from the outset, which makes the decision non-load-bearing either way.
>
> That `DispatchEx` behaviour turned out to be *useful*: it is what lets
> `tools/build_deal_workbook_template.py` drive Excel for the one-time sheet copy
> without the Cap IQ add-in ever loading — the objection
> `add_template_named_ranges.py` correctly raised against COM for template
> surgery.

### What the phase actually cost, and the two things it found

**The four steps landed as planned**, in three commits (steps 2 and 3 are not
separable — the moment a producer writes a tab there is nothing for a merge to
merge, and the aggregator's last two tests merged producer output directly, so
splitting them would have meant a knowingly red commit).

**`write_tab` had to be serialized, which the plan did not call out.** A wave
dispatches stages concurrently and openpyxl rewrites the *whole* workbook on
save, so `comps`, `precedents` and `financial-summary` reaching for one file at
once would silently drop each other's tabs. A six-thread test fails without the
lock.

**Templates arrive as a pre-assembled artefact, not assembled at runtime.**
"Copied in as tabs at init" cannot happen in the runtime: openpyxl cannot move a
sheet between workbooks, and a cell-by-cell copy is the lossy merge this phase
deletes. So `tools/build_deal_workbook_template.py` does it once, with an
add-in-free Excel, and init is a `shutil.copyfile`.

**Finding 1 — Excel AutoSave silently re-saved the shipped templates.** The repo
lives in a OneDrive folder, where Excel 365 enables AutoSave by default and writes
the file back *regardless of* `Close(SaveChanges=False)`. An early revision of the
prep tool opened the four templates in place and re-saved all of them through
Excel, undoing exactly the byte-level preservation Phase C's zip surgery existed
to achieve. Any future tool that opens a repo file in Excel must stage a copy
outside OneDrive first; a source-level test locks it, because reproducing it needs
Excel *and* OneDrive.

**Finding 2 — Phase C shipped a production-breaking regression that Phase D's own
deletions exposed.** v0.5.40 converted the earnings assembler's `slide_index=1`
into a marker lookup and chose `summary_at`; `Rectangle 3` is the **overview**
slide's placeholder, so every earnings-update run with a cap table raised
`KeyError` and failed the deck stage. It shipped green because both covering tests
were invisible on a dev box — one behind the opt-in `excel_com` gate, one behind
`skipif win32` reading "Windows uses the Excel COM path". Deleting COM made both
guards false and removing them failed instantly.

The lesson generalises past this phase, and it is the second instance after
v0.5.36: **a skip guard justified by a platform difference outlives the
difference, and then hides defects on the platform it exempts.** Both are gone;
the suite is now **595 passed, 0 skipped**, where six tests previously never ran
here. A gate that stays off by default is not coverage — it is a promise to
remember, and this repo has now failed to remember twice.

### The xdist LibreOffice-profile item was already done

The brief proposed per-invocation `-env:UserInstallation` profiles on the
hypothesis that `slide_render._libreoffice_render` passed none. Phase C had
already implemented that mechanism, having run the same experiment (2 of 4
concurrent conversions failed on a shared profile; 4 of 4 with one each). An audit
confirms both — and only — `soffice` call sites carry the flag. Nothing changed.

Measured anyway, on the finished phase (`-n 6`, uninterrupted): **277.4 / 235.4 /
181.7 / 240.9 / 249.2 / 197.6 s** — spread **96 s**, median 238 s, against Phase
C's 137–247 s (median 168) on a suite of almost the same size. There is no
before/after, because both halves would have been the same code. The spread is
the box, exactly as Phase C concluded from the same variance — and a first
attempt had to be discarded because prep-tooling Excel work overlapped it,
which shows how little else the machine can be doing for the number to mean
anything.

## Phase E — Conductor as code — ✅ SHIPPED (v0.5.42, 2026-07-28)

Much smaller than D, as forecast — no barrier, fewer stages.

1. **`conductor_cli.py` → `conductor.py`, and the trip finished.** `prep_wave` /
   `collect_wave` resolved and collected; the model still owned the wave
   narration, the checkpoint payloads, and the summary. Added: `plan_overview`
   (plan summary + wave schedule), `prepare_wave` → a typed `WaveDispatch` of
   `PreparedStage`s each carrying a rendered `Task` prompt, `complete_wave` → a
   typed `WaveOutcome` with `ok` / `halt` / `is_final` and one `Checkpoint` per
   stage (the `required` one carrying the code-owned `AskUserQuestion` payload and
   a plain-text fallback), `run_wave(run_dir, n, dispatch)` composing all three,
   and `render_run_summary` / `write_run_summary`. Renamed with no shim, per the
   repo's no-back-compat rule.
2. **The env-var export block is gone.** Three command-line arguments — plugin
   root, `inputs.json`, `outputs.json` — rendered into the prompt as the exact
   invocation and read back by the new `scripts/stage_io.py`; the resolved inputs
   are *also* inlined in the prompt body. All twelve dispatched skills converted.
   `deal_dir` is derived by walking up to `deal.json` rather than passed, so
   finding it is the proof the stage landed somewhere real.
3. **`skills/conductor/SKILL.md`: 216 → 118 lines.** Above the ~80 target, and
   deliberately: what is left is intake (Steps 1–4), dispatch, checkpoints and
   summary, and **the intake is the bulk of it**. The locked-questionnaire
   principle means the deliverable-specific answer mapping — which dialog option
   sets which plan input, which defaults are computed rather than asked — is
   instructions to the model, not code the driver can own. Phase H's `IntakeSpec`
   is what makes that part shrink; compressing it into ambiguity here would have
   traded a real property for a line count.

### The −150 estimate was wrong, and wrong in an instructive way

Actual: **+536 lines of runtime code** (`conductor.py` 374 → 766, `stage_io.py`
144 new), **+453 lines of tests**, **−85 lines of skill docs**. The estimate
assumed the SKILL.md shrink would dominate. It could not: the 98 lines that left
the conductor SKILL.md were *instructions the model re-executed by hand on every
wave*, and replacing them with something that cannot be skipped means code plus
the tests that lock it. Same for the handoff — a one-line `export` block became a
144-line module because "read three paths, derive the deal directory, prove it
exists, write the answer" is what the export block was silently *not* doing.

D's −2,685 came from deleting a subsystem. E's +900 comes from moving behaviour
out of prose. The "net lines" column measures the first kind well and the second
kind not at all; F is the first kind again.

### run_wave's signature, and why the model calls its two halves

The plan wrote `run_wave(run_dir, n)` as though one call could hand back the
`Task` prompts *and* return the checkpoint. It cannot: a function returns once,
and the dispatch in between is a set of `Task` tool calls no Python callable can
issue. So `run_wave(run_dir, n, dispatch)` takes the dispatch as a callback and
composes the whole trip, and the conductor skill calls `prepare_wave` → its `Task`
calls → `complete_wave` — the same sequence with the model standing in for the
callback. The callback form is not vestigial: it is what the tests drive a full
wave through end to end, and it is the seam Phase F's in-process transform stages
plug into.

### Two defects the move to code exposed

- **A `silent` stage that FAILED did not have to surface.** `silent` suppresses a
  routine summary; it was never meant to hide the reason a run stopped. Nothing
  distinguished the two while the behaviour lived in prose. `_checkpoint_for` now
  surfaces a failure whatever the mode.
- **The stale 7-wave pitch example** in the conductor SKILL.md is deleted rather
  than corrected — the skill posts `plan_overview(...).narration()`, so there is
  no second copy of the number left to go stale. The README's two claims are now
  the only prose copy, and `test_readme_wave_counts_match_the_scheduler` parses
  and checks them. Confirmed alongside: the hardcoded aggregator barrier is gone
  from `plan_schedule.py`, and `compute_waves` returns pitch **10 / 6** and
  earnings update **5 / 3**.

### Skip-guard audit — the standing check after v0.5.40

Nine `find_soffice() is None` guards are genuine and kept: LibreOffice is the one
render backend, and a box without it truly cannot run those. Five were stale and
deleted — four `importorskip("pypdfium2")` (Phase A made it an unconditional
declared dependency, so a missing one is a broken install that must fail loudly)
and `test_slide_render`'s `_LIBRARY.exists()` skip, whose **cwd-relative** path
made "ran pytest from the wrong directory" and "the shipped slide library was
deleted" produce the same green run.

Two more were not stale but **too broad, which is the same failure with a
different cause**: `test_slide_library_poc`'s two ownership-slide tests caught
`RuntimeError` and skipped on it. `_render_range_to_png` raises that for a missing
LibreOffice *and* for a conversion that failed or produced no PDF, so a genuine
render defect read as a green skip. Both now carry an explicit `find_soffice`
skipif and let everything else fail. **625 passed, 0 skipped** (595 before).

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

## Phase H — Single-surface analyst intake

Best after E (the conductor is a program by then, so intake is a function that
returns a dict — the shape a form submission already has). **H1 is independent
and may be pulled forward at any time.**

**The problem is delivery, not content.** The question set is already code-owned
and locked (v0.5.27): `deal_init.render_init_dialogs` (Listing / Sector /
Filings) and `deck_spec.render_deck_spec_dialogs` (Notes / CIM / Valuation /
Risk notes / Targets / Highlights) are rendered verbatim — Claude is not
inventing options. What is unpleasant is that a pitch start is **four sequential
`AskUserQuestion` round trips plus three plain-text blocks** (filings note,
defaults echo, documents note) interleaved down the chat. The analyst never sees
the whole intake at once and answers it in four bites.

### H1 — One declarative spec behind every rendering

**Prerequisite, and valuable on its own.** The current "cannot drift" guarantee
is narrower than it looks: `_dialog_item_plan_inputs` derives the numbered-item →
plan-input **mapping** from the dialog order, so the *answer mapping* is safe —
but `_PITCH_SPEC_PROMPT` (`deck_spec.py:540`) is a hand-written string literal
carrying its own copy of every question's wording, defaults, and option labels,
independent of `_PITCH_SPEC_DIALOGS`. Change a dialog option label and nothing
forces the text prompt to follow. **There are already two hand-maintained content
renderings; a widget would make three.**

So H1 collapses them: one declarative `IntakeSpec` (fields, options, defaults,
plan-input target, required/optional) with the dialogs and the text prompt both
*generated* from it. Ship this whether or not H2 ever happens — it removes an
existing drift surface and makes the locked-questionnaire principle structural
rather than conventional.

### H2 — Inline interactive form

A third generated rendering: one fixed-design intake form, every field visible at
once, a Start button that posts all answers back as a single structured payload.
Four round trips → one. Rendered inline via the host's `show_widget` (the
`visualize` MCP server), whose `sendPrompt(text)` global submits the payload as
if the analyst typed it.

Because H1 made the spec declarative, the form is generated, not authored — a
third rendering of one source, not a fork.

> **Open decision — resolve before H2.** `visualize` is a **host-provided** MCP
> server, confirmed present on the desktop session where this was scoped;
> unconfirmed on the Cowork/Linux production runtime. One check on a Cowork shell
> settles it (same class as the Phase A `fc-match`). Note that a plugin cannot
> *depend* on a host MCP server in any case, so the fallback chain
> **widget → dialogs → text is mandatory** and the widget can only ever front the
> dialogs, never replace them. If `visualize` is absent on Cowork, H1 still stands
> on its own and H2 becomes desktop-only ergonomics.

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
