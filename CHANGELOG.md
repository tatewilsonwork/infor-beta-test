# Changelog

All notable changes to `infor-beta` are documented here. Format: [Keep a Changelog](https://keepachangelog.com/en/1.1.0/). The plugin uses a single version across all skills; bump every skill's `version:` frontmatter when bumping the plugin.

## [0.5.17] — 2026-06-29

Two pitch-deck chart follow-ups (pitch only; earnings-update untouched): an LTM-revenue pie on the overview slide, and two formatting fixes to the Financial Summary charts. Both reuse the existing `financial-charts` chart infrastructure (Excel COM build → `Chart.Export` PNG, openpyxl + LibreOffice fallback, `insert_pngs_into_placeholders`).

### Added
- **LTM revenue pie on the overview slide.** `financial_charts.render_ltm_revenue_pie_into_deck` builds an LTM-revenue-by-segment pie on the combined workbook's `ltm-metrics` tab — over the "LTM Revenue Overview" block (categories = Segment column, values = "LTM Revenue (…)" column; the **Total** row excluded), located by its section title via `ltm_revenue_overview_range` (no hardcoded rows; mirrors the aggregator's label-row scan) — and inserts it into the overview slide's (`prs.slides[6]`) deferred `[Pie Chart Placeholder]` (`Rectangle 4`). Excel does the math (the pie references the literal cells). Legend at the **top**, no chart title, no chart border; slice fills from the **INFOR theme accent palette** (new `pptx_helpers.INFOR_ACCENTS` = `0E213F, 46566E, ADB9CA, A4844B, 767171, E5E3E3`, from `INFORFG.thmx` "INFOR (New)" accent1–6), in theme order and cycled past six; Palatino 9 labels with percentage data labels. The chart persists on the `ltm-metrics` tab (mirrors how the FS charts persist on `financial-summary`). Null path: when there is no `ltm-metrics` tab / "LTM Revenue Overview" block, the pie is skipped and the placeholder is left in place. The `financial-charts` stage now renders the pie after the FS charts (chaining the same deck) and its mandatory QA render covers slides `[6, 7]`. (`scripts/financial_charts.py`, `scripts/pptx_helpers.py`, `skills/financial-charts/SKILL.md`)

### Fixed
- **Financial Summary chart border + gray axis.** The legacy `ChartArea.Border.LineStyle = 0` did not suppress the modern chart-area outline; `_format_com_chart` now clears `ChartArea.Format.Line` (+ the plot-area outline) via a shared `_com_strip_chart_border`, and sets the **category-axis line explicitly black** (`cat_axis.Format.Line`, value axis still hidden). The openpyxl fallback (`_make_openpyxl_chart` / `_make_single_value_chart`) mirrors both via `_openpyxl_no_border_black_axis` (no-fill chart-area border, black `x_axis.spPr` line). (`scripts/financial_charts.py`)

### Bumped
- marketplace, plugin manifest, pyproject, README status line, and all shipped skill frontmatter versions to `0.5.17`.

## [0.5.16] — 2026-06-29

Finishes the pitch deck's **Financial Summary** slide: builds the four metric charts on the chart-ready `financial-summary` tab and renders them into the slide. Also closes two `financial-summary`/`ltm-metrics` data-tab defects and a latent aggregator relink bug that left the slide's LTM bars blank.

### Added
- **`financial-charts` skill + `scripts/financial_charts.py`.** A new pitch-only stage that runs **after** `workbook-aggregation`. It builds one INFOR-formatted clustered-column chart per metric (single series = the metric's row, categories = the period header read dynamically so a suppressed-LTM tab charts five columns) on the **combined** workbook's `financial-summary` tab — the only place each flow metric's `=INDEX('ltm-metrics'!…)` LTM link resolves — then renders each chart and inserts it into the slide's four chart placeholders (`Rectangle 17/7/19/18` ← data rows 6–9), stretched to each placeholder's box. Charts persist on the deliverable tab too. Formatting: Palatino 9 black labels, no title/gridlines, hidden value axis, gap width 50%, data labels Outside End, bars filled `46566E`. Excel COM on Windows (full fidelity, builds native charts + exports PNG); openpyxl + LibreOffice fallback off-Windows. A new `financial-charts` plan stage wires `combined_workbook_path` (from `workbook-aggregation`) + `deck_path` (from `deck`). (`plans/pitch.yaml`, `skills/financial-charts/SKILL.md`)
- **Combined-metric values as Excel formulas.** `MetricSeries.fiscal_values` / `ltm_value` now accept an Excel formula string beginning with `"="` (e.g. `"=9000+800"` for a combined metric like "Ending Combined Loan & Advance Bal."), validated and written through as a cell formula — the arithmetic stays auditable (Excel does the math), never pre-summed. (`scripts/financial_summary_workbook.py`, `skills/financial-summary/SKILL.md`)

### Fixed
- **Financial Summary LTM bars were blank / 10⁶ off.** Two root causes, both fixed:
  - **Aggregator left the financial-summary→ltm-metrics link external (COM path).** Copying the financial-summary sheet bound its LTM-link formulas to an *external* workbook relationship (the soon-deleted source); Excel's `.Formula` getter showed an internal-looking `'ltm-metrics'!` but the cell still resolved to `#N/A`. The aggregator now re-assigns each financial-summary LTM-link formula, re-binding it to the sibling `ltm-metrics` tab so it resolves on recalc — mirroring the existing cap-table / ownership relinks. (`scripts/workbook_aggregator.py`: `_relink_financial_summary_com` / `_relink_financial_summary_openpyxl`, `_internalize_external_sheet_ref`)
  - **Units scale mismatch.** Locked both the `financial-summary` and `ltm-metrics` skills onto **millions with an `"MM"` suffix** (`US$MM`, `C$MM`) — never `US$` / full dollars — because each flow metric's LTM cell links the bridge total value-for-value, so a scale mismatch made the LTM bar 10⁶× off. (`skills/financial-summary/SKILL.md`, `skills/ltm-metrics/SKILL.md`)

### Changed
- **Aggregator barrier generalized for post-aggregation stages.** `plan_schedule.stage_dependencies` no longer forces `workbook-aggregator` to depend on *every* stage — it now excludes the aggregator's own downstream consumers (its data-edge descendants). Without this, a stage that consumes the combined workbook (like `financial-charts`) would form an `aggregator → consumer → aggregator` cycle. Earnings-update (no post-aggregation consumer) is unchanged; pitch schedules into **7 waves** (`workbook-aggregation` then `financial-charts` last). (`scripts/plan_schedule.py`)
- **Slide-8 charts no longer deferred in the deck-assembler.** `deck-assembler` still fills only the metric-name tiles; the four chart placeholders are now filled by the downstream `financial-charts` stage. (`skills/deck-assembler/SKILL.md`)

### Bumped
- marketplace, plugin manifest, pyproject, README status line, and all shipped skill frontmatter versions to `0.5.16` (including the new `financial-charts` skill).

## [0.5.15] — 2026-06-29

Adds a chart-ready `financial-summary` data tab behind the pitch deck's Financial Summary slide, and makes that stage the single source of truth for the deck's four financial metrics.

### Added
- **`financial-summary` skill + `financial_summary_workbook.py`.** A new pitch stage that selects the four most relevant metrics for the target (industry-aware: operating company vs. financial institution), gathers their **last five fiscal years** from the latest 10-Ks plus an **LTM** column, and emits a standalone companion `.xlsx` folded into the combined `pitch-<codename>.xlsx` as the `financial-summary` tab. The tab is laid out *chart-ready* (one metric per row, fiscal years + LTM as a single contiguous numeric header axis, no merged cells in the data block, a Units column) so a later task can drop native Excel charts on it with no reshaping — building the charts is out of scope here. The four metric labels are emitted as a typed stage output and become the deck's slide-8 tiles.
- **LTM linkage to the `ltm-metrics` tab.** Each flow metric's LTM cell is a label-keyed lookup (`=INDEX('ltm-metrics'!$B:$B, MATCH("(=) <result_label>", 'ltm-metrics'!$A:$A, 0))`) that resolves in the combined workbook and stays `#N/A` in the standalone file — the same "unresolved in standalone, resolves post-merge" pattern as the cap table's CapIQ formulas. Non-flow metrics (balances / ratios) fall back to the latest reported value.
- **`extra_bridges` on the LTM workbook builder.** `build_ltm_metrics_workbook` gains an optional `extra_bridges` parameter (and a `Bridge` dataclass) that appends one `FY + YTD − prior-YTD` bridge per entry, reusing `_write_bridge`. In the pitch plan `financial-summary` drives these via its `ltm_bridge_specs` output so `ltm-metrics` bridges exactly the selected metrics; the earnings-update plan passes nothing, so its behaviour is unchanged.

### Changed
- **`financial-summary` runs before `ltm-metrics` in `plans/pitch.yaml`.** Because `financial-summary` selects the metrics and tells `ltm-metrics` (via `ltm_bridge_specs`) which extra ones to bridge, it precedes `ltm-metrics`; the `deck` stage reads `financial_metric_labels` from it and the `workbook-aggregation` stage folds its workbook in. The pitch plan now schedules into 6 waves (`financial-summary` joins `wireframe`/`comps`/`precedents` in wave 1; `ltm-metrics` moves to wave 2). No schema change — the references are the DAG.
- **Label selection removed from `pitch-content`.** Dropped `financial_metric_labels` (and its two validators) from the `PitchDeckContent` schema and the `pitch-content` skill; the `deck` stage now takes the labels from the `financial-summary` stage output instead, threaded into `assemble_pitch_deck(financial_metric_labels=...)`. (`scripts/schemas/pitch_deck_content.py`, `skills/pitch-content/SKILL.md`, `scripts/pitch_deck_assembler.py`, `scripts/pitch_deck_wireframe.py`, `skills/deck-assembler/SKILL.md`)
- **Deal-init filings prompt generalized.** The G6 prompt now asks for the latest four fiscal-year 10-Ks (for the five-year financial-summary history) in addition to the interim YTD stubs needed for the LTM bridge, for pitch as well as earnings update. (`scripts/deal_init.py`)

### Bumped
- marketplace, plugin manifest, pyproject, README status line, and all shipped skill frontmatter versions to `0.5.15` (including the new `financial-summary` skill).

## [0.5.14] — 2026-06-09

Adds a precedent-transactions slide to the pitch deck and fixes a batch of companion-workbook and deck issues surfaced by a live pitch run.

### Added
- **Precedent-transactions slide in the shared slide library.** `INFOR Slide Library.pptx` gains a `Precedent Transactions Analysis` slide immediately after the (renamed) `Comparable Companies Analysis` slide — a chart-placeholder slide like comps, carrying a one-line takeaway. The library grows 16 → 17 physical slides and the pitch deck's base order 15 → 16. Wired through every layer: `slide_library_registry.py` (new `precedent-transactions` entry at slide 12, closers renumbered), `pitch_deck_wireframe.py` (`_content_block` + `_section_for`), `pitch_deck_assembler.py` (fills the takeaway; hardcoded post-delete indices bumped for everything after comps; `[Placeholder for Precedents Chart]` added to the verifier), and a new required `precedents_takeaway` field on `PitchDeckContent` (drafted by `pitch-content`). The slide stays a chart placeholder — no Excel→PowerPoint step while Capital IQ can't be refreshed here.

### Fixed
- **Workbook aggregator dropped the precedents source hyperlinks (openpyxl path).** `_copy_sheet` copied values + styles but never `cell.hyperlink`, so the precedents `AB`–`AG` source links vanished from the combined workbook on the openpyxl backend (Cowork/Linux/macOS, or Windows without Excel). It now copies the hyperlink alongside the value/style; the COM backend already preserved them. (`scripts/workbook_aggregator.py`)
- **`Font(color="0000FF")` reset cells to Calibri 11.** A bare openpyxl `Font(color=...)` carries no typeface, so it dropped the template's Palatino and rendered Calibri 11. Fixed the reported cells and their siblings: the cap table's web-sourced `F7`/`F16` and the LTM `D47`/`D48` now use `Font(name="Palatino Linotype", size=9, color="0000FF")` (`skills/captable/SKILL.md`); the ownership insider data cells re-emit the template font blue and `F35` is left at the template's Palatino — no font write, the aggregator relinks it — via a font-preserving helper (`scripts/ownership_workbook.py`); and the precedents target/acquiror (`F`/`G`) are written Palatino 9 explicitly, since the shipped template's name cells were themselves a stray Calibri 11 (`scripts/precedents_workbook.py`).
- **Pitch slide 10 left a blank fifth Considerations/Mitigants row.** `pitch-content` now drafts **five** consideration/mitigant rows (the slide-10 table has five body rows; the schema already allowed up to five). (`skills/pitch-content/SKILL.md`)
- **Market-entry tables (slides 13–16) rendered ~6.3" tall instead of the 5.71" clamp.** The generator writes a 5.71" frame + row heights, but PowerPoint grows a row to fit its text (a stored row height is only a minimum), so the long Overview / Strategic Rationale copy at 10 pt re-expanded the table on open. The value font drops 10 → 9 pt so that copy fits the clamped rows, and `pitch-content` now asks for concise wordy cells. (`scripts/pitch_deck_assembler.py`, `skills/pitch-content/SKILL.md`)
- **Earnings-update assembler kept the wrong closers after the library grew.** The precedents-slide insertion shifted disclaimer/contact from raw indices 14/15 → 15/16; `_KEEP_LIBRARY_INDICES` is updated `(0,6,7,14,15)` → `(0,6,7,15,16)` so the earnings deck still ends with Disclaimer + Contact. (`scripts/earnings_update_assembler.py`)

### Changed
- **Precedents now requires a multiple per deal and targets six deals per group.** The builder rejects a deal that carries only a TEV (no metric and no disclosed multiple → it would just add an empty row), and the skill is directed to find six **valuable** deals per group (12 total), dropping and replacing any deal it can't value. (`scripts/precedents_workbook.py`, `skills/precedents/SKILL.md` + `references/sourcing-criteria.md`)
- **Comps `F3` and precedents `C2` relink to the cap table's output currency.** The aggregator's cross-tab relink pass now points both currency cells at `captable!F5` (restyled Palatino 9) so the combined workbook shows one consistent output currency instead of each skill's standalone literal. (`scripts/workbook_aggregator.py`)

### Bumped
- marketplace, plugin manifest, pyproject, README status line, and all shipped skill frontmatter versions to `0.5.14`.

## [0.5.13] — 2026-06-09

The conductor now runs independent stages in parallel.

### Added
- **`scripts/plan_schedule.py` (`compute_waves` / `stage_dependencies`).** Derives the plan's dependency DAG from the `$stages.<id>.<name>` references already present in each stage's inputs (walking nested dicts/lists, so the aggregator's `workbooks:` map is covered) and topologically sorts the stages into ordered **waves** of mutually-independent stages. The `workbook-aggregator` stage is forced to depend on *every* other stage — a hardcoded final-barrier rule, because it merges and **deletes** the individual companion workbooks and the deck-assembler reads one of them (the standalone cap table) before it is folded in; that ordering is a filesystem side-effect, not a value reference, so it is invisible to the auto-derived data edges. No `depends_on` field is added to the schema — the references *are* the DAG. Cycles raise `PlanCycleError`. (`scripts/plan_schedule.py`; `scripts/tests/test_plan_schedule.py` added — asserts the wave structure of both shipped plans plus barrier / cycle / nested-reference / unknown-reference edge cases)

### Changed
- **Conductor dispatches stages wave-by-wave instead of one at a time.** Step 6 now calls `compute_waves(plan)` and issues one `Task` (Agent) call per stage in a wave **in a single message** so they run concurrently, waits at the wave boundary, then collects each stage's outputs / log / checkpoint before starting the next wave. The pitch plan collapses from 9 sequential stages to **5 waves** — wave 1 overlaps the four research-heavy roots `wireframe` / `ltm-metrics` / `comps` / `precedents` — and earnings-update goes from 6 stages to **4 waves**. `required` checkpoints are now evaluated at the wave boundary, with a documented caveat: a `required` gate stops *downstream* waves, not its own wave-mates (every shipped plan uses `informational`, so behaviour is unchanged). This supersedes the v1 "sequential, no parallel" decision (Obsidian note 12 H1) per analyst direction. (`infor-beta/skills/conductor/SKILL.md`, `references/plan-schema.md`, `scripts/schemas/plan.py` docstrings; `CLAUDE.md` phase status + helper-import list)

### Bumped
- marketplace, plugin manifest, pyproject, README status line, and all shipped skill frontmatter versions to `0.5.13`.

## [0.5.12] — 2026-06-09

Adds the precedent-transactions skill to the pitch pipeline.

### Added
- **`precedents` skill + `scripts/precedents_workbook.py`.** Builds the INFOR precedent-transactions table: researches up to 12 M&A deals (two peer groups of six, labels `E7` / `E16`) and writes each deal's identity (input currency `B`, announce date `E`, target `F`, acquiror `G`, source-FX TEV `I`, 3-letter HQ code `AI` — column `H` left empty), the source-FX $ metric inputs for the **one metric family** the agent picks by the target's industry — operating companies → Revenue (`K`/`L`) + Adj. EBITDA (`O`/`P`) → EV/Revenue + EV/EBITDA; financial institutions → Net Income (`M`/`N`) + Book Value (`Q`) + Tangible Book Value (`R`) → P/E + P/B + P/TBV — and a source hyperlink per metric on the `AB`–`AG` "Link" cells. A **multiple disclosed in the deal PR** is preferred and is written as a literal straight over the template's ratio formula in `S`–`Z`; otherwise the disclosed $ figure is written (using the most recent reported figure as the LTM/NTM proxy — the old multi-filing LTM stub calc is dropped). The column-`C` CapIQ FX rate, the `=+I*C` TEV conversion (`J`), the ratio formulas (`S`–`Z`), and the group/global statistic rows are template-owned and stay un-evaluated until the analyst refreshes Capital IQ in Excel. Activates on `/precedents` and as the pitch plan `precedents` stage. Sourcing criteria carry over from the production `infor-workflows` `precedents-infor` skill. (`infor-beta/skills/precedents/SKILL.md` + `references/sourcing-criteria.md`, `scripts/precedents_workbook.py`; tests added)
- **`templates/INFOR Precedents Template.xlsx`.** The analyst's precedents template, shipped so `precedents` can clone it. Like the ownership template (and unlike comps) it carried heavy cruft — ~58.7k legacy CapIQ defined names + 174 vestigial external-workbook links (6 MB) plus leftover example hyperlinks in the `AB`–`AG` cells — none of which any live formula references; all were stripped so the openpyxl output stays Excel-openable (→ ~9 KB, FX / ratio / statistic formulas intact). Two template edits were applied during prep: `AG4` relabelled `P /E ` → `P / TBV` (AG is the Tangible Book Value source-link column) and `C2` defaulted to `USD` (the output-currency cell the FX formula keys off).

### Changed
- **Pitch plan runs `precedents`.** A new `precedents` stage runs after `comps` and before `deck`; it depends on nothing upstream (only the target's facts). The precedents workbook folds into the combined `pitch-<deal>.xlsx` as the `precedents` tab. The deck's precedents slide stays a placeholder — no Excel→PowerPoint step while Capital IQ can't be refreshed here. (`infor-beta/plans/pitch.yaml`)

### Bumped
- marketplace, plugin manifest, pyproject, README status line, and all shipped skill frontmatter versions to `0.5.12`.

## [0.5.11] — 2026-06-08

Adds the public-comparables (trading comps) skill to the pitch pipeline.

### Added
- **`comps` skill + `scripts/comps_workbook.py`.** Builds the INFOR public-comparables table: finds three verticals (peer groups) relevant to the target and, per vertical, writes six public companies as Capital IQ `Exchange:Ticker` identifiers (`B10:B15` / `B20:B25` / `B30:B35`), the vertical label (`D9` / `D19` / `D29`), and a ≤50-char description each (`AA10:AA15` / `AA20:AA25` / `AA30:AA35`). Every market-data / multiple / statistic column is a Capital IQ array formula keyed off column B, so the skill writes only the inputs CapIQ can't supply and leaves the formulas **un-evaluated** — this environment has no CapIQ connector, so the analyst opens the workbook in Excel with the add-in active and refreshes. Activates on `/comps` and as the pitch plan `comps` stage. The ticker format and peer-selection rules carry over from the production `infor-workflows` `comps-infor` skill. (`infor-beta/skills/comps/SKILL.md`, `scripts/comps_workbook.py`; tests added)
- **`templates/INFOR Comps Template.xlsx`.** The analyst's comps template, shipped so `comps` can clone it. Unlike the ownership template it round-trips cleanly through openpyxl (CapIQ array formulas preserved, output stays openable in Excel and re-openable via Excel COM for the aggregator), so it ships as-is — no defined-name / external-link pre-cleaning needed.

### Changed
- **Pitch plan runs `comps`.** A new `comps` stage runs after `ownership` and before `deck`; the `deck` and `workbook-aggregation` stages now consume `$stages.comps.workbook_path` instead of the manual `comps_workbook_path` plan input, which is removed. The comps workbook folds into the combined `pitch-<deal>.xlsx` as the `comps` tab (the CapIQ `__snloffice` helper sheet is dropped automatically). The deck's comps slide stays a placeholder — no Excel→PowerPoint step while Capital IQ can't be refreshed here. (`infor-beta/plans/pitch.yaml`)

### Fixed
- **Template resolution no longer breaks on Windows.** `scripts/find_template.sh` printed an absolute path via `pwd`, which under Git Bash is a MinGW path (`/c/Users/…`). Captured as a string in Python (no shell arg-conversion), `pathlib` mis-resolved it to a drive-relative `\c\Users\…` and the builder raised `FileNotFoundError`. The helper now emits a **native** path — `pwd -W` on Windows/Git Bash (`C:/Users/…`), falling back to `pwd` on macOS/Linux (`/Users/…`) — which both `cp` and `pathlib` consume everywhere. The `ownership` skill, which fed the helper's output straight into `build_ownership_workbook(template_path=…)`, now resolves the template **in Python** via `CLAUDE_PLUGIN_ROOT / "templates" / …` (matching `deck-assembler` / `comps`), removing the fragile bash→Python handoff. `captable` was not affected — it copies the template with bash `cp` and opens the relative output copy, so the MinGW path never reached Python; the helper hardening makes its path native regardless. (`scripts/find_template.sh`, `ownership` SKILL.md; regression test added to `test_shell_helpers.py`)

### Bumped
- marketplace, plugin manifest, pyproject, README status line, and all shipped skill frontmatter versions to `0.5.11`.

## [0.5.10] — 2026-06-05

Three analyst-reported fixes to the pitch / workbook pipeline.

### Fixed
- **Combined workbook keeps its INFOR theme, clean tab names, and cross-sheet links.** The workbook aggregator's COM merge was silently failing and falling back to the openpyxl backend — which starts from a blank (default-Office-theme) workbook — so the combined `pitch-<deal>.xlsx` lost the INFOR colour scheme, named the ownership tabs `ownership-Ownership` / `ownership-Bloomberg Output`, and broke the `Ownership` sheet's `='Bloomberg Output'!…` lookups (`#REF`). Root cause: the cross-workbook `Worksheet.Copy(After=…)` **named** argument is silently dropped by this Excel build, so the copy lands in a stray new workbook instead of appending — the v0.5.9 "activate before copy" change didn't address it, because activation was never the problem. Fixes (all in `scripts/workbook_aggregator.py`):
  - The copy destination is passed **positionally** (`Copy(None, <last sheet>)`), so sheets append to the base workbook and the COM path no longer falls back.
  - Each source's content sheets are copied **as a group** in one operation, so a source's intra-workbook cross-sheet references stay internal instead of becoming external links to the soon-deleted source — the ownership `Ownership` → `Bloomberg Output` references survive (they were `#REF` before).
  - **Multi-sheet sources keep their original sheet names** (`Ownership`, `Bloomberg Output`) instead of the `<skill>-<sheet>` prefix; the prefix both cluttered the self-describing names and (by renaming) broke the cross-sheet references. Single-sheet sources are still named after the producing skill.
  - The combined workbook is **stamped with the INFOR brand theme** (`templates/INFORFG.thmx`) on both backends — `ApplyTheme` under COM, `loaded_theme` under openpyxl — so it carries INFOR colours/fonts even on a blank base or the openpyxl fallback. Best-effort: a missing/invalid theme never loses the merged workbook.
  - `Workbooks.Open` is hardened: a `None` return is only treated as success when the open count actually rose (else it raises and the caller falls back), with one short retry for transient COM / file-lock races. (`workbook-aggregator` SKILL.md updated; tab-naming test rewritten, cross-sheet-ref-preservation and theme tests added)
- **Market-entry (acquisition-target) tables are clamped to 5.71".** After the cells are filled, each market-entry table is resized to a fixed 5.71" total — graphic-frame extent plus proportionally scaled row heights, mirroring dragging the table's resize handle in PowerPoint — so analyst content longer than the template placeholders can't run the table off the slide. Applied to every market-entry slide, after content (not before). (`scripts/pitch_deck_assembler.py`, `deck-assembler` SKILL.md; test added)

### Added
- **`templates/INFORFG.thmx`** — the INFOR brand theme, shipped so the workbook aggregator can stamp it on the combined workbook.

### Bumped
- marketplace, plugin manifest, pyproject, README status line, and all shipped skill frontmatter versions to `0.5.10`.

## [0.5.9] — 2026-06-05

Pitch-deck and cap-table refinements requested by the analyst, plus a fix for an earnings-update regression introduced in 0.5.8.

### Changed
- **Updated cap table template.** `templates/INFOR Cap Table Template.xlsx` refreshed with the analyst's formatting changes; the `Cap with Links` sheet, all input cells (`F3`/`F5`/`F7`/`F16`/`D47`/`D48`), and the CapIQ array formulas are unchanged. The CapIQ very-hidden `__snloffice` helper sheet — re-introduced when the file was saved with the add-in active — is stripped again per repo convention (lossless, via Excel COM). (`templates/INFOR Cap Table Template.xlsx`)
- **Workbook aggregator merges off the cap table and wires cross-tab links.** The COM merge now opens the `captable` workbook as the **base** (saved as the combined file) instead of starting from a blank workbook, so the cap table's theme, formatting and CapIQ links survive intact — the blank-base merge shifted colours and could drop CapIQ links, which made the combined file hard to format and link. The cross-workbook sheet copy is also made reliable: the destination workbook is activated before each copy (Excel otherwise silently copies into a *new* workbook — a no-op that left earlier-version combined files missing tabs), each sheet is copied individually with an append check, and the `Workbooks.Open`-returns-`None` pywin32 quirk is handled. A post-merge **relink** pass then rewrites the skills' standalone scalar handoffs into live cross-tab formulas: cap table `D47`/`D48` (LTM Revenue / Adj. EBITDA) → `='ltm-metrics'!B<bridge-total>*F7` (located by the `(=) LTM Revenue` / `(=) LTM Adj. EBITDA` labels, since the bridge rows are dynamic), and ownership `F35` (% denominator) → `='captable'!F17*1000000` (cap table millions → ownership full units). The skills still write plain values standalone, so each workbook stays valid on its own and the deck still renders. (`scripts/workbook_aggregator.py`, `workbook-aggregator` SKILL.md; tests added)
- **Acquiror-considerations mitigants may be a short sentence.** The per-mitigant length cap rises from 90 → 160 chars and the `pitch-content` guidance now asks for one very short sentence per mitigant (still exactly three per row). (`scripts/schemas/pitch_deck_content.py` + regenerated JSON schema, `pitch-content` SKILL.md; test added)
- **Insider-ownership slide follows the Financial Summary slide.** In the pitch it previously sat after the market-entry (acquisition-target) slides; the `INFOR Slide Library.pptx` slide is reordered to immediately follow Financial Summary (before Considerations / Mitigants), and `slide_library_registry`, `pitch_deck_wireframe` and `pitch_deck_assembler` follow. The ownership picture insertion is now at a fixed deck index, independent of the market-entry slide count. (`templates/INFOR Slide Library.pptx`, `scripts/slide_library_registry.py`, `scripts/pitch_deck_wireframe.py`, `scripts/pitch_deck_assembler.py`; tests updated)
- **Default of 8 market-entry (acquisition) targets.** When the analyst doesn't specify how many targets they want, the pitch defaults to 8 (4 market-entry slides, two per slide): `pitch_deck_wireframe` defaults the count to 8, `pitch-content` drafts 8 by default (max 8), and the pitch plan gains an optional `market_entry_target_count` input so a specific count can still be requested. (`scripts/pitch_deck_wireframe.py`, `plans/pitch.yaml`, `pitch-content` + `pitch-wireframe` SKILL.md; test updated)

### Fixed
- **Earnings-update deck dropped its Contact slide (regression from 0.5.8).** Inserting the ownership slide into the 16-slide library in 0.5.8 shifted the disclaimer/contact closers to indices 14/15, but the earnings assembler's `_KEEP_LIBRARY_INDICES` stayed `(0, 6, 7, 13, 14)` — so earnings decks shipped as `[Cover, Overview, Earnings Summary, Ownership, Disclaimer]`, with a stray Ownership slide and no Contact slide. Corrected to `(0, 6, 7, 14, 15)`. (`scripts/earnings_update_assembler.py`; regression test added)

### Bumped
- marketplace, plugin manifest, pyproject, README status line, and all shipped skill frontmatter versions to `0.5.9`.

## [0.5.8] — 2026-06-03

Adds an **insider-ownership slide** to the pitch deck for Canadian public targets, sourced from a SEDI "Insider Information by Issuer" report. The earnings-update flow is unchanged.

### Added
- **`ownership` skill + `ownership_workbook.py`.** New skill that parses an analyst-attached SEDI PDF, keeps only current insiders (`Ceased to be Insider: Not Applicable`), sums each one's **common shares** (multiple registered-holder tranches are written as an in-cell `=a+b+c` sum, never hand-totalled), records the latest common-share date, and builds the adjusted `First Last (Role)` name (role from the relationship code + the company site / LinkedIn). `build_ownership_workbook(...)` fills the template's Select-Insiders block (rows 39-65: B/F/G/J) and the `F35` % denominator. SEDI is Canadian-only and **cannot be auto-fetched** (Radware/ShieldSquare bot wall) — the analyst downloads the PDF; see the skill's `references/sedi-extraction.md`. (`skills/ownership/`, `scripts/ownership_workbook.py`; tests added)
- **Ownership template.** `templates/INFOR Ownership Template.xlsx` — the display block `B4:G17` ranks the top 12 insiders by common shares (`LARGE`/`XLOOKUP`); the right "Institutions" side is Bloomberg-sourced and out of scope here.

### Changed
- **Slide library is now 15 entries (was 14).** The updated `INFOR Slide Library.pptx` inserts the **Ownership** slide before the static disclaimer/contact closers; `slide_library_registry` and `pitch_deck_wireframe` register `insider-ownership`. A pitch deck with one market-entry slide is now 15 slides (8 targets → 18). (`scripts/slide_library_registry.py`, `scripts/pitch_deck_wireframe.py`, `pitch-wireframe` SKILL.md; tests updated)
- **Pitch plan gains an `ownership` stage.** It runs **after `captable`** (so `F35` can be sourced from the cap table's Section VII basic shares) and **before `deck`**; the deck stage receives `ownership_workbook_path` and the combined `pitch-<codename>.xlsx` folds in an `ownership` tab. When the target is not a Canadian issuer or no SEDI PDF is attached, the stage emits a null workbook and the deck-assembler leaves the slide's placeholders — the rest of the deck still assembles. (`plans/pitch.yaml`; test updated)
- **deck-assembler inserts the insider-ownership picture.** `assemble_pitch_deck` gains `ownership_workbook_path`; when supplied it pastes `Ownership!B4:G17` into the ownership slide's left `Rectangle 1` "Insiders" placeholder (`slide_index = _MARKET_ENTRY_SLIDE_INDEX + n_market_entry`), mirroring the slide-7 cap-table insertion. The cap-table-specific `insert_cap_table_into_placeholder` is replaced by a single generic `insert_excel_into_placeholder(sheet_name, source_range, placeholder_name, slide_index)` used by both the cap-table and ownership insertions (earnings + pitch assemblers). (`scripts/pitch_deck_assembler.py`, `scripts/earnings_update_assembler.py`, `scripts/excel_to_powerpoint.py`, `deck-assembler` + `excel-to-powerpoint` SKILL.md; tests added)
- **Pre-cleaned the ownership template.** As received it carried 9 dead **external links** and ~5,000 legacy **defined names** (Lotus-era `={#N/A,...}` array literals) that openpyxl re-serializes into a workbook Excel refuses to open — silently breaking the picture render (a plain copy opens; the openpyxl round-trip does not). Both are stripped once from the shipped `templates/INFOR Ownership Template.xlsx` (neither is referenced by the live display formulas; 132 KB → 44 KB), and a regression test guards that the shipped template stays clean. (`templates/INFOR Ownership Template.xlsx`, test added)

### Bumped
- marketplace, plugin manifest, pyproject, README status line, and all shipped skill frontmatter versions to `0.5.8`.

## [0.5.7] — 2026-06-02

Pitch-deck fixes surfaced by the Project PRL2 review. Slides 1–6, 8, 10, 11, 16, 17 were already correct and are unchanged; the earnings-update behaviour is untouched (it remains the reference implementation). All three items below change the pitch deck output.

### Fixed
- **Pitch slide 7 cap table had no LTM Revenue / Adj. EBITDA rows.** The pitch plan ran `captable` with no LTM inputs, so the cap table's LTM valuation cells (`D47`/`D48`) fell back to CapIQ formulas and rendered empty / `#NAME?`. The pitch plan now runs an `ltm-metrics` stage **before** `captable` and feeds its `ltm_revenue` / `ltm_adj_ebitda` totals into `D47`/`D48` — mirroring the earnings-update plan — so the slide-7 picture (range `B15:F40`, already identical to the earnings overview) shows both LTM rows and the EV multiples resolve. The pitch flow now requires two new analyst inputs, `reporting_quarter` and `comparison_quarter`, and the LTM workbook is folded into the combined `pitch-<codename>.xlsx`. (`plans/pitch.yaml`; test updated)
- **Pitch slide 9 (Considerations / Mitigants) rendered too small.** The risk/mitigant table hardcoded a 9 pt header and 8 pt body; the library uses 12 pt / 10 pt. Now driven by named constants (`_RISK_HEADER_SIZE = 12`, `_RISK_BODY_SIZE = 10`) so the table matches the library. (`pitch_deck_assembler.py`; test added)

### Changed
- **Pitch market-entry logo boxes name the company.** Each populated target column's logo box read the generic template string `[Placeholder for Logo]`; it now reads `[<target> Logo]` (e.g. `[Kueski Logo]`) so the analyst knows which logo to drop in. `MarketEntryTarget` gains an optional `name` field (the assembler falls back to a generic `[Company Name Logo]` when it is absent), and `pitch-content` populates it from each target's heading. The unused logo box on an odd final slide is still blanked. (`schemas/pitch_deck_content.py` + regenerated JSON schema, `pitch_deck_assembler.py`, `pitch-content` SKILL.md; tests updated + added)

### Bumped
- marketplace, plugin manifest, pyproject, README status line, and all shipped skill frontmatter versions to `0.5.7`.

## [0.5.6] — 2026-06-02

Repository cleanup, earnings-update fixes surfaced by Test #5 (OTEX), and pitch-deck fixes surfaced by the Project PRL1 review. The cleanup is behaviour-neutral; the **Fixed** / **Changed** items below change the earnings-update and pitch deck output.

### Fixed
- **Pitch exec-summary (slide 2) rendered blue with broken bullets.** `write_bulleted_shape` harvested each bullet level's run `rPr` but never re-applied it, so new runs carried no `<a:solidFill>` and inherited the placeholder list-style `defRPr` — navy `1B2759`, which reads as blue. And `_harvest_bullet_templates` keyed templates one-per-paragraph after a marL sort, so the placeholder's leading `marL=0` `buNone` spacer became "level 0" and main bullets lost their square glyph (falling back to flat `set_text`). The harvest now keys by *distinct bullet level* (skipping `buNone`/glyph-less spacers) and each new run grafts the harvested `rPr`, so body copy keeps the template's dark colour with square main / dash sub bullets. (`pptx_helpers.py`; tests added)
- **Pitch slide 7 cap table was never pasted and the footnote currency was a literal token.** The `captable` stage produced the workbook but nothing inserted it for the pitch deck. `assemble_pitch_deck` now pastes the cap table into slide 7's `Rectangle 3` placeholder (range `B15:F40`, via the same `excel-to-powerpoint` insertion the earnings deck uses) whenever `captable_workbook_path` is supplied, and replaces the `[x]$MM` footnote currency-letter token with the cap table's output currency read from `F5` (`US`/`C`) on **every figure footnote** — slide 7, slide 11 (key investment highlights), and each market-entry slide — so no footnote ships the literal `[x]`. (`pitch_deck_assembler.py`, `excel-to-powerpoint` + `deck-assembler` SKILL.md)
- **Pitch financial-summary tiles (slide 8) showed value-laden strings.** Tiles read e.g. `FY2025 Revenue: US$589.8MM (+31% YoY)`. `financial_metric_labels` now rejects digits, currency/percent tokens, and colons (and over-long phrases), so tiles carry the metric NAME only — the chart shows the amount. (`schemas/pitch_deck_content.py`, `pitch-content` SKILL.md; test added)
- **Pitch market-entry table (slides 12+) had wrong-coloured labels, blank rows, and capped at 2 targets.** `_fill_market_entry_targets` hardcoded `size_pt=8` (labels came out black/8 pt instead of the library's white/11 pt, values 8 instead of 10), and only 7 row labels against a 13-row table left 5 rows blank. The fixed 12-row structure (Overview / HQ / Year Founded → 7 consistent industry metrics → Scale KPIs / Strategic Rationale) now fills every data row, with white 11 pt labels and 10 pt values; the unused column + logo are blanked on an odd final slide. (`pitch_deck_assembler.py`, `schemas/pitch_deck_content.py`)
- **LTM bridge totals were stored as formulas, not text labels.** The bridge result row's label was written as `"= LTM Revenue"` / `"= LTM Adj. EBITDA"`; the leading `=` made openpyxl store the cell as a formula, which Excel rendered as `=@LTM Revenue`. Now written `"(=) LTM Revenue"` — a plain string mirroring the `(+)` / `(−)` component rows. (`ltm_metrics.py`; tests updated.)
- **Cap-table picture showed blank cells down through Enterprise Value.** The Excel→PowerPoint renderer snapshotted the cap-table range without recalculating; openpyxl drops every formula's cached value on save and the template is `calcMode="manual"`, so the EV cascade loaded blank. The renderer now forces `excel.CalculateFull()` (the programmatic F9) before capture. Because a recalc invalidates an *invisible* Excel instance's render buffer (yielding a blank picture), the instance now runs **visible but parked far off-screen**, captures with `CopyPicture(xlScreen)`, and retries with backoff to ride out Office-clipboard contention from any other open Excel. The `B15:F40` picture range is all in-workbook math (no CapIQ functions), so the recalc is safe. (`excel_to_powerpoint.py`)
- **Cap-table picture was still blank under Cowork/Linux.** The fix above only covered the Windows Excel COM path; the headless **LibreOffice** fallback (`_libreoffice_range_to_png`) had no recalc step, so the same openpyxl-stripped, manual-calc formulas printed 0/blank there too. The PDF conversion now runs with a self-contained throwaway LibreOffice profile (`-env:UserInstallation` + a `registrymodifications.xcu` that sets `OOXMLRecalcMode=0` = "Always recalculate"), so formulas recompute on load without touching the user's global profile. CapIQ `_xll.*` cells degrade to `n/a` (unknown to LibreOffice, no add-in) without crashing the recalc; the LTM column and all arithmetic still populate. New `test_excel_to_powerpoint.py` covers the recalc. No new pip dependency (soffice + the existing `pypdfium2`). (`excel_to_powerpoint.py`, `excel-to-powerpoint` SKILL.md)
- **Long KPI tile names overflowed the metric tiles.** A name like "Cloud Services & Subscriptions Rev." wrapped to a second line at the template 9 pt, pushing the 3-line tile past its bounds (PowerPoint ignores scale-less autofit on open). The assembler now shrinks only the name font for over-long names — 8 pt at 25–32 chars, 7 pt beyond — leaving the value at 12 pt so values stay uniform across tiles. (`earnings_update_assembler.py`)
- **Company-overview bullets overflowed into the "LTM Revenue Breakdown" header.** The same scale-less-`<a:normAutofit/>` blind spot let an over-budget block render full size and spill into the pie title. The assembler now resizes `TextBox 9` to the band above the LTM header and, when the copy would still overflow, writes an explicit `fontScale` so the shrink actually happens on open. The `earningsupdate-content` overview budget tightened to 6–8 bullets / 560–820 chars (from 6–10 / 650–1,050). (`earnings_update_assembler.py`, `earningsupdate-content` SKILL.md)

### Removed
- **Dead code.** Deleted `clone_slide` and `fmt_broker_value` (`pptx_helpers.py`), `record_insertion_intent` (`excel_to_powerpoint.py`), and `get_entry` (`slide_library_registry.py`) — none had a production caller.
- **Unadopted typed manifest.** Removed the `SkillManifest` / `SideEffectSpec` schema and `scripts/schemas/json/skill_manifest.schema.json` (no skill ever declared one). `InputSpec` / `OutputSpec` — still used by `Plan` — moved into `plan.py`; `skill_manifest.py` and `test_skill_manifest.py` deleted.
- **Dropped deliverables.** `DeliverableType` trimmed to `pitch` / `earnings-update` / `overview` / `one-off-skill`; `cim`, `teaser`, `fairness-opinion`, and `valuation` removed from the deal-init G7 prompt, conductor, `plan-schema.md`, and the marketplace description.
- **Empty placeholder.** Removed `templates/slide-library/` (the per-entry library is future work; today the library is the single `INFOR Slide Library.pptx`).

### Changed
- **Pitch market-entry now spans multiple slides natively.** `market_entry_targets` accepts up to 8 targets (was capped at 2) and `market_entry_row_labels` is a fixed 12-row contract (schema-enforced). The wireframe emits `ceil(count/2)` market-entry slides when the count is known (`market_entry_target_count` input), and the assembler clones the library's single market-entry slide to the true count from the content bundle — two targets per slide, titled `(N of M)` — regardless of the plan's count. New shared `clone_slide_after` helper in `pptx_helpers.py` (cloning before the earnings-slide delete to avoid duplicate slide part names). (`pitch_deck_wireframe.py`, `pitch_deck_assembler.py`, `pptx_helpers.py`, `schemas/pitch_deck_content.py`, `slide_library_registry.py`, `pitch-content` + `pitch-wireframe` SKILL.md; tests added)
- **Standardised plan naming.** Renamed `plans/pitch-library-poc.yaml` → `plans/pitch.yaml` so the conductor's `plans/<deliverable>.yaml` resolution works for `pitch`. Added a stub `plans/overview.yaml` that registers the overview deck (intentionally references a not-yet-built skill so a premature run fails fast).
- **Test collection fixed.** Moved `test_pptx_helpers.py` and `test_shell_helpers.py` from `scripts/` into `scripts/tests/` so the configured `testpaths` actually collects them — these 42 tests were silently uncollected. Made two pre-existing assertions Windows-portable (`test_deal_context` absolute-path, `test_run_log` Path serialisation) and the `find_template.sh` test independent of the bash/Python temp-dir namespace split.
- **Docs.** Rewrote `README.md` (was frozen at "Phase 0 — no skills implemented yet"); fixed stale `-infor` skill names and the empty-slide-library description in `CLAUDE.md`, the `./infor-workflows` import-path fallback in `pptx_helpers.py`, and the `ltm-revenue`→`ltm-metrics` tab/skill-key examples in `workbook-aggregator`.

### Bumped
- marketplace, plugin manifest, pyproject, README status line, and all shipped skill frontmatter versions to `0.5.6`.

## [0.5.5] — 2026-05-29

### Changed
- **LTM metrics now feed the cap table.** The earnings-update plan runs `ltm-metrics` **before** `captable` (previously parallel siblings). `ltm-metrics` emits its LTM revenue and LTM Adj. EBITDA bridge totals (`ltm_revenue` / `ltm_adj_ebitda`, in millions, filing reporting currency) as typed stage outputs; `captable` reads them and writes the cap table's LTM valuation column — `D47` (Revenue) and `D48` (Adj. EBITDA) — as `=<value>*F7` so the cap table's FX rate converts them into the output currency (F5). New `bridge_total()` helper in `ltm_metrics.py` mirrors the workbook's bridge formula for the handoff.
- **Revised `INFOR Cap Table Template.xlsx`** swapped in: `D47`/`D48` ship empty (the old CapIQ `IQ_REV`/`SP_EBITDA` LTM formulas removed) for the skill to populate. The CapIQ `__snloffice` helper tab was stripped again (it had crept back into the revised file), keeping the standalone workbook clean and consistent with v0.5.3.
- **`captable` Step 6b** added: populate `D47`/`D48` from the `ltm-metrics` handoff (`=<value>*F7`, blue font), or restore the CapIQ `IQ_REV`/`SP_EBITDA` fallback formulas when no LTM values are supplied (direct `/captable` invocation or a plan without an `ltm-metrics` stage), so EV/Revenue and EV/Adj. EBITDA multiples still resolve.
- **Cap-table picture range widened `B15:F36` → `B15:F40`** so the overview-slide picture now also shows the LTM/forward Valuation Metrics rows (EV/Revenue, EV/Adj. EBITDA). Updated `earnings_update_assembler._CAP_TABLE_RANGE`, `excel_to_powerpoint.insert_cap_table_into_placeholder` default `source_range`, and the `deck-assembler` / `excel-to-powerpoint` skill docs.

### Bumped
- marketplace, plugin manifest, pyproject, and all shipped skill frontmatter versions to `0.5.5`.

## [0.5.4] — 2026-05-28

### Changed
- **`ltm-revenue` skill renamed to `ltm-metrics`.** Directory `skills/ltm-revenue`→`skills/ltm-metrics`, helper `scripts/ltm_revenue.py`→`scripts/ltm_metrics.py` (`build_ltm_revenue_workbook`→`build_ltm_metrics_workbook`, `RevenueSegment` kept, new `BridgeComponent`), the earnings-update plan stage id/skill `ltm-revenue`→`ltm-metrics`, and the workbook-aggregator tab key `ltm-revenue`→`ltm-metrics`. Output file is now `<Company> - LTM Metrics.xlsx` on a `LTM Metrics` tab. Updated all references in `plans/`, tests, `workbook_aggregator.py`, `earnings_update_wireframe.py`, CLAUDE.md.
- **LTM metrics tab now stacks three blocks** on one sheet: (1) the existing LTM revenue segment overview, then a spacer row, (2) a new **LTM revenue bridge** that derives the LTM total as `FY + current-year YTD − prior-year YTD` (flexible additive/subtractive component list, total via a cell formula), then a spacer, (3) an **LTM Adj. EBITDA bridge** (falls back to `LTM EBITDA` when no Adjusted figure is disclosed) — bridge only, no segment overview.
- **Earnings-update plan passes `comparison_quarter` to the `ltm-metrics` stage** so the bridge has the prior-year period.

### Added
- **Up-front ask for the statements needed to compute LTM.** The deal-init G7 filings prompt now tells the analyst that LTM deliverables (earnings update, valuation) also need the prior full fiscal year's statements/MD&A — `LTM = FY + current YTD − prior YTD` — and clarifies the cap table is still built off the most recent statement. The `ltm-metrics` SKILL.md gains an "Inputs you must have" section that asks for the specific missing period before computing. The `captable` Step 1 gains a note: build the cap table off the *most recent* reported statement only, never the older FY attached for the LTM bridge.

### Bumped
- marketplace, plugin manifest, pyproject, and all shipped skill frontmatter versions to `0.5.4`.

## [0.5.3] — 2026-05-28

### Changed
- **Revised `INFOR Slide Library.pptx`** swapped in (one small formatting refinement).
- **Cap table now sources FX rate and share price from the web.** The `INFOR Cap Table Template.xlsx` leaves F7 (FX Rate) and F16 (Share Price) empty, each carrying a cell comment with the live CapIQ formula. The `captable` skill gains Step 3b: it fills F7 (Output-currency per Input/filing-currency) and F16 (share price in the Output currency F5) from the web as blue hardcoded values, leaving the commented CapIQ formula intact so the analyst can refresh live in Excel. FX-direction guidance throughout the skill is now definitive (multiply filing-currency figures/strikes by F7) instead of hedging on F7's orientation.
- **Cap-table picture range widened `B15:F31` → `B15:F36`** so the overview-slide capitalization summary now also includes the Financial Metrics section. Updated `earnings_update_assembler._CAP_TABLE_RANGE`, `excel_to_powerpoint.insert_cap_table_into_placeholder` default `source_range`, and the `deck-assembler` / `excel-to-powerpoint` skill docs.

### Fixed
- **Workbook aggregator no longer emits the garbled `captable-__snloffice` tab.** CapIQ's Excel add-in stows formula metadata in a very-hidden `__snloffice` sheet; copied verbatim it surfaced as a CJK-looking tab in the combined workbook. Both the COM and openpyxl merge backends now skip CapIQ helper sheets (`__snl*`), and a single-content-sheet source (cap table) collapses to a tab named just `captable`.

### Bumped
- marketplace, plugin manifest, pyproject, and all shipped skill frontmatter versions to `0.5.3`.

## [0.5.2] — 2026-05-28

### Added
- **`workbook-aggregator` skill** + `workbook_aggregator.py` helper: the final consolidation stage of a deliverable run. Merges every companion `.xlsx` produced during the deliverable into a single combined workbook named `<deliverable>-<deal name>.xlsx` (`earningsupdate-Project Atlas.xlsx`, `pitch-Project Atlas.xlsx`), with each producing skill contributing its sheets under a tab named after that skill (single-sheet sources → one tab named after the skill, multi-sheet → `<skill>-<sheet>`). Preserves formulas, CapIQ links, charts, and formatting via Excel COM on Windows; falls back to a best-effort openpyxl merge off-Windows. The individual source workbooks are deleted once the merge succeeds.
- Wired a final `workbook-aggregation` stage into `plans/earnings-update.yaml` (combines `captable` + `ltm-revenue`) and `plans/pitch-library-poc.yaml` (combines `captable` + an optional `comps` workbook). It runs after the `deck` stage so the deck-assembler can still read the standalone cap-table workbook before it is folded in and removed.

### Bumped
- marketplace, plugin manifest, pyproject, and all shipped skill frontmatter versions to `0.5.2`.

## [0.5.1] — 2026-05-28

POC test-round #3 feedback on the earnings update (Open Text Corporation).

### Changed
- **Skill names standardized.** Dropped the trailing `-infor` suffix from every skill directory and `name:` frontmatter; deliverable-specific skills are now prefixed by deliverable. Renames: `captable-infor`→`captable`, `earningsupdate-wireframe-infor`→`earningsupdate-wireframe`, `earningsupdate-content-infor`→`earningsupdate-content`, `excel-to-powerpoint-infor`→`excel-to-powerpoint`, `ltm-revenue-infor`→`ltm-revenue`, `pitch-content-infor`→`pitch-content`, `pitch-wireframe-infor`→`pitch-wireframe`. Updated all references in `plans/`, tests, `pitch_deck_assembler.py`, the conductor plan-schema reference, and CLAUDE.md (which gains a "Skill naming" convention note).
- **Removed the standalone `earningsupdate-infor` skill.** It was a thin orchestrator duplicating the `earnings-update.yaml` plan; earnings updates now run solely through the conductor plan's decomposed stages.
- **Financial-highlights tiles now carry formatted dollars.** `earnings_update_assembler._fmt_mm` renders KPI tile values as `$XMM` (whole millions, no decimals) and auto-converts a billion or more to `$X.XB` (one decimal): `1283 → $1.3B`, `493 → $493MM`. Percent tiles are left untouched. Content supplies a plain integer in MM; the assembler adds the `$` and suffix.
- **Revised shared slide library swapped in** (`INFOR Slide Library.pptx`): standardized source lines and a currency-letter footnote token `All figures in [x]$MM`. The assembler now preserves the library's footnote and substitutes the `C` / `US` letter (`_fill_footnote` / `_currency_letter`) instead of re-hardcoding the source string on the overview and earnings-summary slides.
- **Content house-style for figures.** `earningsupdate-content-infor` gains a "Number & currency formatting" section: never write a currency code inline (plain `$`, currency is footnoted), `$XMM` / `$X.XB` in prose, plain-integer MM for KPI tiles, and a ~30-char abbreviation rule for KPI tile labels (`Cloud Services & Subscriptions Rev.`).
- **Overflow QA is now mandatory** in `deck-assembler` with explicit checks: slide-2 overview must not overlap the "LTM Revenue Breakdown" title, slide-3 highlight tiles must not clip/wrap to a third line, and no figure may read as `US,…` / `C,…`.

### Fixed
- Root-caused the `US$1,057.8 → US,057.8` corruption to a regex `$1` backreference substitution (PowerShell `-replace` / JS `.replace`) eating `$<digit>` during content-JSON generation — the assembler preserves `$` correctly. Added an explicit warning in `earningsupdate-content-infor` to write values as literals and never build the JSON via regex substitution.

### Bumped
- marketplace, plugin manifest, pyproject, and all shipped skill frontmatter versions to `0.5.1`.

## [0.5.0] — 2026-05-28

### Changed
- **Earnings-update deck now clones the shared INFOR Slide Library** instead of the standalone `INFOR Earnings Update Template.pptx`. `earnings_update_assembler` keeps library slides 1, 7, 8, 14, 15 (cover, overview, earnings summary, disclaimer, contact), deletes the rest tail-first, and fills them. `plans/earnings-update.yaml` and `deck-assembler` now resolve `template_name` to `INFOR Slide Library.pptx` for both deliverables.
- **Overview slide (library slide 7)**: title is now "Introduction to {company}"; the cap-table placeholder is renamed "Capitalization Summary" and inserted into `Rectangle 3` over range `B15:F31` (was `Rectangle 4` / `B13:F31`); the lower-left quadrant reserves an LTM revenue pie placeholder. Company-overview bullet budget tightened to 650–1,050 chars (was 1,200–1,500) and 6–10 bullets (was 7–12) so copy no longer overflows behind the pie.
- **Earnings-summary slide (library slide 8)**: metric boxes now carry the rounded value (whole MM, no decimals) plus the metric name; the reporting/comparison period prints only in the mid-blue bar below the "Financial Highlights" title, not inside the boxes. Broker Reported / Bloomberg Estimate / Variance values are now prefixed with `$`.
- Bumped marketplace, plugin manifest, pyproject, and all shipped skill frontmatter versions to `0.5.0`.

### Added
- **`ltm-revenue-infor` skill** + `ltm_revenue.py` helper: emits a standalone LTM revenue breakdown `.xlsx` (segmented by service/product line, falling back to geography) as the companion to the overview slide's pie placeholder. Added as a sibling `ltm-revenue` stage in `plans/earnings-update.yaml`. The stage builds no chart and does not touch PowerPoint.
- **Slide → PNG overflow QA**: new `slide_render.py` (`render_deck_to_png`) renders overflow-prone slides via PowerPoint COM on Windows / LibreOffice headless elsewhere. `deck-assembler/SKILL.md` gained an Overflow QA section directing the agent to render slides, inspect for text overflow, and shrink via `enable_normal_autofit` until clean.

### Fixed
- `pptx_helpers.delete_slide` now drops the presentation-part relationship (`prs.part.drop_rel(rId)`) in addition to removing the `sldId`, eliminating orphaned slide parts and duplicate zip part-name warnings.
- `pitch_deck_assembler` deletes the earnings slide inserted at library index 7 on open, so the 14-slide pitch indices stay valid after the library grew to 15 slides.

## [0.4.5] — 2026-05-28

### Added
- **Two new slide-library entries**, expanding the INFOR Slide Library POC from 12 to 14 slides. `key-investment-highlights` (slide 11, four numbered quadrants) and `market-entry-targets` (slide 12, two-target comparison table) are inserted before Disclaimer/Contact, which move to slides 13/14.
- **Full fill wiring** for both slides. New `InvestmentHighlight` and `MarketEntryTarget` schema models plus optional `PitchDeckContent` fields (`investment_highlights`, `investment_highlights_tagline`, `market_entry_market`, `market_entry_row_labels`, `market_entry_targets`). `pitch_deck_assembler` fills the highlight quadrants/tagline and the market-entry title/table when content is supplied; target logos remain deferred image placeholders. Fields are optional, so decks that omit them keep the blank placeholders.

### Changed
- `slide_library_registry`, `pitch_deck_wireframe`, and `pitch_deck_assembler` now expect 14 slides; `pitch-content-infor` / `pitch-wireframe-infor` / `deck-assembler` SKILL docs updated.
- Bumped marketplace, plugin manifest, pyproject, and shipped skill frontmatter versions to `0.4.5`.

## [0.4.4] — 2026-05-28

### Added
- **Non-Windows cap-table renderer.** Cap-table picture insertion now falls back to LibreOffice headless on non-Windows hosts (including Claude Cowork's Linux sandbox). `_render_range_to_png` in `excel_to_powerpoint.py` dispatches: Excel COM on Windows, LibreOffice → PDF → `pypdfium2` → PNG elsewhere. Requires `soffice`/`libreoffice` on PATH and the new `pypdfium2` runtime dep (`sys_platform != 'win32'`).
- **Shrink-on-overflow autofit** on slide 2 `TextBox 16` (company overview) and slide 3 `TextBox 1067` (business updates). New helper `enable_normal_autofit(shape)` in `pptx_helpers.py` writes `<a:normAutofit/>` into the text frame's `bodyPr` so PowerPoint scales the font down at render time when the analyst-written copy would otherwise overflow the section divider.
- Tests for the broker-label strip, the autofit XML, and a soffice-gated LibreOffice fallback test that skips on Windows / when LibreOffice isn't installed.

### Changed
- **Broker-estimates row labels** in slide 3 are now stripped of redundant `(US$MM)` / `(C$MM)` / `(MM)` suffixes via the existing `_strip_currency_unit` regex. Non-MM markers such as `EPS (US$)` are preserved so per-share metrics keep their explicit unit. The broker table header already prints "Figures in {currency_short}" — repeating it per row wasted horizontal space.
- `earningsupdate-content-infor/SKILL.md` content rules now require plain broker labels (no inline MM suffix) and abbreviated management-quote roles (`CEO`, `CFO`, `Interim CEO`, `Executive VP and CFO`) — not the spelled-out "Chief Executive Officer".
- Bumped marketplace, plugin manifest, pyproject, and shipped skill frontmatter versions to `0.4.4`.

## [0.4.3] — 2026-05-27

### Changed
- **Cap-table insertion now renders the Excel range as a picture** instead of building an editable PowerPoint table. `insert_cap_table_into_placeholder` opens the workbook via Excel COM, copies `Cap with Links!B13:F31` as a metafile, pastes it into a temporary `ChartObject`, exports the chart to PNG via `Chart.Export`, and inserts the PNG into the slide 2 `Rectangle 4` placeholder stretched to the placeholder's width/height. The chart-export round-trip avoids the system clipboard so Excel stays invisible (no flashing window). Removes ~50 lines of cell-by-cell extraction logic and the `extract_cap_table_rows` helper.
- Tune the cap-table workbook's column widths and row heights so the source range's natural aspect ratio matches the placeholder (~0.84 w/h); the picture is stretched to fit either way.
- Bumped marketplace, plugin manifest, pyproject, and shipped skill frontmatter versions to `0.4.3`.

### Added
- `pywin32` as a Windows-only runtime dependency (`sys_platform == 'win32'` env marker).

### Notes
- Cap-table insertion now requires Microsoft Excel installed on Windows. On non-Windows platforms the earnings-update assembler test is skipped via `pytest.importorskip`.

## [0.4.2] — 2026-05-18

### Added
- Implemented first real Excel-to-PowerPoint cap-table insertion for the earnings-update deck: the deck stage now accepts `captable_workbook_path` and replaces slide 2's Macabacus placeholder with a PowerPoint table extracted from the cap-table workbook summary.
- Added `openpyxl` and `python-pptx` as explicit runtime dependencies for workbook extraction and deck assembly.

### Changed
- Reordered `plans/earnings-update.yaml` to run `wireframe → content → captable → deck` so the deck stage can consume the cap-table workbook.
- Company-overview content guidance now avoids bold `Header:` formatting except for true product/service segment bullets.
- Earnings-summary KPI labels now strip inline currency-unit suffixes such as `(C$MM)` and rely on the slide footnote/table header for units, avoiding label wrap/overlap in KPI tiles.
- Bumped marketplace, plugin manifest, pyproject, and shipped skill frontmatter versions to `0.4.2`.

## [0.4.1] — 2026-05-16

### Added
- Added inner plugin manifest at `infor-beta/.claude-plugin/plugin.json` so Claude Co-work marketplace installs can validate the plugin declared by the root marketplace.

### Changed
- Bumped marketplace, plugin manifest, pyproject, and shipped skill frontmatter versions to `0.4.1`.

## [0.4.0] — 2026-05-16

### Added
- **Phase 3 slide-library POC** — added `plans/pitch-library-poc.yaml`, a conductor-driven 12-slide pitch/deck proof-of-concept for the canonical `INFOR Slide Library.pptx` template.
- **Slide-library POC template** — copied `INFOR Slide Library.pptx` into `infor-beta/templates/` as the canonical blank deck for this POC.
- **`PitchDeckContent` schema** (`scripts/schemas/pitch_deck_content.py`) plus JSON Schema export for the broad typed deck-content handoff.
- **POC helper modules**: `scripts/slide_library_registry.py`, `scripts/pitch_deck_wireframe.py`, `scripts/pitch_deck_assembler.py`, and `scripts/excel_to_powerpoint.py`.
- **New POC skills**: `pitch-wireframe-infor`, `pitch-content-infor`, and `excel-to-powerpoint-infor`.
- **Deck assembler support for pitch POC** — `deck-assembler` now supports both the earnings-update POC and the 12-slide slide-library pitch POC.

### Changed
- All shipped skill `version:` frontmatter, `marketplace.json`, and `pyproject.toml` bumped to `0.4.0` per single-version policy (E3).

### Notes
- Static INFOR credential, disclaimer, and contact slides are preserved. LTM revenue, financial summary charts, and comps/cap-table placement can remain placeholders in this POC while the typed handoff/assembly foundation is proven.

## [0.3.1] — 2026-05-15

### Fixed
- Added missing `INFOR Cap Table Template.xlsx` to `infor-beta/templates/` so the earnings-update plan's sibling `captable-infor` stage can copy the workbook template on a fresh plugin install.

### Changed
- Bumped marketplace, pyproject, and all shipped skill frontmatter versions to `0.3.1`.

## [0.3.0] — 2026-05-15

### Added
- **Phase 3 earnings-update POC decomposition** — `plans/earnings-update.yaml` now runs four typed conductor stages: `wireframe` → `content` → `deck` → `captable`.
- **`earningsupdate-wireframe-infor`** — new conductor-stage skill that emits a fixed five-slide typed `SlidePlan` for the INFOR Earnings Update Template.
- **`earningsupdate-content-infor`** — new conductor-stage skill that emits a strict typed `EarningsUpdateContent` bundle for the deck stage.
- **`deck-assembler`** — new POC assembler skill scoped only to `INFOR Earnings Update Template.pptx`; consumes `SlidePlan` + `EarningsUpdateContent` and writes the earnings update `.pptx`.
- **`EarningsUpdateContent` schema** (`scripts/schemas/earnings_update_content.py`) plus JSON Schema export (`scripts/schemas/json/earnings_update_content.schema.json`).
- **POC helper modules**: `scripts/earnings_update_wireframe.py` and `scripts/earnings_update_assembler.py`.
- **Minimum template port**: `templates/INFOR Earnings Update Template.pptx` copied from production for the POC.

### Changed
- `SlideEntry.layout_variant` removed from `SlidePlan` per I1/I6; concrete `library_entry_id` is the sole selector.
- All shipped skill `version:` frontmatter bumped to `0.3.0` per single-version policy (E3).
- `marketplace.json` and `pyproject.toml` version bumped to `0.3.0`.

### Notes
- Broader slide-library work remains intentionally out of scope: no registry, no CIM/pitch/teaser library entries, no generalized assembler beyond the earnings-update template.

## [0.2.0] — 2026-05-15

### Added
- **`conductor` meta-skill** (`infor-beta/skills/conductor/`) — orchestrates deliverables end-to-end. Workflow: identify deliverable + codename, run deal-init if new, load the plan YAML, collect plan inputs, dispatch each stage via the `Task` (Agent) tool with file-based input / output handoff under `<deal_dir>/runs/<run-id>/stages/<id>/`, run checkpoint behaviour, emit a `summary.md`. Three reference files: `plan-schema.md`, `stage-envelope.md`, `checkpoint-behaviour.md`.
- **Plan schema** (`scripts/schemas/plan.py`) — `Plan` + `Stage` pydantic v2 models with `plan_inputs: list[InputSpec]`, `stages: list[Stage]`, three `CheckpointMode` values. Stage ids must be unique within a plan.
- **Reference resolver** (`scripts/plan_refs.py`) — pure function `resolve_refs(...)` that walks dicts/lists/strings and resolves `$plan_inputs.<name>`, `$deal.<field>` (dotted), `$stages.<id>.<name>` placeholders. Mid-string interpolation deliberately unsupported.
- **Run-log helpers** (`scripts/run_log.py`) — `make_run_id`, `create_run_dir`, `write_plan_snapshot`, `write_stage_inputs`, `read_stage_outputs`, `write_stage_log`, `write_summary`. Run-id format: `YYYY-MM-DD-<plan-id>-<short-uuid>`.
- **Deal-init helper** (`scripts/deal_init.py`) — `render_init_prompt()` (locked G7 7-field prompt), `save_deal_context()` / `load_deal_context()` / `load_or_locate_deal()` for `<deal_dir>/deal.json` persistence and directory bootstrap (`facts/`, `filings/`, `artefacts/`, `runs/`).
- **earnings-update plan** (`plans/earnings-update.yaml`) — 2-stage pilot composing `earningsupdate-infor` and `captable-infor` as sibling stages. Both stages run with `informational` checkpoint.
- **Ported `earningsupdate-infor`** from `infor-workflows` — version bumped to 0.2.0, helper paths retargeted, conductor-mode handoff block added at the top, Step 8 (companion cap table) marked as skip-when-`$STAGE_OUTPUTS`-set so the conductor composes captable as a sibling rather than the skill calling it inline.
- **Ported `captable-infor`** from `infor-workflows` — same surgical edits.
- **Plan JSON Schema** (`scripts/schemas/json/plan.schema.json`) — exported via `python -m schemas.export`.
- **`PyYAML>=6,<7`** added as runtime dependency for plan loading.
- **CLAUDE.md** "Shared helpers" snippet now includes `Plan`, `Stage`, the deal-init / run-log / plan_refs modules. Phase status table updated.

### Changed
- All shipped skill `version:` frontmatter bumped to `0.2.0` per single-version policy (E3).
- `marketplace.json` version bumped `0.1.0 → 0.2.0`.

### Notes
- v1 conductor executes stages **sequentially** in declaration order. Parallel / DAG (`depends_on`, `parallel_with`) deferred until Phase 3+ when CIM/pitch plans justify the complexity.
- Telemetry (`meta.json`: model, tokens, latency) deferred to Phase 5 per locked decision A3.
- Decomposition of `earningsupdate-infor` into `infor-wireframe` + `infor-deck-writing` + `deck-assembler` deferred to Phase 3 per locked decision H3.

## [0.1.0] — 2026-05-15

### Added
- `pyproject.toml` declaring `pydantic>=2,<3` runtime dep + `pytest` dev dep; Python `>=3.11`.
- Typed I/O contract under `infor-beta/scripts/schemas/` (pydantic v2): `Company`, `Filing` + `FilingType` enum, `SlideEntry` + `SlidePlan`, `DealContext` + `DeliverableType`, `SkillManifest` + `InputSpec` / `OutputSpec` / `SideEffectSpec`.
- JSON-Schema export under `infor-beta/scripts/schemas/json/` (one file per model) + idempotent `scripts/schemas/export.py` exporter (run via `python -m schemas.export`).
- Codename resolver `infor-beta/scripts/codename.py` per locked decision G3: case-preserving display, case-insensitive lookup, path-unsafe-char stripping, collision-aware disambiguation suggestions.
- Ported shared helpers from `infor-workflows`: `pptx_helpers.py` (with brand constants `PALATINO`, `COLOR_UP`, `COLOR_DOWN`), `find_template.sh`, `sanitize_name.sh` (retargeted to `infor-beta` install paths).
- Test suite: 56 schema/codename tests (pytest) + 42 ported helper tests (unittest), all green.
- `CLAUDE.md` "Shared helpers" section now documents the concrete import paths.

## [0.0.1] — 2026-05-14

### Added
- Initial scaffold: README, CLAUDE.md, .gitignore, marketplace.json stub, directory skeleton (`infor-beta/skills/`, `infor-beta/plans/`, `infor-beta/scripts/`, `infor-beta/templates/`, `infor-beta/templates/slide-library/`).
- Canonical architecture record referenced in CLAUDE.md (lives in Obsidian at `Hermes-L1/INFOR Platform Architecture/`, note 12 is authoritative).
