# Contributor brief — infor-beta

This file is loaded automatically when Claude Code opens this repo. It orients you (and future contributors) to the plugin's layout, conventions, and the rebuild context.

## Context: this is a clean-break rebuild

The production INFOR plugin today is **`infor-workflows`** (at `~/Desktop/infor-workflows`, v2.10.0). It keeps shipping unchanged until `infor-beta` supersedes it. This repo is a **fresh start, not a refactor in place** — designed directly for the target architecture with no backwards-compatibility shims.

The canonical architecture record lives in Obsidian at `Hermes-L1/INFOR Platform Architecture/`. Note `12 — Locked Decisions` is authoritative.

## What this plugin will be

A single Claude Code plugin, `infor-beta`, containing:

- A **conductor meta-skill** that consumes a deliverable spec (pitch, earnings update, overview, …) and orchestrates sub-skills via the `Agent` tool.
- **Specialised skills** for data tables (comps, precedents, buyer lists, cap tables), modelling (LBO), writing (deck-writing), wireframing (typed `SlidePlan` output), QA (deckcheck), valuation aggregation (football field), and company / industry profiles.
- A **slide library** of reusable `.pptx` entries (one entry per slide concept × layout combination) that the conductor and `deck-assembler` clone and fill.
- A **typed I/O contract** so skills can compose cleanly without prompt glue.

## Layout

```
.claude-plugin/marketplace.json    Marketplace manifest — points at infor-beta/ as plugin root
infor-beta/
├── skills/                        One directory per skill
│   └── <skill-name>/
│       ├── SKILL.md               Workflow + frontmatter (name, description, version, allowed-tools)
│       └── references/            Progressive-disclosure detail loaded on demand
├── plans/                         Conductor plan YAMLs (earnings-update, pitch, overview)
├── scripts/                       Shared helpers + tests (pptx_helpers.py, find_template.sh, sanitize_name.sh, ...)
├── templates/                     Excel + PowerPoint templates shipped with the plugin
│                                   (INFOR Slide Library.pptx, INFOR Cap Table Template.xlsx).
│                                   A per-entry templates/slide-library/ is future work — not yet created.
README.md
CHANGELOG.md
CLAUDE.md                          ← you are here
```

The nesting (`infor-beta/infor-beta/`) exists because `marketplace.json` points `source: "./infor-beta"`. The inner directory is the actual plugin root. Don't flatten without updating the marketplace manifest.

## Locked architectural decisions (summary)

These are the load-bearing decisions made before any code is written. The full record is in Obsidian note 12.

**Tenancy / scope**
- Single-tenant: INFOR Financial Group only. No multi-firm support.
- English only (no French).
- Production runtime: Claude Desktop. Hermes is Tate's dev environment.

**Reproducibility**
- Directional audit only — the analyst audits the deliverable. **No** `pins.json`, **no** prompt-hash capture, **no** model-version pinning per deal.
- Per-stage run logs + source citations on the artefact are sufficient.

**Orchestration**
- Conductor is a Claude Code meta-skill that dispatches via the `Agent` tool. No standalone service, no MCP migration.
- Medium human-in-the-loop initially: confirmation gates at major stage boundaries. Autonomous later via configuration flip (no code change).
- Three checkpoint modes from day one: `required`, `informational`, `silent`.

**Skill contract**
- Skills emit **typed output** in addition to (or instead of) free-form files. Composing skills consume typed input.
- The wireframe skills (`earningsupdate-wireframe`, `pitch-wireframe`) emit a typed `SlidePlan` (markdown view kept for analyst readability).
- `brand-guidelines` is demoted to a **library** consumed by `deck-assembler` and `deckcheck` (both future skills) — not a stage in any plan.
- Earnings updates are delivered solely via `plans/earnings-update.yaml`; the standalone monolith skill was removed (its work is the decomposed `earningsupdate-wireframe` / `earningsupdate-content` / `captable` / `ltm-metrics` / `deck-assembler` stages).

**Slide library**
- Multiple library entries per slide concept × layout combination (e.g. `company-overview-two-col`, `company-overview-full-bleed`). No parameterised `variants:` field. One `.pptx` per entry.
- Realistic library size: 40–80 entries at v1 maturity. Enumerable via a categorised README — no search UX required until well past ~100.
- Will live in `templates/slide-library/` once built; today the only library file is `templates/INFOR Slide Library.pptx`.

**Data**
- Company facts: analyst-provided at deal-init, verified by WebSearch, cached in `deal.facts/company.json`. CapIQ API connector is a future migration target.
- Filings: analyst attaches in chat; conductor persists to the deal directory. SEDAR/EDGAR auto-fetch is a Phase 4 spike.
- URL allow-lists are **per-skill**, calibrated to each skill's data-quality bar.

**Operational**
- Deal directory: `~/Documents/INFOR Deals/<codename>/`. Outputs land there, not in random `cwd`.
- Concurrent deals (different codenames) are supported — directories are scoped by codename, no collision.
- Single plugin version, all skills in lock-step. Revisit at 20+ skills.

## v1 skill portfolio

**Refactored from old repo:**
`comps`, `precedents`, `buyerslist`, `captable`, `lbo-model`, `deck-writing`, wireframe (typed; realised as `earningsupdate-wireframe` + `pitch-wireframe`), `deckcheck` (QA), `brand-guidelines` (→ library). Earnings update is delivered via `plans/earnings-update.yaml` rather than a standalone skill.

**New skills:**
`deck-assembler` (Phase 3 foundational), `workbook-aggregator` (final-stage workbook consolidation), `ownership` (insider-ownership slide from a SEDI report; Canadian public targets), `valuation` (football field over comps + precedents + LBO), `company-profile-public`, `company-profile-private`, `industry-research` (low priority, late Phase 4).

**Conductor plans:**
`earnings-update.yaml` (decomposed Phase 3 POC), `pitch.yaml` (15-slide slide-library deck), `overview.yaml` (stub — not yet implemented).

**Removed from scope:** management presentations, diligence support, research pipelines.

## Conventions

### Skill naming
Skill directory names carry no trailing `-infor` suffix. A skill that is specific to one deliverable is prefixed by that deliverable: `pitch-*` for pitch-deck skills (`pitch-wireframe`, `pitch-content`), `earningsupdate-*` for earnings-update skills (`earningsupdate-wireframe`, `earningsupdate-content`). General-purpose skills reused across deliverables take a plain name (`captable`, `ltm-metrics`, `deck-assembler`, `workbook-aggregator`, `conductor`). The `name:` frontmatter must equal the directory name.

### Versioning — single plugin version, all skills tied to it
One version for the entire plugin. Any skill change bumps the plugin version, and every skill's `version:` frontmatter equals the plugin version. No per-skill drift. Revisit at 20+ skills.

### Progressive disclosure
Long workflows split into a short `SKILL.md` (workflow + step outline) + `references/<topic>.md` loaded on demand. New skills longer than ~250 lines should follow it.

### Shared helpers (don't re-implement)
Templates, filename sanitization, python-pptx formatting helpers, brand constants, and the typed I/O contract live in `infor-beta/scripts/`. Skills import via `CLAUDE_PLUGIN_ROOT`:

```python
import sys, os
sys.path.insert(0, os.environ.get("CLAUDE_PLUGIN_ROOT", "./infor-beta") + "/scripts")

from pptx_helpers import set_text, write_bulleted_shape, set_cell_text, delete_slide, find_shape
from schemas import (
    Company, Filing, FilingType, SlidePlan, EarningsUpdateContent, PitchDeckContent, DealContext,
    Plan, Stage, CheckpointMode, InputSpec, OutputSpec,
)
from earnings_update_wireframe import build_earnings_update_slide_plan, write_slide_plan
from earnings_update_assembler import assemble_earnings_update_deck
from pitch_deck_wireframe import build_pitch_deck_slide_plan
from pitch_deck_assembler import assemble_pitch_deck
from workbook_aggregator import combine_workbooks
from slide_library_registry import load_slide_library_registry
from codename import resolve, find_existing, disambiguate
from deal_init import render_init_prompt, load_or_locate_deal, save_deal_context, load_deal_context
from plan_refs import resolve_refs
from plan_schedule import compute_waves
from run_log import make_run_id, create_run_dir, write_stage_inputs, read_stage_outputs, write_summary
```

For the bash helpers:

```bash
SANITIZED=$(bash "${CLAUDE_PLUGIN_ROOT:-./infor-beta}/scripts/sanitize_name.sh" "$RAW_NAME")
TEMPLATE=$(bash "${CLAUDE_PLUGIN_ROOT:-./infor-beta}/scripts/find_template.sh" "INFOR Cap Table Template.xlsx")
```

Brand constants are in `pptx_helpers` (`PALATINO`, `COLOR_UP`, `COLOR_DOWN`). JSON-Schema views of every typed contract are emitted to `infor-beta/scripts/schemas/json/` — regenerate with `python -m schemas.export` (idempotent).

### Excel does the math, not the LLM
Arithmetic lives in cell formulas for analyst auditability. Skills write inputs and let the workbook compute.

### Output files
Skills write to the **deal directory** (`~/Documents/INFOR Deals/<codename>/`), resolved by the conductor at deal-init. The cwd default is preserved for ad-hoc skill invocation, but the conductor always sets an explicit output target.

## Phase status

- Phase 0 — Stabilise & document. ✅ shipped 2026-05-14.
- Phase 1 — Deal model + typed I/O contract (`Company`, `Filing`, `SlidePlan`, `DealContext`, `SkillManifest`). ✅ shipped 2026-05-14.
- Phase 2 — Conductor v1 meta-skill + earnings-update plan pilot. Adds `Plan` / `Stage` schemas, `plan_refs` resolver, `deal_init` and `run_log` helpers, ports the earnings-update monolith (later removed in v0.5.1) + `captable`. ✅ shipped 2026-05-15.
- Phase 3 — Earnings-update proof-of-concept decomposition + POC `deck-assembler`. Adds `EarningsUpdateContent`, `earningsupdate-wireframe`, `earningsupdate-content`, template-specific `deck-assembler`, and a four-stage `earnings-update.yaml`. ✅ POC shipped 2026-05-15.
- Phase 3 slide-library POC — 14-slide `INFOR Slide Library.pptx` proof-of-concept with `PitchDeckContent`, `pitch-wireframe`, `pitch-content`, `excel-to-powerpoint`, `pitch-library-poc.yaml`, and generalized deck-assembler support. ✅ POC shipped 2026-05-16; expanded to 14 slides (Key Investment Highlights + Potential Market Entry Targets) 2026-05-28.
- Phase 3 earnings-update onto the shared library (v0.5.0) — earnings-update deck now clones `INFOR Slide Library.pptx` (slides 1, 7, 8, 14, 15) instead of the retired standalone template. Overview slide gains an LTM revenue pie placeholder + "Introduction to {company}" title + "Capitalization Summary" cap table (`Rectangle 3`, `B15:F31`); earnings-summary metric boxes show value+name with the period in the mid-blue bar; broker values carry `$`. Adds `ltm-revenue` (standalone LTM revenue workbook) and `slide_render.py` (PNG overflow QA). ✅ shipped 2026-05-28.
- Workbook aggregation (v0.5.2) — adds the `workbook-aggregator` skill + `workbook_aggregator.py`. A final `workbook-aggregation` stage in both the earnings-update and pitch-library-poc plans merges every companion `.xlsx` into one combined workbook named `<deliverable>-<deal name>.xlsx` (one tab per producing skill), preserving formulas/CapIQ links via Excel COM and replacing the individual source files. Runs after `deck` so the deck-assembler can still read the standalone cap table. ✅ shipped 2026-05-28.
- Cap-table web inputs + aggregator cleanup (v0.5.3) — cap table template now leaves FX (F7) and share price (F16) empty with the CapIQ formula stored as a cell comment; the `captable` skill fills both from the web (Step 3b) and the analyst pastes the commented formula to refresh. Cap-table picture range widened to `B15:F36` (adds the Financial Metrics section). The workbook aggregator now skips CapIQ's very-hidden `__snloffice` helper sheet so the combined workbook no longer carries the garbled `captable-__snloffice` tab. ✅ shipped 2026-05-28.
- LTM metrics expansion + rename (v0.5.4) — the `ltm-revenue` skill is renamed **`ltm-metrics`** (`ltm_revenue.py`→`ltm_metrics.py`, `build_ltm_metrics_workbook`, plan stage + aggregator tab key `ltm-revenue`→`ltm-metrics`). Its single "LTM Metrics" tab now stacks three blocks: the existing revenue segment overview, a new **LTM revenue bridge** (`FY + current-YTD − prior-YTD`, flexible component list), and an **LTM Adj. EBITDA (or EBITDA) bridge** (bridge only). The deal-init G7 filings prompt now asks for the prior full-fiscal-year statements needed for the LTM math, and `captable` Step 1 clarifies the cap table is built off the *most recent* statement only — never the older FY attached for the LTM bridge. ✅ shipped 2026-05-28.
- LTM metrics feed the cap table (v0.5.5) — the earnings-update plan now runs `ltm-metrics` **before** `captable` (no longer parallel). `ltm-metrics` emits its LTM revenue and LTM Adj. EBITDA bridge totals as typed outputs (`ltm_revenue` / `ltm_adj_ebitda`, millions, filing reporting currency; new `bridge_total()` helper); `captable` Step 6b writes them to the cap table's LTM valuation column `D47`/`D48` as `=<value>*F7` (FX-converted to the F5 output currency), falling back to the CapIQ `IQ_REV`/`SP_EBITDA` formulas when no LTM values are supplied. The revised `INFOR Cap Table Template.xlsx` ships `D47`/`D48` empty (old CapIQ LTM formulas removed) and re-strips the `__snloffice` helper tab. Cap-table picture range widened `B15:F36`→`B15:F40` to include the Valuation Metrics rows. ✅ shipped 2026-05-29.
- Cleanup pass (unreleased) — scoped the deliverable set to **3 plans** (`earnings-update`, `pitch`, `overview` stub); renamed `pitch-library-poc.yaml`→`pitch.yaml`; trimmed `DeliverableType` to `pitch`/`earnings-update`/`overview`/`one-off-skill` (dropped cim/teaser/fairness-opinion/valuation). Removed dead code (`clone_slide`, `fmt_broker_value`, `record_insertion_intent`, `get_entry`) and the unadopted `SkillManifest`/`SideEffectSpec` schema (`InputSpec`/`OutputSpec` kept, moved into `plan.py`). Relocated `test_pptx_helpers.py` + `test_shell_helpers.py` into `scripts/tests/` so `testpaths` collects them. Removed the empty `templates/slide-library/`.
- Insider ownership slide (v0.5.8) — adds the `ownership` skill + `ownership_workbook.py` and a 15th slide-library entry (`insider-ownership`, before disclaimer/contact). For **Canadian public targets**, the analyst attaches a SEDI "Insider Information by Issuer" PDF (SEDI is bot-walled — no auto-fetch); the skill keeps current insiders only (`Ceased to be Insider: Not Applicable`), sums each one's **common shares** (multi-tranche → in-cell `=a+b+c`), looks up roles (relationship code + company site / LinkedIn), and fills the ownership template (rows 39-65 B/F/G/J + `F35` from the cap table's Section VII basic shares). The pitch plan runs `ownership` **after `captable`, before `deck`**; the deck-assembler pastes `Ownership!B4:G17` into the slide's left "Insiders" placeholder (`Rectangle 1`), leaving the right Bloomberg "Institutions" side a placeholder. `insert_cap_table_into_placeholder` generalised to `insert_excel_into_placeholder`; the shipped ownership template is pre-cleaned of vestigial external links / legacy defined names so the openpyxl output stays Excel-openable for the render. Non-Canadian / no-PDF → null workbook, slide left as placeholder. ✅ shipped 2026-06-03.
- Public comparables (v0.5.11) — adds the `comps` skill + `comps_workbook.py` and ships `templates/INFOR Comps Template.xlsx`. Finds **3 verticals** relevant to the target and writes **6 public `Exchange:Ticker` CapIQ peers** per vertical (`B10:B15`/`B20:B25`/`B30:B35`) with vertical labels (`D9`/`D19`/`D29`) and ≤50-char descriptions (`AA10:AA15`/`AA20:AA25`/`AA30:AA35`); every market-data / multiple / statistic column is a CapIQ array formula keyed off column B, left **un-evaluated** because this environment has no CapIQ connector (the analyst refreshes in Excel). The pitch plan runs `comps` **after `ownership`, before `deck`**; the comps workbook folds into the combined workbook as the `comps` tab (its `__snloffice` helper sheet dropped automatically), and the deck's comps slide stays a placeholder (no Excel→PowerPoint step while CapIQ can't be refreshed). Ticker format + peer-selection rules carry over from the old `comps-infor` skill. Unlike the ownership template, the comps template round-trips cleanly through openpyxl, so it ships as-is — no cruft-stripping. ✅ shipped 2026-06-08.
- Precedent transactions (v0.5.12) — adds the `precedents` skill + `precedents_workbook.py` and ships `templates/INFOR Precedents Template.xlsx`. Researches **≤12 M&A deals** in **2 peer groups of 6** (labels `E7`/`E16`) and writes per deal: identity (input currency `B`, announce date `E`, target `F`, acquiror `G`, source-FX TEV `I`, **3-letter** HQ code `AI` — `H` left empty), the source-FX $ inputs for the **one metric family** the agent picks by industry (operating → Revenue `K`/`L` + Adj. EBITDA `O`/`P` → EV/Rev + EV/EBITDA; financial → Net Income `M`/`N` + Book Value `Q` + Tangible BV `R` → P/E + P/B + P/TBV), and a source hyperlink per metric on the `AB`–`AG` "Link" cells. A **PR-disclosed multiple is preferred** and written as a literal over the ratio formula (`S`–`Z`); else the disclosed $ figure (most recent reported figure as the LTM/NTM proxy — no multi-filing stub calc). Column `C` FX, `J` (`=+I*C`), the `S`–`Z` ratios, and the group/global statistic rows stay un-evaluated until the analyst refreshes CapIQ. The pitch plan runs `precedents` **after `comps`, before `deck`**; it folds into the combined workbook as the `precedents` tab and the slide stays a placeholder. The template carried the ownership-style cruft (≈58.7k defined names + 174 external links + leftover `AB:AG` example hyperlinks, 6 MB → ~9 KB) and was stripped during prep, with `AG4` relabelled `P /E `→`P / TBV` and `C2` defaulted to `USD`. Sourcing criteria carry over from the old `precedents-infor` skill. ✅ shipped 2026-06-09.
- Parallel stage execution (v0.5.13) — the conductor no longer runs stages strictly top-to-bottom. New `scripts/plan_schedule.py` (`compute_waves`) derives the dependency DAG from the `$stages.<id>.<name>` references already in each stage's inputs and groups stages into ordered **waves** of mutually-independent stages; the conductor dispatches each wave's `Task` calls concurrently (one message, multiple tool uses) and waits at the wave boundary. No schema change — there is still no `depends_on` field; the references *are* the DAG. The lone side-effect edge (`workbook-aggregator` must run after `deck`, which reads the standalone cap table before the aggregator folds + deletes it) is encoded as a hardcoded barrier: the aggregator depends on every other stage, so it is always alone in the final wave (generalized in v0.5.16: it depends on every stage *except its own downstream consumers*, so a post-aggregation stage like `financial-charts` can follow it without a cycle). Pitch collapses 9 sequential stages → 5 waves (wave 1 overlaps `wireframe`/`ltm-metrics`/`comps`/`precedents`); earnings-update 6 → 4. `required` checkpoints are evaluated at the wave boundary (documented caveat: they gate downstream waves, not their own wave-mates — all shipped plans use `informational`, so behaviour is unchanged). This supersedes the v1 "sequential, no parallel" decision (note 12 H1) per analyst direction. ✅ shipped 2026-06-09.
- Pitch precedents slide + companion-workbook polish (v0.5.14) — adds a **`precedent-transactions`** slide to the shared `INFOR Slide Library.pptx` (immediately after the renamed `Comparable Companies Analysis` slide; library 16→17 slides, pitch base 15→16), wired through `slide_library_registry`, `pitch_deck_wireframe`, `pitch_deck_assembler` (post-delete index bumps + `precedents_takeaway` fill + verifier placeholder) and a new required `PitchDeckContent.precedents_takeaway`; the earnings-update assembler's `_KEEP_LIBRARY_INDICES` shift `(0,6,7,14,15)`→`(0,6,7,15,16)` for the grown shared library. Companion-workbook + deck fixes from a live pitch run: the workbook aggregator now preserves the precedents `AB`–`AG` source hyperlinks on the openpyxl merge path and relinks comps `F3` / precedents `C2` to the cap table's output currency (`F5`); the `Font(color=…)` Calibri-11 clobber is fixed across captable `F7`/`F16`/`D47`/`D48`, ownership `F35` + insider cells, and precedents target/acquiror (all keep/get Palatino); `precedents` now requires ≥1 multiple per deal and targets 6 per group; `pitch-content` drafts 5 considerations (fills slide 10's fifth row) and the market-entry value font drops 10→9 pt so the table renders at the 5.71" clamp instead of growing to ~6.3". ✅ shipped 2026-06-09.
- Financial Summary data tab (v0.5.15) — adds the `financial-summary` skill + `financial_summary_workbook.py`, the chart-ready companion behind the pitch deck's Financial Summary slide (library entry 8). It **selects the 4 metrics** most relevant to the target (industry-aware: operating vs. financial family) and is now the **single source of truth** for them — it emits `financial_metric_labels` (the deck's slide-8 tiles read it from this stage, not from `pitch-content`; the field + its two validators are dropped from `PitchDeckContent`). The tab carries the **last 5 fiscal years** (from the latest four 10-Ks) **+ LTM** for each metric, laid out *chart-ready* (one metric per row; FY + LTM as a single contiguous numeric header axis; Units column; no merged cells in the data block) so a later task drops native Excel charts on it with no reshaping — charts/PowerPoint are out of scope. **LTM is linked from the `ltm-metrics` tab**, not recomputed: each flow metric's LTM cell is a label-keyed `=INDEX('ltm-metrics'!$B:$B, MATCH("(=) <result_label>", 'ltm-metrics'!$A:$A, 0))` that resolves only in the combined `pitch-<codename>.xlsx` (`#N/A` standalone, like the cap table's CapIQ formulas); non-flow metrics (balances/ratios) fall back to the latest reported value. **The selection↔linking cycle is resolved by "financial-summary drives ltm-metrics"** (analyst-chosen): `financial-summary` runs **before** `ltm-metrics` and passes `ltm_bridge_specs`; `ltm-metrics` gains an optional `extra_bridges` param (+ `Bridge` dataclass, reusing `_write_bridge`) and bridges exactly those metrics plus its always-on Revenue + Adj. EBITDA bridges (earnings-update passes no specs, so it is unchanged). Pitch now schedules into **6 waves** (`financial-summary` joins wave 1; `ltm-metrics`→wave 2); the aggregator folds the workbook in as the `financial-summary` tab with **no aggregator code change** (the link targets the renamed `'ltm-metrics'` tab and survives the openpyxl/COM copy). The deal-init G6 prompt now asks for the latest 4 FY 10-Ks (5-year history) alongside the interim stubs, for pitch as well as earnings update. ✅ shipped 2026-06-29.
- Financial Summary charts (v0.5.16) — finishes the pitch deck's Financial Summary slide. Adds the **`financial-charts`** skill + `scripts/financial_charts.py` and a new pitch plan stage that runs **after `workbook-aggregation`** (the combined workbook is the only place each flow metric's `=INDEX('ltm-metrics'!…)` LTM link resolves). It builds **one INFOR-formatted clustered-column chart per metric** on the combined workbook's `financial-summary` tab (single series = the metric row `B{r}:G{r}`, categories = the period header `B5:G5` read dynamically so a suppressed-LTM tab charts 5 cols), exports each (Excel COM → PNG on Windows, openpyxl + LibreOffice fallback off-Windows), and inserts them into the slide-8 chart placeholders (`Rectangle 17/7/19/18` ← rows 6–9) stretched to box; charts persist on the tab too. Formatting: Palatino 9 black, no title/gridlines, hidden value axis, gap 50%, data labels Outside End, bars `46566E`. The scheduler's aggregator barrier is generalized (excludes the aggregator's downstream consumers) so the new stage can follow it without a cycle → pitch now schedules into **7 waves**. **Fixes a latent v0.5.15 bug:** the v0.5.15 claim that the financial-summary→ltm-metrics link "survives the openpyxl/COM copy" was only true on the openpyxl path — the COM merge bound it to an *external* workbook ref (Excel's `.Formula` getter masked it as internal) so it resolved to `#N/A` and the LTM bar was blank; the aggregator now re-assigns each financial-summary LTM-link formula to re-bind it internally (`_relink_financial_summary_com`/`_openpyxl`). Also: combined metrics may be passed as `"=a+b"` Excel formulas (kept as cell formulas, never pre-summed), and both `financial-summary` + `ltm-metrics` are doc-locked to **millions with an `"MM"` suffix** so the value-for-value LTM link can't be 10⁶× off. ✅ shipped 2026-06-29.
- Pitch chart follow-ups (v0.5.17) — two pitch-only chart changes, both extending the existing `financial-charts` stage / `scripts/financial_charts.py` (no parallel path, no `pitch.yaml` schema change). (1) **LTM revenue pie on the overview slide:** new `render_ltm_revenue_pie_into_deck` builds an LTM-revenue-by-segment pie on the combined workbook's `ltm-metrics` tab over the "LTM Revenue Overview" block (categories = Segment column, values = "LTM Revenue (…)" column, Total excluded; located by section title via `ltm_revenue_overview_range`, mirroring the aggregator's label-row scan — Excel charts the literal cells) and drops it into the overview slide's (`prs.slides[6]`) `[Pie Chart Placeholder]` (`Rectangle 4`). Legend at top, no title/border, percentage labels, Palatino 9; slice fills from the new `pptx_helpers.INFOR_ACCENTS` (the `INFORFG.thmx` "INFOR (New)" accent1–6: `0E213F,46566E,ADB9CA,A4844B,767171,E5E3E3`), in theme order, cycled past six; the chart persists on the `ltm-metrics` tab like the FS charts persist on `financial-summary`. Null path (no `ltm-metrics` tab/block) leaves the placeholder. The deck-assembler is untouched — Rectangle 4 stays a *verified* placeholder through `deck`, then `financial-charts` fills it post-aggregation, and its mandatory QA render now covers slides `[6, 7]`. (2) **FS chart border + axis fix:** the legacy `ChartArea.Border.LineStyle = 0` did not kill the modern outline, so `_format_com_chart` now clears `ChartArea.Format.Line` (+ plot-area) via a shared `_com_strip_chart_border` and sets the category-axis line explicitly black (value axis still hidden); the openpyxl fallback mirrors both via `_openpyxl_no_border_black_axis`. ✅ shipped 2026-06-29.
- Chart-render robustness on Cowork/Linux (v0.5.18) — hardens the pitch `financial-charts` stage for the Excel-less production runtime (Cowork/Linux, openpyxl + LibreOffice only). (1) **Charts now always reach the workbook:** `_build_charts_openpyxl_libreoffice` / `_build_pie_openpyxl_libreoffice` persist the native chart objects on the combined workbook's `financial-summary` / `ltm-metrics` tabs **first**, then attempt the LibreOffice PNG render; when `soffice`/`libreoffice` is missing the workbook charts are still saved and the render degrades gracefully (the orchestrators return `None`, leaving the deck placeholders) instead of aborting the stage — and the `financial-charts` SKILL.md now mandates calling `render_financial_summary_charts_into_deck` / `render_ltm_revenue_pie_into_deck` and forbids hand-rolling charts with matplotlib or any other library. (2) **Visible black category-axis baseline:** the openpyxl axis line (`_openpyxl_no_border_black_axis`, used by `_make_openpyxl_chart` + `_make_single_value_chart`) was width-less, so the LibreOffice PNG dropped it and the bars floated with no baseline — it now carries an explicit visible width (`_AXIS_LINE_WIDTH_EMU` = 1 pt); the COM `_format_com_chart` gets a matching explicit weight. Value axis still hidden, labels Palatino 9 black. Tests cover the persisted charts, the graceful-degradation paths, the axis width, and the `financial-charts`→`deck`/`workbook-aggregation` wiring. ✅ shipped 2026-06-30.
- Phase 4 — New skills (valuation, profiles, industry research).
- Phase 5 — Quality + telemetry + per-skill URL allow-lists.
- Phase 6 — MCP / portability — deferred indefinitely.
