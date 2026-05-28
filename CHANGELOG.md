# Changelog

All notable changes to `infor-beta` are documented here. Format: [Keep a Changelog](https://keepachangelog.com/en/1.1.0/). The plugin uses a single version across all skills; bump every skill's `version:` frontmatter when bumping the plugin.

## [Unreleased]

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
