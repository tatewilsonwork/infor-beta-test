# Changelog

All notable changes to `infor-beta` are documented here. Format: [Keep a Changelog](https://keepachangelog.com/en/1.1.0/). The plugin uses a single version across all skills; bump every skill's `version:` frontmatter when bumping the plugin.

## [Unreleased]

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
