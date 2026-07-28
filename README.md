# infor-beta

INFOR Financial Group's next-generation analyst workflow platform — a Claude Code plugin that orchestrates investment-banking deliverables (earnings updates, pitches, and — soon — overview decks) through a conductor meta-skill, specialised sub-skills, a typed I/O contract, and a shared slide library.

**Status: Phase 3 (proof-of-concept), plugin v0.5.40 — architecture migration Phases A, B, C and I shipped (see `docs/migration-plan.md`).** The conductor, the decomposed earnings-update plan, and the pitch slide-library plan (16-slide base with a configurable slide mix — market-entry, Financial Summary, and Key Investment Highlights; including the insider-ownership slide, a public-comparables companion workbook, a precedent-transactions companion workbook, a chart-ready financial-summary data tab with native Excel charts rendered onto the Financial Summary slide(s), and an LTM-revenue-by-segment pie on the overview slide) all run end-to-end, launchable via the `/pitch` and `/earnings-update` slash commands. The production plugin today is still the existing `infor-workflows` repo; this repo is a clean-break rebuild and will supersede it when ready.

## Vision

One conductor meta-skill (running in Claude Code's `Agent` tool) consumes a deliverable spec, dispatches to specialised data and writing skills with standardised typed I/O, and assembles the final output (`.pptx` / `.xlsx`) by cloning and filling slides from the shared `INFOR Slide Library.pptx`. Initially medium human-in-the-loop with confirmation gates at every major stage; autonomous later via a configuration flip.

## Repo layout

```
.claude-plugin/marketplace.json    Marketplace manifest — points at infor-beta/ as plugin root
infor-beta/
├── commands/                      Slash commands (/pitch, /earnings-update — full deck builds)
├── skills/                        One directory per skill (SKILL.md + references/)
├── plans/                         Conductor plan YAMLs (earnings-update, pitch, overview)
├── scripts/                       Shared helpers, typed schemas (pydantic), and tests
├── templates/                     Excel + PowerPoint templates + brand theme
│                                  (INFOR Slide Library.pptx; Cap Table / Comps /
│                                  Precedents / Ownership .xlsx templates; INFORFG.thmx)
docs/migration-plan.md             Phased plan to the target architecture (temporary)
tools/                             Repo maintenance tooling — NOT shipped in the plugin
README.md
CLAUDE.md                          Contributor brief (loaded by Claude Code)
CHANGELOG.md
```

See `CLAUDE.md` for the full contributor brief and the locked architectural decisions.

## Canonical architecture record

The full design lives in Obsidian at `Hermes-L1/INFOR Platform Architecture/`. Note `12 — Locked Decisions.md` is the canonical record; notes 01–11 are the analytical history. Read 12 first.

## Deliverables & plans

The conductor runs one plan per deliverable, resolved as `plans/<deliverable>.yaml`. The `/pitch <company name>` and `/earnings-update <company name>` slash commands launch a full build directly: they preset the deliverable + subject company, run deal-init, then collect the deck spec through **locked interactive dialogs** (`scripts/deck_spec.py`, rendered via the `AskUserQuestion` tool — clickable options plus a free-text "Other" box, the same questions with the same options every run). Only the judgement items are asked (analyst notes, CIM, valuation range, risk notes, and the pitch slide mix); client name, presentation date, and the LTM quarters are defaulted (computed from deal-init + the attached filings) and echoed once for override, and the required-documents checklist is posted as plain text:

- **`earnings-update.yaml`** — decomposed quarterly earnings update, 5 stages in 3 dependency waves: `wireframe` / `ltm-metrics` (parallel) → `content` / `captable` → `deck`. Clones the shared slide library; the `ltm-metrics` and `captable` stages write their tabs of the deal's single workbook. Fixed 5-slide layout (no slide options).
- **`pitch.yaml`** — the INFOR Slide Library pitch deck (16-slide base), 10 stages scheduled into 6 dependency waves: `wireframe` / `financial-summary` / `comps` / `precedents` (parallel) → `content` / `ltm-metrics` → `captable` → `ownership` → `deck` → `financial-charts` (draws the Financial Summary charts + overview LTM-revenue pie on the deal workbook and inserts them into the deck). Three deck-spec inputs adjust the slide mix: **acquisition-target slides** (1–4 slides, two targets each), **Financial Summary slides** (1 slide / 4 metrics or 2 slides / 8 metrics), and the **Key Investment Highlights slide** (include or omit). Clones the shared library; the cap-table, LTM-metrics, ownership (Canadian targets), comps, precedents and financial-summary stages each write one tab of the deal's single `pitch-<codename>.xlsx`.
- **`overview.yaml`** — stub; the overview deck is registered as a deliverable but not yet implemented.

## Skills

**Implemented:** `conductor`, `earningsupdate-wireframe`, `earningsupdate-content`, `pitch-wireframe`, `pitch-content`, `captable`, `ltm-metrics`, `comps`, `precedents`, `ownership`, `financial-summary`, `financial-charts`, `deck-assembler`.

**Roadmap (not yet built):** `buyerslist`, `lbo-model`, `deck-writing`, `deckcheck` (QA), `brand-guidelines` (library), `valuation` (football field), company / industry profiles.

**Removed from scope:** management presentations, diligence support, research pipelines.

## Conventions

- One plugin version, recorded in `marketplace.json` / `plugin.json` / `pyproject.toml` only — skills carry no `version:` frontmatter. Skill directory names carry no `-infor` suffix; deliverable-specific skills are prefixed by deliverable (`earningsupdate-*`, `pitch-*`).
- Output files land in the **deal directory** (`~/Documents/INFOR Deals/<codename>/`), not the analyst's `cwd`.
- Templates are resolved via the plugin-aware `find_template.sh` helper, never hardcoded paths.
- Inside a template, cells are reached through the `infor_`-prefixed **defined names** it carries and library slides through their **marker shapes** — never a hardcoded address or slide index, so an analyst can re-save a template or insert a library slide without a code change.
- Excel does the math, not the LLM — arithmetic lives in cell formulas for auditability.
- English only (no French). INFOR-only (single-tenant); no multi-tenancy.

## Development

```bash
pip install -e ".[dev]"        # pydantic, PyYAML, python-pptx, openpyxl, pytest, pytest-xdist
python -m pytest               # runs the suite under infor-beta/scripts/tests/
python -m pytest -n0 -k name   # serial, for debugging a single test
cd infor-beta/scripts && python -m schemas.export   # regenerate JSON schemas (idempotent)
```

The suite renders decks through headless LibreOffice (install it — it is the
default renderer on every platform) and is distributed across six workers by
default, since a conversion costs seconds. Expect 2.5-4 minutes for ~650 tests,
depending on machine load.

After re-saving one of the four workbook templates from Excel, restore its
defined names:

```bash
python tools/add_template_named_ranges.py --verify-excel   # --check to dry-run
```

## Distribution

Production target: Claude Desktop, rolled out to INFOR employees via the Claude Code plugin marketplace. Hermes is used by Tate for testing only.
