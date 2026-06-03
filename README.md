# infor-beta

INFOR Financial Group's next-generation analyst workflow platform — a Claude Code plugin that orchestrates investment-banking deliverables (earnings updates, pitches, and — soon — overview decks) through a conductor meta-skill, specialised sub-skills, a typed I/O contract, and a shared slide library.

**Status: Phase 3 (proof-of-concept), plugin v0.5.8.** The conductor, the decomposed earnings-update plan, and the 15-slide pitch slide-library plan (including the insider-ownership slide) all run end-to-end. The production plugin today is still the existing `infor-workflows` repo; this repo is a clean-break rebuild and will supersede it when ready.

## Vision

One conductor meta-skill (running in Claude Code's `Agent` tool) consumes a deliverable spec, dispatches to specialised data and writing skills with standardised typed I/O, and assembles the final output (`.pptx` / `.xlsx`) by cloning and filling slides from the shared `INFOR Slide Library.pptx`. Initially medium human-in-the-loop with confirmation gates at every major stage; autonomous later via a configuration flip.

## Repo layout

```
.claude-plugin/marketplace.json    Marketplace manifest — points at infor-beta/ as plugin root
infor-beta/
├── skills/                        One directory per skill (SKILL.md + references/)
├── plans/                         Conductor plan YAMLs (earnings-update, pitch, overview)
├── scripts/                       Shared helpers, typed schemas (pydantic), and tests
├── templates/                     Excel + PowerPoint templates
│                                  (INFOR Slide Library.pptx, INFOR Cap Table Template.xlsx)
README.md
CLAUDE.md                          Contributor brief (loaded by Claude Code)
CHANGELOG.md
```

See `CLAUDE.md` for the full contributor brief and the locked architectural decisions.

## Canonical architecture record

The full design lives in Obsidian at `Hermes-L1/INFOR Platform Architecture/`. Note `12 — Locked Decisions.md` is the canonical record; notes 01–11 are the analytical history. Read 12 first.

## Deliverables & plans

The conductor runs one plan per deliverable, resolved as `plans/<deliverable>.yaml`:

- **`earnings-update.yaml`** — decomposed quarterly earnings update: `wireframe → content → ltm-metrics → captable → deck → workbook-aggregation`. Clones the shared slide library and emits a companion cap table.
- **`pitch.yaml`** — the 14-slide INFOR Slide Library pitch deck: `wireframe → content → captable → deck → workbook-aggregation`.
- **`overview.yaml`** — stub; the overview deck is registered as a deliverable but not yet implemented.

## Skills

**Implemented:** `conductor`, `earningsupdate-wireframe`, `earningsupdate-content`, `pitch-wireframe`, `pitch-content`, `captable`, `ltm-metrics`, `excel-to-powerpoint`, `deck-assembler`, `workbook-aggregator`.

**Roadmap (not yet built):** `comps`, `precedents`, `buyerslist`, `lbo-model`, `deck-writing`, `deckcheck` (QA), `brand-guidelines` (library), `valuation` (football field), company / industry profiles.

**Removed from scope:** management presentations, diligence support, research pipelines.

## Conventions

- Single plugin version, all skills in lock-step. Skill directory names carry no `-infor` suffix; deliverable-specific skills are prefixed by deliverable (`earningsupdate-*`, `pitch-*`).
- Output files land in the **deal directory** (`~/Documents/INFOR Deals/<codename>/`), not the analyst's `cwd`.
- Templates are resolved via the plugin-aware `find_template.sh` helper, never hardcoded paths.
- Excel does the math, not the LLM — arithmetic lives in cell formulas for auditability.
- English only (no French). INFOR-only (single-tenant); no multi-tenancy.

## Development

```bash
pip install -e ".[dev]"        # pydantic, PyYAML, python-pptx, openpyxl, + pytest
python -m pytest               # runs the suite under infor-beta/scripts/tests/
cd infor-beta/scripts && python -m schemas.export   # regenerate JSON schemas (idempotent)
```

## Distribution

Production target: Claude Desktop, rolled out to INFOR employees via the Claude Code plugin marketplace. Hermes is used by Tate for testing only.
