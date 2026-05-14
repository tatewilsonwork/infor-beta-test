# infor-beta

INFOR Financial Group's next-generation analyst workflow platform — a Claude Code plugin that orchestrates investment-banking deliverables (CIMs, pitches, teasers, earnings updates, fairness opinions, valuation) through a conductor meta-skill, specialised sub-skills, a typed I/O contract, and a parameterised slide library.

**Status: Phase 0 — foundational scaffold.** No skills are implemented yet. The production plugin today is the existing `infor-workflows` repo; this repo is a clean-break rebuild and will supersede it when ready.

## Vision

One conductor meta-skill (running in Claude Code's `Agent` tool) consumes a deliverable spec, dispatches to specialised data and writing skills with standardised typed I/O, and assembles the final output (`.pptx` / `.xlsx`) by cloning and filling slides from a 30–50-entry parameterised slide library. Initially medium HITL with confirmation gates at every major stage; autonomous later via configuration flip.

## Repo layout (planned)

```
.claude-plugin/marketplace.json    Marketplace manifest — points at infor-beta/ as plugin root
infor-beta/
├── skills/                        One directory per skill
│   └── <skill-name>/
│       ├── SKILL.md               Workflow + frontmatter
│       └── references/            Progressive-disclosure detail
├── plans/                         Conductor plan YAMLs (earnings-update, cim, pitch, ...)
├── scripts/                       Shared helpers + tests
├── templates/                     Excel + PowerPoint templates
│   └── slide-library/             Parameterised slide entries (30–50 total, v1)
README.md
CLAUDE.md                          Contributor brief (loaded by Claude Code)
CHANGELOG.md
```

## Canonical architecture record

The full design lives in Obsidian at `Hermes-L1/INFOR Platform Architecture/`. Note `12 — Locked Decisions.md` is the canonical record; notes 01–11 are the analytical history. Read 12 first.

## v1 skill portfolio (planned)

**Refactored from existing repo:** comps, precedents, buyerslist, captable, lbo-model, infor-deck-writing, infor-wireframe (typed SlidePlan), deckcheck (QA stage), brand-guidelines (demoted to library), earningsupdate (decomposed into a plan).

**New:** deck-assembler, valuation-infor (football field aggregator), company-profile-public-infor, company-profile-private-infor, industry-research-infor (low priority).

**Plans:** earnings-update.yaml (Phase 2 pilot), cim.yaml (slim 30–35 slides), pitch.yaml, teaser.yaml, fairness-opinion.yaml, valuation.yaml.

**Removed from scope:** management presentations, diligence support, research pipelines.

## Conventions (carried over from the old repo)

- Single plugin version, all skills in lock-step. Revisit at 20+ skills.
- Output files land in the **deal directory** (`~/Documents/INFOR Deals/<codename>/`), not in the analyst's `cwd`.
- Templates resolved via plugin-aware helper, never hardcoded paths.
- Excel does the math, not the LLM — arithmetic lives in cell formulas for auditability.
- English only (no French).
- INFOR-only (single-tenant). No multi-tenancy.

## Distribution

Production target: Claude Desktop, rolled out to INFOR employees via the Claude Code plugin marketplace. Hermes is used by Tate for testing only.
