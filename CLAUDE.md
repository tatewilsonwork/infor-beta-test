# Contributor brief — infor-beta

This file is loaded automatically when Claude Code opens this repo. It orients you (and future contributors) to the plugin's layout, conventions, and the rebuild context.

## Context: this is a clean-break rebuild

The production INFOR plugin today is **`infor-workflows`** (at `~/Desktop/infor-workflows`, v2.10.0). It keeps shipping unchanged until `infor-beta` supersedes it. This repo is a **fresh start, not a refactor in place** — designed directly for the target architecture with no backwards-compatibility shims.

The canonical architecture record lives in Obsidian at `Hermes-L1/INFOR Platform Architecture/`. Note `12 — Locked Decisions` is authoritative.

## What this plugin will be

A single Claude Code plugin, `infor-beta`, containing:

- A **conductor meta-skill** that consumes a deliverable spec (CIM, pitch, earnings update, …) and orchestrates sub-skills via the `Agent` tool.
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
├── plans/                         Conductor plan YAMLs (earnings-update, cim, pitch, teaser, fairness-opinion, valuation)
├── scripts/                       Shared helpers + tests (pptx_helpers.py, find_template.sh, sanitize_name.sh, ...)
├── templates/                     Excel + PowerPoint templates shipped with the plugin
│   └── slide-library/             One directory per concrete slide entry (concept × layout)
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
- `infor-wireframe` emits a typed `SlidePlan` (markdown view kept for analyst readability).
- `brand-guidelines-infor` is demoted to a **library** consumed by `deck-assembler` and `deckcheck-infor` — not a stage in any plan.
- `earningsupdate-infor` is decomposed into `plans/earnings-update.yaml` (Phase 2 pilot).

**Slide library**
- Multiple library entries per slide concept × layout combination (e.g. `company-overview-two-col`, `company-overview-full-bleed`). No parameterised `variants:` field. One `.pptx` per entry.
- Realistic library size: 40–80 entries at v1 maturity. Enumerable via a categorised README — no search UX required until well past ~100.
- Lives in the same repo (`templates/slide-library/`).

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
`comps-infor`, `precedents-infor`, `buyerslist-infor`, `captable-infor`, `lbo-model`, `infor-deck-writing`, `infor-wireframe` (typed), `deckcheck-infor` (QA), `brand-guidelines-infor` (→ library), `earningsupdate-infor` (→ plan).

**New skills:**
`deck-assembler` (Phase 3 foundational), `valuation-infor` (football field over comps + precedents + LBO), `company-profile-public-infor`, `company-profile-private-infor`, `industry-research-infor` (low priority, late Phase 4).

**Conductor plans:**
`earnings-update.yaml` (decomposed Phase 3 POC), `cim.yaml` (slim 30–35 slides), `pitch.yaml`, `teaser.yaml`, `fairness-opinion.yaml`, `valuation.yaml`.

**Removed from scope:** management presentations, diligence support, research pipelines.

## Conventions

### Versioning — single plugin version, all skills tied to it
One version for the entire plugin. Any skill change bumps the plugin version, and every skill's `version:` frontmatter equals the plugin version. No per-skill drift. Revisit at 20+ skills.

### Progressive disclosure
Long workflows split into a short `SKILL.md` (workflow + step outline) + `references/<topic>.md` loaded on demand. New skills longer than ~250 lines should follow it.

### Shared helpers (don't re-implement)
Templates, filename sanitization, python-pptx formatting helpers, brand constants, and the typed I/O contract live in `infor-beta/scripts/`. Skills import via `CLAUDE_PLUGIN_ROOT`:

```python
import sys, os
sys.path.insert(0, os.environ.get("CLAUDE_PLUGIN_ROOT", "./infor-beta") + "/scripts")

from pptx_helpers import set_text, write_bulleted_shape, set_cell_text, clone_slide, find_shape
from schemas import (
    Company, Filing, FilingType, SlidePlan, EarningsUpdateContent, PitchDeckContent, DealContext,
    Plan, Stage, CheckpointMode, SkillManifest,
)
from earnings_update_wireframe import build_earnings_update_slide_plan, write_slide_plan
from earnings_update_assembler import assemble_earnings_update_deck
from pitch_deck_wireframe import build_pitch_deck_slide_plan
from pitch_deck_assembler import assemble_pitch_deck
from slide_library_registry import load_slide_library_registry
from codename import resolve, find_existing, disambiguate
from deal_init import render_init_prompt, load_or_locate_deal, save_deal_context, load_deal_context
from plan_refs import resolve_refs
from run_log import make_run_id, create_run_dir, write_stage_inputs, read_stage_outputs, write_summary
```

For the bash helpers:

```bash
SANITIZED=$(bash "${CLAUDE_PLUGIN_ROOT:-./infor-beta}/scripts/sanitize_name.sh" "$RAW_NAME")
TEMPLATE=$(bash "${CLAUDE_PLUGIN_ROOT:-./infor-beta}/scripts/find_template.sh" "INFOR Comps Template.xlsx")
```

Brand constants are in `pptx_helpers` (`PALATINO`, `COLOR_UP`, `COLOR_DOWN`). JSON-Schema views of every typed contract are emitted to `infor-beta/scripts/schemas/json/` — regenerate with `python -m schemas.export` (idempotent).

### Excel does the math, not the LLM
Arithmetic lives in cell formulas for analyst auditability. Skills write inputs and let the workbook compute.

### Output files
Skills write to the **deal directory** (`~/Documents/INFOR Deals/<codename>/`), resolved by the conductor at deal-init. The cwd default is preserved for ad-hoc skill invocation, but the conductor always sets an explicit output target.

## Phase status

- Phase 0 — Stabilise & document. ✅ shipped 2026-05-14.
- Phase 1 — Deal model + typed I/O contract (`Company`, `Filing`, `SlidePlan`, `DealContext`, `SkillManifest`). ✅ shipped 2026-05-14.
- Phase 2 — Conductor v1 meta-skill + earnings-update plan pilot. Adds `Plan` / `Stage` schemas, `plan_refs` resolver, `deal_init` and `run_log` helpers, ports `earningsupdate-infor` + `captable-infor`. ✅ shipped 2026-05-15.
- Phase 3 — Earnings-update proof-of-concept decomposition + POC `deck-assembler`. Adds `EarningsUpdateContent`, `earningsupdate-wireframe-infor`, `earningsupdate-content-infor`, template-specific `deck-assembler`, and a four-stage `earnings-update.yaml`. ✅ POC shipped 2026-05-15.
- Phase 3 slide-library POC — 14-slide `INFOR Slide Library.pptx` proof-of-concept with `PitchDeckContent`, `pitch-wireframe-infor`, `pitch-content-infor`, `excel-to-powerpoint-infor`, `pitch-library-poc.yaml`, and generalized deck-assembler support. ✅ POC shipped 2026-05-16; expanded to 14 slides (Key Investment Highlights + Potential Market Entry Targets) 2026-05-28.
- Phase 3 earnings-update onto the shared library (v0.5.0) — earnings-update deck now clones `INFOR Slide Library.pptx` (slides 1, 7, 8, 14, 15) instead of the retired standalone template. Overview slide gains an LTM revenue pie placeholder + "Introduction to {company}" title + "Capitalization Summary" cap table (`Rectangle 3`, `B15:F31`); earnings-summary metric boxes show value+name with the period in the mid-blue bar; broker values carry `$`. Adds `ltm-revenue-infor` (standalone LTM revenue workbook) and `slide_render.py` (PNG overflow QA). ✅ shipped 2026-05-28.
- Phase 4 — New skills (valuation, profiles, industry research).
- Phase 5 — Quality + telemetry + per-skill URL allow-lists.
- Phase 6 — MCP / portability — deferred indefinitely.
