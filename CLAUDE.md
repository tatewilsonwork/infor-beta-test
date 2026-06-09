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
- Phase 4 — New skills (valuation, profiles, industry research).
- Phase 5 — Quality + telemetry + per-skill URL allow-lists.
- Phase 6 — MCP / portability — deferred indefinitely.
