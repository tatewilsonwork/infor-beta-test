---
name: deck-assembler
description: >
  Use this skill as the deck assembly stage. It consumes a typed SlidePlan and typed content bundle
  and writes either the earnings-update deck or the pitch deck, both cloned from the shared INFOR
  slide library.
allowed-tools: [Read, Write, Bash]
---

# Deck Assembler

This stage assembles typed slide/content handoffs into PowerPoint decks.

## Supported POC paths

1. **Earnings update POC**
   - `SlidePlan.deliverable_type = "earnings-update"`
   - content bundle schema: `EarningsUpdateContent`
   - template: `INFOR Slide Library.pptx` (shared library)
   - helper: `scripts/earnings_update_assembler.py`
   - clones library slides 1, 7, 8, 16, 17 (cover, overview, earnings summary, plus the two static closers — indices shifted when the shared library grew to 17 slides in v0.5.14; the helper's `_KEEP_LIBRARY_INDICES` is authoritative) and deletes the rest; cap table replaces `Rectangle 3` on the overview slide (range `B15:F40`, including the Financial Metrics section and the LTM/forward Valuation Metrics rows).

2. **Slide-library pitch POC**
   - `SlidePlan.deliverable_type = "pitch"`
   - content bundle schema: `PitchDeckContent`
   - template: `INFOR Slide Library.pptx`
   - helper: `scripts/pitch_deck_assembler.py`

## Conductor mode

When invoked by the conductor, read:

- `$STAGE_INPUTS` — JSON with `slide_plan_path`, `content_bundle_path`, `template_name`, and `output_dir`; may also include `captable_workbook_path`, `ownership_workbook_path`, and `financial_metric_labels` (the Financial Summary tile names, selected by the `financial-summary` stage — four per FS slide in the plan, so 4 or 8)
- `$STAGE_OUTPUTS` — path where this stage must write its structured handoff
- `$DEAL_DIR` — deal directory root

## Workflow

1. Read `$STAGE_INPUTS`.
2. Resolve `template_name` under `${CLAUDE_PLUGIN_ROOT:-./infor-beta}/templates/`. Both deliverables now resolve to `INFOR Slide Library.pptx`.
3. Inspect the `SlidePlan.deliverable_type`.
4. If `earnings-update`, call `assemble_earnings_update_deck(...)` and pass `captable_workbook_path` when supplied so slide 2's cap-table placeholder is replaced.
5. If `pitch`, call `assemble_pitch_deck(...)`.
   Both assemblers verify the library and workbook layouts before touching hardcoded indices/ranges (shared `template_layout` map): every library slide about to be kept, cloned, deleted, or filled is checked against its marker shape (extending the `Rectangle 17` FS self-discovery pattern), and each Excel picture range's sentinel anchors (`B15`/`B40` on the cap table; `B4`/`B17`, `B19`/`B35` on ownership) are checked before pasting. A re-ordered library or re-saved workbook raises `TemplateLayoutError` naming what moved — report it to the analyst; do not work around it.
6. Write the deck under `$DEAL_DIR/artefacts/` when `$DEAL_DIR` is set; otherwise use the supplied `output_dir`.
7. **Geometry QA runs inside the assembler — do not re-do it by hand.** Both assemblers call `deck_repair.converge_deck` (write → verify → repair → re-verify against `deck_contract`) and raise `DeckNotConvergedError` if the deck still breaks the contract. If that raises, **report it — do not hand-patch the deck**: the message names the shape, the measured overflow, and every shrink that was tried, and the remedy is editorial (shorter copy from the content stage). What remains for you is the **visual review** below, which is judgement a measurement cannot make.
8. Write `$STAGE_OUTPUTS` as:

```json
{
  "deck_path": "/absolute/path/to/Pitch Deck - Company.pptx"
}
```

## Pitch boundaries

The slide mix is **SlidePlan-driven**: the plan's repeated `financial-summary` entries grow that section (one slide per four metrics), a missing `key-investment-highlights` entry deletes that slide, and the market-entry section is cloned to the true content-bundle target count. Every slide after the Financial Summary section therefore shifts with it — the helper computes all indices (`_PitchLayout`); never hardcode a slide number past 7. The slide numbers below are the **default** mix (1 FS slide, KIH included):

- Slide 1: client name/date only.
- Slide 2: executive summary bullets — square main / dash sub bullets, template body colour (not the navy list default).
- Slides 3–5: static credentials, do not touch.
- Slide 6: section divider labels.
- Slide 7: company overview bullets — the assembler sizes the bullets box to the band above the "LTM Revenue Breakdown" header (`pptx_helpers.fit_overview_textbox`) and the converge loop sets the autofit `fontScale` from the **measured** render. Do not hand-resize this box or hand-tune its scale. The cap table is pasted into the `Rectangle 3` placeholder when `captable_workbook_path` is supplied (`slide_index=6`, range `B15:F40`), and the `[x]$MM` footnote token is set to the cap table's output currency (`US` / `C`, read from `F5`). The LTM revenue pie stays a deferred placeholder.
- Slide 8 (+9 when the plan carries two FS slides): financial metric **NAME** labels only, taken from the `financial_metric_labels` stage input (the `financial-summary` stage's output, not the content bundle) — four tiles per slide, in order; with two slides they are retitled `Financial Summary (1 of 2)` / `(2 of 2)` and the extra slide is cloned from the library's FS slide. The chart placeholders are left as-is here; they are filled by the downstream **`financial-charts`** stage, which runs after `workbook-aggregation` (the charts can only be drawn once the combined workbook's `financial-summary` LTM links resolve against the folded-in `ltm-metrics` tab).
- Ownership slide (Canadian public targets) — **follows the Financial Summary section** (slide 9 in the default mix; the helper's `layout.ownership` is authoritative). When `ownership_workbook_path` is supplied, the left **"Insiders"** placeholder (`Rectangle 1`) is replaced by a picture of the ownership workbook's Select-Insiders block (sheet `Ownership`, range `B4:G17`). When the ownership workbook also carries Bloomberg institutional data (`Bloomberg Output` tab holder in `C14` — the ownership stage ingested an analyst-attached Bloomberg export), the right **"Institutions"** placeholder (`Rectangle 3`) is replaced by the Select-Institutions block too (range `B19:G35`: top-12 + subtotal + Other Shareholders + Total); the assembler detects this itself (`_ownership_has_bloomberg`) — no extra stage input. Without Bloomberg data the right side stays a Bloomberg-sourced placeholder.
- Considerations/Mitigants slide (default 10): concise risks/mitigants + tagline. Cells go in at the library's sizes (12 pt header, 10 pt body) and the table is clamped back to the library's **5.18"** total height. If the rendered table then overruns, the converge loop caps the body font until it fits (the header stays 12 pt). Do not hand-resize this table or pre-shrink its copy.
- Comps slide (default 11): comps takeaway; chart placeholder stays unless insertion replaces it later.
- Precedents slide (default 12): precedent-transactions takeaway (`precedents_takeaway` from the content bundle); the chart area stays a placeholder like comps (no Excel→PowerPoint while CapIQ can't be refreshed).
- Key Investment Highlights slide (default 13): filled when content supplies them; **deleted entirely when the SlidePlan omits its entry** (the deck spec's "omit" option) — content-bundle highlights are then ignored.
- Market-entry slides (default 14+): potential market-entry targets — the fixed 12-row comparison table (Overview / HQ / Year Founded → 7 industry metrics → Scale KPIs / Strategic Rationale), **two targets per slide**. The assembler clones the library's market-entry slide to `ceil(len(market_entry_targets) / 2)` slides, titles them `(N of M)`, writes the label column white at 11 pt and target values at 9 pt, blanks the unused column + logo on an odd final slide, and — after the cells are filled — clamps each table to a fixed **5.71"** total height. A stored row height is only a render-time MINIMUM, so the clamp alone cannot stop the layout engine re-growing a row; the converge loop measures the rendered table and caps the body font when it does, taking the 11 pt labels down before the 9 pt values. Disclaimer/contact follow.

## Visual QA (mandatory)

**Geometry is already measured.** `deck_repair.converge_deck` runs inside the
assembler: it renders the deck, measures rendered ink against every shape's
declared box (including per-shape attribution for overflows a neighbour masks and
for autofit shapes the renderer would otherwise hide), repairs what a font size can
fix, and **fails the stage** if it cannot converge. You do not need to re-measure
overflow, and you must not hand-patch a shape to work around a
`DeckNotConvergedError` — report it, naming the shape and the depth from the
message.

What is left is the part measurement cannot do: **reading the slides.** Render the
slides below, look at each one, and check content and appearance — values that are
wrong, labels that are unreadable, pictures that landed in the wrong place, text
drawn over text. If `verify_deck` also produced advisory `vision-review` findings,
they name the slides worth the closest look and why.

Do not skip this, and do not claim a deck is clean if the renderer was unavailable —
say that QA could not run.

- **Earnings update**: render slides 2 and 3, and check specifically for:
  - Slide 2 — the company-overview bullet block must not run under or overlap
    the **"LTM Revenue Breakdown"** title or the pie/cap-table region.
  - Slide 3 — every **Financial Highlights** metric tile must show its value and
    label on at most two lines with no text touching the tile edge. A label like
    `Cloud Services & Subscriptions Revenue` that wraps to a third line or clips
    means the label is too long for the tile — the content stage must abbreviate
    it (e.g. `Cloud Services & Subscriptions Rev.`). Report it; do not shrink the
    tile's font here.
  - Both slides — confirm no figure reads as `US,…` / `C,…` (a dropped `$` and
    digit from a regex-substitution bug upstream); every dollar figure must
    start with a plain `$`.
- **Pitch**: render the executive summary (slide 2), company overview (slide 7),
  risks/mitigants, ownership, and every market-entry slide. Compute the
  zero-based indices from the slide mix instead of hardcoding: with `nfs` =
  the plan's Financial Summary slide count and `kih` = 1 when the plan carries
  the Key Investment Highlights entry else 0 — risks = `7 + nfs + 1`, ownership
  = `6 + nfs + 1`, market-entry = `10 + nfs + kih + 1 ..` (defaults: 9, 8, 13+).
  On slide 2 confirm the body text is the template dark colour with square/dash
  bullets (not blue); on slide 7 confirm the cap-table picture landed and the
  footnote currency is correct; with two FS slides confirm both are titled
  `(1 of 2)` / `(2 of 2)` and each shows four metric tiles; on the market-entry
  slides confirm no row label or value clips and no data row is blank; on the
  ownership slide confirm the insider picture landed on the left and the
  institutions side matches the data: the Select-Institutions picture when a
  Bloomberg export was ingested (no `####` overflow, names resolved), the
  placeholder otherwise.

```python
import sys, os
sys.path.insert(0, os.environ.get("CLAUDE_PLUGIN_ROOT", "./infor-beta") + "/scripts")
from slide_render import render_deck_to_png

pngs = render_deck_to_png(deck_path, output_dir, slide_indices=[1, 2])  # zero-based
```

Read each PNG with the `Read` tool. **Never hand-edit shape geometry
(position/size) or font sizes to force a fit** — geometry is the converge loop's,
and a manual nudge both fights it and is invisible to it on the next run. If you
believe a shape still overflows after the loop reported convergence, that is a bug
in the contract worth reporting, not something to patch here.

The renderer is **LibreOffice headless on every platform** (v0.5.35), so dev and
Cowork render identically; PowerPoint COM is reachable only by explicit opt-in
(`backend="powerpoint"` / `INFOR_SLIDE_RENDER_BACKEND`). If LibreOffice is absent it
raises `RuntimeError` — note that QA could not run rather than claiming the deck is
clean.

## Reference command

```python
import json, os, sys
from pathlib import Path

plugin_root = Path(os.environ.get("CLAUDE_PLUGIN_ROOT", "./infor-beta"))
sys.path.insert(0, str(plugin_root / "scripts"))

from schemas import SlidePlan
from earnings_update_assembler import assemble_earnings_update_deck
from pitch_deck_assembler import assemble_pitch_deck

inputs = json.loads(Path(os.environ["STAGE_INPUTS"]).read_text())
slide_plan = SlidePlan.model_validate_json(Path(inputs["slide_plan_path"]).read_text())
template_path = plugin_root / "templates" / inputs["template_name"]
output_dir = Path(os.environ.get("DEAL_DIR", inputs["output_dir"])) / "artefacts"

if slide_plan.deliverable_type == "earnings-update":
    deck_path = assemble_earnings_update_deck(
        slide_plan_path=inputs["slide_plan_path"],
        content_path=inputs["content_bundle_path"],
        template_path=template_path,
        output_dir=output_dir,
        captable_workbook_path=inputs.get("captable_workbook_path"),
    )
elif slide_plan.deliverable_type == "pitch":
    deck_path = assemble_pitch_deck(
        slide_plan_path=inputs["slide_plan_path"],
        content_path=inputs["content_bundle_path"],
        template_path=template_path,
        output_dir=output_dir,
        captable_workbook_path=inputs.get("captable_workbook_path"),
        ownership_workbook_path=inputs.get("ownership_workbook_path"),
        financial_metric_labels=inputs.get("financial_metric_labels"),
    )
else:
    raise ValueError(f"unsupported deliverable_type: {slide_plan.deliverable_type}")

Path(os.environ["STAGE_OUTPUTS"]).write_text(json.dumps({"deck_path": str(deck_path)}, indent=2) + "\n")
```
