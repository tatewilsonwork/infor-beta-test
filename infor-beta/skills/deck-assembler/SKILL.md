---
name: deck-assembler
description: >
  Use this skill as the deck assembly stage. It consumes a typed SlidePlan and typed content bundle
  and writes either the earnings-update deck or the pitch deck, both cloned from the shared INFOR
  slide library.
version: 0.5.15
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
   - clones library slides 1, 7, 8, 14, 15 (cover, overview, earnings summary, plus the two static closers) and deletes the rest; cap table replaces `Rectangle 3` on the overview slide (range `B15:F40`, including the Financial Metrics section and the LTM/forward Valuation Metrics rows).

2. **Slide-library pitch POC**
   - `SlidePlan.deliverable_type = "pitch"`
   - content bundle schema: `PitchDeckContent`
   - template: `INFOR Slide Library.pptx`
   - helper: `scripts/pitch_deck_assembler.py`

## Conductor mode

When invoked by the conductor, read:

- `$STAGE_INPUTS` — JSON with `slide_plan_path`, `content_bundle_path`, `template_name`, and `output_dir`; may also include `captable_workbook_path`, `ownership_workbook_path`, and `financial_metric_labels` (the four Financial Summary tile names, selected by the `financial-summary` stage)
- `$STAGE_OUTPUTS` — path where this stage must write its structured handoff
- `$DEAL_DIR` — deal directory root

## Workflow

1. Read `$STAGE_INPUTS`.
2. Resolve `template_name` under `${CLAUDE_PLUGIN_ROOT:-./infor-beta}/templates/`. Both deliverables now resolve to `INFOR Slide Library.pptx`.
3. Inspect the `SlidePlan.deliverable_type`.
4. If `earnings-update`, call `assemble_earnings_update_deck(...)` and pass `captable_workbook_path` when supplied so slide 2's cap-table placeholder is replaced.
5. If `pitch`, call `assemble_pitch_deck(...)`.
6. Write the deck under `$DEAL_DIR/artefacts/` when `$DEAL_DIR` is set; otherwise use the supplied `output_dir`.
7. **Overflow QA — mandatory, do not skip** (see below). Render the overflow-prone slides to PNG, read each PNG, and autofit until text is clean. This stage is not complete until the QA has run; if the renderer is unavailable, say so explicitly rather than skipping silently.
8. Write `$STAGE_OUTPUTS` as:

```json
{
  "deck_path": "/absolute/path/to/Pitch Deck - Company.pptx"
}
```

## Pitch boundaries

- Slide 1: client name/date only.
- Slide 2: executive summary bullets — square main / dash sub bullets, template body colour (not the navy list default).
- Slides 3–5: static credentials, do not touch.
- Slide 6: section divider labels.
- Slide 7: company overview bullets; the cap table is pasted into the `Rectangle 3` placeholder when `captable_workbook_path` is supplied (`slide_index=6`, range `B15:F40`), and the `[x]$MM` footnote token is set to the cap table's output currency (`US` / `C`, read from `F5`). The LTM revenue pie stays a deferred placeholder.
- Slide 8: financial metric **NAME** labels only, taken from the `financial_metric_labels` stage input (the `financial-summary` stage's output, not the content bundle); charts stay placeholders.
- Slide 9: ownership slide (Canadian public targets) — **follows the Financial Summary slide**. When `ownership_workbook_path` is supplied, the left **"Insiders"** placeholder (`Rectangle 1`) is replaced by a picture of the ownership workbook's Select-Insiders block (sheet `Ownership`, range `B4:G17`, fixed `slide_index = _OWNERSHIP_SLIDE_INDEX = 8`). The right **"Institutions"** side stays a Bloomberg-sourced placeholder.
- Slide 10: concise risks/mitigants + tagline.
- Slide 11: comps takeaway; chart placeholder stays unless insertion replaces it later.
- Slide 12: key investment highlights (filled when content supplies them).
- Slides 13+: potential market-entry targets — the fixed 12-row comparison table (Overview / HQ / Year Founded → 7 industry metrics → Scale KPIs / Strategic Rationale), **two targets per slide**. The assembler clones the library's market-entry slide to `ceil(len(market_entry_targets) / 2)` slides, titles them `(N of M)`, writes the label column white at 11 pt and target values at 10 pt, blanks the unused column + logo on an odd final slide, and — after the cells are filled — clamps each table to a fixed **5.71"** total height (frame + proportional row heights) so long content can't run the table off the slide. Disclaimer/contact follow.

## Overflow QA (mandatory)

Text overflow (bullets spilling past a divider, a metric label wrapping onto a
third line) is invisible to python-pptx — the XML is valid, the text just
doesn't fit. After assembling the deck you **must** render the overflow-prone
slides to PNG and visually inspect them — this step is not optional and the
stage is incomplete without it:

- **Earnings update**: render slides 2 and 3, and check specifically for:
  - Slide 2 — the company-overview bullet block must not run under or overlap
    the **"LTM Revenue Breakdown"** title or the pie/cap-table region.
  - Slide 3 — every **Financial Highlights** metric tile must show its value and
    label on at most two lines with no text touching the tile edge. A label like
    `Cloud Services & Subscriptions Revenue` that wraps to a third line or clips
    is a failure: shrink it via `enable_normal_autofit`, and if it still
    overflows, the label is too long — the content stage must abbreviate it
    (e.g. `Cloud Services & Subscriptions Rev.`).
  - Both slides — confirm no figure reads as `US,…` / `C,…` (a dropped `$` and
    digit from a regex-substitution bug upstream); every dollar figure must
    start with a plain `$`.
- **Pitch**: render slides 2, 7, 10 (executive summary, company overview,
  risks/mitigants), the ownership slide (9), and every market-entry slide (13+).
  On slide 2 confirm the body text is the template dark colour with square/dash
  bullets (not blue); on slide 7 confirm the cap-table picture landed and the
  footnote currency is correct; on the market-entry slides confirm no row label
  or value clips and no data row is blank; on the ownership slide confirm the
  insider picture landed on the left and the institutional side is still a
  placeholder.

```python
import sys, os
sys.path.insert(0, os.environ.get("CLAUDE_PLUGIN_ROOT", "./infor-beta") + "/scripts")
from slide_render import render_deck_to_png
from pptx_helpers import enable_normal_autofit
from pptx import Presentation

pngs = render_deck_to_png(deck_path, output_dir, slide_indices=[1, 2])  # zero-based
```

Read each PNG with the `Read` tool and check for text touching or crossing a
shape boundary. If a shape overflows, reopen the deck, shrink the offending
shape with `enable_normal_autofit(shape, font_scale=...)` (start ~0.9 and step
down), save, re-render, and re-inspect. Repeat until every slide is clean.

The renderer uses PowerPoint COM on Windows and LibreOffice headless elsewhere,
so it works in Cowork. If neither backend is available it raises `RuntimeError` —
note that overflow QA could not run rather than claiming the deck is clean.

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
