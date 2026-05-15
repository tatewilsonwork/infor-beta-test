---
name: earningsupdate-infor
description: >
  Use this skill when the user invokes /earningsupdate-infor or asks to build a quarterly earnings
  update deck for a public company. Populates the INFOR Earnings Update Template with a company
  overview, four KPI tiles (actual vs prior-year quarter), a Broker Estimates vs Actuals table
  sourced from a Bloomberg EEO snip, business-update bullets, management quotes, and a short
  performance summary. Activates on "earnings update", "earnings deck", "quarterly earnings",
  "earnings summary deck", or any request to build a branded update deck off a recent 10-Q/10-K
  and Bloomberg EEO snip.
version: 0.2.0
allowed-tools: [Read, Bash, Write, Glob, WebSearch, WebFetch]
---

# INFOR Earnings Update — Workflow

This skill builds a branded 5-slide earnings update deck from a company's most recent quarter of financials, a Bloomberg EEO screenshot, and optionally an earnings call transcript. It also produces a companion capitalization table XLSX that the analyst manually inserts into the deck in place of the Macabacus placeholder on slide 2.

Today's date is available from the system context (`currentDate`) — do not shell out to `date`. Earnings template, cap table template, and working directory are resolved inline in their respective steps.

**Detailed references** (loaded on demand — not in this file's main flow):
- [`references/slide2-company-overview.md`](references/slide2-company-overview.md) — bullet structure, density caps, segment-name bolding, content focus rules for Step 5
- [`references/slide3-earnings-summary.md`](references/slide3-earnings-summary.md) — KPI tiles, broker table, management quotes, summary box rules for Step 6
- [`references/python-implementation.md`](references/python-implementation.md) — reference python-pptx driver (imports helpers from the plugin's shared `pptx_helpers` module)
- [`references/common-pitfalls.md`](references/common-pitfalls.md) — quick-reference table of failure modes and fixes

The shared formatting helpers — `set_text`, `write_bulleted_shape`, `set_cell_text`, `find_shape`, brand constants — live at [`infor-beta/scripts/pptx_helpers.py`](../../scripts/pptx_helpers.py) and are imported in [`references/python-implementation.md`](references/python-implementation.md).

## Conductor-mode handoff (read first when running under the conductor)

When invoked as a stage of a conductor plan, the environment carries `$STAGE_INPUTS`, `$STAGE_OUTPUTS`, and `$DEAL_DIR`:

- Read company/ticker/reporting-quarter/comparison-quarter/EEO-snip-path from `$STAGE_INPUTS` instead of asking the analyst at Step 1. Skip the Step 1 prompt entirely.
- Write the deck to `$DEAL_DIR/artefacts/Earnings Update - <SANITIZED>.pptx` (NOT cwd). Bootstrap `$DEAL_DIR/artefacts/` if it doesn't exist.
- **Skip Step 8 (Generate Companion Cap Table) entirely.** The earnings-update conductor plan dispatches `captable-infor` as a sibling stage; the conductor handles composition, not this skill. If you invoke captable from within this skill while running under the conductor, you'll produce a duplicate.
- At the end of Step 9 (Save and Verify), write the structured handoff:
  ```bash
  python -c "import json,os; json.dump({'deck_path': os.environ['OUTPUT']}, open(os.environ['STAGE_OUTPUTS'], 'w'))"
  ```
  The conductor will not proceed past this stage until `$STAGE_OUTPUTS` exists and parses as JSON.

When `$STAGE_OUTPUTS` is **unset** (direct `/earningsupdate-infor` invocation), follow the workflow below as-is — including Step 8 — and output lands in cwd.

---

## Workflow Steps

### Step 1 — Collect Inputs

Required:
- Company name and CapIQ ticker (e.g., `NasdaqGS:VFF`)
- Reporting quarter (e.g., Q4 2025) and prior-year comparison quarter (Q4 2024)
- Most recent 10-Q / 10-K / annual report + MD&A
- Bloomberg EEO snip (image/screenshot showing consensus vs. actual for the quarter)

Optional but highly useful:
- Earnings call transcript
- Company earnings press release URL (if not web-searchable from company name + quarter)

If any required input is missing, ask in a single message:

> "To build the earnings update, I need:
> - **Company name + CapIQ ticker** (e.g., NasdaqGS:VFF)
> - **Reporting quarter** (e.g., Q4 2025)
> - **10-Q / 10-K / MD&A** attached
> - **Bloomberg EEO snip** attached (image)
>
> Optional: earnings call transcript (attach) for extra detail."

Wait for all required inputs before proceeding.

---

### Step 2 — Locate and Copy the Template

Resolve the earnings template via the plugin's shared helper:
```bash
bash "${CLAUDE_PLUGIN_ROOT:-./infor-beta}/scripts/find_template.sh" "INFOR Earnings Update Template.pptx"
```

Sanitize the company name via the shared helper, then copy:
```bash
SANITIZED_COMPANY=$(bash "${CLAUDE_PLUGIN_ROOT:-./infor-beta}/scripts/sanitize_name.sh" "$COMPANY_NAME")
TEMPLATE=$(bash "${CLAUDE_PLUGIN_ROOT:-./infor-beta}/scripts/find_template.sh" "INFOR Earnings Update Template.pptx")
OUTPUT="./Earnings Update - $SANITIZED_COMPANY.pptx"
cp "$TEMPLATE" "$OUTPUT" && echo "COPY_OK" || echo "COPY_FAILED"
ls -lh "$OUTPUT"
```

If the copy fails or the file is 0 bytes, STOP and tell the user:
> "I could not copy the INFOR Earnings Update Template. Please confirm the file `INFOR Earnings Update Template.pptx` exists in the `templates/` folder of the infor-beta plugin."

---

### Step 3 — Determine Reporting Currency

Read the 10-Q/10-K to identify the reporting currency. Village Farms reports in US$ despite being Canadian-listed — do NOT infer currency from exchange. Read the cover page or the "Basis of Presentation" footnote in the financial statements.

Output the currency code as one of `US$MM`, `C$MM`, `€MM`, `£MM`, `A$MM`, etc.

**Currency only lives in the footnote.** Use the currency code (`US$MM`, `C$MM`, etc.) **only** in:
- Slide 2 and Slide 3 footnote text (`Note: All figures in C$MM, except where indicated otherwise`)
- The Broker Estimates table header cell (`Figures in C$MM`)

**Everywhere else on the slide, values use a plain `$` prefix** — never `C$`, `US$`, `€`, etc. The footnote scopes the currency; repeating it on every value is redundant and visually noisy. Applies to KPI tile values/deltas, gold summary box, Business Updates bullets, Slide 2 description bullets, and Broker table cells.

For non-dollar currencies (€ / £ / ¥), use the symbol once in the footnote (`Figures in €MM`) and plain `$` prefix on values, or omit the prefix entirely.

---

### Step 4 — Populate Slide 1 (Cover)

Open the deck with python-pptx and update ONLY the date placeholder in the bottom-right. Leave the title, subtitle, logo, and all other cover elements alone.

Target shape: on slide 1 (index 0), the `Subtitle 2` PLACEHOLDER at L≈5.55in, T≈6.9in contains `[Current Month] 2026`.

Replace with `[Month YYYY]` using **today's month and year** — always the current month, never the earnings-release month. E.g., if today is 2026-04-23, use `April 2026`.

---

### Step 5 — Populate Slide 2 (Company Overview)

Slide 2 is indexed 1. Update three shapes:
- `Title 1` placeholder → `<Company Name> Overview`
- `TextBox 16` → 7–12 bullet company description (durable profile, not quarterly performance)
- `Text Placeholder 1` footnote line 2 → replace `[x]` with reporting currency code

**Do NOT touch `Rectangle 4`** — the Macabacus placeholder for the analyst-pasted cap table.

**Density caps (enforce programmatically):** 7–12 bullets, ≤250 chars/bullet, 1,200–1,500 chars total, no trailing periods. Vary 2–4 line bullets — uniform 2-line bullets read as mechanical.

Use `write_bulleted_shape(textbox16, description_items)` from [`pptx_helpers`](../../scripts/pptx_helpers.py). It harvests the template's seed pPr (level 0 main square bullet 10.5 pt, level 1 sub dash bullet 10 pt) BEFORE wiping, so bullet glyphs survive. Each item is `(bold_prefix, rest, level)` — `bold_prefix=""` for plain bullets, populated for bold `SegmentName:` sub-bullets.

See [`references/slide2-company-overview.md`](references/slide2-company-overview.md) for the full shape map, target structure of 12 bullets, segment-bolding examples, line-math, sourcing guidance, and the bullet-formatting bug this helper prevents.

---

### Step 6 — Populate Slide 3 (Earnings Summary)

Slide 3 is indexed 2 — the densest slide. Six sub-regions, each with its own shape map:

- **Title + section headers** (`Title 1`, `Rectangle 7`, footnote currency)
- **KPI tiles** (4 rows × 6 shapes — prior box, current box, delta box, two triangles per row)
- **Business Updates bullets** (`TextBox 1067`, 4–6 narrative bullets ≤900 chars total)
- **Broker Estimates table** (PPTX table — exactly 5 metric rows, every cell Palatino 9 pt, Variance column colored by sign)
- **Management Quotes** (`Group 1070` + `Group 1086`, each with quote + attribution TextBoxes — must address the key item of the quarter)
- **Performance Summary box** (`Rectangle 1111`, ≤25 words/150 chars, one sentence beat/miss + qualifier)

Key invariants (failure modes from prior runs):
- **NEVER rotate triangles** — direction lives in the `+`/`-` sign, not the arrow
- **Delta box font: fixed 10 pt** — shorten the number (`+$0.9B` not `+$911MM`) before stepping down
- **Delta color: positive=green `#00B050`, negative=red `#C00000`** by arithmetic direction (not "good/bad")
- **Rate deltas in `%` not bps** — `+14.6%`, never `+1,460 bps`
- **Broker table: always 5 rows** — swap labels for metrics the EEO snip covers; never `N/A`, never delete rows
- **Quotes ≤200 chars / ≤30 words** — overflowed groupboxes on goeasy v1.9.13
- **All on-slide values use plain `$`** — `C$` only in the footnote and table header

See [`references/slide3-earnings-summary.md`](references/slide3-earnings-summary.md) for the full shape maps (sub-sections 6a–6f), KPI metric substitutes by sector, broker table number formatting, quote selection rules, and the gold summary box budget.

---

### Step 7 — Leave Slides 4 and 5 Untouched

Slide 4 (Disclaimer) and Slide 5 (Contact — Neil Selfe + three placeholder tables) must remain exactly as shipped in the template. Do not touch them.

---

### Step 8 — Generate Companion Cap Table (skip when running under the conductor)

**If `$STAGE_OUTPUTS` is set, skip this step entirely.** The earnings-update conductor plan runs `captable-infor` as a sibling stage immediately after this one; the conductor composes the two skills. Running Step 8 here produces a duplicate cap table at a different path.

When invoked directly (no conductor — `$STAGE_OUTPUTS` unset), continue as below.

Once the deck is populated, invoke the **captable-infor** skill's workflow (Steps 2–8 of that skill) using the same 10-Q/10-K/MD&A attachments and the CapIQ ticker provided in Step 1. The cap table will be saved alongside the deck as:

```
./<SANITIZED_TICKER> - Capitalization Table.xlsx
```

The analyst will open this file, refresh CapIQ, and use Macabacus to link the cap table range into `Rectangle 4` (the Macabacus placeholder) on slide 2 of the deck. Do NOT try to embed the xlsx into the deck programmatically — manual insertion is intentional because the analyst refreshes CapIQ market data before linking.

---

### Step 8b — Visual Overflow Check (optional)

Claude **cannot directly render PowerPoint** to see whether text overflows a shape visually. The character / word caps in Steps 5 and 6 are the primary defense — enforce them with asserts before writing.

If the user explicitly requests a visual check, or if you've tightened wording to the caps and are unsure, convert the deck to PDF and read page 3 as an image:
```bash
soffice --headless --convert-to pdf "./Earnings Update - $SANITIZED_COMPANY.pptx" --outdir .
```
Then `Read` the resulting `.pdf` to check slides 2 and 3 for overflow. If `soffice` / LibreOffice is not installed, skip this step and trust the character caps.

---

### Step 9 — Save and Verify

Save the deck. Reopen it with python-pptx in read mode and verify:
- Slide 1: date placeholder no longer contains `[Current Month]`
- Slide 2: title no longer contains `[Client Name]`, description TextBox has no `[x]` placeholders, footnote has no `[x]` in currency string
- Slide 3: title has real quarter, top-right section header has real quarters, all four KPI rows populated, broker table has no blank cells, both management quote groups updated, summary box updated
- Placeholder `Rectangle 4` on slide 2 still reads `[Macabacus Placeholder]` (yes, on purpose — analyst pastes a screenshot of the cap table here using the Macabacus add-in)

If any placeholder `[x]` or `[Client Name]` or `Qx 202x` token remains anywhere on slides 1–3, fix it and re-save.

---

### Step 10 — Report to User

Output a brief summary:

1. **Deck file:** absolute path to saved `.pptx`
2. **Cap table file:** absolute path to saved `.xlsx`
3. **Reporting currency used:** e.g., `US$MM`
4. **KPIs selected for slide 3:** list the four metrics and why (one phrase each)
5. **Sources used for bullets / quotes:** which file or URL each major section came from
6. **Manual steps remaining:**
   - Refresh CapIQ in the cap table file and paste a screenshot over `Rectangle 4` on slide 2
   - Optionally swap the cap table paste for a Bloomberg tearsheet screenshot if the analyst prefers that layout
   - Review the Performance Summary box wording before sending

---

## Reference

The full Python driver, shape maps, content rules, and pitfall reference live in `references/` next to this file — load them on demand when working on the corresponding step. See [`references/common-pitfalls.md`](references/common-pitfalls.md) for the consolidated failure-mode table.
