# Slide 3 — Earnings Summary (Detailed Rules)

Slide 3 is indexed 2. This is the densest slide. Each sub-region below maps to specific shapes by name.

## 6a — Title and Section Headers

| Shape | Update |
|-------|--------|
| `Title 1` (PLACEHOLDER) | `<Company Name> Q<x> 202<x> Earnings Summary` (current quarter) |
| `Rectangle 7` (top-right section header) | `Q<x> 202<prior> vs. Q<x> 202<current> Financial Highlights` (e.g., `Q4 2024 vs. Q4 2025 Financial Highlights`) |
| `Text Placeholder 1` (footnote, line 2) | Replace `[x]` with reporting currency code |

Leave section header shapes `Rectangle 5` ("Business Updates"), `Rectangle 10` ("Broker Estimates vs Actuals"), and `Rectangle 11` ("Management Guidance") unchanged.

## 6b — KPI Tiles (top-right quadrant)

The template ships with four KPI rows, defaulted to Revenue / Net Income / Operating Cash Flow / Gross Margin. **Pick the four metrics that best reflect this company's story** — not necessarily the defaults. Examples:
- Asset managers: AUM, Management Fees, Adj. EBITDA, Operating Margin
- Energy: Production (boe/d), Netback, FFO, Net Debt
- Banks: Net Interest Income, PCL, ROE, CET1 Ratio
- REITs: FFO/share, SSNOI Growth, Occupancy, Debt/EBITDA
- SaaS: ARR, Net Revenue Retention, Adj. EBITDA, FCF Margin

Each of the four rows has six shapes. Map by vertical row:

| Row | Prior box | Current box | Delta $ box | Delta % box / 2nd delta | Left triangle | Right triangle |
|-----|-----------|-------------|-------------|-------------------------|---------------|----------------|
| 1 | `Rectangle 1032` | `Rectangle 1034` | `Rectangle 1041` | — | `Isosceles Triangle 1039` | `Isosceles Triangle 1040` |
| 2 | `Rectangle 1043` | `Rectangle 1037` | `Rectangle 1042` | — | `Isosceles Triangle 1045` | `Isosceles Triangle 1046` |
| 3 | `Rectangle 1035` | `Rectangle 1036` | `Rectangle 1061` | — | `Isosceles Triangle 1062` | `Isosceles Triangle 1063` |
| 4 | `Rectangle 1057` | `Rectangle 1058` | `Rectangle 1064` | — | `Isosceles Triangle 1065` | `Isosceles Triangle 1066` |

For each row, set:
- Prior box line 1: prior-period value (e.g., `$45.4MM`). Wrap negatives in parentheses: `($8.7MM)`.
- Prior box line 2: `Q<x> 202<prior> <Metric Name>` (e.g., `Q4 2024 Revenue`).
- Current box line 1: current-period value.
- Current box line 2: `Q<x> 202<current> <Metric Name>`.
- Delta box line 1: signed delta with `+` or `-` sign (e.g., `+$4.2MM`, `-$0.3MM`, `+22.6%` for margin deltas).

**Triangle orientation — NEVER rotate.** The gold triangles ship in the template's intended orientation. Do **not** set `shape.rotation` on any of the `Isosceles Triangle 10xx` shapes, ever — regardless of whether the metric moved up, down, or is a "good vs. bad" direction-flipped metric like charge-offs or expenses. Negative deltas are signalled by the `-` sign / parentheses on the delta value itself; the triangles are decorative.

**Delta-box font size — fixed 10 pt.** All four delta rectangles (`Rectangle 1041`, `1042`, `1061`, `1064`) use Palatino Linotype at a fixed **10 pt** — **minimum and only** size. Every delta box on the slide must match. Do not step down to 9 pt. If a delta string doesn't fit in the ≈1.0–1.1 in wide box at 10 pt, **shorten the number format** before anything else:
- `+$0.9B` instead of `+$911MM`
- `+$1.2B` instead of `+$1,234MM`
- `+14.6%` instead of `+1,460 bps` (bps is banned anyway — see next rule)

If the formatted number is already compact and still doesn't fit, drop the trailing `MM`/`B` unit (readers infer from the prior/current boxes). Never decrease font size below 10 pt.

**Delta-box color — green up, red down.** Color the delta value text based purely on the **direction of movement**, not whether the change is "good" or "bad" for the company:
- Positive delta (current > prior) → **green** `#00B050`
- Negative delta (current < prior) → **red** `#C00000`
- Zero change → leave template default (black)

This applies even when the "good" direction is inverted — for charge-off rates, cost ratios, or expense metrics, a positive delta is still green because the metric went up. The color reflects arithmetic direction, not business interpretation. Set the color explicitly via `run.font.color.rgb = RGBColor.from_string("00B050")` / `"C00000"`.

**Percent deltas only — never BPS.** For margin / rate / yield metrics (gross margin, charge-off %, CET1 ratio, occupancy, etc.), the delta value is always in percentage points expressed as `%`, not basis points. Write `+14.6%`, not `+1,460 bps`. Compute delta as `current_pct − prior_pct` and format with one decimal and a `%` suffix. This applies across every delta box in all four rows.

**Currency prefix — always plain `$`.** Do NOT use `C$`, `US$`, `€`, etc. inside any value on the slide — the footnote already scopes the currency. Write `$406.3MM`, `+$4.2MM`, `($8.7MM)`. For percentages, no prefix. For counts (e.g., production volumes), use the appropriate unit (`boe/d`, `MMcf/d`, `MW`, etc.).

## 6c — Business Updates Bullets (top-left)

Target shape: `TextBox 1067` at L≈0.35, T≈1.44. Replace all `[x]` bullets with concise bullets covering the quarter's operational story.

**Must fit inside the box.** The Business Updates area runs from T≈1.44 to T≈4.13 (the Broker Estimates section header starts at T=4.18). That's ~2.69 in of vertical space at Palatino 10 pt. **Bullets may run 2, 3, or 4 lines** — vary by topic, don't force every bullet to 2 lines (v1.9.15 did this and read as mechanical).

**Caps — enforce programmatically:**
- **Bullet count:** 4–6 depending on how the char budget is distributed
- **≤ 250 characters per bullet** (≈ 4 lines at Palatino 10 pt in a 4.53 in column)
- **≤ 900 characters total** across all bullets — text must end above T≈4.13
- **No trailing periods.** Bullets are fragment-style; do not end with `.` or `;`

Claude cannot visually check overflow — the caps are the only defense. Line math at Palatino 10 pt in a 4.53 in column:

- ≤ 70 chars → 1 line
- 71–140 chars → 2 lines
- 141–210 chars → 3 lines
- 211–280 chars → 4 lines

Budget: each bullet takes its line count × 0.17 in + 0.04 in spacing. Don't let the sum exceed 2.55 in (leave ~0.1 in buffer above the section header).

**Bullet formatting — mandatory.** All bullets must use the INFOR square bullet character at **level 0 (main)**, **Palatino Linotype 10 pt**. The template's paragraph 0 has the correct `pPr` (bullet character, indent, spacing). When you add paragraphs beyond paragraph 0, python-pptx creates them with an empty `<a:pPr/>` and no `rPr` — the result is Calibri with **no bullet character**. This is the bug that left bullets 3+ on goeasy without bullets.

Fix: use `write_bulleted_shape(shape, items)` with items as `[(text, 0), (text, 0), ...]` — all bullets at level 0, same square glyph. The helper asserts post-write that every paragraph has a `buChar` — fails loudly if the pPr got dropped. Never hand-roll `tf.add_paragraph()` + `p.text = "..."`.

After writing, re-open the deck with python-pptx and measure the bullet text length. If 5 bullets × ~30 words still visually overflows (a safe proxy is total rendered characters > ~800 for this shape at 10 pt), drop to 4 bullets or shorten wording — do not shrink font below 10 pt.

**Content style — narrative, not metric listings.** The Business Updates section on the LEFT should read like a pithy quarterly narrative, not a bulleted list of financial metrics. The RIGHT-side KPI tiles and Broker table already carry the numbers. The LEFT side is for:

1. **Key events / decisions of the quarter** — strategic actions, management changes, capital allocation choices, one-time charges, impairments, guidance revisions, segment wind-downs, major contracts won/lost, acquisitions or divestitures
2. **Operating commentary on segment performance** — "why" the numbers moved, not "what" the numbers are (the tiles show "what"). One bullet per material segment is typical.
3. **Forward outlook** — guidance for next quarter / full year, upcoming catalysts (product launches, facility ramps, regulatory decisions), capital plan

Aim for roughly **2 event bullets + 1–2 segment commentary bullets + 1 outlook bullet**. Light numerical anchoring is fine (one or two figures per bullet for context) but a bullet that is *primarily* a metric statement — e.g., *"Revenue of $406.3MM was effectively flat vs. Q4 2024 (-0.2%)"* — belongs on the right side, not the left. Instead, write *"Revenue stayed flat as yield compression from the 35% rate cap and LendCare charge-offs offset growth in the easyfinancial direct channel"*

Sourcing priority:
1. Earnings call transcript (if attached) — highest quality, management's own framing
2. Earnings press release (WebSearch: `"<Company Name>" Q<x> 202<x> earnings press release`) — use the company's site or a wire that reproduces the company release verbatim
3. MD&A "Highlights" or "Overview" section

Each bullet should:
- Lead with the driver or event (not the metric)
- Explain cause and consequence, not just the fact
- Be written in full prose sentences — no sentence fragments or metric-first phrasing

Good example (Village Farms): *"International medical exports surged on strong demand across European markets, with Germany and the UK both scaling materially."*

Bad example (goeasy v3): *"Gross consumer loans receivable grew 19.8% YoY to $5.51B, with Q4 originations of $951.5MM (+16.9% YoY) demonstrating continued robust customer demand"* — reads as two metrics and a platitude; belongs on the right side.

Include one bullet covering forward outlook / management guidance for the coming year.

## 6d — Broker Estimates vs Actuals Table (bottom-left)

Target shape: the PPTX table (shape type 19) on slide 3 — the one with header row `['Figures in x$MM', 'Reported', 'Bloomberg Estimate', 'Variance']`.

Update the header cell `Figures in x$MM` → `Figures in <currency code>$MM` (e.g., `Figures in US$MM`, `Figures in C$MM`).

**Font — mandatory.** Every cell in the table (header row and all metric rows) must be **Palatino Linotype at 9 pt**. After writing each cell, explicitly iterate every run and set `run.font.name = "Palatino Linotype"` and `run.font.size = Pt(9)`. Do not trust inherited formatting — PowerPoint's default run fallback is Calibri, which has been observed to slip in when cells are re-written.

**Always exactly 5 metric rows — never delete.** The table must always have 5 populated metric rows (6 rows total including header). **Never** write `N/A`. **Never** delete rows. If the Bloomberg EEO snip only shows consensus for 3 of the template's default metrics, **swap** the other two row labels for different metrics the snip *does* cover — pick substitutions so all 5 rows have real Reported / Estimate / Variance values.

The template's 5 default metric rows are Revenue, Gross Margin, Adj. EBITDA, Net Income, UPS. Change the label on any row where the EEO snip lacks consensus. Common substitute metrics by sector:

| Sector | Preferred 5 metrics (all typically on EEO) |
|--------|-------------------------------------------|
| Consumer / industrial | Revenue, Gross Margin %, Adj. EBITDA, Adj. EBITDA Margin %, Adj. EPS |
| Financials / lenders | Revenue, PPNR, Net Income, ROE %, EPS |
| Asset managers | Revenue, Adj. EBITDA, AUM, Management Fees, Adj. EPS |
| Energy | Revenue, Adj. EBITDA, Production, Netback, CFPS |
| REITs / yield vehicles | Revenue, NOI, FFO, AFFO / share, Occupancy % |
| SaaS | Revenue, ARR, Adj. EBITDA, FCF, Adj. EPS |

Before writing, read the EEO snip and inventory which metrics the snip actually shows consensus for. Build a list of 5 metrics all present on the snip — combine template defaults + sector substitutes. Only then begin writing.

If the EEO snip genuinely shows fewer than 5 metrics (rare — most EEOs cover 6-8), use the lowest-quality filler from the sector row above that the analyst can still manually source (e.g., street consensus from a research note) — but prefer to expand your read of the snip first. Never leave a cell empty or `N/A`.

Populate each row:

- **Reported** — from the 10-Q/10-K or the company's earnings press release
- **Bloomberg Estimate** — read directly from the EEO snip
- **Variance** — `Reported − Estimate`

**Number formatting — include $ or % on every value.** Bare numbers like `406.3` or `(8.93)` are ambiguous — even though the header row says `Figures in C$MM`, every individual cell should carry a `$` prefix or `%` suffix that matches the metric.

| Metric type | Reported / Estimate / Variance format | Examples |
|-------------|---------------------------------------|----------|
| Dollar (revenue, EBITDA, net income) | `$` prefix, one decimal, `()` for negatives | `$406.3`, `($121.1)`, `($31.2)` |
| Per-share (EPS, FFO/share, AFFO/share) | `$` prefix, two decimals, `()` for negatives | `$1.24`, `($8.93)`, `$0.06` |
| Margin / rate / ratio (%) | `%` suffix, one decimal, `-` prefix or `()` for negatives | `38.7%`, `(2.3%)`, `+0.8%` |
| Volume / count (production, units, AUM in $B) | Unit suffix or prefix matching the metric | `950 MMcf/d`, `$12.4B AUM`, `1.6M customers` |

Variance keeps the same prefix/suffix as the value: a `+$1.2` or `($31.2)` variance on a dollar metric; `+0.8%` or `(2.3%)` on a margin. Never a bare `(31.2)` with no `$` or `%`.

**Variance column color.** Color the **Variance** cell text based on the sign — same scheme as the KPI delta boxes on the right side of the slide:
- Positive variance (beat) → **green `#00B050`**
- Negative variance (miss) → **red `#C00000`**
- Zero → template default (black)

Set on the variance cell's run via `run.font.color.rgb = RGBColor.from_string("00B050" if variance >= 0 else "C00000")`. The Reported and Bloomberg Estimate cells stay template-default (black). Only the Variance column is colored.

**Currency-in-header vs currency-on-value.** Even with `Figures in C$MM` in the header, each cell carries `$` (not `C$` — the header already scopes it to Canadian). Never `C$406.3` in-cell; just `$406.3`. For an EPS row in the same table, `$1.24` implicitly means C$1.24 because the header already said so.

## 6e — Management Quotes (bottom-right)

Two quote groups on slide 3: `Group 1070` and `Group 1086`. Each group contains three shapes — the swoosh freeform (leave alone), a quote TextBox, and an attribution TextBox.

| Group | Quote TextBox | Attribution TextBox |
|-------|---------------|---------------------|
| `Group 1070` | `TextBox 1072` | `TextBox 1073` |
| `Group 1086` | `TextBox 1088` | `TextBox 1089` |

Set the quote TextBox to a 1–3 sentence management quote, wrapped in curly quotes `"..."` (the template uses `"Quote"` with slanted double quotes). Set the attribution TextBox to `<Name> – <Role>` (en-dash, not hyphen) — e.g., `Michael DeGiglio – CEO`.

**Formatting preservation — just call `set_text(shape, [line])`.** The template ships the quote TextBoxes (`TextBox 1072`, `TextBox 1088`) with **Palatino Linotype 10.5 pt italic**, and the attribution TextBoxes (`TextBox 1073`, `TextBox 1089`) with **Palatino Linotype 10.5 pt bold**. The `set_text` helper preserves the template's run-level `rPr` (font, size, italic, bold, color) automatically — just mutate the text without passing `size_pt` / `color_hex`. Do not add an explicit `run.font.italic = True` or similar; trust the template.

**Quotes must address THE key item of the quarter — not general strategy.** The section header says "Management Guidance" — the analyst wants to see management's own words about whatever made this quarter noteworthy: the goodwill impairment, the credit-loss spike, the guidance cut, the reorg, the acquisition close, the margin inflection. Generic confidence language ("we're committed to long-term shareholder value", "we have a clear plan to execute with urgency") fails this test even if the CEO said it verbatim.

Before picking quotes, identify the **single most important event or result** driving this quarter's narrative. That is usually:
1. The largest negative surprise vs. BBG consensus (from Slide 3 Broker table), OR
2. The largest one-time charge or impairment (from MD&A / press release), OR
3. The biggest operational shift (segment wind-down, management change, guidance revision)

Then find management's words on **that specific item**. Example for goeasy Q4 2025 (LendCare goodwill impairment + charge-off spike):
- Good: *"We are taking decisive action on LendCare — accelerating the integration, right-sizing the credit book, and recognizing the $160MM goodwill impairment today rather than extending the problem into 2026."*
- Bad (too generic): *"My top priority is to ensure we manage credit well and deliver the strong performance we expect of ourselves."*

If the transcript / press release doesn't contain a direct quote addressing the key item, it is acceptable to use a closely adjacent quote (e.g., one quote on the item + one on the response plan) — but do not fall back to generic mission-statement language.

Sourcing:
- If earnings call transcript attached → pick the two most insight-dense quotes that directly address the key item (CEO + CFO typically). Scan Q&A as well as prepared remarks — analyst questions often pull out the most pointed management responses.
- Otherwise → WebSearch for `"<Company Name>" Q<x> 202<x> earnings press release` and pull direct quotes from the company's press release. If the press release lacks a pointed quote, WebSearch for a post-earnings interview or sell-side note that quotes management.

**Hard quote length caps — enforce programmatically:**
- **≤ 30 words per quote**
- **≤ 200 characters per quote** (including curly quote marks)

goeasy v1.9.13's CFO quote was 52 words / 351 chars and overflowed the group box. Each quote group (`Group 1070`, `Group 1086`) is 4.51–4.53 in wide × 1.16–1.18 in tall. At Palatino 10.5 pt italic the text wraps at roughly 55 chars/line; with the swoosh freeform taking ~0.95 in of vertical space, the quote gets at most ~4 lines before visibly overflowing the group. 200 chars ≈ 3.6 lines fits safely.

Assert `len(q) <= 200 and len(q.split()) <= 30` for each quote before writing. If a transcript pull is too long, either trim to the single most pointed clause or use an ellipsis (`"...it reflects our conviction that..."`).

## 6f — Performance Summary Box (bottom-left, below the table)

Target shape: `Rectangle 1111` at L≈0.35, T≈6.19 — the gold/gilded summary box. Dimensions: **4.53 in wide × 0.63 in tall**.

**Must fit inside the box.** At 11 pt Palatino Linotype the box holds roughly **2 lines ≈ 28 words / 170 characters** before overflowing. Write **one sentence ≤ 25 words / 150 characters** summarizing the quarter's overall performance relative to expectations. The sentence should:
- Mention whether the company beat, missed, or met Bloomberg consensus (pull from 6d)
- Include one qualifier highlighting the underlying story (growth, margin, mix, one-time items)
- Be readable as a standalone callout

Good example (Village Farms, 20 words): *"Village Farms reported metrics that were below Bloomberg estimates, but results still demonstrated strong growth and margin profile"*

Bad example (goeasy, 26 words, overflowed): *"Q4 2025 results were dominated by a $159.6MM LendCare goodwill impairment and incremental $177.9MM of charge-offs"* — too many specific figures for the box; move concrete numbers to the bullets and keep the summary high-level.

If your draft exceeds 25 words, rewrite more tightly — do not shrink the font below 11 pt. Verify with `len(sentence.split()) <= 25 and len(sentence) <= 150` before writing.
