# Common Pitfalls — Quick Reference

| Issue | Guidance |
|-------|----------|
| Macabacus placeholder on slide 2 | Leave `Rectangle 4` (`[Macabacus Placeholder]`) alone — analyst links the cap table here via the Macabacus add-in |
| Currency | Read from filing "Basis of Presentation" — don't infer from exchange (VFF is NASDAQ but reports US$; BMO is NYSE-listed but reports C$) |
| Currency on values | **Plain `$` on every slide value** — never `C$` / `US$` / `€` etc. The footnote (`Note: All figures in C$MM...`) and the Broker table header cell (`Figures in C$MM`) carry the currency code; values just say `$406.3MM`, `+$4.2MM`, `($8.7MM)`. |
| KPI metric choice | Pick 4 metrics that reflect the *company's* story, not the template defaults |
| Triangle rotation | **NEVER** rotate triangles. Leave every `Isosceles Triangle 10xx` at its template rotation. Delta sign goes on the value, not the arrow. |
| Delta font size | **Fixed 10 pt**, all four delta boxes. No step-down. If text doesn't fit, shorten the number format (`+$0.9B` not `+$911MM`) — never drop below 10 pt. |
| Delta color | Positive delta → green `#00B050`; negative delta → red `#C00000`. Direction-based, not "good/bad" — charge-off rate going up is green. |
| Rate / margin deltas | Always `%` (e.g., `+14.6%`), never `bps` (`+1,460 bps` is wrong) |
| `set_text` must preserve formatting | The helper mutates `paragraph.runs[0].text` in place (preserving the template's `rPr` — font, size, bold, italic, color). It only creates fresh runs for brand-new paragraphs beyond the template's seed count, and when it does, it grafts a copy of paragraph 0's `rPr` onto the new run. **Do not pass `size_pt` / `color_hex` unless you intentionally want to override.** Overriding on shapes like `Title 1`, `Rectangle 7`, `Rectangle 1111`, `Subtitle 2` (date), or the quote/attribution TextBoxes would wipe out the template's bold/italic/color formatting. |
| Bullet formatting (slide 2 + slide 3 business updates) | Use `write_bulleted_shape(shape, items)` — single call that harvests both level `pPr` templates before wiping, inserts a deepcopy on each new paragraph, sets explicit Palatino run font/size, and asserts post-write that every paragraph has a `buChar`. Hand-rolled `tf.add_paragraph()` + `p.text = "..."` produces empty `<a:pPr/>` → no bullet glyph renders. |
| Slide 2 main vs. sub bullets | Main bullets (level 0) = 10.5 pt, sub-bullets (level 1) = 10 pt. Description bullets use both levels to group segment-level detail under summary bullets. |
| Business Updates overflow | Hard cap at 5 bullets (4 preferred), ≤30 words each, 10 pt Palatino, **all at level 0** (square bullet). Text must end above T≈4.13 to not collide with the Broker Estimates header at T=4.18. |
| Broker table font | Every cell forced to Palatino Linotype 9 pt via `run.font.name` + `run.font.size` — do not trust inherited formatting |
| Broker table — always 5 rows | **Never delete rows. Never write N/A.** If the EEO snip doesn't cover a template default metric, swap that row's label for a different metric the snip *does* cover. 5 real rows, no exceptions. |
| Management quote focus | Quotes must address THE key item of the quarter (the largest surprise, charge, or inflection) — not generic strategy language. If the transcript/press release lacks a pointed quote on the key item, expand the search (Q&A section, post-earnings interviews). |
| Business Updates tone | Narrative prose, not metric listings. Left side is events + segment commentary + outlook; right side carries the numbers. A bullet that is primarily a metric (`"Revenue grew 19.8% YoY to $5.51B"`) belongs in the KPI tiles, not here. |
| Slide 2 density | **7–12 bullets, 1,200–1,500 chars total, ≤250 chars per bullet.** Let bullets run 2–4 lines each — variable length looks natural; uniformly 2-line bullets look mechanical (v1.9.15). Height budget 5.4 in; overflow past footer at T=7.03 is the failure mode. |
| Slide 2 content focus | Description is a durable company profile — what the company DOES. Do NOT include quarterly earnings performance ("FY 2025 revenue of ...", "Q4 net loss of ..."). That belongs on slide 3. |
| Segment bullet bolding | Sub-bullets naming a segment use a bold `SegmentName:` prefix + regular rest. Pass the item to `write_bulleted_shape` as a 3-tuple `(bold_prefix, rest, level)` — two-run paragraph. |
| Slide 3 Business Updates density | **4–6 bullets, ≤250 chars each, ≤900 chars total.** Bullets may run 2–4 lines. Height budget 2.55 in; must end above Broker Estimates header at T=4.18. |
| No trailing periods | Bullets on slide 2 (description) and slide 3 (business updates) are fragment-style — no trailing `.` or `;`. Asserts fail the run if any bullet ends with punctuation. |
| Broker table number format | Dollar metrics: `$` prefix, 1 decimal, `()` for negatives (`$406.3`, `($121.1)`). Per-share: `$` prefix, 2 decimals. Margin/rate: `%` suffix, 1 decimal. Never a bare `406.3` with no `$` or `%`. |
| Broker variance color | Variance cell text colored green (`#00B050`) on positive / beats, red (`#C00000`) on negative / misses. Reported and Estimate cells stay black. |
| Quote length | **≤200 chars / ≤30 words per quote.** goeasy v1.9.13 CFO quote was 52 words / 351 chars and overflowed the 1.18 in group box. |
| Visual overflow detection | Claude cannot render PPTX; hard char caps are the primary defense. Optional: `soffice --convert-to pdf` + Read the PDF to spot-check after writing. |
| Gold summary box overflow | ≤25 words / ≤150 chars. Keep high-level; move specific figures to the bullets |
| Negative numbers | Wrap in parentheses: `($8.7MM)`, not `-$8.7MM` — consistent with template's Village Farms example |
| Curly quotes | Use `"..."` (U+201C / U+201D), not straight `"..."` — preserves template typography |
| Attribution dash | Use en-dash `–` (U+2013), not hyphen `-` |
| Quarter label | `Q4 2025`, not `Q4'25` or `4Q25` — template uses full format |
| EEO variance sign | Reported − Estimate — a beat is `+`, a miss is `-` (or parentheses) |
| Slide 4 / 5 | NEVER modify — disclaimer and contact page are fixed |
| Cap table | Invoke captable-infor workflow as Step 8; save alongside but do not embed |
