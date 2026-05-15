# Slide 2 — Company Overview (Detailed Rules)

Slide 2 is indexed 1. Update these shapes by name:

| Shape name | Current value | Update to |
|------------|---------------|-----------|
| `Title 1` (PLACEHOLDER) | `[Client Name] Overview` | `<Company Name> Overview` |
| `TextBox 16` (L≈0.35, T≈1.45) | `[x]` / `[x]` | 7–9 bullet company description, filling the left column down to ~T=6.85 |
| `Text Placeholder 1` (footnote, line 2) | `Note: All figures in [x]$MM, ...` | Replace `[x]` with `US`, `C`, etc. |

**Do NOT touch `Rectangle 4` — the Macabacus placeholder at L≈5.12, T≈1.48** (renders as `[Macabacus Placeholder]`). The analyst pastes the cap table here using the Macabacus Excel-to-PowerPoint linker; leave the placeholder text as shipped in the template.

## Company Description Bullets — Style

See the Village Farms slide 2 left panel for tone. Available vertical space (T=1.45 → T=7.03) is ~5.58 in at Palatino 10.5 pt. **Bullets may run 2, 3, or 4 lines** — vary the length by topic. Uniformly-2-line bullets look mechanical; mix them up so shorter points (ratings, leadership) sit next to longer descriptive points (segments, history).

**Total-length budget:** **1,200–1,500 characters total** across all bullets, with a **max of 250 characters per bullet** (≈ 4 lines at Palatino 10.5 pt in a 4.53 in column). Bullet count can vary 7–12 depending on how the character budget is distributed. v1.9.13 overflowed with 1,641 chars; v1.9.14 under-filled with 852 chars; v1.9.15's rigid 120-char cap per bullet produced a uniform 2-line look that read as mechanical.

**No trailing periods.** Financial-deck bullets are fragment-style — do not end any bullet with `.` or `;`. Each bullet ends on its last word. Asserts in the reference code strip and flag trailing periods.

## Content Focus — What the Company DOES, Not How It Performed

The description is a durable company profile that could have been written any time in the last year. Do NOT include quarterly financial performance here (that belongs on slide 3). Specifically:
- ❌ "FY 2025 revenue of $1.7B, up 20% YoY, with 26.6% yield" — performance
- ❌ "Q4 2025 net loss of $337MM driven by LendCare impairment" — performance
- ✅ "~475,000 active customers, 2,600+ employees, 400+ locations" — durable scale
- ✅ "Senior unsecured notes rated B- (S&P) / B1 (Moody's)" — durable credit profile
- ✅ "Acquired LendCare in 2021 to extend into point-of-sale financing" — history

## Structure (12 bullets as a target)

1. Main (10.5 pt): Company name, founding year, exchange/ticker, one-sentence business description
2. Main (10.5 pt): Headquarters + geographic footprint
3. Main (10.5 pt): Scale anchors (employees, customers, locations, annualized volume)
4. Main (10.5 pt): Balance-sheet / book anchor (loan book size, AUM, production capacity, installed base)
5. Main (10.5 pt): Segment overview intro — "Operates X reportable segments" or "organized around Y business lines"
6–8. **Sub-bullets (10 pt)** — one per reportable segment. **Bold the segment name prefix** (see below).
9. Main (10.5 pt): Distribution channels / go-to-market (branch footprint, digital, wholesale, merchant partners)
10. Main (10.5 pt): Key historical transactions (acquisitions, divestitures, JVs) in chronological order
11. Main (10.5 pt): Strategic partnerships, licenses, or exclusive arrangements
12. Main (10.5 pt): Governance / leadership / credit rating / ESG note

Combine or drop slots if the filing doesn't support 12 real bullets, but aim for 10+.

## Segment-Name Bolding

On the sub-bullets naming each segment (e.g. `easyfinancial: direct-to-consumer unsecured...`), render the segment name + colon **bold** and the rest of the bullet regular weight. Two-run paragraph pattern — pass the item as a 3-tuple `(bold_prefix, rest, level)` to `write_bulleted_shape`. Examples of bold prefixes:
- `easyfinancial:` / `LendCare:` / `easyhome:` (goeasy)
- `Cannabis (Canada):` / `Cannabis (International):` / `Produce:` / `Clean Energy:` (Village Farms)
- `Commercial Banking:` / `Wealth Management:` / `Capital Markets:` (a bank)

## Line Math (for reference)

Shape width 4.53 in fits ~65 characters per line at Palatino 10.5 pt:
- ≤ 65 chars → 1 line (rare — use only for very short bullets like `"Dual-listed on NYSE and TSX"`)
- 66–130 chars → 2 lines
- 131–195 chars → 3 lines
- 196–250 chars → 4 lines

Budget your bullets to the 5.58-in available height. A rough rule: assume each bullet renders its line count × 0.18 in + 0.11 in spacing. Don't let the sum exceed 5.4 in (leave a small margin above the footer).

Claude cannot visually render the slide to check overflow; **enforce the total-char and per-bullet caps programmatically**.

## Sourcing

Source content from the MD&A's "Overview" / "Our Business" / "Operating Segments" sections, the 10-K Item 1 "Business" section, and the company website if needed via WebSearch. Keep each bullet tight — one idea per line.

## Bullet Formatting — Mandatory

The template ships `TextBox 16` with two seed paragraphs: paragraph 0 is a **main bullet** (Palatino Linotype 10.5 pt, square bullet character, `marL="180975" indent="-180975"`) and paragraph 1 is a **sub-bullet** (Palatino Linotype 10 pt, dash bullet character, `marL="360000" indent="-180000"`). When you add a new paragraph beyond those two, python-pptx creates it with an empty `<a:pPr/>` and no run properties — the result inherits PowerPoint's theme default, which is **Calibri 18 pt with no bullet character**. This is the bug that showed up on goeasy bullets 3+.

Fix: use the `write_bulleted_shape(shape, items)` helper from [`pptx_helpers`](../../../scripts/pptx_helpers.py). It harvests both seed paragraphs' `pPr` templates **before** wiping the shape, inserts a deepcopy of the level-appropriate `pPr` on every new paragraph, sets explicit `run.font.name = "Palatino Linotype"` and `run.font.size` on every run, and **asserts post-write that every paragraph has a `buChar`** (raises `RuntimeError` if any bullet is missing its glyph). Never hand-roll `tf.add_paragraph()` + `p.text = "..."` — that produces empty `<a:pPr/>` and PowerPoint renders no bullet.
