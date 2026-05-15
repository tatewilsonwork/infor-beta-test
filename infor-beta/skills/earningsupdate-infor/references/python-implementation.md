# Reference Python Implementation

Driver code for populating the earnings update deck. All formatting helpers (`set_text`, `write_bulleted_shape`, `set_cell_text`, `find_shape`, `find_shape_in_group`, `PALATINO`, `COLOR_UP`, `COLOR_DOWN`) live in the plugin's shared module at [`infor-beta/scripts/pptx_helpers.py`](../../../scripts/pptx_helpers.py).

The template uses distinctive shape names (e.g., `Rectangle 1032`, `Group 1070`, `TextBox 1067`) that are stable across the template's lifetime. Target shapes by name, not by index, because shape order can vary.

```python
import os
import sys
from pptx import Presentation

# Load the plugin's shared pptx helpers
sys.path.insert(0, os.environ.get("CLAUDE_PLUGIN_ROOT", "./infor-beta") + "/scripts")
from pptx_helpers import (
    PALATINO, COLOR_UP, COLOR_DOWN,
    find_shape, find_shape_in_group,
    set_text, set_cell_text, write_bulleted_shape,
)

prs = Presentation(output_path)

# Slide 1 — cover date
slide1 = prs.slides[0]
# The date placeholder is the second "Subtitle 2" — pick by text content match
for s in slide1.shapes:
    if s.name == "Subtitle 2" and s.has_text_frame and "[Current Month]" in s.text_frame.text:
        set_text(s, [f"{month_name} {year}"])

# Slide 2
slide2 = prs.slides[1]
set_text(find_shape(slide2, "Title 1"), [f"{company_name} Overview"])

# Slide 2 description — pass each bullet with its level (0 = main 10.5 pt, 1 = sub 10 pt).
# Each item is (prefix_bold_text, rest_text, level). prefix_bold_text == "" for plain bullets.
# Examples:
#   ("", "Founded 1990, goeasy (TSX:GSY) is a leading Canadian non-prime consumer lender", 0)
#   ("easyfinancial:", " direct-to-consumer unsecured and home-equity loans...", 1)
def _full_text(item):
    prefix, rest, _ = item
    return (prefix + rest) if prefix else rest

def _lines_at_10_5(text, chars_per_line=65):
    """Rough line count for a bullet at Palatino 10.5 pt in a 4.53 in column."""
    return max(1, -(-len(text) // chars_per_line))

_desc_full = [_full_text(x) for x in description_items]
assert 7 <= len(description_items) <= 12, "Slide 2 description 7-12 bullets"
assert all(len(t) <= 250 for t in _desc_full), "Slide 2 bullets <= 250 chars (4 lines)"
assert 1200 <= sum(len(t) for t in _desc_full) <= 1500, "Slide 2 total 1,200-1,500 chars"
assert not any(t.rstrip().endswith((".", ";")) for t in _desc_full), "Slide 2 bullets must not end with . or ;"
_height = sum(_lines_at_10_5(t) * 0.18 + 0.11 for t in _desc_full)
assert _height <= 5.4, f"Slide 2 estimated height {_height:.2f} in exceeds 5.4 in budget"

# SINGLE CALL handles the entire description. write_bulleted_shape:
#   1. Harvests both seed pPr templates from the template shape (level 0 = square glyph
#      with marL=180975; level 1 = dash glyph with marL=360000)
#   2. Wipes all paragraphs and rewrites each with a deepcopy of the level-appropriate pPr
#   3. Sets explicit Palatino font + size on every run (no inherited formatting)
#   4. Post-write asserts every paragraph has a buChar — if not, raises RuntimeError
textbox16 = find_shape(slide2, "TextBox 16")
write_bulleted_shape(textbox16, description_items)

footnote = find_shape(slide2, "Text Placeholder 1")
set_text(footnote, ["Source: Company filings, S&P CapIQ, equity research ",
                    f"Note: All figures in {currency}, except where indicated otherwise"])

# Slide 3 — title, headers, footnote
slide3 = prs.slides[2]
set_text(find_shape(slide3, "Title 1"), [f"{company_name} Q{q} {year_c} Earnings Summary"])
set_text(find_shape(slide3, "Rectangle 7"),
         [f"Q{q} {year_p} vs. Q{q} {year_c} Financial Highlights"])
footnote3 = find_shape(slide3, "Text Placeholder 1")
set_text(footnote3, ["Source: Company filings, S&P CapIQ, equity research ",
                     f"Note: All figures in {currency}, except where indicated otherwise"])

# Slide 3 — Business updates. ALL bullets at level 0 (main), square bullet, 10 pt.
def _lines_at_10(text, chars_per_line=70):
    return max(1, -(-len(text) // chars_per_line))

assert 4 <= len(business_updates) <= 6, "Business Updates 4-6 bullets"
assert all(len(b) <= 250 for b in business_updates), "Each Business Update <= 250 chars (4 lines)"
assert sum(len(b) for b in business_updates) <= 900, "Total Business Updates <= 900 chars"
assert not any(b.rstrip().endswith((".", ";")) for b in business_updates), \
    "Business Updates must not end with . or ;"
_bu_height = sum(_lines_at_10(b) * 0.17 + 0.04 for b in business_updates)
assert _bu_height <= 2.55, f"Business Updates estimated height {_bu_height:.2f} in exceeds 2.55 in budget"

write_bulleted_shape(
    find_shape(slide3, "TextBox 1067"),
    [(text, 0) for text in business_updates],
)

# Slide 3 — KPI tiles (rows list each hold (prior_box, current_box, delta_box, tri_l, tri_r, metric))
kpi_rows = [
    ("Rectangle 1032", "Rectangle 1034", "Rectangle 1041", "Isosceles Triangle 1039", "Isosceles Triangle 1040", kpi[0]),
    ("Rectangle 1043", "Rectangle 1037", "Rectangle 1042", "Isosceles Triangle 1045", "Isosceles Triangle 1046", kpi[1]),
    ("Rectangle 1035", "Rectangle 1036", "Rectangle 1061", "Isosceles Triangle 1062", "Isosceles Triangle 1063", kpi[2]),
    ("Rectangle 1057", "Rectangle 1058", "Rectangle 1064", "Isosceles Triangle 1065", "Isosceles Triangle 1066", kpi[3]),
]
for prior, current, delta, tl, tr, m in kpi_rows:
    set_text(find_shape(slide3, prior), [m["prior_value"], f"Q{q} {year_p} {m['name']}"])
    set_text(find_shape(slide3, current), [m["current_value"], f"Q{q} {year_c} {m['name']}"])
    # Delta box — FIXED 10 pt for every delta. No step-down.
    # Rate deltas must be %, never bps.
    # Color based purely on direction: positive = green, negative = red.
    delta_str = m["delta_str"]
    assert "bps" not in delta_str.lower(), "Rate deltas must be % not bps"
    sign = m["delta_sign"]  # +1 / -1 / 0
    color = COLOR_UP if sign > 0 else (COLOR_DOWN if sign < 0 else None)
    set_text(find_shape(slide3, delta), [delta_str], size_pt=10, color_hex=color)
    # NEVER rotate triangles. Do not set shape.rotation on tl or tr.

# Slide 3 — Broker table
assert len(broker_rows) == 5, "Broker table must have exactly 5 rows"
for r in broker_rows:
    for v in r[1:4]:
        assert v and v.strip().lower() not in ("n/a", "na", "-"), f"Broker cell cannot be N/A: {r}"

for shape in slide3.shapes:
    if shape.shape_type == 19:  # TABLE
        tbl = shape.table
        set_cell_text(tbl.cell(0, 0), f"Figures in {currency_short}", size_pt=9)
        set_cell_text(tbl.cell(0, 1), "Reported", size_pt=9)
        set_cell_text(tbl.cell(0, 2), "Bloomberg Estimate", size_pt=9)
        set_cell_text(tbl.cell(0, 3), "Variance", size_pt=9)
        # broker_rows items: (label, reported_str, estimate_str, variance_str, variance_sign)
        for i, (label, reported, estimate, variance, vsign) in enumerate(broker_rows, start=1):
            set_cell_text(tbl.cell(i, 0), label, size_pt=9)
            set_cell_text(tbl.cell(i, 1), reported, size_pt=9)
            set_cell_text(tbl.cell(i, 2), estimate, size_pt=9)
            v_color = COLOR_UP if vsign > 0 else (COLOR_DOWN if vsign < 0 else None)
            set_cell_text(tbl.cell(i, 3), variance, size_pt=9, color_hex=v_color)
        break

# Slide 3 — Quotes. Hard caps enforced (quote text only, excluding curly quote marks).
for q in (quote1_text, quote2_text):
    assert len(q) <= 200, "Quote <= 200 chars"
    assert len(q.split()) <= 30, "Quote <= 30 words"
g1070 = find_shape(slide3, "Group 1070")
# Call set_text with no size_pt / color_hex overrides — preserves template italic / bold.
set_text(find_shape_in_group(g1070, "TextBox 1072"), [f"“{quote1_text}”"])
set_text(find_shape_in_group(g1070, "TextBox 1073"), [f"{quote1_name} – {quote1_role}"])
g1086 = find_shape(slide3, "Group 1086")
set_text(find_shape_in_group(g1086, "TextBox 1088"), [f"“{quote2_text}”"])
set_text(find_shape_in_group(g1086, "TextBox 1089"), [f"{quote2_name} – {quote2_role}"])

# Slide 3 — Performance summary box (must fit: <= 25 words / <= 150 chars)
assert len(performance_summary_sentence.split()) <= 25, "Summary too long — rewrite"
assert len(performance_summary_sentence) <= 150, "Summary too long — rewrite"
set_text(find_shape(slide3, "Rectangle 1111"), [performance_summary_sentence])

prs.save(output_path)
```

Use `"“"` / `"”"` for curly quotes and `"–"` for en-dash — the template uses these characters, and reverting to plain `"` or `-` would be a formatting regression.
