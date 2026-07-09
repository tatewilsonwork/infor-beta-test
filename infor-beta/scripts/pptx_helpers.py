"""python-pptx helpers shared across INFOR skills.

These helpers encode the formatting rules that recur across every INFOR deck
work and that have a history of regressing when re-derived inline in a skill:

  - `set_text(shape, lines)` preserves run-level rPr (font/size/bold/italic/color)
    by mutating runs[0].text in place rather than wiping and re-adding runs.
  - `write_bulleted_shape(shape, items)` harvests bullet pPr templates from the
    template's seed paragraphs BEFORE wiping, so new bullets keep the square /
    dash glyphs that python-pptx would otherwise drop.
  - `set_cell_text(cell, text, size_pt, color_hex)` forces Palatino on table
    cells (PowerPoint's default fallback is Calibri, which has been observed
    to slip in when cells are rewritten).
  - `delete_slide(prs, index)` removes a slide and drops its presentation-part
    relationship, so reducing a cloned library deck to the wanted entries
    leaves no orphaned slide parts behind.

Skills can load these via:

    import sys, os
    sys.path.insert(0, os.environ.get("CLAUDE_PLUGIN_ROOT", "./infor-beta") + "/scripts")
    from pptx_helpers import set_text, write_bulleted_shape, set_cell_text, delete_slide, find_shape

Tests live in tests/test_pptx_helpers.py and build fresh in-memory
decks so they don't depend on the INFOR template files.
"""

import math
from copy import deepcopy

from lxml import etree
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt


# ─── Brand constants ─────────────────────────────────────────────────────────

PALATINO = "Palatino Linotype"
COLOR_UP = "00B050"    # green — positive delta / beat
COLOR_DOWN = "C00000"  # red   — negative delta / miss

# INFOR theme accent palette (INFORFG.thmx "INFOR (New)", accent1–6) — categorical
# fills for pie / segment charts, used in theme order and cycled past six. accent2
# (46566E) is also the clustered-column bar colour used by the Financial Summary charts.
INFOR_ACCENTS = ["0E213F", "46566E", "ADB9CA", "A4844B", "767171", "E5E3E3"]


# ─── Palatino text measurement ───────────────────────────────────────────────
# Per-character advance widths for Palatino Linotype (regular), in inches per
# point of font size, measured from the font file (PIL `getlength`, kerning
# excluded — so string sums err ~2-5% wide, the safe direction for "does this
# fit on one line" checks). Used to size table labels / estimate cell wrap
# without needing the font installed at runtime (the Cowork/Linux runtime has
# no Palatino).
_PALATINO_CHAR_WIDTH_PER_PT = {
    " ": 0.00347, "!": 0.00386, '"': 0.00515, "#": 0.00665, "$": 0.00694,
    "%": 0.01167, "&": 0.01081, "'": 0.00289, "(": 0.00463, ")": 0.00463,
    "*": 0.0054, "+": 0.00694, ",": 0.00347, "-": 0.00463, ".": 0.00347,
    "/": 0.00444, "0": 0.00694, "1": 0.00694, "2": 0.00694, "3": 0.00694,
    "4": 0.00694, "5": 0.00694, "6": 0.00694, "7": 0.00694, "8": 0.00694,
    "9": 0.00694, ":": 0.00347, ";": 0.00347, "<": 0.00694, "=": 0.00694,
    ">": 0.00694, "?": 0.00617, "@": 0.00949, "A": 0.01081, "B": 0.00849,
    "C": 0.00985, "D": 0.01075, "E": 0.00849, "F": 0.00772, "G": 0.0106,
    "H": 0.01156, "I": 0.00468, "J": 0.00463, "K": 0.01008, "L": 0.00849,
    "M": 0.01314, "N": 0.01154, "O": 0.01092, "P": 0.00839, "Q": 0.01092,
    "R": 0.00928, "S": 0.00729, "T": 0.00851, "U": 0.01081, "V": 0.01003,
    "W": 0.01389, "X": 0.00926, "Y": 0.00926, "Z": 0.00926, "[": 0.00463,
    "\\": 0.00842, "]": 0.00463, "^": 0.00694, "_": 0.00694, "`": 0.00463,
    "a": 0.00694, "b": 0.00768, "c": 0.00617, "d": 0.00849, "e": 0.00665,
    "f": 0.00463, "g": 0.00772, "h": 0.00808, "i": 0.00404, "j": 0.00325,
    "k": 0.00772, "l": 0.00404, "m": 0.01226, "n": 0.00808, "o": 0.00758,
    "p": 0.00835, "q": 0.00778, "r": 0.00549, "s": 0.00589, "t": 0.00453,
    "u": 0.00838, "v": 0.00785, "w": 0.01158, "x": 0.00717, "y": 0.00772,
    "z": 0.00694, "{": 0.00463, "|": 0.00694, "}": 0.00463, "~": 0.00694,
}
_PALATINO_DEFAULT_CHAR_WIDTH_PER_PT = 0.00712  # lowercase average — non-ASCII fallback


def palatino_text_width_in(text, font_pt):
    """Estimated single-line width (inches) of Palatino text at `font_pt`."""
    return font_pt * sum(
        _PALATINO_CHAR_WIDTH_PER_PT.get(ch, _PALATINO_DEFAULT_CHAR_WIDTH_PER_PT)
        for ch in text
    )


# ─── Shape lookup ────────────────────────────────────────────────────────────

def find_shape(slide, name):
    """Return the first shape on the slide whose .name matches."""
    for s in slide.shapes:
        if s.name == name:
            return s
    raise KeyError(f"Shape {name!r} not found on slide")


def find_shape_in_group(group, name):
    """Return the first child of the group whose .name matches."""
    for s in group.shapes:
        if s.name == name:
            return s
    raise KeyError(f"Shape {name!r} not found in group {group.name!r}")


# ─── XML helpers (private) ───────────────────────────────────────────────────

def _pPr_of(paragraph):
    """Return the paragraph-level <a:pPr> element, or None if absent."""
    for child in paragraph._p:
        if child.tag.endswith("}pPr"):
            return child
    return None


def _first_run_rPr(paragraph):
    """Return a deepcopy of the first run's <a:rPr>, or None if absent."""
    for r in paragraph.runs:
        for child in r._r:
            if child.tag.endswith("}rPr"):
                return deepcopy(child)
        return None
    return None


def _replace_run_rPr(run, template_rPr):
    """Strip the run's existing rPr children and graft a fresh deepcopy in."""
    if template_rPr is None:
        return
    for child in [c for c in run._r if c.tag.endswith("}rPr")]:
        run._r.remove(child)
    run._r.insert(0, deepcopy(template_rPr))


# ─── set_text ────────────────────────────────────────────────────────────────

def set_text(shape, lines, size_pt=None, color_hex=None):
    """Replace shape text while preserving the template's run formatting.

    Strategy:
      - For existing paragraphs (i < len(tf.paragraphs)): mutate runs[0].text
        in place to keep its rPr (font/size/bold/italic/color). Remove later
        runs on the same paragraph.
      - For new paragraphs beyond the existing count: add_paragraph, copy
        pPr from paragraph 0, copy first run's rPr from paragraph 0 so the
        new run inherits the template's formatting.

    `size_pt` and `color_hex` are explicit overrides applied AFTER the
    template formatting is restored. Pass them only when you intentionally
    want to override — e.g., the delta boxes on the earnings deck need
    forced 10 pt + green/red. Title bars / quote boxes should NOT receive
    overrides (the template's bold/italic/color would be wiped).
    """
    tf = shape.text_frame
    template_pPr = _pPr_of(tf.paragraphs[0])
    template_rPr = _first_run_rPr(tf.paragraphs[0])

    for i, line in enumerate(lines):
        if i < len(tf.paragraphs):
            p = tf.paragraphs[i]
            for r in list(p.runs[1:]):
                r._r.getparent().remove(r._r)
            if p.runs:
                # In-place mutation preserves the run's rPr
                p.runs[0].text = line
                run = p.runs[0]
            else:
                run = p.add_run()
                run.text = line
                _replace_run_rPr(run, template_rPr)
        else:
            p = tf.add_paragraph()
            if template_pPr is not None:
                for child in list(p._p):
                    if child.tag.endswith("}pPr"):
                        p._p.remove(child)
                p._p.insert(0, deepcopy(template_pPr))
            run = p.add_run()
            run.text = line
            _replace_run_rPr(run, template_rPr)

        if size_pt is not None:
            run.font.name = PALATINO
            run.font.size = Pt(size_pt)
        if color_hex is not None:
            run.font.color.rgb = RGBColor.from_string(color_hex)

    while len(tf.paragraphs) > len(lines):
        p = tf.paragraphs[-1]
        p._p.getparent().remove(p._p)


# ─── set_cell_text ───────────────────────────────────────────────────────────

def set_cell_text(cell, text, size_pt=9, color_hex=None):
    """Overwrite a table cell as a single Palatino run at size_pt.

    Unlike `set_text`, this DOES set font.name + font.size explicitly because
    PowerPoint's table-cell default fallback is Calibri — inheriting from the
    template has been observed to slip back to Calibri across rewrites.
    """
    tf = cell.text_frame
    while len(tf.paragraphs) > 1:
        last = tf.paragraphs[-1]
        last._p.getparent().remove(last._p)
    p = tf.paragraphs[0]
    for r in list(p.runs):
        r._r.getparent().remove(r._r)
    run = p.add_run()
    run.text = text
    run.font.name = PALATINO
    run.font.size = Pt(size_pt)
    if color_hex is not None:
        run.font.color.rgb = RGBColor.from_string(color_hex)


# ─── write_bulleted_shape ────────────────────────────────────────────────────

def _has_bullet_glyph(pPr):
    """True if the paragraph properties carry a real bullet glyph (not buNone)."""
    return any(
        child.tag.endswith("}buChar") or child.tag.endswith("}buAutoNum")
        for child in pPr
    )


def _harvest_bullet_templates(shape):
    """Capture pPr + rPr templates from the shape's bulleted seed paragraphs.

    Returns {level_index: (pPr_copy, rPr_copy)} keyed 0, 1, 2... by ascending
    marL (smallest indent = main bullet = level 0).

    Only paragraphs that carry a real bullet glyph (`<a:buChar>` / `<a:buAutoNum>`)
    define a level. Empty spacer paragraphs and explicit `<a:buNone>` paragraphs
    are skipped, and multiple seed paragraphs sharing an indent collapse to one
    level — keyed by distinct marL, not one-entry-per-paragraph. (The INFOR
    library's exec-summary placeholder ships ~14 paragraphs, most empty or
    buNone; the old per-paragraph enumeration mapped level 0 to the marL=0
    buNone spacer, so main bullets lost their glyph and fell back to the
    placeholder's navy list colour.) When several seed paragraphs share a marL,
    the first wins, preferring one whose run carries an rPr so the level keeps
    the template's colour / size.
    """
    tf = shape.text_frame
    by_marL: dict[int, tuple] = {}
    for para in tf.paragraphs:
        pPr = _pPr_of(para)
        if pPr is None or not _has_bullet_glyph(pPr):
            continue
        marL = int(pPr.get("marL") or "0")
        rPr = _first_run_rPr(para)
        if marL not in by_marL:
            by_marL[marL] = (deepcopy(pPr), deepcopy(rPr) if rPr is not None else None)
        elif by_marL[marL][1] is None and rPr is not None:
            by_marL[marL] = (deepcopy(pPr), deepcopy(rPr))
    ordered = sorted(by_marL.items(), key=lambda kv: kv[0])
    return {i: tpl for i, (_marL, tpl) in enumerate(ordered)}


def _emit_bullet_run(paragraph, text, tmpl_rPr, size, *, bold):
    """Append a run, grafting the harvested template rPr so it keeps the
    template's colour (and italic / etc.), then re-assert Palatino/size/bold.

    Grafting the harvested rPr is what keeps body copy the template's intended
    colour: without it the run carries no `<a:solidFill>` and inherits the
    placeholder list-style `defRPr`, which on the INFOR library is navy
    (`1B2759`) and renders as blue body text. name/size/bold are re-applied on
    top because PowerPoint's list/table fallback typeface is Calibri."""
    run = paragraph.add_run()
    run.text = text
    if tmpl_rPr is not None:
        _replace_run_rPr(run, tmpl_rPr)
    run.font.name = PALATINO
    run.font.size = size
    run.font.bold = bold
    return run


def write_bulleted_shape(shape, items):
    """Wipe the shape and rewrite bullets with pPr + rPr correctly preserved.

    `items` is a list of tuples:
        (text, level)                       — single-run bullet
        (prefix_bold, rest_regular, level)  — two-run bullet (bold prefix +
                                              regular tail, used for segment
                                              names like 'easyfinancial: ...')

    `level` 0 = main (square glyph, larger font); 1 = sub (dash, smaller).

    Harvests the seed pPr templates BEFORE wiping so the bullet characters
    survive. Sets font.name = Palatino Linotype and font.size explicitly on
    every run. After writing, asserts every paragraph has a buChar element —
    raises RuntimeError if any bullet is missing its glyph, so a broken deck
    fails at write-time instead of shipping silently.
    """
    tf = shape.text_frame
    templates = _harvest_bullet_templates(shape)  # must happen BEFORE we wipe
    if not templates:
        raise RuntimeError(
            f"Shape {shape.name!r} has no bullet templates to harvest; "
            f"write_bulleted_shape requires the template to ship at least one "
            f"seed paragraph with a pPr (square or dash bullet)."
        )

    def _size_for(level):
        _, rPr = templates.get(level, (None, None))
        if rPr is not None:
            sz = rPr.get("sz")
            if sz is not None:
                return Pt(int(sz) / 100)
        return Pt(10.5 if level == 0 else 10.0)

    # Wipe: leave one paragraph behind, clear its runs and pPr
    while len(tf.paragraphs) > 1:
        last = tf.paragraphs[-1]
        last._p.getparent().remove(last._p)
    first = tf.paragraphs[0]
    for r in list(first.runs):
        r._r.getparent().remove(r._r)
    for child in list(first._p):
        if child.tag.endswith("}pPr"):
            first._p.remove(child)

    for i, item in enumerate(items):
        if len(item) == 2:
            prefix, rest, level = "", item[0], item[1]
        elif len(item) == 3:
            prefix, rest, level = item
        else:
            raise ValueError(
                "items must be (text, level) or (prefix_bold, rest_regular, level)"
            )

        p = first if i == 0 else tf.add_paragraph()
        if i != 0:
            for child in list(p._p):
                if child.tag.endswith("}pPr"):
                    p._p.remove(child)

        tmpl_pPr, tmpl_rPr = templates.get(level, templates[0])
        p._p.insert(0, deepcopy(tmpl_pPr))

        size = _size_for(level)
        if prefix:
            _emit_bullet_run(p, prefix, tmpl_rPr, size, bold=True)
            _emit_bullet_run(p, rest, tmpl_rPr, size, bold=False)
        else:
            _emit_bullet_run(p, rest, tmpl_rPr, size, bold=False)

    # Post-write: every paragraph must have a bullet character
    for i, para in enumerate(tf.paragraphs):
        has_bu = False
        for elem in para._p.iter():
            if elem.tag.endswith("}buChar") or elem.tag.endswith("}buAutoNum"):
                has_bu = True
                break
        if not has_bu:
            raise RuntimeError(
                f"Shape {shape.name!r} paragraph {i} has no bullet character — "
                f"pPr template was not propagated. Refusing to ship a broken deck."
            )


# ─── Autofit (shrink text on overflow) ───────────────────────────────────────

def enable_normal_autofit(shape, font_scale=None, line_space_reduction=None):
    """Set the text frame's body autofit to 'Shrink text on overflow'.

    PowerPoint's `<a:normAutofit/>` autofit makes the renderer scale the font
    down at display time when the text would overflow the frame. Use this on
    any shape whose copy length varies from deck to deck — bulleted overview
    blocks, business-updates lists, etc. — so an over-budget LLM run doesn't
    silently render text past the divider line.

    Pass `font_scale` (0-100, percent of original) and `line_space_reduction`
    (0-100, percent reduction in line spacing) to pre-apply a specific
    scaling instead of letting PowerPoint compute one on first render. The
    OOXML stores both as integer thousandths-of-a-percent.
    """
    tf = shape.text_frame
    txBody = tf._txBody
    bodyPr = txBody.find(qn("a:bodyPr"))
    if bodyPr is None:
        bodyPr = etree.SubElement(txBody, qn("a:bodyPr"))
        txBody.insert(0, bodyPr)
    for child in list(bodyPr):
        if child.tag in (qn("a:noAutofit"), qn("a:normAutofit"), qn("a:spAutoFit")):
            bodyPr.remove(child)
    norm = etree.SubElement(bodyPr, qn("a:normAutofit"))
    if font_scale is not None:
        norm.set("fontScale", str(int(round(font_scale * 1000))))
    if line_space_reduction is not None:
        norm.set("lnSpcReduction", str(int(round(line_space_reduction * 1000))))


# ─── Overview-bullets fit (shared by the earnings-update + pitch assemblers) ─
# Empirical Palatino Linotype layout constants for the shared overview slide's
# ~4.5"-wide TextBox 9 (zero side insets), calibrated against PowerPoint's own
# rendered line counts / text extents (TextRange.Lines / BoundHeight on a live
# 1,235-char 8-bullet block, cross-checked against the 1,055-char / 7-bullet
# block the earnings-update assembler was originally tuned on):
#   - average prose character width ≈ 0.485 × the font size;
#   - a wrapped line is ≈ 1.2 × the font size tall;
#   - each bullet paragraph adds ≈ 6 pt of paragraph spacing.
# The pre-v0.5.23 earnings-update constants (64 chars/line, 0.182"/line, no
# paragraph-spacing term) under-estimated the rendered height by ~15%, which is
# why a "fitted" overview block could still spill into the LTM revenue section.
_FIT_CHAR_WIDTH_EM = 0.485
_FIT_LINE_HEIGHT_EM = 1.2
_FIT_PARA_SPACING_IN = 0.083
_FIT_DEFAULT_FONT_PT = 10.5
_FIT_MIN_SCALE = 70.0
_FIT_SCALE_STEP = 2.5
_FIT_BAND_GAP_IN = 0.12
_OVERVIEW_BAND_PREFIX = "LTM Revenue"


def _shape_text_width_in(shape):
    """Usable text width of a shape in inches (box width minus side insets)."""
    tf = shape.text_frame
    left = tf.margin_left if tf.margin_left is not None else Inches(0.1)
    right = tf.margin_right if tf.margin_right is not None else Inches(0.1)
    return max(0.5, Emu(shape.width - left - right).inches)


def estimate_text_height_in(paragraph_texts, width_in, font_pt):
    """Estimated rendered height (inches) of Palatino paragraphs in a box.

    Per paragraph: wrapped line count at the average-character-width estimate,
    times the line height, plus the paragraph-spacing allowance. Deliberately
    ignores ``lnSpcReduction`` so the estimate errs tall (extra safety margin).
    """
    chars_per_line = width_in / (_FIT_CHAR_WIDTH_EM * font_pt / 72.0)
    height = 0.0
    for text in paragraph_texts:
        lines = max(1, math.ceil(len(text) / chars_per_line))
        height += lines * (font_pt / 72.0) * _FIT_LINE_HEIGHT_EM + _FIT_PARA_SPACING_IN
    return height


def fit_text_scale(paragraph_texts, width_in, avail_in, font_pt=_FIT_DEFAULT_FONT_PT):
    """Largest normAutofit fontScale (percent) at which the text fits ``avail_in``.

    Walks down from 100% in small steps, re-estimating the wrap at each scale
    (a smaller font also fits more characters per line, so the height falls
    faster than linearly). Returns 100.0 when the text already fits, else the
    first fitting scale, floored at ``_FIT_MIN_SCALE``.
    """
    scale = 100.0
    while scale > _FIT_MIN_SCALE:
        if estimate_text_height_in(paragraph_texts, width_in, font_pt * scale / 100.0) <= avail_in:
            break
        scale -= _FIT_SCALE_STEP
    return max(scale, _FIT_MIN_SCALE)


def fit_overview_textbox(slide, shape, *, band_prefix=_OVERVIEW_BAND_PREFIX):
    """Keep the overview bullets above the 'LTM Revenue Breakdown' header.

    The shared library ships the overview TextBox 9 one line tall and relies on
    autofit, but PowerPoint ignores a scale-less ``<a:normAutofit/>`` on open,
    so an over-long run renders at full size and spills into the LTM revenue
    section below. Size the box to the available band, and when the copy would
    still overflow at the template font, write an **explicit** ``fontScale``
    (plus a modest line-space reduction) so the shrink actually happens on
    open. Returns the applied scale (100.0 when no shrink was needed).

    Used by both deck assemblers — the pitch and earnings-update overview
    slides are the same library entry, and both previously relied on autofit
    alone (the recurring "overview text overlaps the LTM revenue breakdown"
    analyst report).
    """
    top_in = Emu(shape.top).inches
    band_bottom = None
    for other in slide.shapes:
        if getattr(other, "has_text_frame", False) and band_prefix in other.text_frame.text:
            band_bottom = Emu(other.top).inches
            break
    if band_bottom is not None and band_bottom - top_in > 0.5:
        avail_in = band_bottom - _FIT_BAND_GAP_IN - top_in
        shape.height = Inches(avail_in)
    else:
        avail_in = Emu(shape.height).inches

    paras = [p.text for p in shape.text_frame.paragraphs if p.text.strip()]
    font_pt = _FIT_DEFAULT_FONT_PT
    for p in shape.text_frame.paragraphs:
        run_size = next((r.font.size for r in p.runs if r.font.size is not None), None)
        if run_size is not None:
            font_pt = run_size.pt
            break
    scale = fit_text_scale(paras, _shape_text_width_in(shape), avail_in, font_pt)
    if scale < 100.0:
        enable_normal_autofit(shape, font_scale=scale, line_space_reduction=8)
    else:
        enable_normal_autofit(shape)
    return scale


# ─── delete_slide ────────────────────────────────────────────────────────────

def delete_slide(prs, index):
    """Remove a slide from a presentation by zero-based index.

    Drops both the `<p:sldId>` entry AND the presentation-part relationship that
    points at the slide. Removing only the sldId leaves the slide part orphaned
    in the package, which python-pptx still serialises — producing duplicate
    part-name warnings and stray slides on re-open. Dropping the relationship
    makes the part unreachable so it is not written out.

    Used to reduce a cloned library deck to the slides you want: open the full
    `INFOR Slide Library.pptx`, then delete the entries you don't need.
    """
    xml_slides = prs.slides._sldIdLst
    sldId = list(xml_slides)[index]
    rId = sldId.get(qn("r:id"))
    xml_slides.remove(sldId)
    if rId is not None:
        prs.part.drop_rel(rId)


# ─── clone_slide_after ────────────────────────────────────────────────────────

def clone_slide_after(prs, index):
    """Duplicate the slide at `index`, inserting the copy immediately after it.

    Returns the new slide. Creates a fresh slide from the source's layout, drops
    the placeholders `add_slide` seeds, deep-copies every shape element from the
    source `spTree`, then moves the new `sldId` from the tail to position
    `index + 1`.

    Used to grow a fixed-size library section to a variable count — e.g. the
    pitch market-entry slide, which the deck repeats (two targets per slide) when
    a deal has more than two targets.

    Caveat: this copies shape XML only. Slide-level relationships (embedded
    pictures, charts, hyperlinks) are NOT re-created, so the source slide must
    not depend on them. The INFOR market-entry slide carries only a table, text,
    and shape-based logo placeholders, so an spTree copy is sufficient.
    """
    source = prs.slides[index]
    new_slide = prs.slides.add_slide(source.slide_layout)
    for shape in list(new_slide.shapes):
        shape._element.getparent().remove(shape._element)
    for shape in source.shapes:
        new_slide.shapes._spTree.append(deepcopy(shape._element))
    sldIdLst = prs.slides._sldIdLst
    new_sldId = list(sldIdLst)[-1]
    sldIdLst.remove(new_sldId)
    sldIdLst.insert(index + 1, new_sldId)
    return new_slide


# ─── Shape traversal / lookup shared by the deck assemblers ──────────────────

def iter_all_shapes(shapes):
    """Yield every shape in `shapes`, recursing into groups depth-first."""
    for shape in shapes:
        yield shape
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from iter_all_shapes(shape.shapes)


def find_table_shape(slide):
    """Return the first shape on `slide` that holds a table; raise KeyError if none."""
    for shape in slide.shapes:
        if getattr(shape, "has_table", False):
            return shape
    raise KeyError("no table shape found on slide")


# ─── Footnote currency token ─────────────────────────────────────────────────

def fill_footnote_token(shape, letter, *, token="[x]"):
    """Substitute the currency-letter token in a library footnote, in place.

    The shared INFOR library footnotes ship a 'All figures in [x]$MM' line where
    `[x]` is the currency letter ('US' / 'C'). Replace the token across every
    paragraph and rewrite via `set_text`, so the rest of the standardized
    source/note lines (and their formatting) are preserved rather than
    re-hardcoded.
    """
    lines = [p.text.replace(token, letter) for p in shape.text_frame.paragraphs]
    set_text(shape, lines)


# ─── Bullet writer with plain-text fallback ──────────────────────────────────

def write_bullets_or_plain(shape, items, *, autofit=False):
    """Write bullets, falling back to plain paragraphs when the shape has no
    bullet-glyph template to harvest.

    `items` is the tuple list `write_bulleted_shape` expects — `(text, level)` or
    `(prefix_bold, rest, level)`. If the template ships no seed bullet (so
    `write_bulleted_shape` raises RuntimeError), degrade to a flat `set_text` of
    each item's text. Pass `autofit=True` to also enable shrink-on-overflow
    autofit, for variable-length blocks (overview bullets, business updates).
    """
    try:
        write_bulleted_shape(shape, items)
    except RuntimeError:
        set_text(shape, [item[0] for item in items])
    if autofit:
        enable_normal_autofit(shape)
