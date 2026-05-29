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

from copy import deepcopy

from lxml import etree
from pptx.dml.color import RGBColor
from pptx.oxml.ns import qn
from pptx.util import Pt


# ─── Brand constants ─────────────────────────────────────────────────────────

PALATINO = "Palatino Linotype"
COLOR_UP = "00B050"    # green — positive delta / beat
COLOR_DOWN = "C00000"  # red   — negative delta / miss


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

def _harvest_bullet_templates(shape):
    """Capture pPr + rPr templates from the shape's seed paragraphs BEFORE wiping.

    Returns {level_index: (pPr_copy, rPr_copy)} keyed 0, 1, 2... by ascending
    marL (smallest indent = main bullet = level 0).
    """
    tf = shape.text_frame
    harvested = []
    for para in tf.paragraphs:
        pPr = _pPr_of(para)
        rPr = _first_run_rPr(para)
        if pPr is None:
            continue
        marL = int(pPr.get("marL") or "0")
        harvested.append((marL, deepcopy(pPr), deepcopy(rPr) if rPr is not None else None))
    harvested.sort(key=lambda t: t[0])
    return {i: (pPr, rPr) for i, (_, pPr, rPr) in enumerate(harvested)}


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

        tmpl_pPr, _ = templates.get(level, templates[0])
        p._p.insert(0, deepcopy(tmpl_pPr))

        size = _size_for(level)
        if prefix:
            r1 = p.add_run()
            r1.text = prefix
            r1.font.name = PALATINO
            r1.font.size = size
            r1.font.bold = True
            r2 = p.add_run()
            r2.text = rest
            r2.font.name = PALATINO
            r2.font.size = size
            r2.font.bold = False
        else:
            r = p.add_run()
            r.text = rest
            r.font.name = PALATINO
            r.font.size = size
            r.font.bold = False

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
