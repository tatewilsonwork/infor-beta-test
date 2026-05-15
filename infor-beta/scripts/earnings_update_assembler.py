"""Template-specific earnings-update deck assembler for the Phase 3 POC."""

from __future__ import annotations

import re
from pathlib import Path

from pptx import Presentation

from pptx_helpers import (
    COLOR_DOWN,
    COLOR_UP,
    find_shape,
    find_shape_in_group,
    set_cell_text,
    set_text,
    write_bulleted_shape,
)
from schemas import EarningsUpdateContent, SlidePlan


def _safe_name(value: str) -> str:
    safe = re.sub(r"[/\\:*?\"<>|]+", "-", value).strip()
    safe = re.sub(r"\s+", " ", safe)
    return safe or "Company"


def _quarter_parts(quarter: str) -> tuple[str, str]:
    match = re.search(r"Q\s*([1-4])\s*(?:FY|CY)?\s*(20\d{2}|\d{2})", quarter, re.IGNORECASE)
    if not match:
        return quarter, ""
    q, year = match.groups()
    if len(year) == 2:
        year = "20" + year
    return f"Q{q}", year


def _bullet_tuple(bullet) -> tuple[str, str, int] | tuple[str, int]:
    if bullet.bold_prefix:
        return (bullet.bold_prefix, bullet.text, bullet.level)
    return (bullet.text, bullet.level)


def _table_shape(slide):
    for shape in slide.shapes:
        if shape.shape_type == 19:  # TABLE
            return shape
    raise KeyError("Broker table not found on earnings summary slide")


def assemble_earnings_update_deck(
    *,
    slide_plan_path: Path | str,
    content_path: Path | str,
    template_path: Path | str,
    output_dir: Path | str,
) -> Path:
    """Fill the existing INFOR Earnings Update Template from typed inputs.

    This is intentionally template-specific. It is not the generalized Phase 3+
    slide-library assembler.
    """
    slide_plan = SlidePlan.model_validate_json(Path(slide_plan_path).read_text(encoding="utf-8"))
    content = EarningsUpdateContent.model_validate_json(Path(content_path).read_text(encoding="utf-8"))
    if slide_plan.deliverable_type != "earnings-update":
        raise ValueError("deck assembler POC only supports earnings-update SlidePlan objects")

    template = Path(template_path)
    if not template.exists():
        raise FileNotFoundError(f"earnings update template not found: {template}")

    out_dir = Path(output_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    output_path = out_dir / f"Earnings Update - {_safe_name(content.company_name)}.pptx"

    prs = Presentation(template)

    # Slide 1 — cover date.
    slide1 = prs.slides[0]
    for shape in slide1.shapes:
        if shape.name == "Subtitle 2" and getattr(shape, "has_text_frame", False) and "[Current Month]" in shape.text_frame.text:
            set_text(shape, [content.cover_date])

    # Slide 2 — company overview. Leave Rectangle 4 untouched.
    slide2 = prs.slides[1]
    set_text(find_shape(slide2, "Title 1"), [f"{content.company_name} Overview"])
    write_bulleted_shape(find_shape(slide2, "TextBox 16"), [_bullet_tuple(b) for b in content.company_overview_bullets])
    set_text(
        find_shape(slide2, "Text Placeholder 1"),
        [
            "Source: Company filings, S&P CapIQ, equity research ",
            f"Note: All figures in {content.currency}, except where indicated otherwise",
        ],
    )

    # Slide 3 — earnings summary.
    rq, ryear = _quarter_parts(content.reporting_quarter)
    cq, cyear = _quarter_parts(content.comparison_quarter)
    slide3 = prs.slides[2]
    set_text(find_shape(slide3, "Title 1"), [f"{content.company_name} {content.reporting_quarter} Earnings Summary"])
    set_text(find_shape(slide3, "Rectangle 7"), [f"{content.comparison_quarter} vs. {content.reporting_quarter} Financial Highlights"])
    set_text(
        find_shape(slide3, "Text Placeholder 1"),
        [
            "Source: Company filings, S&P CapIQ, equity research ",
            f"Note: All figures in {content.currency}, except where indicated otherwise",
        ],
    )
    write_bulleted_shape(find_shape(slide3, "TextBox 1067"), [(b, 0) for b in content.business_updates])

    rows = [
        ("Rectangle 1032", "Rectangle 1034", "Rectangle 1041", content.kpi_rows[0]),
        ("Rectangle 1043", "Rectangle 1037", "Rectangle 1042", content.kpi_rows[1]),
        ("Rectangle 1035", "Rectangle 1036", "Rectangle 1061", content.kpi_rows[2]),
        ("Rectangle 1057", "Rectangle 1058", "Rectangle 1064", content.kpi_rows[3]),
    ]
    for prior_shape, current_shape, delta_shape, kpi in rows:
        prior_label = f"{content.comparison_quarter} {kpi.name}"
        current_label = f"{content.reporting_quarter} {kpi.name}"
        # Preserve template's two-line value + metric label pattern.
        set_text(find_shape(slide3, prior_shape), [kpi.prior_value, prior_label])
        set_text(find_shape(slide3, current_shape), [kpi.current_value, current_label])
        color = COLOR_UP if kpi.delta_sign > 0 else (COLOR_DOWN if kpi.delta_sign < 0 else None)
        set_text(find_shape(slide3, delta_shape), [kpi.delta_str], size_pt=10, color_hex=color)

    tbl = _table_shape(slide3).table
    set_cell_text(tbl.cell(0, 0), f"Figures in {content.currency_short}", size_pt=9)
    set_cell_text(tbl.cell(0, 1), "Reported", size_pt=9)
    set_cell_text(tbl.cell(0, 2), "Bloomberg Estimate", size_pt=9)
    set_cell_text(tbl.cell(0, 3), "Variance", size_pt=9)
    for i, row in enumerate(content.broker_rows, start=1):
        set_cell_text(tbl.cell(i, 0), row.label, size_pt=9)
        set_cell_text(tbl.cell(i, 1), row.reported, size_pt=9)
        set_cell_text(tbl.cell(i, 2), row.estimate, size_pt=9)
        color = COLOR_UP if row.variance_sign > 0 else (COLOR_DOWN if row.variance_sign < 0 else None)
        set_cell_text(tbl.cell(i, 3), row.variance, size_pt=9, color_hex=color)

    q1, q2 = content.management_quotes
    g1070 = find_shape(slide3, "Group 1070")
    set_text(find_shape_in_group(g1070, "TextBox 1072"), [f"“{q1.quote}”"])
    set_text(find_shape_in_group(g1070, "TextBox 1073"), [f"{q1.speaker} – {q1.role}"])
    g1086 = find_shape(slide3, "Group 1086")
    set_text(find_shape_in_group(g1086, "TextBox 1088"), [f"“{q2.quote}”"])
    set_text(find_shape_in_group(g1086, "TextBox 1089"), [f"{q2.speaker} – {q2.role}"])
    set_text(find_shape(slide3, "Rectangle 1111"), [content.performance_summary])

    prs.save(output_path)
    _verify_output(output_path)
    return output_path


def _verify_output(path: Path) -> None:
    prs = Presentation(path)
    all_text = "\n".join(
        shape.text
        for slide in list(prs.slides)[:3]
        for shape in slide.shapes
        if getattr(shape, "has_text_frame", False)
    )
    forbidden = ["[Current Month]", "[Client Name]", "Q[x]", "Qx 202x"]
    leftovers = [token for token in forbidden if token in all_text]
    if leftovers:
        raise ValueError(f"assembled earnings update deck still contains placeholders: {leftovers}")
    slide2_text = "\n".join(
        shape.text for shape in prs.slides[1].shapes if getattr(shape, "has_text_frame", False)
    )
    if "[Macabacus Placeholder]" not in slide2_text:
        raise ValueError("slide 2 Macabacus placeholder was modified; it must remain for manual cap-table paste")
