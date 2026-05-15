"""Unit tests for the deal_init helper."""

from pathlib import Path

import pytest

from deal_init import (
    DEAL_SUBDIRS,
    load_deal_context,
    load_or_locate_deal,
    render_init_prompt,
    save_deal_context,
)
from schemas import Company, DealContext


def _ctx(tmp_root: Path, codename: str = "Project OpenText", **overrides) -> DealContext:
    kwargs = dict(
        codename=codename,
        deal_dir=tmp_root / codename,
        deliverable_type="earnings-update",
        subject_company=Company(legal_name="OpenText Corporation", ticker="OTEX"),
    )
    kwargs.update(overrides)
    return DealContext(**kwargs)


def test_render_init_prompt_contains_seven_questions():
    """G7 lists exactly seven numbered prompts — verify all are present."""
    prompt = render_init_prompt()
    for n in range(1, 8):
        assert f"{n}." in prompt, f"prompt missing item {n}"
    # The locked field names should appear verbatim
    for label in (
        "Codename:",
        "Deliverable type:",
        "Subject company name:",
        "Public or private?:",
        "Sector / industry:",
        "Filings / attachments:",
        "Anything else?:",
    ):
        assert label in prompt, f"prompt missing label {label!r}"


def test_save_deal_context_bootstraps_dirs(tmp_path: Path):
    ctx = _ctx(tmp_path)
    save_deal_context(ctx)
    deal_dir = tmp_path / "Project OpenText"
    assert (deal_dir / "deal.json").is_file()
    for sub in DEAL_SUBDIRS:
        assert (deal_dir / sub).is_dir(), f"missing subdir {sub}"


def test_save_then_load_round_trip(tmp_path: Path):
    ctx = _ctx(tmp_path)
    save_deal_context(ctx)
    loaded = load_deal_context(tmp_path / "Project OpenText")
    assert loaded.codename == ctx.codename
    assert loaded.subject_company == ctx.subject_company


def test_load_missing_raises(tmp_path: Path):
    (tmp_path / "Project OpenText").mkdir()  # dir exists, no deal.json yet
    with pytest.raises(FileNotFoundError):
        load_deal_context(tmp_path / "Project OpenText")


def test_load_or_locate_existing_deal(tmp_path: Path):
    save_deal_context(_ctx(tmp_path))
    loaded, path = load_or_locate_deal("Project OpenText", deals_root=tmp_path)
    assert loaded is not None
    assert path == tmp_path / "Project OpenText"


def test_load_or_locate_case_insensitive(tmp_path: Path):
    save_deal_context(_ctx(tmp_path))
    loaded, path = load_or_locate_deal("project opentext", deals_root=tmp_path)
    assert loaded is not None
    assert path == tmp_path / "Project OpenText"


def test_load_or_locate_new_deal(tmp_path: Path):
    ctx, path = load_or_locate_deal("Project Brand New", deals_root=tmp_path)
    assert ctx is None
    assert path == tmp_path / "Project Brand New"
    # Must NOT mutate disk for a brand-new deal
    assert not path.exists()


def test_load_or_locate_dir_exists_no_deal_json(tmp_path: Path):
    """Empty directory shouldn't count as an existing deal."""
    (tmp_path / "Project Half-Built").mkdir()
    ctx, path = load_or_locate_deal("Project Half-Built", deals_root=tmp_path)
    assert ctx is None
    # Path should be the existing directory (case-insensitive match)
    assert path == tmp_path / "Project Half-Built"
