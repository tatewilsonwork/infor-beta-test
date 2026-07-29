"""Unit tests for the codename resolver (Phase 1, locked decision G3)."""

from pathlib import Path

import pytest

from codename import (
    ORIGIN_DOCUMENTS,
    ORIGIN_EXPLICIT,
    ORIGIN_MOUNTED,
    ORIGIN_NEW,
    Listing,
    codename_from_company,
    disambiguate,
    find_existing,
    resolve,
    resolve_deals_root,
    split_listing,
)


def test_strip_unsafe_chars():
    display, _ = resolve("Project Open/Text*?", deals_root="/tmp/deals")
    assert display == "Project OpenText"


def test_preserves_case():
    display, _ = resolve("Project OpenText", deals_root="/tmp/deals")
    assert display == "Project OpenText"
    display2, _ = resolve("project opentext", deals_root="/tmp/deals")
    assert display2 == "project opentext"


def test_preserves_common_chars():
    """& , . spaces and apostrophes are all macOS-safe and must survive."""
    display, _ = resolve("Project Smith & Co., Inc.", deals_root="/tmp/deals")
    assert display == "Project Smith & Co., Inc."


def test_dir_path_under_deals_root():
    _, path = resolve("Project OpenText", deals_root="/tmp/deals")
    assert path == Path("/tmp/deals/Project OpenText")


def test_empty_after_strip_raises():
    with pytest.raises(ValueError):
        resolve("///", deals_root="/tmp/deals")


def test_none_raises():
    with pytest.raises(ValueError):
        resolve(None)  # type: ignore[arg-type]


@pytest.mark.parametrize(
    ("company", "expected"),
    [
        ("OpenText Corporation", "Project OpenText"),
        ("ACME Corp", "Project ACME"),
        ("ACME Corp.", "Project ACME"),
        ("ACME, Inc.", "Project ACME"),
        ("acme incorporated", "Project acme"),
        ("ACME Holdings Inc.", "Project ACME"),  # suffixes strip repeatedly
        ("Northern Rail Ltd", "Project Northern Rail"),
        ("Northern Rail Limited", "Project Northern Rail"),
        ("Maple Leaf Foods LLC", "Project Maple Leaf Foods"),
        ("Brookfield LP", "Project Brookfield"),
        ("Vodafone PLC", "Project Vodafone"),
        ("Boston Consulting Group", "Project Boston Consulting"),
        ("The Coca-Cola Company", "Project The Coca-Cola"),
        ("Smith & Co", "Project Smith & Co"),  # "& Co" brands kept whole
        ("Kelso & Co LP", "Project Kelso & Co"),
        ("NoSuffix", "Project NoSuffix"),
        ("Cobalt", "Project Cobalt"),  # "Co" only strips as a whole word
    ],
)
def test_codename_from_company_strips_corporate_suffixes(company, expected):
    assert codename_from_company(company) == expected


def test_codename_from_company_strips_unsafe_and_collapses_whitespace():
    assert codename_from_company("Open/Text*  Corp?") == "Project OpenText"


def test_codename_from_company_all_suffix_name_keeps_sanitised_form():
    # A name that is nothing but suffix words must not empty out.
    assert codename_from_company("Limited") == "Project Limited"


def test_codename_from_company_rejects_empty():
    with pytest.raises(ValueError):
        codename_from_company(None)  # type: ignore[arg-type]
    with pytest.raises(ValueError):
        codename_from_company("///")


def test_find_existing_case_insensitive(tmp_path: Path):
    (tmp_path / "Project OpenText").mkdir()
    (tmp_path / "Project OpenText" / "comps.xlsx").write_bytes(b"")
    found = find_existing(tmp_path, "project opentext")
    assert found is not None
    assert found.name == "Project OpenText"


def test_find_existing_returns_none_when_missing(tmp_path: Path):
    (tmp_path / "Project OpenText").mkdir()
    assert find_existing(tmp_path, "Project Atlas") is None


def test_find_existing_handles_missing_root(tmp_path: Path):
    missing = tmp_path / "does-not-exist"
    assert find_existing(missing, "Project OpenText") is None


def test_find_existing_ignores_files(tmp_path: Path):
    """A *file* named like the codename must not be returned — we only look at directories."""
    (tmp_path / "Project OpenText").write_bytes(b"")
    assert find_existing(tmp_path, "Project OpenText") is None


def test_disambiguate_suggestions_avoid_existing(tmp_path: Path):
    (tmp_path / "Project OpenText").mkdir()
    (tmp_path / "Project OpenText II").mkdir()  # already taken
    suggestions = disambiguate(tmp_path, "Project OpenText")
    assert "Project OpenText II" not in suggestions
    assert "Project OpenText 2026" in suggestions
    assert len(suggestions) <= 4


def test_disambiguate_returns_at_least_one(tmp_path: Path):
    (tmp_path / "Project OpenText").mkdir()
    suggestions = disambiguate(tmp_path, "Project OpenText")
    assert len(suggestions) >= 1


# ─── The listing parenthetical (B14) ─────────────────────────────────────────


@pytest.mark.parametrize(
    ("company", "base", "listing"),
    [
        # The documented /pitch invocation.
        ("Open Text (TSX:OTEX)", "Open Text", Listing("TSX", "OTEX")),
        ("Example Target Inc. (NYSE:AAAA)", "Example Target Inc.", Listing("NYSE", "AAAA")),
        # Typed with a space after the colon.
        ("Open Text (TSX: OTEX)", "Open Text", Listing("TSX", "OTEX")),
        ("Open Text ( TSX : OTEX )", "Open Text", Listing("TSX", "OTEX")),
        # Mixed-case exchange codes and dotted / numeric tickers.
        ("Berkshire (NYSE:BRK.B)", "Berkshire", Listing("NYSE", "BRK.B")),
        ("Microsoft (NasdaqGS:MSFT)", "Microsoft", Listing("NasdaqGS", "MSFT")),
        ("Tencent (SEHK:700)", "Tencent", Listing("SEHK", "700")),
        # No parenthetical at all.
        ("Open Text Corporation", "Open Text Corporation", None),
        ("ACME Corp", "ACME Corp", None),
        # A parenthetical that is NOT a listing — no colon, so it is part of the
        # name and must survive verbatim.
        ("Acme (Canada)", "Acme (Canada)", None),
        ("Acme (Holdings) Ltd", "Acme (Holdings) Ltd", None),
        # A colon inside, but not `EXCHANGE:TICKER` shape (multi-word tokens).
        ("Acme (Series A: 2024)", "Acme (Series A: 2024)", None),
        # Nothing left if the parenthetical is stripped — keep the whole string
        # and let the caller's own validation complain.
        ("(TSX:OTEX)", "(TSX:OTEX)", None),
    ],
)
def test_split_listing(company, base, listing):
    assert split_listing(company) == (base, listing)


def test_split_listing_rejects_none():
    with pytest.raises(ValueError):
        split_listing(None)  # type: ignore[arg-type]


def test_listing_renders_the_capiq_form():
    assert Listing("TSX", "OTEX").capiq == "TSX:OTEX"


@pytest.mark.parametrize(
    ("company", "expected"),
    [
        # The defect: stripping only the path-unsafe colon left
        # "Project Open Text (TSXOTEX)" as the directory name.
        ("Open Text (TSX:OTEX)", "Project Open Text"),
        ("Open Text Corporation (TSX:OTEX)", "Project Open Text"),
        ("Example Target Inc. (NYSE:AAAA)", "Project Example Target"),
        ("Open Text (TSX: OTEX)", "Project Open Text"),
        # Not a listing — the parenthetical is part of the name.
        ("Acme (Canada)", "Project Acme (Canada)"),
    ],
)
def test_codename_from_company_splits_a_listing_parenthetical(company, expected):
    assert codename_from_company(company) == expected


# ─── Where the deals live (B7) ───────────────────────────────────────────────


def _deals_root(parent: Path, *deals: str) -> Path:
    """An `INFOR Deals` root under `parent`, holding `deals`."""
    root = parent / "INFOR Deals"
    root.mkdir(parents=True)
    for deal in deals:
        (root / deal).mkdir()
    return root


def test_resolve_deals_root_takes_an_explicit_argument_as_given(tmp_path: Path):
    named = tmp_path / "somewhere else"
    root = resolve_deals_root(named, home=tmp_path / "home")
    assert root.path == named
    assert root.origin == ORIGIN_EXPLICIT
    assert not named.exists(), "resolving must create nothing"


def test_resolve_deals_root_discovers_the_mounted_workspace_folder(tmp_path: Path):
    """The production shape: $HOME/mnt/<mounted folder>/INFOR Deals/<codename>."""
    home = tmp_path / "sessions" / "determined-eager-archimedes"
    mounted = _deals_root(home / "mnt" / "Claude Access Folder", "Project Open Text", "Project Atlas")

    root = resolve_deals_root(home=home)

    assert root.path == mounted
    assert root.origin == ORIGIN_MOUNTED
    assert root.deal_count == 2
    assert root.alternatives == ()


def test_resolve_deals_root_falls_back_to_documents_when_it_exists(tmp_path: Path):
    documents = _deals_root(tmp_path / "Documents", "Project Atlas")
    root = resolve_deals_root(home=tmp_path)
    assert root.path == documents
    assert root.origin == ORIGIN_DOCUMENTS
    assert root.deal_count == 1


def test_resolve_deals_root_returns_a_fresh_root_without_creating_it(tmp_path: Path):
    root = resolve_deals_root(home=tmp_path)
    assert root.path == tmp_path / "Documents" / "INFOR Deals"
    assert root.origin == ORIGIN_NEW
    assert root.deal_count == 0
    assert not root.exists, "the first save_deal_context creates it, not the resolver"


def test_resolve_deals_root_never_passes_over_a_root_holding_deals(tmp_path: Path):
    """The whole defect: a fresh/empty root must not win over ten real deals.

    The mounted shape outranks ~/Documents on precedence alone, but an EMPTY
    mounted folder beside a populated Documents root would leave `find_existing`
    blind to every prior deal — so holding deals wins, and the runner-up is
    reported rather than dropped.
    """
    empty_mount = _deals_root(tmp_path / "mnt" / "Claude Access Folder")
    documents = _deals_root(tmp_path / "Documents", *[f"Project {n}" for n in range(10)])

    root = resolve_deals_root(home=tmp_path)

    assert root.path == documents
    assert root.origin == ORIGIN_DOCUMENTS
    assert root.deal_count == 10
    assert root.alternatives == (empty_mount,)


def test_resolve_deals_root_prefers_the_mount_when_both_hold_deals(tmp_path: Path):
    mounted = _deals_root(tmp_path / "mnt" / "Claude Access Folder", "Project Open Text")
    documents = _deals_root(tmp_path / "Documents", "Project Atlas")

    root = resolve_deals_root(home=tmp_path)

    assert root.path == mounted, "precedence: the mounted workspace folder first"
    assert root.alternatives == (documents,)


def test_resolve_deals_root_matches_the_mounted_folder_name_case_insensitively(tmp_path: Path):
    mount = tmp_path / "mnt" / "Claude Access Folder"
    mount.mkdir(parents=True)
    (mount / "infor deals").mkdir()
    (mount / "infor deals" / "Project Atlas").mkdir()

    root = resolve_deals_root(home=tmp_path)

    assert root.path == mount / "infor deals"
    assert root.origin == ORIGIN_MOUNTED


def test_resolve_deals_root_ignores_hidden_entries_and_files(tmp_path: Path):
    documents = _deals_root(tmp_path / "Documents", "Project Atlas", ".DS_Store_dir")
    (documents / "notes.txt").write_text("x", encoding="utf-8")
    assert resolve_deals_root(home=tmp_path).deal_count == 1


def test_deals_root_describes_the_decision_it_made(tmp_path: Path):
    mounted = _deals_root(tmp_path / "mnt" / "Claude Access Folder", "Project Open Text")
    documents = _deals_root(tmp_path / "Documents", "Project Atlas")

    described = resolve_deals_root(home=tmp_path).describe()

    assert str(mounted) in described
    assert "1 existing deal" in described
    assert str(documents) in described, "the runner-up is offered, not hidden"


def test_a_resolved_deals_root_is_path_like(tmp_path: Path):
    """It goes straight into every `deals_root=`-taking helper."""
    mounted = _deals_root(tmp_path / "mnt" / "Workspace", "Project OpenText")
    root = resolve_deals_root(home=tmp_path)

    assert Path(root) == mounted
    assert find_existing(root, "project opentext") == mounted / "Project OpenText"
    _, deal_dir = resolve("Project Atlas", root)
    assert deal_dir == mounted / "Project Atlas"
