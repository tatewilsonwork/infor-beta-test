"""Codename resolver — Phase 1.

Implements the G3 convention from Obsidian note `12 — Locked Decisions`:

    Deal codename is the literal string the analyst types (e.g. "Project OpenText").

    - Case-preserving for display
    - Case-insensitive for lookup
    - Path-unsafe characters stripped: /, \\, :, *, ?, ", <, >, |
    - Common characters (&, ',', '.', spaces) preserved — macOS handles them fine
    - Collision-aware: caller (the conductor) decides what to do on hit; this module
      provides find_existing() and disambiguate() helpers

It also owns the two things that have to happen *before* a codename exists:

  - **Where the deals live** (`resolve_deals_root`). The E1 default
    `~/Documents/INFOR Deals` does not exist on the production runtime, where
    every deal sits under the mounted workspace folder
    (`$HOME/mnt/<mounted folder>/INFOR Deals/<codename>`, `$HOME` being
    `/sessions/<session>`). Both variable parts mean the path cannot be
    hardcoded, so it is *discovered* — and a root that already holds deals is
    never passed over for a fresh one.
  - **Splitting a listing parenthetical off the company string**
    (`split_listing`). `/pitch Open Text (TSX:OTEX)` is the documented
    invocation, and the parenthetical is a listing hint rather than part of the
    name: `codename_from_company` used to strip only the colon and leave
    `"Project Open Text (TSXOTEX)"` as the directory name.

This module is intentionally schema-free (no pydantic import) — it is used at deal-init
before the typed contract is constructed.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path

# Characters that are unsafe in a filesystem path on macOS / Linux / Windows
# (the conductor targets all three indirectly). Stripped, not replaced.
_PATH_UNSAFE = re.compile(r'[\\/:*?"<>|]')

# A trailing `(EXCHANGE:TICKER)` parenthetical — the listing hint the /pitch and
# /earnings-update commands document (`/pitch Open Text (TSX:OTEX)`). The COLON
# is the discriminator: a trailing parenthetical without one is part of the name
# ("Acme (Canada)") and is left alone. Both tokens are single words, so a
# parenthetical that merely contains a colon ("Acme (Series A: 2024)") does not
# match either. Whitespace either side of the colon is tolerated — "TSX: OTEX"
# is how it gets typed.
_LISTING_PARENTHETICAL = re.compile(
    r"\s*\(\s*([A-Za-z][A-Za-z0-9.\-]{0,15})\s*:\s*([A-Za-z0-9.\-]{1,15})\s*\)\s*$"
)

# Trailing corporate suffixes stripped when auto-deriving a codename from the
# subject company name (case-insensitive, with or without a period). Longer
# forms listed first for readability; the end-of-string anchor makes the match
# unambiguous either way.
_CORP_SUFFIXES = (
    "Incorporated",
    "Inc",
    "Corporation",
    "Corp",
    "Company",
    "Co",
    "Group",
    "Limited",
    "Ltd",
    "LLC",
    "LP",
    "PLC",
    "Holdings",
)
# The `(?<!&)` guard keeps "& Co"-style brand names whole ("Smith & Co" stays
# "Smith & Co", never "Smith &") — same rule as ownership's
# strip_legal_suffixes.
_CORP_SUFFIX_RE = re.compile(
    r"(?<!&)[\s,]+(?:" + "|".join(_CORP_SUFFIXES) + r")\.?\s*$",
    re.IGNORECASE,
)

# ─── Where the deals live ────────────────────────────────────────────────────
#
# The E1 deal-directory decision names `~/Documents/INFOR Deals`. That is the
# LAST RESORT, not the answer: the production runtime has no `~/Documents` at
# all, and every real deal lives under the mounted workspace folder. Never
# default to the constant below — call `resolve_deals_root()` and pass its
# `.path` explicitly, so a run cannot write a client deliverable somewhere the
# analyst cannot see it while `find_existing` reports no prior deals.
DEALS_DIR_NAME = "INFOR Deals"

#: Where mounted workspace folders appear, relative to `$HOME`. On the
#: production runtime `$HOME` is `/sessions/<session>` and the mounted folder's
#: own name is analyst-chosen, so `$HOME/mnt/*/INFOR Deals` is the only fixed
#: part of the shape.
MOUNT_DIR_NAME = "mnt"

#: The E1 default, expanded once at import — the fallback branch of
#: `resolve_deals_root`, and the only branch that may not exist yet.
DEFAULT_DEALS_ROOT = (Path("~/Documents") / DEALS_DIR_NAME).expanduser()

#: `DealsRoot.origin` values.
ORIGIN_EXPLICIT = "explicit"
ORIGIN_MOUNTED = "mounted-workspace"
ORIGIN_DOCUMENTS = "documents"
ORIGIN_NEW = "new"


@dataclass(frozen=True)
class DealsRoot:
    """A resolved deals root, plus how it was chosen — a *reported* decision.

    `os.PathLike`, so it drops straight into anything taking `deals_root`
    (`load_or_locate_deal(codename, deals_root=root)`); use `.path` when a real
    `Path` is wanted.

    `origin` is one of the `ORIGIN_*` constants above. `deal_count` is how many
    deal directories the root holds — 0 for a root that does not exist yet.
    `alternatives` are the other existing candidate roots, in precedence order,
    so the conductor can offer them instead of silently picking for the analyst.
    """

    path: Path
    origin: str
    deal_count: int = 0
    alternatives: tuple[Path, ...] = ()

    def __fspath__(self) -> str:
        return str(self.path)

    @property
    def exists(self) -> bool:
        return self.path.is_dir()

    def describe(self) -> str:
        """One line the conductor states before it writes anything."""
        if self.origin == ORIGIN_NEW:
            line = (
                f"Deals root: {self.path} — new. No existing "
                f'"{DEALS_DIR_NAME}" folder was found under $HOME/{MOUNT_DIR_NAME}/ '
                f"or in ~/Documents, so this one will be created."
            )
        else:
            why = {
                ORIGIN_EXPLICIT: "as supplied",
                ORIGIN_MOUNTED: f"discovered under $HOME/{MOUNT_DIR_NAME}/",
                ORIGIN_DOCUMENTS: "the documented default",
            }.get(self.origin, self.origin)
            held = (
                f"holding {self.deal_count} existing deal"
                f"{'' if self.deal_count == 1 else 's'}"
                if self.deal_count
                else "holding no deals yet"
            )
            line = f"Deals root: {self.path} — {why}, {held}."
        if self.alternatives:
            others = "; ".join(str(p) for p in self.alternatives)
            line += f" Other candidates found (say the word to use one instead): {others}."
        return line


def _deal_count(root: Path) -> int:
    """How many deal directories a root holds. 0 if it does not exist."""
    try:
        return sum(1 for c in root.iterdir() if c.is_dir() and not c.name.startswith("."))
    except OSError:
        return 0


def _mounted_candidates(home: Path) -> list[Path]:
    """Every `$HOME/mnt/*/INFOR Deals` directory, ordered by mounted-folder name.

    Matched case-insensitively on the directory name: the mounted folder comes
    from a case-preserving host filesystem onto a case-sensitive one, so
    globbing the literal string is not safe.
    """
    needle = DEALS_DIR_NAME.casefold()
    out: list[Path] = []
    try:
        mounts = sorted((home / MOUNT_DIR_NAME).iterdir(), key=lambda p: p.name)
    except OSError:
        return out
    for mount in mounts:
        if not mount.is_dir():
            continue
        try:
            children = sorted(mount.iterdir(), key=lambda p: p.name)
        except OSError:
            continue
        out.extend(c for c in children if c.is_dir() and c.name.casefold() == needle)
    return out


def resolve_deals_root(
    explicit: Path | str | DealsRoot | None = None,
    *,
    home: Path | str | None = None,
) -> DealsRoot:
    """Decide where this box keeps its deals. Creates nothing.

    Precedence, in order:

    1. **`explicit`** — an argument the analyst or the caller supplied. Taken as
       given, existing or not.
    2. **An existing `$HOME/mnt/*/INFOR Deals`** — the mounted-workspace shape
       the production runtime uses. `$HOME` is `/sessions/<session>` there and
       the mounted folder's name is analyst-chosen, so it is discovered rather
       than hardcoded.
    3. **An existing `~/Documents/INFOR Deals`** — the E1 default, which is what
       a dev box actually has.
    4. **`~/Documents/INFOR Deals` as a fresh root** — returned but NOT created;
       the first `save_deal_context` makes it, as it always did.

    With one override on that order, which is the whole point of the function: a
    candidate **holding deals** outranks an empty one. A run must never quietly
    start a fresh root beside ten existing deals — that is the defect this
    resolves, and it makes `find_existing` / `disambiguate` see the deals the
    analyst can see. Every other existing candidate comes back in
    `alternatives`, so the choice is reported rather than assumed.

    `home` is for tests; it defaults to `Path.home()`.
    """
    if explicit is not None:
        path = Path(explicit).expanduser()
        return DealsRoot(path, ORIGIN_EXPLICIT, _deal_count(path))

    home_dir = Path(home).expanduser() if home is not None else Path.home()
    documents = home_dir / "Documents" / DEALS_DIR_NAME

    # Precedence order (2 then 3), before the holds-deals override.
    ranked: list[tuple[Path, str, int]] = [
        (path, ORIGIN_MOUNTED, _deal_count(path)) for path in _mounted_candidates(home_dir)
    ]
    if documents.is_dir():
        ranked.append((documents, ORIGIN_DOCUMENTS, _deal_count(documents)))

    if not ranked:
        return DealsRoot(documents, ORIGIN_NEW, 0)

    # Stable, so a populated root wins and precedence order survives inside each
    # group. `sorted` on the flag alone is exactly that.
    ranked.sort(key=lambda item: 0 if item[2] else 1)
    path, origin, count = ranked[0]
    return DealsRoot(path, origin, count, tuple(other[0] for other in ranked[1:]))


# ─── Listing parenthetical ───────────────────────────────────────────────────


@dataclass(frozen=True)
class Listing:
    """An `EXCHANGE:TICKER` listing hint parsed off a company string."""

    exchange: str
    ticker: str

    @property
    def capiq(self) -> str:
        """The Capital IQ `Exchange:Ticker` form, e.g. `"TSX:OTEX"`."""
        return f"{self.exchange}:{self.ticker}"


def split_listing(name: str) -> tuple[str, Listing | None]:
    """Split a trailing `(EXCHANGE:TICKER)` hint off a company string.

    `"Open Text (TSX:OTEX)"` -> `("Open Text", Listing("TSX", "OTEX"))`, and
    `"TSX: OTEX"` parses the same way. A company string with no such
    parenthetical comes back unchanged with `None`, and so does one whose
    parenthetical is **part of the name** — `"Acme (Canada)"` has no colon in it,
    so it is preserved verbatim rather than mangled into a directory name.

    The `Listing` is not decoration: it answers the deal-init Listing question,
    so the conductor pre-fills `subject_company.exchange` / `.ticker` from it and
    omits that question from the run's one dialog.

    Raises ValueError if `name` is None.
    """
    if name is None:
        raise ValueError("company name must be a non-empty string")
    match = _LISTING_PARENTHETICAL.search(name)
    if match is None:
        return name.strip(), None
    base = name[: match.start()].strip()
    if not base:
        # The whole string was the parenthetical — there is no company name to
        # keep, so leave it for the caller's own validation to reject.
        return name.strip(), None
    return base, Listing(exchange=match.group(1), ticker=match.group(2))


def _strip_unsafe(name: str) -> str:
    """Strip path-unsafe characters and collapse the resulting whitespace runs."""
    stripped = _PATH_UNSAFE.sub("", name)
    # Collapse any runs of whitespace caused by stripped characters.
    collapsed = re.sub(r"\s+", " ", stripped).strip()
    return collapsed


def codename_from_company(name: str) -> str:
    """Auto-derive the default codename from the subject company name.

    ``"Project " + <name with trailing corporate suffixes stripped>`` —
    e.g. ``"OpenText Corporation"`` -> ``"Project OpenText"``,
    ``"ACME Holdings Inc."`` -> ``"Project ACME"``. Suffixes are stripped
    repeatedly (case-insensitive, with or without periods; "& Co" brands are
    kept whole), path-unsafe characters are removed, and whitespace is
    collapsed. A name that is nothing but suffixes keeps its sanitised form
    rather than emptying.

    A trailing ``(EXCHANGE:TICKER)`` listing hint is split off first
    (`split_listing`), because it is not part of the name: ``"Open Text
    (TSX:OTEX)"`` -> ``"Project Open Text"``, where stripping only the
    path-unsafe colon gave ``"Project Open Text (TSXOTEX)"``. A parenthetical
    that is not a listing is kept — ``"Acme (Canada)"`` -> ``"Project Acme
    (Canada)"``.

    This is the silent deal-init default — the analyst is not asked, but can
    override the codename in chat at any point before the deal directory is
    created.

    Raises ValueError if the name is empty after sanitisation.
    """
    if name is None:
        raise ValueError("company name must be a non-empty string")
    base, _listing = split_listing(name)
    base = _strip_unsafe(base)
    if not base:
        raise ValueError(f"company name {name!r} contains only path-unsafe characters")
    while True:
        stripped = _strip_unsafe(_CORP_SUFFIX_RE.sub("", base).rstrip(" ,."))
        if not stripped or stripped == base:
            break
        base = stripped
    return f"Project {base}"


def resolve(
    codename: str, deals_root: Path | str | DealsRoot | None = None
) -> tuple[str, Path]:
    """Return (display_name, dir_path) for a freshly-typed codename.

    The display name preserves the analyst's casing but with path-unsafe chars
    removed and whitespace collapsed. The dir_path is `<deals_root>/<display_name>`.

    `deals_root=None` resolves it via `resolve_deals_root()` rather than assuming
    `DEFAULT_DEALS_ROOT`, so a caller that omits it still lands where the deals
    actually are. Pass it explicitly — and report it — when you know it.

    Raises ValueError if the codename is empty after sanitisation.
    """
    if codename is None:
        raise ValueError("codename must be a non-empty string")
    display = _strip_unsafe(codename)
    if not display:
        raise ValueError(f"codename {codename!r} contains only path-unsafe characters")
    root = Path(resolve_deals_root() if deals_root is None else deals_root).expanduser()
    return display, root / display


def find_existing(deals_root: Path | str | DealsRoot, codename: str) -> Path | None:
    """Case-insensitive directory lookup under `deals_root`.

    Returns the existing Path if a folder whose name matches `codename` (after
    sanitisation, case-insensitive) is found; otherwise None.

    Non-existent or non-directory `deals_root` returns None — the conductor will
    create the root lazily on first write.
    """
    root = Path(deals_root).expanduser()
    if not root.exists() or not root.is_dir():
        return None
    display, _ = resolve(codename, root)
    needle = display.casefold()
    for child in root.iterdir():
        if child.is_dir() and child.name.casefold() == needle:
            return child
    return None


def disambiguate(
    deals_root: Path | str | DealsRoot, codename: str, max_suggestions: int = 4
) -> list[str]:
    """Suggest disambiguating codenames when `codename`'s directory already exists.

    Returns a list of display-form suggestions (e.g. `Project OpenText II`,
    `Project OpenText 2026`, `Project OpenText III`). The conductor presents
    these to the analyst — this helper does NOT pick one.

    Suggestions that already collide on disk are skipped.
    """
    root = Path(deals_root).expanduser()
    display, _ = resolve(codename, root)

    # Order matters: roman numerals first (analyst convention), then current year,
    # then -B / -C suffixes for the rare third/fourth collision.
    candidates = [
        f"{display} II",
        f"{display} 2026",
        f"{display} III",
        f"{display} 2025",
        f"{display}-B",
        f"{display}-C",
    ]
    out: list[str] = []
    for cand in candidates:
        if find_existing(root, cand) is None:
            out.append(cand)
        if len(out) >= max_suggestions:
            break
    return out
