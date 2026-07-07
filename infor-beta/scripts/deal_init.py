"""Deal-init helper — Phase 2.

Owns the conductor's deal-init flow (Obsidian note 12, G7 + H5):

  - The G7 7-field prompt is rendered here, ONCE per deal.
  - DealContext is persisted as `<deal_dir>/deal.json`.
  - The deal directory tree is bootstrapped: `facts/`, `filings/`,
    `artefacts/`, `runs/`.

Plans cannot ask any of the G7 questions — those belong to deal-init.

This module is a thin orchestrator over `codename.py` + the pydantic
`DealContext` model. No Agent dispatch, no LLM calls.
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from codename import DEFAULT_DEALS_ROOT, find_existing, resolve
from schemas import DealContext


# Subdirectories every deal directory carries.
DEAL_SUBDIRS = ("facts", "filings", "artefacts", "runs")

# Where the persisted DealContext lives within a deal directory.
DEAL_JSON_NAME = "deal.json"


_INIT_PROMPT = """\
What deal is this for?

1. Codename:                  (e.g. "Project Atlas")
2. Deliverable type:          (pitch / earnings update / overview / one-off skill)
3. Subject company name:      (e.g. "ACME Corp")
4. Public or private?:        (public → ask for ticker + exchange; private → skip)
5. Sector / industry:         (one line, free-form)
6. Filings / attachments:     (drop now or "none for now")
                              For pitch and earnings-update deliverables I need
                              the latest four annual financial statements / 10-Ks
                              (they cover five fiscal years for the
                              financial-summary history), plus the current-year
                              YTD interim statements / 10-Q and the prior-year
                              same-period interim statements / 10-Q for the LTM
                              bridge (LTM = full fiscal year + current YTD −
                              prior-year YTD, so a single filing isn't enough).
                              The cap table is still built off the most recent
                              statement; the older filings are only for the
                              5-year history and the LTM math.
7. Anything else?:            (optional analyst notes)
"""


def render_init_prompt() -> str:
    """Return the locked G7 deal-init prompt verbatim."""
    return _INIT_PROMPT


def _bootstrap_dirs(deal_dir: Path) -> None:
    """Create the deal directory and the four standard subdirs. Idempotent."""
    deal_dir.mkdir(parents=True, exist_ok=True)
    for sub in DEAL_SUBDIRS:
        (deal_dir / sub).mkdir(exist_ok=True)


def deal_json_path(deal_dir: Path | str) -> Path:
    """Path to the persisted DealContext within a deal directory."""
    return Path(deal_dir).expanduser() / DEAL_JSON_NAME


def save_deal_context(ctx: DealContext) -> Path:
    """Write `ctx` to `<deal_dir>/deal.json`. Returns the path written.

    Caller is responsible for having set `ctx.deal_dir` to the desired
    location (typically via `resolve(codename, deals_root)`).
    """
    deal_dir = Path(str(ctx.deal_dir)).expanduser()
    _bootstrap_dirs(deal_dir)
    target = deal_json_path(deal_dir)
    target.write_text(
        ctx.model_dump_json(indent=2) + "\n",
        encoding="utf-8",
    )
    return target


def load_deal_context(deal_dir: Path | str) -> DealContext:
    """Read `<deal_dir>/deal.json` and return the parsed DealContext.

    Raises FileNotFoundError if the file does not exist.
    """
    path = deal_json_path(deal_dir)
    if not path.exists():
        raise FileNotFoundError(
            f"no deal.json at {path} — the deal has not been initialised yet"
        )
    return DealContext.model_validate_json(path.read_text(encoding="utf-8"))


def load_or_locate_deal(
    codename: str,
    *,
    deals_root: Path | str = DEFAULT_DEALS_ROOT,
) -> tuple[DealContext | None, Path]:
    """Look up an existing deal by codename.

    Returns `(ctx, deal_dir)`:
      - `ctx` is the loaded DealContext if the directory exists AND contains
        a valid `deal.json`, else None.
      - `deal_dir` is the canonical Path for this codename — either the
        existing directory found by case-insensitive lookup, or the
        fresh path that would be created if the analyst proceeds.

    This helper does NOT mutate disk. It is intended to be called by the
    conductor right after the analyst types a codename, so the conductor
    can decide whether to render the G7 prompt (no existing deal) or skip
    to "what would you like to do for `<codename>`?" (existing deal).
    """
    root = Path(deals_root).expanduser()
    existing = find_existing(root, codename)
    if existing is not None:
        try:
            return load_deal_context(existing), existing
        except FileNotFoundError:
            # Directory exists but no deal.json yet — treat as new deal.
            return None, existing
    _, fresh_path = resolve(codename, root)
    return None, fresh_path
