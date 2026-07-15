"""Deal-init helper — Phase 2.

Owns the conductor's deal-init flow (Obsidian note 12, G7 + H5):

  - The G7 questions are rendered here, ONCE per deal — as interactive
    dialogs (`render_init_dialogs`, the `AskUserQuestion` payloads) with the
    legacy single-message text prompt (`render_init_prompt`) kept as the
    fallback for surfaces without the interactive question UI.
  - DealContext is persisted as `<deal_dir>/deal.json`.
  - The deal directory tree is bootstrapped: `facts/`, `filings/`,
    `artefacts/`, `runs/`.

Plans cannot ask any of the G7 questions — those belong to deal-init.

This module is a thin orchestrator over `codename.py` + the pydantic
`DealContext` model. No Agent dispatch, no LLM calls.
"""

from __future__ import annotations

import copy
import json
from pathlib import Path
from typing import Any

from codename import DEFAULT_DEALS_ROOT, find_existing, resolve
from schemas import DealContext


# Subdirectories every deal directory carries.
DEAL_SUBDIRS = ("facts", "filings", "artefacts", "runs")

# Where the persisted DealContext lives within a deal directory.
DEAL_JSON_NAME = "deal.json"


# G7 item 6 — the filings checklist. Attachments cannot come through the
# interactive dialogs, so this is always posted as plain text (alongside the
# dialogs, and embedded in the text-fallback prompt below).
INIT_FILINGS_NOTE = """\
Filings / attachments (drop in this chat now, or say "none for now"):
For pitch and earnings-update deliverables I need the latest four annual
financial statements / 10-Ks (they cover five fiscal years for the
financial-summary history), plus the current-year YTD interim statements /
10-Q and the prior-year same-period interim statements / 10-Q for the LTM
bridge (LTM = full fiscal year + current YTD − prior-year YTD, so a single
filing isn't enough). The cap table is still built off the most recent
statement; the older filings are only for the 5-year history and the LTM math.
"""

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

# ---------------------------------------------------------------------------
# Interactive deal-init dialogs (AskUserQuestion payloads)
#
# Shape contract (enforced by tests): at most 4 questions per dialog, 2–4
# options per question, header at most 12 characters, multiSelect always
# False. Each question's `header` doubles as its key in INIT_DIALOG_FIELDS.
#
# Not every G7 item is a dialog question:
#   - item 3 (subject company name) is pure free text with nothing to
#     suggest — when the slash command / analyst message did not supply it,
#     the conductor asks for it as a plain chat question;
#   - item 6 (filings) is an attachments checklist — INIT_FILINGS_NOTE is
#     posted as plain text because files cannot come through the dialogs.
# ---------------------------------------------------------------------------

_DELIVERABLE_QUESTION: dict = {
    "question": "Which deliverable is this?",
    "header": "Deliverable",
    "multiSelect": False,
    "options": [
        {
            "label": "Pitch",
            "description": (
                "The 16-slide slide-library pitch deck (configurable slide mix)."
            ),
        },
        {
            "label": "Earnings update",
            "description": "The fixed 5-slide quarterly earnings-update deck.",
        },
        {
            "label": "Overview",
            "description": (
                "Stub — the overview plan is not yet implemented; the "
                "conductor will say so and stop."
            ),
        },
        {
            "label": "One-off skill",
            "description": (
                "No deliverable plan — invoke the individual skill directly "
                "instead of the conductor."
            ),
        },
    ],
}

_INIT_QUESTIONS: list[dict] = [
    {
        "question": 'Codename for the deal (e.g. "Project Atlas")?',
        "header": "Codename",
        "multiSelect": False,
        "options": [
            {
                "label": "Propose one for me",
                "description": (
                    "I'll suggest a `Project <single word>` codename and "
                    "confirm it with you before creating the deal directory."
                ),
            },
            {
                "label": "Use the company name",
                "description": (
                    "Skip the confidentiality codename — the deal directory "
                    "is named after the company itself."
                ),
            },
        ],
    },
    {
        "question": "Is the subject company public or private?",
        "header": "Listing",
        "multiSelect": False,
        "options": [
            {
                "label": "Public — I'll give the ticker",
                "description": (
                    'Type ticker + exchange in the Other box (e.g. "Public — '
                    'TSX:ABC"), or I\'ll ask right after this dialog.'
                ),
            },
            {
                "label": "Private",
                "description": "No ticker or exchange.",
            },
        ],
    },
    {
        "question": "Sector / industry (one line)?",
        "header": "Sector",
        "multiSelect": False,
        "options": [
            {
                "label": "Infer from the web — I'll confirm",
                "description": (
                    "I'll look it up, verify by web search, and confirm the "
                    "one-liner with you."
                ),
            },
            {
                "label": "I'll type it",
                "description": (
                    "Put the one-line sector / industry in the Other box."
                ),
            },
        ],
    },
    {
        "question": "Anything else I should know for this deal?",
        "header": "Extras",
        "multiSelect": False,
        "options": [
            {
                "label": "Nothing else",
                "description": "Default.",
            },
            {
                "label": "I'll add notes",
                "description": (
                    "Type them in the Other box (or reply right after this "
                    "dialog)."
                ),
            },
        ],
    },
]

# Question header -> where the answer lands on the DealContext.
INIT_DIALOG_FIELDS: dict[str, str] = {
    "Codename": "codename",
    "Deliverable": "deliverable_type",
    "Listing": "subject_company.ticker + subject_company.exchange (Private -> both None)",
    "Sector": "subject_company.sector / subject_company.industry",
    "Extras": "notes",
}

_DIALOG_MAX_QUESTIONS = 4


def render_init_dialogs(*, include_deliverable: bool = False) -> list[list[dict]]:
    """Return the locked deal-init dialogs, verbatim.

    Each inner list is one `AskUserQuestion` call's `questions` payload —
    render them in order, unchanged. The slash-command entry points preset
    the deliverable type, so the deliverable question is only included when
    `include_deliverable=True` (generic conductor entry with no deliverable
    named). Returns deep copies so callers cannot mutate the locked
    constants.

    Post INIT_FILINGS_NOTE (plain text) alongside these — attachments cannot
    come through the dialogs — and ask for the subject company name as a
    plain chat question when it was not preset.
    """
    questions = list(_INIT_QUESTIONS)
    if include_deliverable:
        questions.insert(1, _DELIVERABLE_QUESTION)
    return [
        copy.deepcopy(questions[i : i + _DIALOG_MAX_QUESTIONS])
        for i in range(0, len(questions), _DIALOG_MAX_QUESTIONS)
    ]


def render_init_prompt() -> str:
    """Return the locked G7 deal-init text prompt verbatim.

    This is the FALLBACK for surfaces where the interactive question UI
    (`AskUserQuestion`) is unavailable; `render_init_dialogs` is the primary
    rendering.
    """
    return _INIT_PROMPT


def render_init_filings_note() -> str:
    """Return the G7 filings checklist (item 6), verbatim plain text."""
    return INIT_FILINGS_NOTE


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
