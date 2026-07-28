"""Deal-init helper — Phase 2.

Owns the conductor's deal-init flow (Obsidian note 12, G7 + H5):

  - The G7 questions are declared ONCE, as the `INIT_INTAKE` spec, and asked
    ONCE per deal. Both renderings are *generated* from that spec (Phase H1):
    the interactive dialogs (`render_init_dialogs`, the `AskUserQuestion`
    payloads), the filings checklist posted alongside them
    (`render_init_filings_note`), and the single-message text prompt
    (`render_init_prompt`) kept as the fallback for surfaces without the
    interactive question UI. The codename is not asked: it is auto-derived via
    `codename.codename_from_company` ("Project <company>" with corporate
    suffixes stripped), overridable by the analyst in chat.
  - DealContext is persisted as `<deal_dir>/deal.json`.
  - The deal directory tree is bootstrapped: `facts/`, `filings/`,
    `artefacts/`, `runs/`.

Plans cannot ask any of the G7 questions — those belong to deal-init.

This module is a thin orchestrator over `intake_spec.py`, `codename.py` and
the pydantic `DealContext` model. No Agent dispatch, no LLM calls.
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from codename import DEFAULT_DEALS_ROOT, find_existing, resolve
from intake_spec import (
    IntakeField,
    IntakeNote,
    IntakeOption,
    IntakeSpec,
    render_dialogs,
    render_note,
    render_prompt,
)
from schemas import DealContext


# Subdirectories every deal directory carries.
DEAL_SUBDIRS = ("facts", "filings", "artefacts", "runs")

# Where the persisted DealContext lives within a deal directory.
DEAL_JSON_NAME = "deal.json"


# ---------------------------------------------------------------------------
# The G7 questionnaire — declared once (Phase H1)
#
# Both renderings below are generated from INIT_INTAKE, so a changed option
# label cannot leave the text fallback describing a choice the dialog no
# longer offers.
#
# Shape contract (enforced by `intake_spec` and the tests): at most 4
# questions per dialog, 2–4 options per question, header at most 12
# characters, multiSelect always False. Each question's `header` doubles as
# its key in INIT_DIALOG_FIELDS.
#
# Not every G7 item is a dialog question:
#   - the codename is never asked — it is derived silently via
#     `codename.codename_from_company` ("Project <company>" with corporate
#     suffixes stripped); the analyst can still override it in chat, which is
#     why it is not a field at all;
#   - the subject company name is pure free text with nothing to suggest, so
#     it is declared with a `hint` and no options — a numbered item in the
#     text prompt, a plain chat question in the interactive flow;
#   - the filings item is a fixed STATUS question ("Filings" below) — files
#     cannot come through a dialog, so the analyst answers attached / will
#     drop next message / none, and the files themselves arrive through the
#     chat input; INIT_FILINGS_NOTE is posted alongside as the checklist
#     detail (which filings are needed and why);
#   - free-form analyst notes are no longer asked — DealContext.notes stays
#     settable when the analyst volunteers notes in chat.
# ---------------------------------------------------------------------------

# The deliverable question, kept separable: the slash-command entry points
# pre-answer it, so `render_init_dialogs` omits it by default.
DELIVERABLE_FIELD_KEY = "Deliverable"

INIT_INTAKE = IntakeSpec(
    name="deal-init",
    title="What deal is this for?",
    preamble=(
        'The codename is derived automatically: "Project <company>" with '
        "corporate suffixes stripped — say otherwise to override it.",
    ),
    fields=(
        IntakeField(
            key=DELIVERABLE_FIELD_KEY,
            prompt_label="Deliverable type",
            target="deliverable_type",
            target_kind="deal-context",
            required=True,
            group="deal",
            question="Which deliverable is this?",
            options=(
                IntakeOption(
                    "Pitch",
                    "The 16-slide slide-library pitch deck (configurable slide mix).",
                ),
                IntakeOption(
                    "Earnings update",
                    "The fixed 5-slide quarterly earnings-update deck.",
                ),
                IntakeOption(
                    "Overview",
                    "Stub — the overview plan is not yet implemented; the "
                    "conductor will say so and stop.",
                ),
                IntakeOption(
                    "One-off skill",
                    "No deliverable plan — invoke the individual skill directly "
                    "instead of the conductor.",
                ),
            ),
        ),
        IntakeField(
            key="Company",
            prompt_label="Subject company name",
            target="subject_company.legal_name",
            target_kind="deal-context",
            required=True,
            group="deal",
            hint='e.g. "ACME Corp"',
        ),
        IntakeField(
            key="Listing",
            prompt_label="Public or private?",
            target=(
                "subject_company.ticker + subject_company.exchange "
                "(Private -> both None)"
            ),
            target_kind="deal-context",
            required=True,
            group="deal",
            question="Is the subject company public or private?",
            options=(
                IntakeOption(
                    "Public — I'll give the ticker",
                    'Type ticker + exchange in the Other box (e.g. "Public — '
                    'TSX:ABC"), or I\'ll ask right after this dialog.',
                ),
                IntakeOption("Private", "No ticker or exchange."),
            ),
        ),
        IntakeField(
            key="Sector",
            prompt_label="Sector / industry",
            target="subject_company.sector / subject_company.industry",
            target_kind="deal-context",
            group="deal",
            question="Sector / industry (one line)?",
            options=(
                IntakeOption(
                    "Infer from the web",
                    "I'll look it up and use it — no confirmation needed.",
                    default=True,
                ),
                IntakeOption(
                    "I'll type it",
                    "Put the one-line sector / industry in the Other box.",
                ),
            ),
        ),
        IntakeField(
            key="Filings",
            prompt_label="Filings / attachments",
            target=(
                "filings (attachments persisted to <deal_dir>/filings/; "
                "'I'll drop them in my next message' -> the run waits for them)"
            ),
            target_kind="deal-context",
            group="deal",
            question="Filings — how will you provide them?",
            options=(
                IntakeOption(
                    "Attached in this chat",
                    "I'll save them under the deal's filings/ directory now.",
                ),
                IntakeOption(
                    "I'll drop them in my next message",
                    "The run waits here for the attachments — the checklist note "
                    "lists what I need.",
                ),
                IntakeOption(
                    "None for now",
                    "Continue without filings; attach them in chat at any later "
                    "point.",
                    default=True,
                ),
            ),
        ),
    ),
    note=IntakeNote(
        header='Filings / attachments (drop in this chat now, or say "none for now"):',
        body=(
            "For pitch and earnings-update deliverables I need the latest four "
            "annual financial statements / 10-Ks (they cover five fiscal years "
            "for the financial-summary history), plus the current-year YTD "
            "interim statements / 10-Q and the prior-year same-period interim "
            "statements / 10-Q for the LTM bridge (LTM = full fiscal year + "
            "current YTD − prior-year YTD, so a single filing isn't enough). "
            "The cap table is still built off the most recent statement; the "
            "older filings are only for the 5-year history and the LTM math.",
        ),
    ),
)

# G7 item 6 — the filings checklist. Attachments cannot come through the
# interactive dialogs, so this is always posted as plain text (alongside the
# dialogs, and embedded in the text-fallback prompt).
INIT_FILINGS_NOTE = render_note(INIT_INTAKE)

# Question header -> where the answer lands on the DealContext. Derived from
# the spec's dialog fields, so it cannot list a question that is not asked.
INIT_DIALOG_FIELDS: dict[str, str] = INIT_INTAKE.targets("deal-context")


def render_init_dialogs(*, include_deliverable: bool = False) -> list[list[dict]]:
    """Return the locked deal-init dialogs, generated from `INIT_INTAKE`.

    Each inner list is one `AskUserQuestion` call's `questions` payload —
    render them in order, unchanged. The slash-command entry points preset
    the deliverable type, so the deliverable question is only included when
    `include_deliverable=True` (generic conductor entry with no deliverable
    named). Fresh payloads every call, so a caller mutating the result cannot
    affect the next render.

    Post INIT_FILINGS_NOTE (plain text) alongside these — it is the
    checklist detail behind the "Filings" status question; file bytes cannot
    come through a dialog, so the attachments themselves arrive via the chat
    input — and ask for the subject company name as a plain chat question
    when it was not preset. The codename is never a dialog question: derive
    it silently with `codename.codename_from_company(<subject company
    name>)` (the analyst can override it in chat).
    """
    omit = () if include_deliverable else (DELIVERABLE_FIELD_KEY,)
    return render_dialogs(INIT_INTAKE, omit=omit)


def render_init_prompt() -> str:
    """Return the locked G7 deal-init text prompt, generated from `INIT_INTAKE`.

    This is the FALLBACK for surfaces where the interactive question UI
    (`AskUserQuestion`) is unavailable; `render_init_dialogs` is the primary
    rendering. Both come from the same spec, so they ask the same items in the
    same order with the same options.
    """
    return render_prompt(INIT_INTAKE)


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
