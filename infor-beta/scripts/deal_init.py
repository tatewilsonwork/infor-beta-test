"""Deal-init helper — Phase 2.

Owns the conductor's deal-init flow (Obsidian note 12, G7 + H5):

  - The G7 questions are declared ONCE, as the `INIT_INTAKE` spec, and asked
    ONCE per deal. Both renderings are *generated* from that spec (Phase H1):
    the interactive dialogs (`render_init_dialogs`, the `AskUserQuestion`
    payloads) and the single-message text prompt (`render_init_prompt`) kept as
    the fallback for surfaces without the interactive question UI. The codename
    is not asked: it is auto-derived via `codename.codename_from_company`
    ("Project <company>" with corporate suffixes stripped), overridable by the
    analyst in chat.
  - Only ONE G7 item is still a dialog question — the public/private listing.
    A slash-command run merges it with the deliverable's questions into a
    single `AskUserQuestion` call (`deck_spec.render_run_dialogs`);
    `render_init_dialogs` renders deal-init's half alone, which is what generic
    conductor entry asks in its first of two rounds.
  - The G7 filings are an **attachment**, not a question (v0.5.50). They are
    one bullet of the run's single attachment request, which
    `deck_spec.render_run_attachment_request` renders from this spec plus the
    deliverable's — so the requirement is declared here, once, and the
    conductor posts it once.
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

from codename import DealsRoot, find_existing, resolve, resolve_deals_root
from intake_spec import (
    IntakeDefault,
    IntakeField,
    IntakeOption,
    IntakeSpec,
    render_dialogs,
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
# Not every G7 item is a dialog question — by v0.5.51 only ONE is:
#   - the codename is never asked — it is derived silently via
#     `codename.codename_from_company` ("Project <company>" with corporate
#     suffixes stripped); the analyst can still override it in chat, which is
#     why it is not a field at all;
#   - the subject company name is pure free text with nothing to suggest, so
#     it is declared with a `hint` and no options — a numbered item in the
#     text prompt, a plain chat question in the interactive flow;
#   - the sector / industry is a DEFAULT (v0.5.51), not a question. Its dialog
#     already defaulted to "Infer from the web — I'll look it up and use it, no
#     confirmation needed", so the question's whole content was an invitation to
#     type instead, which a reply does. It is now `deal-context`-targeted
#     `defaults`, echoed once with the inferred one-liner for override, and
#     dropping it is what got a slash-command run's question count to four —
#     one dialog;
#   - the filings item is an ATTACHMENT ("Filings" below), so it is asked about
#     in no rendering at all. It was a status question through v0.5.49
#     (attached / will drop next message / none), which asked the analyst to
#     assert something the deal's filings/ directory already knew and could
#     contradict. It is now one REQUIRED bullet of the run's single attachment
#     request, and the analyst answers by dropping the files into chat;
#   - free-form analyst notes are no longer asked — DealContext.notes stays
#     settable when the analyst volunteers notes in chat.
#
# Which leaves the listing question, and it cannot be defaulted away too: it is
# required=True, and a required field may declare no default option.
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
            key="Filings",
            prompt_label="Financial statements / filings",
            target=(
                "persisted to <deal_dir>/filings/ and appended to ctx.filings as "
                "Filing entries; every data stage reads them from there"
            ),
            target_kind="attachment",
            required=True,
            checklist=(
                "the latest four annual financial statements / 10-Ks (they cover "
                "five fiscal years for the financial-summary history), plus the "
                "current-year YTD interim statements / 10-Q and the prior-year "
                "same-period interim statements / 10-Q for the LTM bridge (LTM = "
                "full fiscal year + current YTD − prior-year YTD, so a single "
                "filing isn't enough). The cap table is built off the most recent "
                "statement; the older filings are only for the 5-year history and "
                "the LTM math. Without them no data stage has a source and the "
                "deck's figures stay placeholders — and I cannot infer the "
                "reporting quarter, so I will have to ask you for it."
            ),
        ),
    ),
    # Deal-init's first default (v0.5.51). `name` IS the target — the same
    # string the Sector question carried — so a spec that re-added the question
    # while keeping this entry is rejected rather than collecting the sector
    # twice.
    defaults=(
        IntakeDefault(
            name="subject_company.sector / subject_company.industry",
            label="Sector / industry",
            rule=(
                "researched on the web and verified by search, then used without "
                "confirmation — the one-line sector / industry"
            ),
            supplied=True,
            target_kind="deal-context",
            echo="{sector}",
        ),
    ),
)

# Question header -> where the answer lands on the DealContext. Derived from
# the spec's dialog fields, so it cannot list a question that is not asked.
INIT_DIALOG_FIELDS: dict[str, str] = INIT_INTAKE.targets("deal-context")

# DealContext field -> the default rule, for the items the conductor computes
# and sets rather than asking about. Kept apart from the plan-input default
# tables (`deck_spec.*_DEFAULT_SUPPLIED_INPUTS`) by `target_kind`: these land on
# the DealContext, and one leaking into `plan_inputs` would have the conductor
# set an input named after a DealContext path.
INIT_DEFAULT_FIELDS: dict[str, str] = INIT_INTAKE.default_rules(
    supplied=True, target_kind="deal-context"
)


def render_init_dialogs(*, include_deliverable: bool = False) -> list[list[dict]]:
    """Return deal-init's half of the run's dialogs, generated from `INIT_INTAKE`.

    **Prefer `deck_spec.render_run_dialogs(<deliverable>)`** when the deliverable
    is known — a slash-command run asks deal-init's questions and the
    deliverable's in ONE `AskUserQuestion` call, and that is the merged
    renderer. This function renders deal-init's questions alone, which is what
    **generic** conductor entry needs: the Deliverable answer decides which deck
    spec exists, so with no deliverable named the run is inherently two rounds —
    this call, then `deck_spec.render_deck_spec_dialogs(<answer>)`.

    Each inner list is one call's `questions` payload — render them in order,
    unchanged. The slash-command entry points preset the deliverable type, so
    the deliverable question is only included when `include_deliverable=True`.
    Fresh payloads every call, so a caller mutating the result cannot affect the
    next render.

    Ask for the subject company name as a plain chat question when it was not
    preset. The codename is never a dialog question: derive it silently with
    `codename.codename_from_company(<subject company name>)` (the analyst can
    override it in chat). The sector is never a dialog question either — it is
    researched and echoed for override (`INIT_DEFAULT_FIELDS`).

    Nothing here asks about attachments. The G7 filings reach the analyst
    through the run's one attachment request
    (`deck_spec.render_run_attachment_request`), posted after every question in
    the run has been answered — not as a status dialog per document.
    """
    omit = () if include_deliverable else (DELIVERABLE_FIELD_KEY,)
    return render_dialogs(INIT_INTAKE, omit=omit)


def render_init_prompt() -> str:
    """Return the locked G7 deal-init text prompt, generated from `INIT_INTAKE`.

    This is the FALLBACK for surfaces where the interactive question UI
    (`AskUserQuestion`) is unavailable; `render_init_dialogs` is the primary
    rendering. Both come from the same spec, so they ask the same items in the
    same order with the same options, and the prompt ends with deal-init's half
    of the attachment request — the G7 filings, in the words the live request
    uses.
    """
    return render_prompt(INIT_INTAKE)


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
    deals_root: Path | str | DealsRoot | None = None,
) -> tuple[DealContext | None, Path]:
    """Look up an existing deal by codename.

    Returns `(ctx, deal_dir)`:
      - `ctx` is the loaded DealContext if the directory exists AND contains
        a valid `deal.json`, else None.
      - `deal_dir` is the canonical Path for this codename — either the
        existing directory found by case-insensitive lookup, or the
        fresh path that would be created if the analyst proceeds.

    `deals_root=None` resolves it via `codename.resolve_deals_root()` — the
    mounted workspace folder in production, `~/Documents/INFOR Deals` on a dev
    box — rather than assuming the E1 default, which does not exist on the
    production runtime. The conductor passes the resolved root explicitly and
    states it (SKILL.md Step 2); the default is here so a caller that omits it
    still finds the deals the analyst can see instead of starting a fresh root
    beside them.

    This helper does NOT mutate disk. It is intended to be called by the
    conductor right after the analyst types a codename, so the conductor
    can decide whether to render the G7 prompt (no existing deal) or skip
    to "what would you like to do for `<codename>`?" (existing deal).
    """
    root = Path(resolve_deals_root() if deals_root is None else deals_root).expanduser()
    existing = find_existing(root, codename)
    if existing is not None:
        try:
            return load_deal_context(existing), existing
        except FileNotFoundError:
            # Directory exists but no deal.json yet — treat as new deal.
            return None, existing
    _, fresh_path = resolve(codename, root)
    return None, fresh_path
