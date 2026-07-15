"""Deck-spec questionnaires — the locked per-deliverable analyst dialogs (Step 4).

The conductor collects the deck spec right after deal-init through the
interactive question UI (the `AskUserQuestion` tool): it renders each dialog in
:func:`render_deck_spec_dialogs` verbatim — one `AskUserQuestion` call per
dialog, payload unchanged — so every run of a deliverable asks the analyst the
same questions with the same options in the same order (mirrors
`deal_init.render_init_dialogs`, which owns the G7 questions — a deck spec must
never re-ask those). The UI adds an "Other" free-text box to every question
automatically, so the analyst can always answer manually.

Only the judgement items are asked. Everything with a sensible default is
defaulted instead and echoed once via :func:`render_deck_spec_defaults` so the
analyst can override by replying:

  - client name on the cover      -> the subject company name from deal-init
  - presentation date             -> current month + year
                                     (:func:`default_presentation_date`)
  - reporting quarter             -> the latest attached interim filing's
                                     quarter (conductor-inferred — fiscal
                                     labels depend on the company's fiscal
                                     calendar, so this cannot be computed from
                                     the calendar date alone)
  - comparison quarter            -> prior-year same quarter
                                     (:func:`prior_year_quarter`)
  - Financial Summary slides      -> 1 slide / 4 metrics (input left unset)
  - section divider labels        -> wireframe defaults (input left unset)

Override answers and dialog answers convert deterministically, never
improvised:

  - "N Financial Summary slides"  -> ``financial_metric_count = 4 * N``
    (:func:`metric_count_from_slides`; each FS slide shows four metric tiles)
  - "N acquisition-target slides" -> ``market_entry_target_count = 2 * N``
    (:func:`market_entry_targets_from_slides`; two targets per slide)
  - Key Investment Highlights "Omit" -> ``include_investment_highlights = False``
    (any include variant leaves the input unset = included; analyst-dictated
    highlight copy belongs in the analyst notes, not a separate input)
  - analyst notes "Draft from the attached filings + web"
    -> ``analyst_notes = NO_NOTES_ANALYST_NOTES`` (a code-owned literal, so the
    no-notes run is reproducible too)

Deliverable-specific attachments get their own locked status dialogs
(:func:`render_deck_spec_documents_dialogs` — pitch: SEDI PDF + Bloomberg
export; earnings-update: none), rendered alongside the plain-text checklist
note (:func:`render_deck_spec_documents_note`). File bytes cannot come
through a dialog, so the questions gate the run (attached / will drop next
message / none) while the file itself arrives via the chat input; the
answers never land in ``plan_inputs``.

The legacy single-message text prompts are kept (:func:`render_deck_spec_prompt`)
as the fallback for surfaces where the interactive question UI is unavailable;
they ask the same items and list the same defaults as the dialogs.

No LLM calls, no dispatch — this module only owns the locked dialog payloads,
prompts, and answer converters, so the questionnaire (and therefore the deck
layout the answers produce) is reproducible run over run.
"""

from __future__ import annotations

import copy
import re
from datetime import date

# ---------------------------------------------------------------------------
# Interactive dialogs (AskUserQuestion payloads)
#
# Shape contract (enforced by tests): at most 4 questions per dialog, 2–4
# options per question, header at most 12 characters, multiSelect always
# False. Each question's `header` doubles as its key in the
# *_DIALOG_PLAN_INPUTS tables below.
# ---------------------------------------------------------------------------

_PITCH_SPEC_DIALOGS: list[list[dict]] = [
    # Dialog 1 — content inputs.
    [
        {
            "question": (
                "Analyst notes — what should drive the executive summary, "
                "company overview, risks, and takeaways?"
            ),
            "header": "Notes",
            "multiSelect": False,
            "options": [
                {
                    "label": "I'll paste notes in my next message",
                    "description": (
                        "The run waits here for your notes. Specific "
                        "acquisition-target names and Key Investment "
                        "Highlights copy belong in these notes too."
                    ),
                },
                {
                    "label": "Draft from the attached filings + web",
                    "description": (
                        "No analyst notes — the content stage drafts "
                        "everything from the deal's filings and public "
                        "sources."
                    ),
                },
            ],
        },
        {
            "question": "Is there a CIM or management presentation to draw from?",
            "header": "CIM",
            "multiSelect": False,
            "options": [
                {
                    "label": "None",
                    "description": "Default — the deck drafts without one.",
                },
                {
                    "label": "Attached in this chat",
                    "description": (
                        "I'll save the attachment under the deal's filings/ "
                        "directory. Use Other to give an absolute path instead."
                    ),
                },
            ],
        },
        {
            "question": "Valuation range language for the executive summary?",
            "header": "Valuation",
            "multiSelect": False,
            "options": [
                {
                    "label": "None",
                    "description": (
                        "Default — the executive summary carries no valuation "
                        "range."
                    ),
                },
                {
                    "label": "I'll provide it",
                    "description": (
                        "Type the range language in the Other box (or reply "
                        "right after this dialog)."
                    ),
                },
            ],
        },
        {
            "question": (
                "Any specific risks / mitigants for the Considerations / "
                "Mitigants slide?"
            ),
            "header": "Risk notes",
            "multiSelect": False,
            "options": [
                {
                    "label": "None",
                    "description": (
                        "Default — risks and mitigants are drafted from the "
                        "filings and your notes."
                    ),
                },
                {
                    "label": "I'll provide specific risks / mitigants",
                    "description": (
                        "Type them in the Other box (or reply right after "
                        "this dialog)."
                    ),
                },
            ],
        },
    ],
    # Dialog 2 — slide mix.
    [
        {
            "question": (
                "How many Potential Market Entry Targets slides "
                "(two targets per slide)?"
            ),
            "header": "Targets",
            "multiSelect": False,
            "options": [
                {
                    "label": "4 slides — 8 targets",
                    "description": "Default.",
                },
                {"label": "3 slides — 6 targets", "description": "Six targets."},
                {"label": "2 slides — 4 targets", "description": "Four targets."},
                {"label": "1 slide — 2 targets", "description": "Two targets."},
            ],
        },
        {
            "question": "Key Investment Highlights slide?",
            "header": "Highlights",
            "multiSelect": False,
            "options": [
                {
                    "label": "Include — draft from my notes",
                    "description": (
                        "Default — the content stage drafts the highlight copy."
                    ),
                },
                {
                    "label": "Include — draft from attached filings + web",
                    "description": (
                        "The content stage drafts the highlights from the "
                        "deal's filings and public sources."
                    ),
                },
                {
                    "label": "Omit",
                    "description": "Drops the slide from the deck.",
                },
            ],
        },
    ],
]

_EARNINGS_UPDATE_SPEC_DIALOGS: list[list[dict]] = [
    [
        {
            "question": (
                "Bloomberg EEO snip (the broker estimates vs. actuals "
                "screenshot)?"
            ),
            "header": "EEO snip",
            "multiSelect": False,
            "options": [
                {
                    "label": "Attached in this chat",
                    "description": (
                        "I'll save it under the deal's filings/ directory."
                    ),
                },
                {
                    "label": "I'll attach it in my next message",
                    "description": (
                        "The run waits here for the snip. Use Other to give "
                        "an absolute path instead."
                    ),
                },
            ],
        },
    ],
]

_SPEC_DIALOGS: dict[str, list[list[dict]]] = {
    "pitch": _PITCH_SPEC_DIALOGS,
    "earnings-update": _EARNINGS_UPDATE_SPEC_DIALOGS,
}

# ---------------------------------------------------------------------------
# Attachment-status dialogs (AskUserQuestion payloads)
#
# One fixed status question per deliverable-specific document. File bytes
# cannot come through a dialog — the attachment itself always arrives via the
# chat input (or an absolute path in the Other box); these questions are the
# locked gate that pauses the run until the analyst says attached / will drop
# next message / none. Their answers are NOT plan inputs (the consuming
# stages discover the saved files under <deal_dir>/filings/), so their
# headers deliberately do not appear in the *_DIALOG_PLAN_INPUTS tables.
# The G7 filings have their own status question at deal-init; the CIM (pitch)
# and EEO snip (earnings-update) are plan inputs and stay in the spec dialogs
# above.
# ---------------------------------------------------------------------------

_PITCH_DOCUMENTS_DIALOGS: list[list[dict]] = [
    [
        {
            "question": (
                'SEDI "Insider Information by Issuer" PDF '
                "(Canadian public targets)?"
            ),
            "header": "SEDI PDF",
            "multiSelect": False,
            "options": [
                {
                    "label": "Attached in this chat",
                    "description": (
                        "I'll save it under the deal's filings/ directory now."
                    ),
                },
                {
                    "label": "I'll drop it in my next message",
                    "description": (
                        "The run waits here for the PDF — SEDI is bot-walled, "
                        "so I cannot fetch it myself."
                    ),
                },
                {
                    "label": "Not applicable / none",
                    "description": (
                        "Non-Canadian target or no report — the ownership "
                        "slide's insider side stays a placeholder."
                    ),
                },
            ],
        },
        {
            "question": "Bloomberg ownership export (.xlsm)?",
            "header": "BBG export",
            "multiSelect": False,
            "options": [
                {
                    "label": "Attached in this chat",
                    "description": (
                        "I'll save it under the deal's filings/ directory now."
                    ),
                },
                {
                    "label": "I'll drop it in my next message",
                    "description": "The run waits here for the export.",
                },
                {
                    "label": "None",
                    "description": (
                        "The ownership slide's institutions side stays a "
                        "placeholder."
                    ),
                },
            ],
        },
    ],
]

# Earnings-update has no deliverable-specific attachments beyond the EEO snip
# (a plan input, asked in the spec dialogs) and the G7 filings (deal-init's
# status question) — nothing to ask here.
_DOCUMENTS_DIALOGS: dict[str, list[list[dict]]] = {
    "pitch": _PITCH_DOCUMENTS_DIALOGS,
    "earnings-update": [],
}

# Documents-dialog question header -> where the answer leads (never a plan
# input; the files land under <deal_dir>/filings/ and the consuming stage
# discovers them there).
PITCH_DOCUMENTS_DIALOG_TARGETS: dict[str, str] = {
    "SEDI PDF": (
        "saved under <deal_dir>/filings/ — consumed by the ownership stage "
        "(insider side)"
    ),
    "BBG export": (
        "saved under <deal_dir>/filings/ — consumed by the ownership stage "
        "(institutions side)"
    ),
}

EARNINGS_UPDATE_DOCUMENTS_DIALOG_TARGETS: dict[str, str] = {}

# Question header -> plan_inputs name. Every dialog answer maps through one of
# these; the converters below turn slide-count / include-omit answers into the
# typed values.
PITCH_DIALOG_PLAN_INPUTS: dict[str, str] = {
    "Notes": "analyst_notes",
    "CIM": "cim_path",
    "Valuation": "valuation_range",
    "Risk notes": "risk_notes",
    "Targets": "market_entry_target_count",       # market_entry_targets_from_slides(answer)
    "Highlights": "include_investment_highlights",  # False only on "Omit"
}

EARNINGS_UPDATE_DIALOG_PLAN_INPUTS: dict[str, str] = {
    "EEO snip": "eeo_snip_path",
}

# The code-owned analyst_notes literal for the "Draft from the attached
# filings + web" choice (analyst_notes is a required plan input, so the
# no-notes run still needs a reproducible value).
NO_NOTES_ANALYST_NOTES = (
    "No analyst notes provided — draft the executive summary, company "
    "overview, risks, and takeaways from the attached filings and public "
    "sources."
)

# ---------------------------------------------------------------------------
# Defaulted items (not asked; echoed via render_deck_spec_defaults)
# ---------------------------------------------------------------------------

# Required plan inputs the conductor computes and supplies on every run
# (name -> the default rule, human-readable).
PITCH_DEFAULT_SUPPLIED_INPUTS: dict[str, str] = {
    "client_name": "the subject company name from deal-init",
    "presentation_date": (
        'the current month + year — default_presentation_date(date.today()), '
        'e.g. "July 2026"'
    ),
    "reporting_quarter": (
        "the latest attached interim filing's quarter (conductor-inferred "
        "from the statements — fiscal quarter labels depend on the company's "
        "fiscal calendar)"
    ),
    "comparison_quarter": (
        "prior-year same quarter — prior_year_quarter(reporting_quarter)"
    ),
}

# Optional plan inputs left OUT of plan_inputs on default (the consuming
# skills apply their own defaults).
PITCH_DEFAULT_UNSET_INPUTS: dict[str, str] = {
    "financial_metric_count": (
        'unset -> one Financial Summary slide, 4 metrics ("2 slides" '
        "override -> metric_count_from_slides(2) = 8)"
    ),
    "section_labels": (
        "unset -> wireframe defaults (Overview, Financial Summary, "
        "Valuation, Process)"
    ),
}

EARNINGS_UPDATE_DEFAULT_SUPPLIED_INPUTS: dict[str, str] = {
    "reporting_quarter": PITCH_DEFAULT_SUPPLIED_INPUTS["reporting_quarter"],
    "comparison_quarter": PITCH_DEFAULT_SUPPLIED_INPUTS["comparison_quarter"],
}

EARNINGS_UPDATE_DEFAULT_UNSET_INPUTS: dict[str, str] = {}


def render_deck_spec_dialogs(deliverable_type: str) -> list[list[dict]]:
    """Return the locked deck-spec dialogs for a deliverable, verbatim.

    Each inner list is one `AskUserQuestion` call's `questions` payload —
    render them in order, unchanged. Returns a deep copy so callers cannot
    mutate the locked constants. Raises ValueError for a deliverable with no
    questionnaire (e.g. the `overview` stub) — the conductor then falls back
    to prompting from the plan's `plan_inputs` specs directly.
    """
    try:
        return copy.deepcopy(_SPEC_DIALOGS[deliverable_type])
    except KeyError:
        raise ValueError(
            f"no deck-spec questionnaire for deliverable type {deliverable_type!r}; "
            f"known: {sorted(_SPEC_DIALOGS)}"
        ) from None


def render_deck_spec_documents_dialogs(deliverable_type: str) -> list[list[dict]]:
    """Return the locked attachment-status dialogs for a deliverable, verbatim.

    Each inner list is one `AskUserQuestion` call's `questions` payload —
    render them in order, unchanged, alongside
    :func:`render_deck_spec_documents_note` (the note carries the checklist
    detail; these questions are the fixed attached / will-drop / none gate).
    File bytes cannot come through a dialog: the attachment itself arrives
    via the chat input, or as an absolute path in the Other box. The answers
    never land in ``plan_inputs`` — "Attached in this chat" -> save under
    `<deal_dir>/filings/`; "I'll drop it in my next message" -> wait for the
    attachment before dispatching; "Not applicable / None" -> proceed (the
    consuming slide side stays a placeholder).

    Returns an EMPTY list for a deliverable with no deliverable-specific
    documents (earnings-update — render nothing); returns deep copies so
    callers cannot mutate the locked constants. Raises ValueError for an
    unknown deliverable type.
    """
    try:
        return copy.deepcopy(_DOCUMENTS_DIALOGS[deliverable_type])
    except KeyError:
        raise ValueError(
            f"no documents dialogs for deliverable type {deliverable_type!r}; "
            f"known: {sorted(_DOCUMENTS_DIALOGS)}"
        ) from None


def render_deck_spec_defaults(
    deliverable_type: str,
    *,
    client_name: str | None = None,
    presentation_date: str | None = None,
    reporting_quarter: str | None = None,
    comparison_quarter: str | None = None,
) -> str:
    """Render the one-shot defaults echo the conductor posts with the dialogs.

    The caller passes the computed default values (pitch: all four; earnings
    update: the two quarters). The returned text lists every defaulted item
    with an override invitation — the analyst overrides by replying, not
    through the dialogs.
    """
    if deliverable_type == "pitch":
        missing = [
            name
            for name, value in (
                ("client_name", client_name),
                ("presentation_date", presentation_date),
                ("reporting_quarter", reporting_quarter),
                ("comparison_quarter", comparison_quarter),
            )
            if value is None
        ]
        if missing:
            raise ValueError(f"pitch defaults echo needs {', '.join(missing)}")
        return (
            "Defaults in effect — reply to override any of these before the "
            "run starts:\n"
            f"- Client name on the cover: {client_name}\n"
            f"- Presentation date: {presentation_date}\n"
            f"- LTM bridge quarters: {reporting_quarter} vs {comparison_quarter} "
            "(from the latest attached interim filing)\n"
            '- Financial Summary slides: 1 slide — 4 metrics ("2 slides" gives 8)\n'
            "- Section divider labels: Overview, Financial Summary, Valuation, "
            "Process\n"
        )
    if deliverable_type == "earnings-update":
        if reporting_quarter is None or comparison_quarter is None:
            raise ValueError(
                "earnings-update defaults echo needs reporting_quarter and "
                "comparison_quarter"
            )
        return (
            "Defaults in effect — reply to override before the run starts:\n"
            f"- Reporting vs comparison quarter: {reporting_quarter} vs "
            f"{comparison_quarter} (from the latest attached interim filing)\n"
        )
    raise ValueError(
        f"no deck-spec defaults for deliverable type {deliverable_type!r}; "
        f"known: {sorted(_SPEC_DIALOGS)}"
    )


# ---------------------------------------------------------------------------
# Text fallback (surfaces without the interactive question UI)
# ---------------------------------------------------------------------------

_PITCH_DOCUMENTS_NOTE = """\
Documents (attach in this chat if not already attached at deal-init):
- The G7 filings: latest four annual statements / 10-Ks plus the current-year
  and prior-year interim statements (5-year history + LTM bridge).
- SEDI "Insider Information by Issuer" PDF — Canadian public targets only;
  without it the ownership slide's insider side stays a placeholder.
- Bloomberg ownership export (.xlsm) — without it the ownership slide's
  institutions side stays a placeholder.
"""

_EARNINGS_UPDATE_DOCUMENTS_NOTE = """\
Documents (attach in this chat if not already attached at deal-init):
- The G7 filings: latest four annual statements / 10-Ks plus the current-year
  and prior-year interim statements (5-year history + LTM bridge).
"""

_DOCUMENTS_NOTES: dict[str, str] = {
    "pitch": _PITCH_DOCUMENTS_NOTE,
    "earnings-update": _EARNINGS_UPDATE_DOCUMENTS_NOTE,
}

_PITCH_SPEC_PROMPT = """\
Deck spec — pitch

Answer by item number; reply "defaults" to accept every [bracketed] default.
Items marked REQUIRED have no default.

1. Analyst notes:             REQUIRED — the raw notes behind the executive
                              summary, company overview, risks, and takeaways;
                              "draft from filings + web" is acceptable
2. CIM / management pres.:    [none] — attach the file or give its path
3. Valuation range:           [none] — optional executive-summary language
4. Risk notes:                [none] — optional specific risks / mitigants for
                              the Considerations / Mitigants slide
5. Acquisition-target slides: [4 slides — 8 targets] — 1 to 4 slides, two
                              targets per slide; name specific targets if you
                              have them
6. Key Investment Highlights: [include — drafted from your notes] /
                              "include — draft from attached filings + web" /
                              "omit" — drops the slide

Defaulted unless you override here (no need to answer):
- Client name on the cover:   the subject company name from deal-init
- Presentation date:          the current month + year, e.g. "July 2026"
- Reporting quarter:          the latest attached interim filing's quarter
- Comparison quarter:         prior-year same quarter
- Financial Summary slides:   1 slide — 4 metrics ("2 slides" — 8 metrics)
- Section divider labels:     Overview, Financial Summary, Valuation, Process

""" + _PITCH_DOCUMENTS_NOTE

_EARNINGS_UPDATE_SPEC_PROMPT = """\
Deck spec — earnings update

The deck itself is the fixed 5-slide earnings-update layout (no slide options).

1. Bloomberg EEO snip:        REQUIRED — attach the screenshot or give its
                              absolute path

Defaulted unless you override here (no need to answer):
- Reporting quarter:          the latest attached interim filing's quarter
- Comparison quarter:         prior-year same quarter

""" + _EARNINGS_UPDATE_DOCUMENTS_NOTE

_SPEC_PROMPTS: dict[str, str] = {
    "pitch": _PITCH_SPEC_PROMPT,
    "earnings-update": _EARNINGS_UPDATE_SPEC_PROMPT,
}


def _dialog_item_plan_inputs(
    dialogs: list[list[dict]], header_table: dict[str, str]
) -> dict[int, str]:
    """Derive the numbered-item table from the dialog order (single source)."""
    return {
        i + 1: header_table[q["header"]]
        for i, q in enumerate(q for dialog in dialogs for q in dialog)
    }


# Fallback-prompt item number -> plan_inputs name. Derived from the dialog
# order so the text fallback and the dialogs can never drift apart.
PITCH_ITEM_PLAN_INPUTS: dict[int, str] = _dialog_item_plan_inputs(
    _PITCH_SPEC_DIALOGS, PITCH_DIALOG_PLAN_INPUTS
)

EARNINGS_UPDATE_ITEM_PLAN_INPUTS: dict[int, str] = _dialog_item_plan_inputs(
    _EARNINGS_UPDATE_SPEC_DIALOGS, EARNINGS_UPDATE_DIALOG_PLAN_INPUTS
)


def render_deck_spec_prompt(deliverable_type: str) -> str:
    """Return the locked deck-spec text prompt for a deliverable, verbatim.

    This is the FALLBACK for surfaces where the interactive question UI
    (`AskUserQuestion`) is unavailable; it asks the same items and lists the
    same defaults as :func:`render_deck_spec_dialogs`. Raises ValueError for
    a deliverable with no questionnaire (e.g. the `overview` stub) — the
    conductor then falls back to prompting from the plan's `plan_inputs`
    specs directly.
    """
    try:
        return _SPEC_PROMPTS[deliverable_type]
    except KeyError:
        raise ValueError(
            f"no deck-spec questionnaire for deliverable type {deliverable_type!r}; "
            f"known: {sorted(_SPEC_PROMPTS)}"
        ) from None


def render_deck_spec_documents_note(deliverable_type: str) -> str:
    """Return the deliverable's documents checklist, verbatim.

    Attachments cannot come through the interactive dialogs, so the conductor
    posts this as plain text alongside them (the text fallback prompt already
    embeds it).
    """
    try:
        return _DOCUMENTS_NOTES[deliverable_type]
    except KeyError:
        raise ValueError(
            f"no documents note for deliverable type {deliverable_type!r}; "
            f"known: {sorted(_DOCUMENTS_NOTES)}"
        ) from None


# ---------------------------------------------------------------------------
# Deterministic converters
# ---------------------------------------------------------------------------

# Slide-geometry constants behind the converters (and the questionnaire copy).
METRICS_PER_FINANCIAL_SUMMARY_SLIDE = 4
TARGETS_PER_MARKET_ENTRY_SLIDE = 2
MAX_MARKET_ENTRY_SLIDES = 4  # PitchDeckContent caps market_entry_targets at 8

_QUARTER_RE = re.compile(r"^\s*[Qq]([1-4])\s+(\d{4})\s*$")


def default_presentation_date(today: date) -> str:
    """The presentation-date default: spelled-out month + four-digit year."""
    return f"{today.strftime('%B')} {today.year}"


def prior_year_quarter(quarter: str) -> str:
    """The comparison-quarter default: same quarter, one year earlier.

    "Q2 2026" -> "Q2 2025". Raises ValueError for anything not shaped
    "Q<1-4> <year>".
    """
    match = _QUARTER_RE.match(quarter)
    if match is None:
        raise ValueError(
            f'quarter {quarter!r} is not shaped "Q<1-4> <year>" (e.g. "Q2 2026")'
        )
    q, year = match.groups()
    return f"Q{q} {int(year) - 1}"


def metric_count_from_slides(financial_summary_slides: int) -> int:
    """`financial_metric_count` for an "N Financial Summary slides" answer."""
    n = int(financial_summary_slides)
    if n < 1:
        raise ValueError("the deck needs at least one Financial Summary slide")
    return METRICS_PER_FINANCIAL_SUMMARY_SLIDE * n


def market_entry_targets_from_slides(market_entry_slides: int) -> int:
    """`market_entry_target_count` for an "N acquisition-target slides" answer."""
    n = int(market_entry_slides)
    if n < 1:
        raise ValueError("the deck needs at least one acquisition-target slide")
    if n > MAX_MARKET_ENTRY_SLIDES:
        raise ValueError(
            f"at most {MAX_MARKET_ENTRY_SLIDES} acquisition-target slides "
            f"(PitchDeckContent caps targets at "
            f"{MAX_MARKET_ENTRY_SLIDES * TARGETS_PER_MARKET_ENTRY_SLIDE})"
        )
    return TARGETS_PER_MARKET_ENTRY_SLIDE * n
