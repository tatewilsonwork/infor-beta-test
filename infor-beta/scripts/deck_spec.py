"""Deck-spec questionnaires — the locked per-deliverable analyst dialogs (Step 4).

Each deliverable's questionnaire is declared ONCE as an
:class:`intake_spec.IntakeSpec` (Phase H1) and every rendering is generated
from it: the interactive dialogs, the attachment request, the defaults echo,
and the single-message text fallback. A changed option label therefore reaches
every surface, which is what makes the locked-questionnaire principle
structural rather than conventional.

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

**The analyst is asked nothing about attachments** (v0.5.50). Every document
the run needs — deal-init's G7 filings and the deliverable's own — is one
bullet of a single plain-text request, rendered by
:func:`render_run_attachment_request` and posted *after* every question has
been answered; the analyst then drops the files into chat and the run continues.
Attachments were three `AskUserQuestion` status gates per pitch run through
v0.5.49, asked one dialog at a time: three pauses to collect three assertions
about files, each of which the deal's ``filings/`` directory already knew and
could contradict. There is now one request and one pause.

Two of those documents carry a path into ``plan_inputs`` — the pitch CIM
(``cim_path``, optional) and the earnings-update Bloomberg EEO snip
(``eeo_snip_path``, REQUIRED). The conductor resolves both from the file it
saved under ``<deal_dir>/filings/``, through
:data:`PITCH_ATTACHMENT_PLAN_INPUTS` / :data:`EARNINGS_UPDATE_ATTACHMENT_PLAN_INPUTS`;
a missing optional one stays out of ``plan_inputs`` entirely, and a missing
REQUIRED one halts the run rather than resolving to None.

The single-message text prompt is generated too (:func:`render_deck_spec_prompt`)
as the fallback for surfaces where the interactive question UI is unavailable;
it asks the same items in the same order, quotes the same question wording,
option labels and option descriptions as the dialogs, and ends with the same
attachment bullets, because all of it comes from one spec.

No LLM calls, no dispatch — this module only owns the locked questionnaires and
answer converters, so the questionnaire (and therefore the deck layout the
answers produce) is reproducible run over run.
"""

from __future__ import annotations

import re
from datetime import date
from typing import get_args

from deal_init import INIT_INTAKE
from intake_spec import (
    IntakeDefault,
    IntakeField,
    IntakeOption,
    IntakeSpec,
    render_attachment_request,
    render_defaults_echo,
    render_dialogs,
    render_prompt,
)
from schemas import DeliverableType

# The declared deliverables, from the one place they are declared. Only two
# have a questionnaire; the rest still need the deal's filings.
DELIVERABLE_TYPES: tuple[str, ...] = get_args(DeliverableType)

# ---------------------------------------------------------------------------
# The declared questionnaires (Phase H1)
#
# Shape contract (enforced by `intake_spec` and the tests): at most 4
# questions per dialog, 2–4 options per question, header at most 12
# characters, multiSelect always False. Each field's `key` is its dialog
# `header` and doubles as its key in the *_DIALOG_PLAN_INPUTS tables below.
#
# Fields with target_kind="attachment" are files, and are asked about in no
# rendering: each is one bullet of the attachment request, carrying a
# `checklist` line that states what the run loses without it (the warning that
# used to live in a status dialog's option description). A `plan_input` names
# the plan input the saved path becomes — the CIM and the EEO snip have one;
# the SEDI report and the Bloomberg export do not, because the ownership stage
# finds them under <deal_dir>/filings/ itself. The G7 filings are declared once
# more, on `deal_init.INIT_INTAKE`, and reach the analyst through the same
# request; they used to be described a second time here, in different words,
# which is exactly the drift H1 exists to remove.
# ---------------------------------------------------------------------------

# The two quarter defaults are identical for both deliverables — declared once
# so the pitch and earnings-update specs cannot describe them differently.
_REPORTING_QUARTER_DEFAULT = IntakeDefault(
    name="reporting_quarter",
    label="Reporting quarter",
    rule=(
        "the latest attached interim filing's quarter — conductor-inferred from "
        "the statements, since fiscal quarter labels depend on the company's "
        "fiscal calendar, not the calendar date"
    ),
    supplied=True,
    echo_label="Reporting vs comparison quarter",
    echo=(
        "{reporting_quarter} vs {comparison_quarter} (from the latest attached "
        "interim filing — the LTM bridge)"
    ),
)

# Echoed on the reporting quarter's line, as one range.
_COMPARISON_QUARTER_DEFAULT = IntakeDefault(
    name="comparison_quarter",
    label="Comparison quarter",
    rule="prior-year same quarter",
    supplied=True,
    echoed=False,
)

PITCH_INTAKE = IntakeSpec(
    name="pitch",
    title="Deck spec — pitch",
    fields=(
        # -- content inputs -------------------------------------------------
        IntakeField(
            key="Notes",
            prompt_label="Analyst notes",
            target="analyst_notes",
            required=True,
            group="content",
            question=(
                "Analyst notes — what should drive the executive summary, "
                "company overview, risks, and takeaways?"
            ),
            options=(
                IntakeOption(
                    "I'll paste notes in my next message",
                    "Paste them alongside the attachments I ask for next — the run "
                    "waits there for both. Specific acquisition-target names and "
                    "Key Investment Highlights copy belong in these notes too.",
                ),
                IntakeOption(
                    "Draft from the attached filings + web",
                    "No analyst notes — the content stage drafts everything from "
                    "the deal's filings and public sources.",
                ),
            ),
        ),
        IntakeField(
            key="Valuation",
            prompt_label="Valuation range",
            target="valuation_range",
            group="content",
            question="Valuation range language for the executive summary?",
            options=(
                IntakeOption(
                    "None",
                    "Default — the executive summary carries no valuation range.",
                    default=True,
                ),
                IntakeOption(
                    "I'll provide it",
                    "Type the range language in the Other box (or reply right "
                    "after this dialog).",
                ),
            ),
        ),
        IntakeField(
            key="Risk notes",
            prompt_label="Risk notes",
            target="risk_notes",
            group="content",
            question=(
                "Any specific risks / mitigants for the Considerations / "
                "Mitigants slide?"
            ),
            options=(
                IntakeOption(
                    "None",
                    "Default — risks and mitigants are drafted from the filings "
                    "and your notes.",
                    default=True,
                ),
                IntakeOption(
                    "I'll provide specific risks / mitigants",
                    "Type them in the Other box (or reply right after this "
                    "dialog).",
                ),
            ),
        ),
        # -- slide mix ------------------------------------------------------
        IntakeField(
            key="Targets",
            prompt_label="Acquisition-target slides",
            # market_entry_targets_from_slides(answer)
            target="market_entry_target_count",
            group="slide mix",
            question=(
                "How many Potential Market Entry Targets slides "
                "(two targets per slide)?"
            ),
            options=(
                IntakeOption("4 slides — 8 targets", "Default.", default=True),
                IntakeOption("3 slides — 6 targets", "Six targets."),
                IntakeOption("2 slides — 4 targets", "Four targets."),
                IntakeOption("1 slide — 2 targets", "Two targets."),
            ),
        ),
        IntakeField(
            key="Highlights",
            prompt_label="Key Investment Highlights",
            target="include_investment_highlights",  # False only on "Omit"
            group="slide mix",
            question="Key Investment Highlights slide?",
            options=(
                IntakeOption(
                    "Include — draft from my notes",
                    "Default — the content stage drafts the highlight copy.",
                    default=True,
                ),
                IntakeOption(
                    "Include — draft from attached filings + web",
                    "The content stage drafts the highlights from the deal's "
                    "filings and public sources.",
                ),
                IntakeOption("Omit", "Drops the slide from the deck."),
            ),
        ),
        # -- attachments (asked about in no rendering) ----------------------
        IntakeField(
            key="CIM",
            prompt_label="CIM / management presentation",
            target="consumed by the pitch-content, comps and precedents stages",
            target_kind="attachment",
            plan_input="cim_path",
            group="documents",
            checklist=(
                "if one exists; it is the richest single source for the company "
                "overview and for choosing the comps and precedents verticals. "
                "Without it the deck drafts from the filings and public sources "
                "only."
            ),
        ),
        IntakeField(
            key="SEDI PDF",
            prompt_label='SEDI "Insider Information by Issuer" PDF',
            target="consumed by the ownership stage (insider side)",
            target_kind="attachment",
            group="documents",
            checklist=(
                "Canadian public targets only. SEDI is bot-walled, so I cannot "
                "fetch it myself; without it the ownership slide's insider side "
                "stays a placeholder."
            ),
        ),
        IntakeField(
            key="BBG export",
            prompt_label="Bloomberg ownership export (.xlsm)",
            target="consumed by the ownership stage (institutions side)",
            target_kind="attachment",
            group="documents",
            checklist=(
                "the institutional holders export. Without it the ownership "
                "slide's institutions side stays a placeholder."
            ),
        ),
    ),
    defaults=(
        IntakeDefault(
            name="client_name",
            label="Client name on the cover",
            rule="the subject company name from deal-init",
            supplied=True,
            echo="{client_name}",
        ),
        IntakeDefault(
            name="presentation_date",
            label="Presentation date",
            rule='the current month + year, e.g. "July 2026"',
            supplied=True,
            echo="{presentation_date}",
        ),
        _REPORTING_QUARTER_DEFAULT,
        _COMPARISON_QUARTER_DEFAULT,
        IntakeDefault(
            name="financial_metric_count",
            label="Financial Summary slides",
            rule='1 slide — 4 metrics ("2 slides" — 8 metrics)',
            supplied=False,
        ),
        IntakeDefault(
            name="section_labels",
            label="Section divider labels",
            rule="Overview, Financial Summary, Valuation, Process",
            supplied=False,
        ),
    ),
)

EARNINGS_UPDATE_INTAKE = IntakeSpec(
    name="earnings-update",
    title="Deck spec — earnings update",
    preamble=(
        "The deck itself is the fixed 5-slide earnings-update layout (no slide "
        "options).",
    ),
    # The earnings-update deck spec asks NOTHING: both quarters are defaulted
    # (inferred from the attached interim filing) and the EEO snip is an
    # attachment, so `render_deck_spec_dialogs("earnings-update")` is empty and
    # Step 4 for this deliverable is the defaults echo plus the request.
    fields=(
        IntakeField(
            key="EEO snip",
            prompt_label="Bloomberg EEO snip",
            target="consumed by the earningsupdate-content stage",
            target_kind="attachment",
            plan_input="eeo_snip_path",
            required=True,
            group="documents",
            checklist=(
                "the broker estimates vs. actuals screenshot (an image file). "
                "There is no other source for the estimates-vs-actuals slide and "
                "no sensible default, so the run halts here until it arrives "
                "rather than building the deck without it."
            ),
        ),
    ),
    defaults=(_REPORTING_QUARTER_DEFAULT, _COMPARISON_QUARTER_DEFAULT),
)

_INTAKE_SPECS: dict[str, IntakeSpec] = {
    "pitch": PITCH_INTAKE,
    "earnings-update": EARNINGS_UPDATE_INTAKE,
}


def _spec(deliverable_type: str) -> IntakeSpec:
    """The declared questionnaire for a deliverable.

    Raises ValueError for a deliverable with no questionnaire (e.g. the
    `overview` stub) — the conductor then falls back to prompting from the
    plan's `plan_inputs` specs directly.
    """
    try:
        return _INTAKE_SPECS[deliverable_type]
    except KeyError:
        raise ValueError(
            f"no deck-spec questionnaire for deliverable type {deliverable_type!r}; "
            f"known: {sorted(_INTAKE_SPECS)}"
        ) from None


# ---------------------------------------------------------------------------
# Answer-mapping tables — all derived from the specs above
# ---------------------------------------------------------------------------

# plan_inputs name -> the attached document whose saved path becomes it. The
# conductor's post-drop resolution table: it saved each file under
# <deal_dir>/filings/, so it knows which path is which, and this is what stops
# a dropped CIM from never reaching `cim_path`. An OPTIONAL one that never
# arrived stays out of plan_inputs entirely (never pre-seeded None); a REQUIRED
# one that never arrived halts the run — `attachment_inputs(required=True)`.
PITCH_ATTACHMENT_PLAN_INPUTS: dict[str, str] = PITCH_INTAKE.attachment_inputs()

EARNINGS_UPDATE_ATTACHMENT_PLAN_INPUTS: dict[str, str] = (
    EARNINGS_UPDATE_INTAKE.attachment_inputs()
)

# Question header -> plan_inputs name. Every dialog answer maps through one of
# these; the converters below turn slide-count / include-omit answers into the
# typed values. Empty for earnings-update, which asks nothing.
PITCH_DIALOG_PLAN_INPUTS: dict[str, str] = PITCH_INTAKE.targets("plan-input")

EARNINGS_UPDATE_DIALOG_PLAN_INPUTS: dict[str, str] = (
    EARNINGS_UPDATE_INTAKE.targets("plan-input")
)

# Fallback-prompt item number -> plan_inputs name. Numbered from the same
# field order the dialogs are rendered in, so the text fallback and the
# dialogs can never drift apart.
PITCH_ITEM_PLAN_INPUTS: dict[int, str] = PITCH_INTAKE.item_targets()

EARNINGS_UPDATE_ITEM_PLAN_INPUTS: dict[int, str] = (
    EARNINGS_UPDATE_INTAKE.item_targets()
)

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
# (name -> the default rule, human-readable). The rule is the same string the
# text prompt lists and the echo falls back to.
PITCH_DEFAULT_SUPPLIED_INPUTS: dict[str, str] = PITCH_INTAKE.default_rules(
    supplied=True
)

# Optional plan inputs left OUT of plan_inputs on default (the consuming
# skills apply their own defaults).
PITCH_DEFAULT_UNSET_INPUTS: dict[str, str] = PITCH_INTAKE.default_rules(supplied=False)

EARNINGS_UPDATE_DEFAULT_SUPPLIED_INPUTS: dict[str, str] = (
    EARNINGS_UPDATE_INTAKE.default_rules(supplied=True)
)

EARNINGS_UPDATE_DEFAULT_UNSET_INPUTS: dict[str, str] = (
    EARNINGS_UPDATE_INTAKE.default_rules(supplied=False)
)


def render_deck_spec_dialogs(deliverable_type: str) -> list[list[dict]]:
    """Return the locked deck-spec dialogs for a deliverable, generated.

    Each inner list is one `AskUserQuestion` call's `questions` payload —
    render them in order, unchanged. Fresh payloads every call, so a caller
    mutating the result cannot affect the next render. Raises ValueError for a
    deliverable with no questionnaire (e.g. the `overview` stub).

    Returns an EMPTY list for **earnings-update**, which has no deck-spec
    questions left: both quarters are defaulted and the EEO snip is an
    attachment. Attachments are never here, and never a dialog anywhere — see
    :func:`render_run_attachment_request`.
    """
    return render_dialogs(_spec(deliverable_type), target_kinds=("plan-input",))


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
    through the dialogs. Which computed values a deliverable needs comes from
    its spec's echo templates, so an absent one is named rather than rendered
    as a blank.
    """
    return render_defaults_echo(
        _spec(deliverable_type),
        {
            "client_name": client_name,
            "presentation_date": presentation_date,
            "reporting_quarter": reporting_quarter,
            "comparison_quarter": comparison_quarter,
        },
    )


# ---------------------------------------------------------------------------
# Text fallback (surfaces without the interactive question UI)
# ---------------------------------------------------------------------------


def render_deck_spec_prompt(deliverable_type: str) -> str:
    """Return the locked deck-spec text prompt for a deliverable, generated.

    This is the FALLBACK for surfaces where the interactive question UI
    (`AskUserQuestion`) is unavailable. It is generated from the same spec as
    :func:`render_deck_spec_dialogs`, so it asks the same items in the same
    order, quotes the same question wording and option labels/descriptions,
    lists the same defaults, and embeds the same documents checklist. Raises
    ValueError for a deliverable with no questionnaire (e.g. the `overview`
    stub) — the conductor then falls back to prompting from the plan's
    `plan_inputs` specs directly.
    """
    return render_prompt(_spec(deliverable_type))


def render_run_attachment_request(deliverable_type: str | None = None) -> str:
    """Return THE attachment request for a run: one message, posted once.

    The conductor posts this **after every question in the run has been
    answered** — deal-init's and the deck spec's — and then waits once for the
    analyst to drop the files into chat. It is the whole of the attachment
    conversation: there is no dialog to render, no status to collect, and no
    second pause.

    The list merges `deal_init.INIT_INTAKE`'s G7 filings with the deliverable's
    own documents, in that order, split into REQUIRED and OPTIONAL. It reaches
    across into deal-init deliberately: the filings are the one attachment every
    deliverable needs, they were formerly re-described in each deck spec in
    different words, and merging from the declarations is what removed that
    second wording.

    `deliverable_type=None`, or a declared deliverable with no questionnaire
    (the `overview` stub, `one-off-skill`), returns deal-init's filings alone
    rather than raising: a deal still needs its filings when there is no deck
    spec to add to them. A deliverable type that is not a `DeliverableType` at
    all still raises — that is a typo, not a stub.
    """
    specs = [INIT_INTAKE]
    if deliverable_type is not None:
        if deliverable_type not in DELIVERABLE_TYPES:
            raise ValueError(
                f"unknown deliverable type {deliverable_type!r}; known: "
                f"{sorted(DELIVERABLE_TYPES)}"
            )
        if deliverable_type in _INTAKE_SPECS:
            specs.append(_INTAKE_SPECS[deliverable_type])
    return render_attachment_request(*specs)


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
