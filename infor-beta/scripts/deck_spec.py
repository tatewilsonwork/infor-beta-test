"""Deck-spec questionnaires — the locked per-deliverable prompts (Step 4).

The conductor renders ONE of these prompts verbatim right after deal-init, in a
single message, so every run of a deliverable asks the analyst the same
questions in the same order (mirrors `deal_init.render_init_prompt`, which owns
the G7 questions — a deck spec must never re-ask those). The numbered items map
1:1 onto the plan's `plan_inputs` via the `*_ITEM_PLAN_INPUTS` tables below;
unanswered optional items are simply left out of the collected dict (the
resolver hands the consuming stage `None`, and the wireframe/skills apply the
bracketed defaults).

Slide-mix answers are converted deterministically, never improvised:

  - "N Financial Summary slides"  -> ``financial_metric_count = 4 * N``
    (:func:`metric_count_from_slides`; each FS slide shows four metric tiles)
  - "N acquisition-target slides" -> ``market_entry_target_count = 2 * N``
    (:func:`market_entry_targets_from_slides`; two targets per slide)
  - Key Investment Highlights "omit" -> ``include_investment_highlights = False``
    (any include variant leaves the input unset = included; analyst-dictated
    highlight copy belongs in the analyst notes, not a separate input)

No LLM calls, no dispatch — this module only owns the locked text + the answer
converters, so the questionnaire (and therefore the deck layout the answers
produce) is reproducible run over run.
"""

from __future__ import annotations

_PITCH_SPEC_PROMPT = """\
Deck spec — pitch

Answer by item number; reply "defaults" to accept every [bracketed] default.
Items marked REQUIRED have no default.

1. Client name on the cover:  [the subject company name from deal-init]
2. Presentation date:         REQUIRED — spelled-out month + four-digit year,
                              e.g. "April 2026"
3. Analyst notes:             REQUIRED — the raw notes behind the executive
                              summary, company overview, risks, and takeaways
4. Reporting quarter:         REQUIRED — the LTM bridge quarter, e.g. "Q2 2026"
5. Comparison quarter:        REQUIRED — prior-year same quarter, e.g. "Q2 2025"
6. Financial Summary slides:  [1 slide — 4 metrics] / "2 slides" — 8 metrics
7. Acquisition-target slides: [4 slides — 8 targets] — 1 to 4 slides, two
                              targets per slide; name specific targets if you
                              have them
8. Key Investment Highlights: [include — drafted from your notes] /
                              "include — I'll provide them" (put the highlight
                              copy in item 3) / "omit" — drops the slide
9. Section divider labels:    [Overview, Financial Summary, Valuation, Process]
10. Valuation range:          [none] — optional executive-summary language
11. Risk notes:               [none] — optional steer for the Considerations /
                              Mitigants slide
12. CIM / management pres.:   [none] — attach the file or give its path

Documents (attach in this chat if not already attached at deal-init):
- The G7 filings: latest four annual statements / 10-Ks plus the current-year
  and prior-year interim statements (5-year history + LTM bridge).
- SEDI "Insider Information by Issuer" PDF — Canadian public targets only;
  without it the ownership slide's insider side stays a placeholder.
- Bloomberg ownership export (.xlsm) — without it the ownership slide's
  institutions side stays a placeholder.
"""

_EARNINGS_UPDATE_SPEC_PROMPT = """\
Deck spec — earnings update

Answer by item number. All three items are REQUIRED; the deck itself is the
fixed 5-slide earnings-update layout (no slide options).

1. Reporting quarter:         e.g. "Q2 2026"
2. Comparison quarter:        prior-year same quarter, e.g. "Q2 2025"
3. Bloomberg EEO snip:        attach the screenshot or give its absolute path

Documents (attach in this chat if not already attached at deal-init):
- The G7 filings: latest four annual statements / 10-Ks plus the current-year
  and prior-year interim statements (5-year history + LTM bridge).
"""

_SPEC_PROMPTS: dict[str, str] = {
    "pitch": _PITCH_SPEC_PROMPT,
    "earnings-update": _EARNINGS_UPDATE_SPEC_PROMPT,
}

# Questionnaire item number -> plan_inputs name. Items 6/7/8 pass through the
# converters below; every other answer is the plan-input value verbatim.
PITCH_ITEM_PLAN_INPUTS: dict[int, str] = {
    1: "client_name",
    2: "presentation_date",
    3: "analyst_notes",
    4: "reporting_quarter",
    5: "comparison_quarter",
    6: "financial_metric_count",        # metric_count_from_slides(answer)
    7: "market_entry_target_count",     # market_entry_targets_from_slides(answer)
    8: "include_investment_highlights", # False only on "omit"
    9: "section_labels",
    10: "valuation_range",
    11: "risk_notes",
    12: "cim_path",
}

EARNINGS_UPDATE_ITEM_PLAN_INPUTS: dict[int, str] = {
    1: "reporting_quarter",
    2: "comparison_quarter",
    3: "eeo_snip_path",
}

# Slide-geometry constants behind the converters (and the questionnaire copy).
METRICS_PER_FINANCIAL_SUMMARY_SLIDE = 4
TARGETS_PER_MARKET_ENTRY_SLIDE = 2
MAX_MARKET_ENTRY_SLIDES = 4  # PitchDeckContent caps market_entry_targets at 8


def render_deck_spec_prompt(deliverable_type: str) -> str:
    """Return the locked deck-spec prompt for a deliverable, verbatim.

    Raises ValueError for a deliverable with no questionnaire (e.g. the
    `overview` stub) — the conductor then falls back to prompting from the
    plan's `plan_inputs` specs directly.
    """
    try:
        return _SPEC_PROMPTS[deliverable_type]
    except KeyError:
        raise ValueError(
            f"no deck-spec questionnaire for deliverable type {deliverable_type!r}; "
            f"known: {sorted(_SPEC_PROMPTS)}"
        ) from None


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
