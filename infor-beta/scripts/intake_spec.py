"""Declarative analyst intake — one spec behind every rendering (Phase H1).

Every analyst-facing question the conductor asks before a run starts is
declared **once**, here, as an :class:`IntakeSpec`. Both renderings are
generated from it:

  - the interactive **dialogs** — one `AskUserQuestion` `questions` payload per
    call (:func:`render_dialogs`);
  - the single-message **text prompt** used on surfaces where the interactive
    question UI is unavailable (:func:`render_prompt`), plus the attachment
    request (:func:`render_attachment_request`) and the defaults echo
    (:func:`render_defaults_echo`).

Why this module exists. The locked-questionnaire principle — *every run asks
the same questions, in the same order, with the same options* — used to be
half structural and half conventional. The item -> plan-input **mapping** was
derived from the dialog order, so answer mapping could not drift; but the text
prompt was a hand-written string literal carrying its own copy of every
question's wording, defaults and option labels. Changing a dialog option label
left the prompt describing a choice the dialog no longer offered, and nothing
failed. Two hand-maintained content renderings existed; a third (an inline
form, H2) would have made it worse. Now there is one, and
:mod:`tests.test_intake_spec` asserts each rendering is exactly what the
generators produce from the spec.

Generation is **deterministic**: the spec is frozen dataclasses of plain
strings, wrapping is `textwrap` at a fixed width, and nothing consults the
clock, the environment, or a set/dict iteration order. Two calls return equal
values; two releases return the same values unless the spec changed.

What is declared per field:

  - ``key``          — the dialog `header`; also the key in the
                      `*_DIALOG_PLAN_INPUTS` / `*_DIALOG_FIELDS` tables the
                      conductor maps answers through.
  - ``prompt_label`` — the item's label: the numbered heading in the text
                      prompt, or an attachment's short handle.
  - ``target``       — where the answer lands (a `plan_inputs` name, a
                      `DealContext` field, or — for an attachment — prose
                      naming the stage that consumes the file).
  - ``target_kind``  — which of those three, and therefore which table the
                      field appears in.
  - ``required``     — no default; the prompt marks the item REQUIRED, and an
                      attachment is listed under the request's REQUIRED
                      section.
  - ``question`` + ``options`` — the dialog wording, verbatim in both
                      renderings. Exactly one option may be marked
                      ``default=True``; the prompt renders it ``[bracketed]``.
  - ``checklist`` + ``plan_input`` — attachments only (below).

Three kinds of field, and the validator checks each as its own shape:

  - **dialog** — `question` + 2-4 `options`.
  - **free-text** — no options, a `hint` instead (the subject company name —
    pure free text with nothing to suggest). A numbered item in the text
    prompt and a plain chat question in the interactive flow, never a dialog
    question.
  - **attachment** — a file. Nothing is *asked*: the analyst is told what to
    attach and drops the files into chat.

An attachment field is never a question in any rendering. It carries no
`question`, no `options` and no `hint`; it carries a ``checklist`` line, which
is the bullet :func:`render_attachment_request` lists it as, and which must
state what a missing file costs — that warning has nowhere else to live now
that there is no dialog option description to hold it. Attachments were
formerly *status gates* (attached / will-drop-next-message / none), which
asked the analyst to assert something the filesystem already knew and could
contradict; v0.5.50 replaced three such dialogs per pitch run with one request
message and one pause.

``plan_input`` names the plan input the saved file's path becomes (the CIM,
the Bloomberg EEO snip). Empty — the common case — means the consuming stage
discovers the file under `<deal_dir>/filings/` for itself (the SEDI report,
the Bloomberg ownership export). A **required** attachment feeding a plan
input is the one case where a missing file must halt the run rather than
resolve to None, so the requiredness declared here is checked against the plan
YAML.

No LLM calls, no dispatch, no I/O — this module is the questionnaire and its
renderers, nothing else.
"""

from __future__ import annotations

import textwrap
from dataclasses import dataclass
from string import Formatter
from typing import Iterable, Literal, Mapping, Sequence

# ---------------------------------------------------------------------------
# Shape contract (mirrored by the AskUserQuestion payload tests)
# ---------------------------------------------------------------------------

DIALOG_MAX_QUESTIONS = 4
DIALOG_MIN_OPTIONS = 2
DIALOG_MAX_OPTIONS = 4
HEADER_MAX_CHARS = 12

# Text-prompt line width. Fixed, so wrapping is reproducible.
PROMPT_WIDTH = 78

TargetKind = Literal["plan-input", "deal-context", "attachment"]
TARGET_KINDS: tuple[TargetKind, ...] = ("plan-input", "deal-context", "attachment")

DEFAULTS_ECHO_HEADER = (
    "Defaults in effect — reply to override any of these before the run starts:"
)
DEFAULTS_PROMPT_HEADER = "Defaulted unless you override here (no need to answer):"

# The attachment request — one message, posted once, after every question has
# been answered. Code-owned so the merged live message and the per-spec text
# fallback are framed identically; the bullets underneath come from the fields.
ATTACHMENT_REQUEST_HEADER = (
    "Attachments — drop these into this chat now. Nothing left to answer."
)
ATTACHMENT_REQUEST_BODY = (
    "I wait here until the files arrive, then carry on without asking anything "
    "else. If one of the OPTIONAL documents does not exist or you would rather "
    "skip it, say so and I proceed with whatever was attached — each bullet says "
    "what the run loses without that file. A REQUIRED document that never "
    "arrives halts the run instead of degrading it."
)
ATTACHMENT_REQUIRED_HEADER = "REQUIRED:"
ATTACHMENT_OPTIONAL_HEADER = "OPTIONAL:"


def _fill(text: str, *, initial_indent: str = "", subsequent_indent: str = "") -> str:
    """Wrap one paragraph deterministically at :data:`PROMPT_WIDTH`."""
    return textwrap.fill(
        text,
        width=PROMPT_WIDTH,
        initial_indent=initial_indent,
        subsequent_indent=subsequent_indent,
        break_long_words=False,
        break_on_hyphens=False,
    )


# ---------------------------------------------------------------------------
# The declarative model
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class IntakeOption:
    """One dialog choice. `label` + `description` are the analyst-facing text.

    `default` marks the choice taken when the analyst says nothing; the text
    prompt renders it as the item's `[bracketed]` default. At most one option
    per field may carry it, and a `required` field may carry none.
    """

    label: str
    description: str
    default: bool = False

    def __post_init__(self) -> None:
        if not self.label.strip():
            raise ValueError("an intake option needs a label")
        if not self.description.strip():
            raise ValueError(f"option {self.label!r} needs a description")

    def as_dialog_option(self) -> dict[str, str]:
        """The `AskUserQuestion` option payload — never carries `default`."""
        return {"label": self.label, "description": self.description}


@dataclass(frozen=True)
class IntakeField:
    """One intake item, in every rendering.

    Three kinds, each validated as its own shape:

      - **dialog** — `question` + 2-4 `options`; rendered as one
        `AskUserQuestion` question and as a numbered prompt item listing the
        same question and the same option labels/descriptions.
      - **free-text** — no options, a `hint` instead; a numbered prompt item
        and a plain chat question, never a dialog question (nothing to
        suggest).
      - **attachment** — a file, never a question: one bullet of the
        attachment request, carrying its `checklist` line and, when the saved
        path becomes a plan input, its `plan_input` name.
    """

    key: str
    prompt_label: str
    target: str
    target_kind: TargetKind = "plan-input"
    required: bool = False
    group: str = "main"
    question: str = ""
    options: tuple[IntakeOption, ...] = ()
    hint: str = ""
    checklist: str = ""
    plan_input: str = ""

    def __post_init__(self) -> None:
        object.__setattr__(self, "options", tuple(self.options))
        if not self.key.strip() or self.key != self.key.strip():
            raise ValueError(f"intake field key {self.key!r} must be non-blank, untrimmed")
        if len(self.key) > HEADER_MAX_CHARS:
            raise ValueError(
                f"dialog header {self.key!r} is {len(self.key)} chars; "
                f"AskUserQuestion allows {HEADER_MAX_CHARS}"
            )
        if not self.prompt_label.strip():
            raise ValueError(f"field {self.key!r} needs a prompt_label")
        if not self.target.strip():
            raise ValueError(f"field {self.key!r} needs a target")
        if self.target_kind not in TARGET_KINDS:
            raise ValueError(
                f"field {self.key!r} has target_kind {self.target_kind!r}; "
                f"known: {list(TARGET_KINDS)}"
            )
        if self.is_attachment:
            # An attachment is its own kind, not a dialog field about a file.
            # Nothing is asked, so there is no question wording to carry — the
            # request names the document and the analyst drops it into chat.
            if self.options or self.question or self.hint:
                raise ValueError(
                    f"attachment field {self.key!r} carries dialog wording "
                    f"(question / options / hint) — an attachment is never asked "
                    f"about: it is one bullet of the attachment request, and the "
                    f"analyst answers by dropping the file into chat"
                )
            if not self.checklist.strip():
                raise ValueError(
                    f"attachment field {self.key!r} needs a checklist line — it is "
                    f"the bullet the attachment request lists it as, and the only "
                    f"place the analyst reads what a missing file costs"
                )
            return
        if self.checklist:
            raise ValueError(
                f"field {self.key!r} is not an attachment; only attachment fields "
                f"contribute checklist bullets"
            )
        if self.plan_input:
            raise ValueError(
                f"field {self.key!r} is not an attachment; a question's plan input "
                f"is its `target`, and `plan_input` is where an attachment's saved "
                f"path lands"
            )
        if self.options:
            if not self.question.strip().endswith("?"):
                raise ValueError(
                    f"field {self.key!r} question must read as a question: "
                    f"{self.question!r}"
                )
            if not DIALOG_MIN_OPTIONS <= len(self.options) <= DIALOG_MAX_OPTIONS:
                raise ValueError(
                    f"field {self.key!r} has {len(self.options)} options; "
                    f"AskUserQuestion allows {DIALOG_MIN_OPTIONS}-{DIALOG_MAX_OPTIONS}"
                )
            if self.hint:
                raise ValueError(
                    f"field {self.key!r} is a dialog question — its wording lives in "
                    f"`question` and the options, not in a prompt-only hint (that is "
                    f"the drift surface H1 removed)"
                )
            labels = [opt.label for opt in self.options]
            if len(labels) != len(set(labels)):
                raise ValueError(f"field {self.key!r} has duplicate option labels")
            if len([opt for opt in self.options if opt.default]) > 1:
                raise ValueError(f"field {self.key!r} marks more than one default option")
            if self.required and self.default_option is not None:
                raise ValueError(
                    f"field {self.key!r} is REQUIRED but marks a default option — "
                    f"REQUIRED means there is no default"
                )
        else:
            if self.question:
                raise ValueError(
                    f"field {self.key!r} has no options, so it is a plain chat "
                    f"question — put its wording in `hint`"
                )
            if not self.hint.strip():
                raise ValueError(f"free-text field {self.key!r} needs a hint")

    @property
    def is_attachment(self) -> bool:
        """True when this field is a file the analyst drops into chat."""
        return self.target_kind == "attachment"

    @property
    def is_dialog(self) -> bool:
        """True when this field renders as an `AskUserQuestion` question."""
        return bool(self.options)

    @property
    def is_free_text(self) -> bool:
        """True when this field is a plain chat question with nothing to suggest."""
        return not self.options and not self.is_attachment

    @property
    def default_option(self) -> IntakeOption | None:
        """The choice taken when the analyst says nothing, if there is one."""
        for opt in self.options:
            if opt.default:
                return opt
        return None

    def as_dialog_question(self) -> dict:
        """The `AskUserQuestion` question payload for this field."""
        if not self.is_dialog:
            raise ValueError(f"field {self.key!r} is free text, not a dialog question")
        return {
            "question": self.question,
            "header": self.key,
            "multiSelect": False,
            "options": [opt.as_dialog_option() for opt in self.options],
        }


@dataclass(frozen=True)
class IntakeDefault:
    """A value the conductor supplies or leaves unset instead of asking for it.

    `rule` is the one human-readable statement of the default — it is what the
    text prompt lists and what the `*_DEFAULT_*_INPUTS` tables carry, so the
    two cannot disagree.

    `supplied=True` means the conductor computes the value and puts it in
    `plan_inputs`; `supplied=False` means the input is left OUT and the
    consuming skill applies its own default.

    The defaults echo (:func:`render_defaults_echo`) posts one line per
    `echoed` default: `echo` is a `str.format` template over the computed
    values, or None to echo the static `rule`. `echoed=False` folds this
    default onto another entry's line (the comparison quarter is echoed
    alongside the reporting quarter, as one range).
    """

    name: str
    label: str
    rule: str
    supplied: bool
    echoed: bool = True
    echo: str | None = None
    echo_label: str | None = None

    def __post_init__(self) -> None:
        if not self.name.strip():
            raise ValueError("an intake default needs a plan_inputs name")
        if not self.label.strip():
            raise ValueError(f"default {self.name!r} needs a label")
        if not self.rule.strip():
            raise ValueError(f"default {self.name!r} needs a rule")
        if not self.echoed and (self.echo or self.echo_label):
            raise ValueError(
                f"default {self.name!r} is not echoed on its own line, so it cannot "
                f"carry echo text"
            )

    @property
    def echo_line_label(self) -> str:
        return self.echo_label or self.label

    @property
    def echo_values(self) -> tuple[str, ...]:
        """The computed value names this default's echo template needs."""
        if self.echo is None:
            return ()
        return tuple(
            name for _, name, _, _ in Formatter().parse(self.echo) if name
        )


@dataclass(frozen=True)
class IntakeSpec:
    """One locked questionnaire: deal-init, or one deliverable's deck spec."""

    name: str
    title: str
    preamble: tuple[str, ...] = ()
    fields: tuple[IntakeField, ...] = ()
    defaults: tuple[IntakeDefault, ...] = ()

    def __post_init__(self) -> None:
        object.__setattr__(self, "preamble", tuple(self.preamble))
        object.__setattr__(self, "fields", tuple(self.fields))
        object.__setattr__(self, "defaults", tuple(self.defaults))
        if not self.name.strip():
            raise ValueError("an intake spec needs a name")
        if not self.title.strip():
            raise ValueError(f"intake spec {self.name!r} needs a title")
        keys = [f.key for f in self.fields]
        if len(keys) != len(set(keys)):
            # Dialog headers are the answer-mapping keys, and the attachments
            # share the namespace with the plan-input questions — a collision
            # would route one question's answer through another's target.
            dupes = sorted({k for k in keys if keys.count(k) > 1})
            raise ValueError(
                f"intake spec {self.name!r} reuses dialog header(s) {dupes}"
            )
        seen_groups: list[str] = []
        for f in self.fields:
            if f.group not in seen_groups:
                seen_groups.append(f.group)
            elif seen_groups[-1] != f.group:
                raise ValueError(
                    f"intake spec {self.name!r} splits group {f.group!r} — a group's "
                    f"fields must be contiguous so dialog batching is deterministic"
                )
        names = [d.name for d in self.defaults]
        if len(names) != len(set(names)):
            raise ValueError(f"intake spec {self.name!r} declares a default twice")
        # A plan input is supplied exactly one way: asked as a question, carried
        # in by an attachment, or defaulted.
        supplied = list(self.targets("plan-input").values()) + list(
            self.attachment_inputs()
        )
        if len(supplied) != len(set(supplied)):
            dupes = sorted({n for n in supplied if supplied.count(n) > 1})
            raise ValueError(
                f"intake spec {self.name!r} collects plan input(s) {dupes} twice"
            )
        clash = set(supplied) & set(names)
        if clash:
            raise ValueError(
                f"intake spec {self.name!r} both asks for and defaults "
                f"{sorted(clash)}"
            )

    # -- field views ------------------------------------------------------

    def dialog_fields(
        self,
        *,
        target_kinds: Sequence[TargetKind] | None = None,
        omit: Iterable[str] = (),
    ) -> tuple[IntakeField, ...]:
        """Dialog fields, in spec order, optionally filtered.

        `omit` drops questions by `key` — the slash commands pre-answer the
        deliverable, and the conductor drops any question an earlier message
        already answered.
        """
        skip = set(omit)
        kinds = None if target_kinds is None else set(target_kinds)
        return tuple(
            f
            for f in self.fields
            if f.is_dialog
            and f.key not in skip
            and (kinds is None or f.target_kind in kinds)
        )

    def prompt_fields(self) -> tuple[IntakeField, ...]:
        """The numbered items of the text prompt, in order.

        Attachments are excluded: nothing is asked about them, so they are not
        items competing with the questions. They are the request's bullets.
        """
        return tuple(f for f in self.fields if not f.is_attachment)

    def attachment_fields(
        self, *, required: bool | None = None
    ) -> tuple[IntakeField, ...]:
        """The attachments this spec needs, in spec order.

        `required=True` / `False` selects one section of the request;
        `required=None` (the default) returns both.
        """
        return tuple(
            f
            for f in self.fields
            if f.is_attachment and (required is None or f.required is required)
        )

    def attachment_inputs(self, *, required: bool | None = None) -> dict[str, str]:
        """`{plan_inputs name: the document's label}` for path-carrying attachments.

        The conductor's resolution table: after the drop, each of these saved
        paths becomes that plan input. Attachments the consuming stage finds on
        disk for itself declare no `plan_input` and are absent here.
        """
        return {
            f.plan_input: f.prompt_label
            for f in self.attachment_fields(required=required)
            if f.plan_input
        }

    def targets(self, target_kind: TargetKind) -> dict[str, str]:
        """`{dialog header: target}` for one kind — the answer-mapping table.

        Dialog fields only. A free-text item is answered in chat, so it has no
        header to map through; an attachment is not asked about at all, so
        `"attachment"` is rejected rather than silently answered with `{}`.
        """
        if target_kind == "attachment":
            raise ValueError(
                "attachments are not dialog questions — use attachment_fields() "
                "or attachment_inputs()"
            )
        if target_kind not in TARGET_KINDS:
            raise ValueError(f"unknown target kind {target_kind!r}")
        return {
            f.key: f.target
            for f in self.fields
            if f.is_dialog and f.target_kind == target_kind
        }

    def item_targets(self, target_kind: TargetKind = "plan-input") -> dict[int, str]:
        """`{prompt item number: target}` — the text fallback's answer mapping.

        Numbering follows :meth:`prompt_fields`, which is the order the
        dialogs are rendered in, so the two renderings agree by construction.
        """
        if target_kind == "attachment":
            raise ValueError(
                "attachments are not numbered prompt items — use "
                "attachment_fields() or attachment_inputs()"
            )
        if target_kind not in TARGET_KINDS:
            raise ValueError(f"unknown target kind {target_kind!r}")
        return {
            number: f.target
            for number, f in enumerate(self.prompt_fields(), start=1)
            if f.target_kind == target_kind
        }

    def default_rules(self, *, supplied: bool) -> dict[str, str]:
        """`{plan_inputs name: rule}` for the supplied or the unset defaults."""
        return {d.name: d.rule for d in self.defaults if d.supplied is supplied}


# ---------------------------------------------------------------------------
# Renderers — the only place either rendering is produced
# ---------------------------------------------------------------------------


def render_dialogs(
    spec: IntakeSpec,
    *,
    target_kinds: Sequence[TargetKind] | None = None,
    omit: Iterable[str] = (),
) -> list[list[dict]]:
    """Generate the `AskUserQuestion` payloads for `spec`.

    Each inner list is one call's `questions` payload: fields are batched by
    declared `group`, then chunked at :data:`DIALOG_MAX_QUESTIONS`. Fresh
    dicts every call, so a caller mutating the result cannot affect the next
    render. Returns `[]` when the filter selects nothing (the earnings-update
    deck spec has no questions left at all — every input is defaulted or
    attached).
    """
    fields = spec.dialog_fields(target_kinds=target_kinds, omit=omit)
    batches: list[list[IntakeField]] = []
    for f in fields:
        if batches and batches[-1][0].group == f.group and len(
            batches[-1]
        ) < DIALOG_MAX_QUESTIONS:
            batches[-1].append(f)
        else:
            batches.append([f])
    return [[f.as_dialog_question() for f in batch] for batch in batches]


def render_attachment_request(*specs: IntakeSpec) -> str:
    """Generate THE attachment request: one message, two sections, one pause.

    Every attachment across `specs`, in spec order then field order, split into
    a REQUIRED and an OPTIONAL section. Each bullet is one field's
    `prompt_label` and its `checklist` line, so the consequence of a missing
    file travels with the request for it.

    Called with **one** spec by :func:`render_prompt`, which embeds that spec's
    own list in the text fallback, and with **deal-init plus the deliverable's
    spec** by the conductor, which posts them as a single message after the
    questions. One function, so the live message and the fallback cannot come
    to describe different documents. Returns `""` for a spec with no
    attachments.

    Raises ValueError if two specs declare the same field `key`: the G7 filings
    used to be described once in deal-init's prose and again, in different
    words, in each deck spec's checklist — merging from the declarations is
    what removed that, and a repeated key would quietly bring it back.
    """
    fields = [f for spec in specs for f in spec.attachment_fields()]
    keys = [f.key for f in fields]
    if len(keys) != len(set(keys)):
        dupes = sorted({k for k in keys if keys.count(k) > 1})
        raise ValueError(
            f"attachment request repeats {dupes} — one document, one declaration"
        )
    if not fields:
        return ""
    lines = [ATTACHMENT_REQUEST_HEADER, "", _fill(ATTACHMENT_REQUEST_BODY)]
    for header, required in (
        (ATTACHMENT_REQUIRED_HEADER, True),
        (ATTACHMENT_OPTIONAL_HEADER, False),
    ):
        section = [f for f in fields if f.required is required]
        if not section:
            continue
        lines.append("")
        lines.append(header)
        lines.extend(
            _fill(
                f"{f.prompt_label} — {f.checklist}",
                initial_indent="- ",
                subsequent_indent="  ",
            )
            for f in section
        )
    return "\n".join(lines) + "\n"


def _render_answer_instructions(spec: IntakeSpec) -> str:
    items = spec.prompt_fields()
    sentences = []
    if not items:
        # The earnings-update deck spec: every input is defaulted or attached.
        sentences.append("Nothing to answer here.")
        if spec.defaults:
            sentences.append(
                "Every input is defaulted — reply to override any of the defaults "
                "below."
            )
        if spec.attachment_fields():
            sentences.append("The attachments listed at the end are what I need.")
        return " ".join(sentences)
    if any(f.default_option is not None for f in items):
        sentences.append(
            'Answer by item number; reply "defaults" to accept every [bracketed] '
            "default."
        )
    else:
        sentences.append("Answer by item number.")
    if any(f.required for f in items):
        sentences.append("Items marked REQUIRED have no default.")
    sentences.append(
        "The listed choices are exactly what the interactive dialogs offer — any "
        "other answer is fine too."
    )
    return " ".join(sentences)


def _render_item(number: int, field: IntakeField) -> str:
    head = f"{number}. {field.prompt_label}:"
    if field.required:
        head += " REQUIRED"
    elif field.default_option is not None:
        head += f" [{field.default_option.label}]"
    lines = [head]
    if field.is_dialog:
        lines.append(_fill(field.question, initial_indent="   ", subsequent_indent="   "))
        lines.extend(
            _fill(
                f"{opt.label}: {opt.description}",
                initial_indent="   - ",
                subsequent_indent="     ",
            )
            for opt in field.options
        )
    else:
        lines.append(_fill(field.hint, initial_indent="   ", subsequent_indent="   "))
    return "\n".join(lines)


def _render_defaults_list(spec: IntakeSpec) -> str:
    width = max(len(d.label) for d in spec.defaults) + 4
    lines = [DEFAULTS_PROMPT_HEADER]
    for d in spec.defaults:
        prefix = f"- {d.label}:".ljust(width)
        lines.append(_fill(d.rule, initial_indent=prefix, subsequent_indent=" " * width))
    return "\n".join(lines)


def render_prompt(spec: IntakeSpec) -> str:
    """Generate the single-message text prompt for `spec`.

    The fallback for surfaces without the interactive question UI. It asks the
    same items in the same order as :func:`render_dialogs`, quoting each
    question and every option label and description verbatim, then lists the
    defaults and ends with this spec's half of the attachment request — same
    documents, same order, generated by the same function the live message uses.
    """
    sections = [spec.title]
    sections.extend(_fill(para) for para in spec.preamble)
    sections.append(_fill(_render_answer_instructions(spec)))
    sections.extend(
        _render_item(number, f)
        for number, f in enumerate(spec.prompt_fields(), start=1)
    )
    if spec.defaults:
        sections.append(_render_defaults_list(spec))
    request = render_attachment_request(spec)
    if request:
        sections.append(request.rstrip("\n"))
    return "\n\n".join(sections) + "\n"


def render_defaults_echo(spec: IntakeSpec, values: Mapping[str, str | None]) -> str:
    """Generate the one-shot defaults echo, filling in the computed values.

    Raises ValueError naming every computed value the spec's echo templates
    need and `values` did not supply — a half-filled echo would tell the
    analyst a default is in effect without saying what it is.
    """
    echoed = [d for d in spec.defaults if d.echoed]
    needed: list[str] = []
    for d in echoed:
        for name in d.echo_values:
            if name not in needed:
                needed.append(name)
    missing = [name for name in needed if values.get(name) is None]
    if missing:
        raise ValueError(f"{spec.name} defaults echo needs {', '.join(missing)}")
    lines = [DEFAULTS_ECHO_HEADER]
    for d in echoed:
        text = d.rule if d.echo is None else d.echo.format(**dict(values))
        lines.append(f"- {d.echo_line_label}: {text}")
    return "\n".join(lines) + "\n"
