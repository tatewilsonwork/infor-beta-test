"""Declarative analyst intake — one spec behind every rendering (Phase H1).

Every analyst-facing question the conductor asks before a run starts is
declared **once**, here, as an :class:`IntakeSpec`. Both renderings are
generated from it:

  - the interactive **dialogs** — one `AskUserQuestion` `questions` payload per
    call (:func:`render_dialogs`);
  - the single-message **text prompt** used on surfaces where the interactive
    question UI is unavailable (:func:`render_prompt`), plus its attachment
    checklist (:func:`render_note`) and its defaults echo
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
                      `*_DIALOG_PLAN_INPUTS` / `*_DIALOG_FIELDS` /
                      `*_DOCUMENTS_DIALOG_TARGETS` tables the conductor maps
                      answers through.
  - ``prompt_label`` — the numbered item's label in the text prompt.
  - ``target``       — where the answer lands (a `plan_inputs` name, a
                      `DealContext` field, or prose for an attachment gate).
  - ``target_kind``  — which of those three, and therefore which table the
                      field appears in.
  - ``required``     — no default; the prompt marks the item REQUIRED.
  - ``question`` + ``options`` — the dialog wording, verbatim in both
                      renderings. Exactly one option may be marked
                      ``default=True``; the prompt renders it ``[bracketed]``.

A field with no options is a **free-text** item (the subject company name —
pure free text with nothing to suggest). It is a numbered item in the text
prompt and a plain chat question in the interactive flow, never a dialog
question.

An ``attachment`` field is a **status gate**: file bytes cannot come through a
dialog, so the question only asks attached / will-drop-next-message / none
while the file itself arrives via the chat input. Its answer is deliberately
NOT a plan input — the consuming stage discovers the saved file under
`<deal_dir>/filings/` — so it never appears in a plan-input table, and it is
not a numbered prompt item either: its ``checklist`` line goes into the spec's
:class:`IntakeNote` instead, which is where an analyst looks for what to
attach.

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
    """One intake item, in both renderings.

    Two kinds, distinguished by whether `options` is empty:

      - **dialog** — `question` + 2-4 `options`; rendered as one
        `AskUserQuestion` question and as a numbered prompt item listing the
        same question and the same option labels/descriptions.
      - **free-text** — no options, a `hint` instead; a numbered prompt item
        and a plain chat question, never a dialog question (nothing to
        suggest).
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

    def __post_init__(self) -> None:
        object.__setattr__(self, "options", tuple(self.options))
        if not self.key.strip() or self.key != self.key.strip():
            raise ValueError(f"intake field key {self.key!r} must be non-blank, untrimmed")
        if len(self.key) > HEADER_MAX_CHARS:
            raise ValueError(
                f"dialog header {self.key!r} is {len(self.key)} chars; "
                f"AskUserQuestion allows {HEADER_MAX_CHARS}"
            )
        if not self.target.strip():
            raise ValueError(f"field {self.key!r} needs a target")
        if self.target_kind not in TARGET_KINDS:
            raise ValueError(
                f"field {self.key!r} has target_kind {self.target_kind!r}; "
                f"known: {list(TARGET_KINDS)}"
            )
        if self.target_kind == "attachment":
            if not self.checklist.strip():
                raise ValueError(
                    f"attachment field {self.key!r} needs a checklist line — it is "
                    f"not a numbered prompt item, so the note is where the analyst "
                    f"reads what to attach"
                )
        else:
            if self.checklist:
                raise ValueError(
                    f"field {self.key!r} is not an attachment gate; only attachment "
                    f"fields contribute checklist bullets"
                )
            if not self.prompt_label.strip():
                raise ValueError(f"field {self.key!r} needs a prompt_label")
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
    def is_dialog(self) -> bool:
        """True when this field renders as an `AskUserQuestion` question."""
        return bool(self.options)

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
class IntakeNote:
    """The attachment checklist posted alongside the dialogs as plain text.

    File bytes cannot come through a dialog, so what the analyst must attach —
    and what degrades without it — is prose, not a question. `header` is the
    first line; `body` paragraphs and `bullets` follow, then one bullet per
    attachment field's `checklist`, in field order.
    """

    header: str
    body: tuple[str, ...] = ()
    bullets: tuple[str, ...] = ()

    def __post_init__(self) -> None:
        object.__setattr__(self, "body", tuple(self.body))
        object.__setattr__(self, "bullets", tuple(self.bullets))
        if not self.header.strip():
            raise ValueError("an intake note needs a header line")


@dataclass(frozen=True)
class IntakeSpec:
    """One locked questionnaire: deal-init, or one deliverable's deck spec."""

    name: str
    title: str
    preamble: tuple[str, ...] = ()
    fields: tuple[IntakeField, ...] = ()
    defaults: tuple[IntakeDefault, ...] = ()
    note: IntakeNote | None = None

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
            # Dialog headers are the answer-mapping keys, and the attachment
            # gates share the namespace with the plan-input questions — a
            # collision would route a status answer into plan_inputs.
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
        asked = set(self.targets("plan-input").values())
        clash = asked & set(names)
        if clash:
            raise ValueError(
                f"intake spec {self.name!r} both asks for and defaults "
                f"{sorted(clash)}"
            )
        if self.note is None and any(
            f.target_kind == "attachment" for f in self.fields
        ):
            raise ValueError(
                f"intake spec {self.name!r} has attachment gates but no note to "
                f"carry their checklist lines"
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

        Attachment gates are excluded: they are described by the note, which
        is where an analyst reads what to attach.
        """
        return tuple(f for f in self.fields if f.target_kind != "attachment")

    def targets(self, target_kind: TargetKind) -> dict[str, str]:
        """`{dialog header: target}` for one kind — the answer-mapping table.

        Dialog fields only. A free-text item is answered in chat, so it has no
        header to map through.
        """
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
    render. Returns `[]` when the filter selects nothing (a deliverable with
    no attachment gates has no documents dialog to render).
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


def render_note(spec: IntakeSpec) -> str:
    """Generate the attachment checklist: header, prose, bullets, gates."""
    if spec.note is None:
        raise ValueError(f"intake spec {spec.name!r} has no note")
    parts = [spec.note.header]
    parts.extend(_fill(para) for para in spec.note.body)
    bullets = list(spec.note.bullets) + [
        f.checklist for f in spec.fields if f.target_kind == "attachment"
    ]
    parts.extend(
        _fill(bullet, initial_indent="- ", subsequent_indent="  ")
        for bullet in bullets
    )
    return "\n".join(parts) + "\n"


def _render_answer_instructions(spec: IntakeSpec) -> str:
    items = spec.prompt_fields()
    sentences = []
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
    defaults and embeds the attachment checklist.
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
    if spec.note is not None:
        sections.append(render_note(spec).rstrip("\n"))
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
