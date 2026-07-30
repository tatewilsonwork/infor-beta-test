"""The read-the-slides pass over a finished deck — the `deckread` stage.

`deck_repair` measured the geometry and repaired it. `deckcheck` asks whether the
numbers are true. This asks the one question neither of them can: **do the slides
read correctly** — is text drawn over other text, are shapes colliding, is anything
too faint to read, are chart and table labels piling up, is a rasterised range
legible at the size it was placed.

Why it is a stage at all
------------------------
The `deck` transform already writes a review to `vision_review_path`, and through
v0.5.51 that file was the whole of this pass. It is an **agenda**, not a review: 19
KB of questions addressed to nobody, containing no assertion about the deck in
front of it. Nothing in either plan referenced it, so nothing ever answered one of
its questions — a run's own defect sat on a slide whose two pictures that file
listed by name, and the run reported clean. An agenda no stage consumes is an
output the plan believes it has.

Reading slides is judgment: it needs a model looking at pixels and saying what it
sees. A transform cannot do it — a transform has no `Task` tool by construction and
no eyes — so this is a **dispatched** stage with its own SKILL.md, scheduled in the
final wave beside `deckcheck`, reading the artefact the analyst receives.

The split of labour, same as everywhere else
--------------------------------------------
**Mechanical, here.** Render the finished deck, extract every embedded picture at
native resolution, and build the work list: which slides carry content a string
scan cannot reach, which declared boxes overlap, where the fill tightened a
clearance the library left wider. That is `deck_contract.vision_pass`, unchanged and
un-duplicated — this module re-runs it on the FINAL artefact, because a pitch deck
gains its charts after assembly and a work list built before them lists the wrong
pictures.

**Judgement, in the SKILL.md.** Whether slide 8's tile labels actually collide,
whether the cap-table paste is legible at 4.5", whether the pie's segment labels
overlap. No measurement answers those; `deck_contract` says so itself, which is why
its vision tier emits an agenda and no verdict.

Advisory, always
----------------
Every finding is `SEVERITY_ADVISORY` and :class:`SlideFinding` refuses to be
constructed otherwise — the same construction-time refusal `deckcheck.CheckFinding`
carries, for the same reason. No shipped stage gates (v0.5.49), and a pass whose
verdict is "this looks wrong to me" is the last one that should start: the deck
already converged against measured renders, so a halt here would be an opinion
overruling a measurement.

Not a finding
-------------
:data:`NOT_YOURS` is the list, rendered into the work list and the report rather
than only stated in prose, so it sits in front of whoever is reading at the moment
they would otherwise report one. Geometry is `deck_repair`'s and it has already
converged; unsubstituted `[Placeholder for …]` tokens are `deck_contract`'s and it
has already run; error values in CapIQ-dependent cells are the normal state of a
shipped artefact here; and figures being *wrong* is `deckcheck`'s question, running
concurrently in this same wave.

Read nothing, report nothing
----------------------------
The one failure this module is built to make impossible is a clean report from a
pass that never looked. `slides_read` is a declared output, `rendered` is on the
evidence, and a report built without renders opens by saying so — because "no
findings" and "no renderer" are the same text file otherwise.
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from deck_contract import SEVERITY_ADVISORY, VISION_CHECKLIST, VisionPass, vision_pass

#: What a reading finding can be *about*. The first four are `VISION_CHECKLIST`'s
#: four, one token each, so a finding names the thing the reader was asked to look
#: for; the rest are what a reader reports that the checklist does not name.
#: Constrained for the same reason `deckcheck.VERDICTS` is: a free-text kind is a
#: field nothing can group, sort, or count.
READ_KINDS = (
    "text-over-text",
    "collision",
    "faint-contrast",
    "label-pileup",
    "illegible-figure",
    "clipped-text",
    "wrong-content",
)

#: What this pass must not report, and who owns each. Rendered into the work list
#: and the report (`_not_yours_block`) so the boundary is in front of the reader.
NOT_YOURS = (
    (
        "geometry — overflow, font sizes, autofit scales",
        "`deck_repair` converged the deck from measured renders inside the assembler, "
        "and the stage would have failed if it had not. Never re-measure, and never "
        "edit a shape.",
    ),
    (
        "unsubstituted `[Placeholder for …]` tokens",
        "`deck_contract`'s substitution check has already run; a placeholder region the "
        "plan defers on purpose (the comps and precedents slides) is expected.",
    ),
    (
        "error values in CapIQ-dependent cells",
        "Capital IQ cannot be refreshed in this environment, so `#VALUE!` / `n/a` in the "
        "cap table's forward estimates and the comps / precedents formulas is the normal "
        "state of a shipped artefact. Re-flagging them is how a review gets ignored.",
    ),
    (
        "whether a figure is TRUE",
        "that is `deckcheck`, running in this same wave against the filings. Report that "
        "a number is unreadable; never that it is wrong.",
    ),
    (
        "spelling, wording, and brand formatting",
        "not this pass, and not automated at all yet.",
    ),
)


class DeckReadError(RuntimeError):
    """The pass was handed something it cannot read."""


# ─── Findings ────────────────────────────────────────────────────────────────


@dataclass(frozen=True)
class SlideFinding:
    """One thing a reader saw on one slide. Advisory by construction.

    ``severity`` exists only to be pinned. This is a review an analyst reads on a
    finished deliverable, not a gate, and the deck it is reviewing already
    converged against measured renders — so a finding that claimed to be blocking
    would be an opinion overruling a measurement. Constructing one raises.

    ``evidence`` is the render or crop the reader was looking at. It is what turns
    "slide 8's labels collide" into something the analyst can check in one click,
    and a finding without it is a claim about a picture nobody else can find.
    """

    slide: int
    kind: str
    issue: str
    detail: str
    shape: str | None = None
    evidence: Path | str | None = None
    severity: str = SEVERITY_ADVISORY

    def __post_init__(self) -> None:
        if self.kind not in READ_KINDS:
            raise ValueError(
                f"unknown finding kind {self.kind!r}; expected one of {', '.join(READ_KINDS)}"
            )
        if self.severity != SEVERITY_ADVISORY:
            raise ValueError(
                f"a deckread finding is always {SEVERITY_ADVISORY!r} — this pass reads a "
                f"deck that has already converged against measured renders, and no shipped "
                f"stage gates. Report what you saw; the fix belongs to the stage that "
                f"produced it."
            )

    @property
    def slide_number(self) -> int:
        """The 1-based number the analyst sees in PowerPoint."""
        return self.slide + 1


# ─── The mechanical half ─────────────────────────────────────────────────────


@dataclass(frozen=True)
class ReadEvidence:
    """The rendered artefact plus the work list, for one deck.

    Thin on purpose: everything here comes from `deck_contract.vision_pass`, which
    is the same agenda builder the `deck` stage's own review uses. What this adds is
    the deck it belongs to, the slide count (so "read every slide" is checkable),
    and :attr:`rendered` — the difference between a clean deck and a deck nobody
    looked at.
    """

    deck: Path
    vision: VisionPass
    slide_count: int

    @property
    def renders(self) -> dict[int, Path]:
        """slide index -> PNG. Empty when no renderer was available."""
        return self.vision.review_images

    @property
    def crops(self) -> list[tuple[int, str, Path]]:
        """(slide index, shape name, picture at native resolution)."""
        return self.vision.picture_crops

    @property
    def rendered(self) -> bool:
        """False when LibreOffice was absent — nothing was looked at."""
        return bool(self.vision.review_images)

    @property
    def flagged_slides(self) -> tuple[int, ...]:
        """Slide indices the work list singles out, ascending."""
        return tuple(sorted({t.slide for t in self.vision.targets}))

    @property
    def checklist(self) -> str:
        """What the reader is asked to look for — `deck_contract`'s own wording."""
        return VISION_CHECKLIST


def existing_renders(out_dir: Path | str) -> dict[int, Path]:
    """The slide PNGs a previous :func:`read_deck` already wrote under `out_dir`.

    The stage runs as two snippets — build the evidence, read it, then write the
    report — and those are two `python` invocations, so the second one is a fresh
    process with a cold render cache. Passing this back into `read_deck(renders=…)`
    rebuilds the same evidence without converting the deck a second time, which on
    a 19-slide deck over a cloud-synced deal directory is not a rounding error.

    `{}` when nothing was rendered, which `read_deck` treats exactly as it treats a
    host with no LibreOffice: the report says the visual half did not run.
    """
    slides = Path(out_dir) / "slides"
    if not slides.is_dir():
        return {}
    found: dict[int, Path] = {}
    for png in slides.glob("slide_*.png"):
        try:
            found[int(png.stem.split("_")[-1]) - 1] = png
        except ValueError:  # not one of ours
            continue
    return found


def read_deck(
    deck: Path | str,
    *,
    out_dir: Path | str,
    renders: dict[int, Path] | None = None,
) -> ReadEvidence:
    """Render `deck`, extract its pictures, and build the reader's work list.

    One call, so the stage cannot half-do it: the renders, the native-resolution
    crops and the agenda all come out of `deck_contract.vision_pass`, which is the
    function the `deck` stage's own review is built from. Re-run here on the
    FINISHED artefact — in the pitch plan the charts land after assembly, so the
    `deck` stage's crops are of a deck that no longer exists on disk.

    Pass `renders={}` to skip rendering (a test, or a host with no LibreOffice);
    :attr:`ReadEvidence.rendered` is then False and the report says so instead of
    reporting a clean deck.

    **Its own evidence directory, not `deckcheck`'s.** The two stages are wave-mates
    and run concurrently, so one shared directory would be a read-modify-write race
    between sub-agents — the same reason each stage writes its own `provenance.json`
    fragment. The cost is one extra set of slide PNGs under the run directory, which
    is the size `deckcheck`'s evidence already is, and it buys findings that cite a
    path the analyst can open. That is a different thing from `converge_deck`'s QA
    scratch (~170 files per pass, deleted on the way out): these renders are read by
    a person, so they are durable on purpose.
    """
    from pptx import Presentation

    deck = Path(deck)
    if not deck.is_file():
        raise DeckReadError(f"no deck to read at {deck}")
    vision = vision_pass(deck, out_dir=Path(out_dir), renders=renders)
    return ReadEvidence(
        deck=deck, vision=vision, slide_count=len(Presentation(deck).slides)
    )


# ─── Rendering ───────────────────────────────────────────────────────────────


def _not_yours_block() -> list[str]:
    return [
        "## Not yours — do not report these",
        "",
        *(f"- **{what}** — {who}" for what, who in NOT_YOURS),
        "",
    ]


def render_worklist(evidence: ReadEvidence, *, checklist_path: Path | str | None = None) -> str:
    """The reader's work list for the finished deck.

    The body is `stage_transforms.render_vision_review` — the `deck` stage's own
    review, re-rendered on the artefact the analyst receives. Deliberately not a
    second copy: the checklist has one wording and one place (`VISION_CHECKLIST`),
    and this pass exists because that agenda had no reader, not because it was
    written wrong.

    `checklist_path` is the `deck` stage's `vision_review_path`, named so the reader
    can see what the pre-chart deck was flagged for and what the charts changed.
    """
    from stage_transforms import render_vision_review  # lazy: no cycle, no cost

    head = [
        f"# Deck read — work list for `{evidence.deck.name}`",
        "",
        "This is the **final artefact**: the deck as the analyst receives it, charts "
        "included. Read every render below and answer each question with what you see.",
        "",
    ]
    if checklist_path:
        head += [
            f"The `deck` stage's own review of the pre-chart deck: `{checklist_path}`. "
            "Anything it flagged that the charts then covered is worth a second look.",
            "",
        ]
    if not evidence.rendered:
        head += [
            "**No renders — the visual half of this pass did not run.** Say so in the "
            "report; do not report a clean deck.",
            "",
        ]
    head += _not_yours_block()
    return "\n".join(head) + "\n" + render_vision_review(evidence.deck, evidence.vision)


def write_worklist(
    path: Path | str, evidence: ReadEvidence, *, checklist_path: Path | str | None = None
) -> Path:
    """Render the work list and write it. Returns the path."""
    target = Path(path)
    target.parent.mkdir(parents=True, exist_ok=True)
    target.write_text(
        render_worklist(evidence, checklist_path=checklist_path), encoding="utf-8"
    )
    return target


def _findings_table(findings) -> list[str]:
    lines = [
        "| Slide | Kind | What is wrong | What was seen | Shape | Evidence |",
        "|---|---|---|---|---|---|",
    ]
    for f in sorted(findings, key=lambda f: (f.slide, f.kind, f.issue)):
        lines.append(
            f"| {f.slide_number} | `{f.kind}` | **{f.issue}** | {f.detail} "
            f"| {f.shape or '—'} | {f.evidence or '—'} |"
        )
    return lines


def render_report(
    evidence: ReadEvidence,
    findings: "list[SlideFinding] | tuple[SlideFinding, ...]" = (),
    *,
    company: str,
    checklist_path: Path | str | None = None,
    notes: "list[str] | tuple[str, ...]" = (),
) -> str:
    """The analyst-facing read: what was looked at, then what was seen.

    "What was looked at" comes first because it is the claim the rest depends on. A
    pass with no renders behind it says that in its second line, above any finding
    and above any statement that nothing was found.
    """
    read = len(evidence.renders)
    lines = [
        f"# Deck read — {company}",
        "",
        "**Advisory review, not a gate.** Every item below is something to look at "
        "before this deck leaves the building. Nothing here halts a run, and the deck's "
        "geometry already converged against measured renders — a finding here is what a "
        "reader saw, not a measurement.",
        "",
    ]
    if not evidence.rendered:
        lines += [
            "> **The visual half did not run.** No renderer was available, so no slide of "
            "this deck was read. Nothing below is evidence that the deck is clean.",
            "",
        ]
    lines += [
        f"- Deck: `{evidence.deck}`",
        f"- Slides read: {read} of {evidence.slide_count}",
        f"- Slides on the work list: {len(evidence.flagged_slides)} "
        f"({', '.join(str(i + 1) for i in evidence.flagged_slides) or 'none'})",
        f"- Pictures read at native resolution: {len(evidence.crops)}",
        (
            f"- `deck` stage's pre-chart review: `{checklist_path}`"
            if checklist_path
            else "- `deck` stage's pre-chart review: (not supplied)"
        ),
        f"- Findings: {len(findings)}",
        "",
        "## Findings",
        "",
    ]
    if findings:
        lines += _findings_table(findings)
    elif evidence.rendered:
        lines.append(
            f"Every slide was read and none of {evidence.checklist} was found."
        )
    else:
        lines.append("(nothing was read, so nothing was found — see the warning above)")
    lines.append("")
    if notes:
        lines += ["## Notes", ""] + [f"- {n}" for n in notes] + [""]
    lines += _not_yours_block()
    lines += ["## What was read", ""]
    lines += [
        f"- slide {i + 1}: `{evidence.renders[i]}`" for i in sorted(evidence.renders)
    ] or ["- (no renders)"]
    lines.append("")
    if evidence.crops:
        lines += ["### Pictures at native resolution", ""]
        lines += [
            f"- slide {index + 1} `{name}`: `{path}`" for index, name, path in evidence.crops
        ]
        lines.append("")
    return "\n".join(lines)


def write_report(
    path: Path | str,
    evidence: ReadEvidence,
    findings: "list[SlideFinding] | tuple[SlideFinding, ...]" = (),
    *,
    company: str,
    checklist_path: Path | str | None = None,
    notes: "list[str] | tuple[str, ...]" = (),
) -> Path:
    """Render the report and write it. Returns the path."""
    target = Path(path)
    target.parent.mkdir(parents=True, exist_ok=True)
    target.write_text(
        render_report(
            evidence,
            findings,
            company=company,
            checklist_path=checklist_path,
            notes=notes,
        ),
        encoding="utf-8",
    )
    return target
