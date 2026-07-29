"""Structured provenance for every cited figure — the Phase G record.

v0.5.31 / v0.5.34 gave every headline figure an in-artefact citation, and the
mechanism *was* the record: a skill composed the sentence
``"FY2025 10-K, Consolidated Statements of Operations"`` and
`comment_citations` glued ``"Source: "`` on the front. Nothing else ever knew
what a figure's source was, because the only copy of it was prose inside an
openpyxl `Comment`, on a cell, inside one tab, inside the deal workbook.

That is enough for an analyst reading the workbook and not enough for anything
else. Three consequences, all of which this module removes:

- **A figure on the deck cannot be traced.** The deck shows ``$1,234.5MM``; the
  citation is a comment on a cell in a different file. Nothing can join them, so
  nothing can ask "which filing says that?" — which is the entire question the
  `deckcheck` stage exists to ask.
- **The fields are not fields.** ``filing``, ``statement`` and ``page`` were a
  comma-joined string, so "cite the page" was a convention a skill either
  followed or didn't, unenforceably.
- **There were two wordings for one thing.** `source_line` rendered the URL form
  and every skill hand-wrote the filing form, which is the same drift pair H1
  collapsed on the intake side.

So the record is a record — :class:`FigureSource` (filing → statement → page, or
URL → retrieval date) and :class:`FigureProvenance` (the figure, its value, where
it was written, and its sources) — and **the cell comment is RENDERED FROM it**
(:meth:`FigureSource.render`, wired into `comment_citations`). One direction of
travel: a record exists first, the comment is a view of it. A citation that is
not a record cannot be written any more, because the appender takes a record.

Per run, per stage
------------------
Each stage writes its own fragment at ``<run_dir>/stages/<id>/provenance.json``,
beside the ``outputs.json`` it already writes. The consolidated per-run record
is the merge of every fragment, written to ``<run_dir>/provenance.json`` by the
`deckcheck` stage (:func:`write_run_provenance`).

A fragment per stage rather than one shared file is not bookkeeping taste: two
stages in the same wave run **concurrently**, and a single shared
``provenance.json`` would be a read-modify-write race between sub-agents. Each
stage owning its own file makes that impossible by construction, and needs
nothing from the conductor — the stage directory already exists.

Two things a record must be able to say
---------------------------------------
The first cut of this module got the *sources* right and left two claims as
prose, which is the same failure one level down — a field nothing can follow is
a field nothing enforces.

- **What a derived figure was derived FROM.** ``derivation`` is a sentence for a
  reader ("FY2025 + Q3 2026 YTD − Q3 2025 YTD"); :attr:`FigureProvenance.derived_from`
  is the same claim as :class:`FigureRef` pointers a machine can resolve
  (:meth:`ProvenanceLedger.trace`). Without it "derived from an upstream record"
  and "unsourced" read identically to `deckcheck` — and to the reviewer, who had
  no way to check that the components named in the prose have records at all.
- **Where the figure LANDS.** :class:`DeckPlacement` names the slide and the
  typed content field (or shape) a figure was written to, so `deckcheck` can join
  a number on a slide to its record **by identity**. Joining on value alone is a
  coincidence generator: ARR of 4,190.5 on the executive summary matched the
  FY2024 gross profit record of 4,191.0 on a real run, and the report printed it
  as traced. A value agreement with no placement behind it is still worth
  showing, but it is a *lead*, not provenance, and it has to be labelled as one.
"""

from __future__ import annotations

import datetime as _dt
import json
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Iterable, Mapping, Sequence

#: Fragment name inside a stage directory, and the consolidated name in the run
#: directory. Deliberately the same word in both places — the second is the merge
#: of the first, not a different kind of file.
PROVENANCE_FILENAME = "provenance.json"


class ProvenanceError(ValueError):
    """A provenance record is incomplete — a half-citation, or a figure with none."""


def _iso(value: "str | _dt.date | None") -> str | None:
    if value is None:
        return None
    if isinstance(value, _dt.date):
        return value.isoformat()
    text = str(value).strip()
    return text or None


def _clean(value: Any) -> str | None:
    if value is None:
        return None
    text = str(value).strip()
    return text or None


@dataclass(frozen=True)
class FigureSource:
    """Where one figure came from: a document location, or a web page.

    - ``filing`` — the document, as the analyst would name it: ``"FY2025 10-K"``,
      ``"Q3 2026 10-Q"``, ``"Bloomberg ownership export"``, or an attached
      filename.
    - ``statement`` — the statement, note or section inside it:
      ``"Consolidated Statements of Operations"``, ``"Note 12: Long-Term Debt"``,
      ``"revenue disaggregation note"``.
    - ``page`` — the page it is on. ``87`` or ``"F-12"``; rendered ``"p. 87"``.
    - ``url`` / ``retrieved`` — the web form, for a value typed in from a market
      data page (the cap table's FX rate and share price). A URL without a
      retrieval date is not a citation, so ``retrieved`` is required with it.

    Validated on construction, because a half-citation is worse than none: it
    reads as provenance and cannot be followed. ``statement`` or ``page`` without
    the ``filing`` they are inside is rejected, and a record naming neither a
    filing nor a URL is rejected.
    """

    filing: str | None = None
    statement: str | None = None
    page: int | str | None = None
    url: str | None = None
    retrieved: str | None = None

    def __post_init__(self) -> None:
        # Normalise before validating so " " counts as absent.
        for name in ("filing", "statement", "url"):
            object.__setattr__(self, name, _clean(getattr(self, name)))
        object.__setattr__(self, "page", _clean(self.page))
        object.__setattr__(self, "retrieved", _iso(self.retrieved))

        if not self.filing and not self.url:
            raise ProvenanceError(
                "a figure source must name a filing (document) or a url; got "
                f"{self!r}. Pass filing=… for a document figure, or url=… + "
                f"retrieved=… for a web-sourced one."
            )
        if (self.statement or self.page) and not self.filing:
            raise ProvenanceError(
                f"statement/page without a filing is a half-citation: {self!r}. "
                f"Name the document the statement or page is in."
            )
        if self.url and not self.retrieved:
            raise ProvenanceError(
                f"url without a retrieval date is not a citation: {self.url!r}. "
                f"Pass retrieved=<YYYY-MM-DD> (the date you read the page)."
            )

    def render(self) -> str:
        """The citation text — the ONE place a citation's wording is decided.

        Both established forms fall out of one rule (segments joined by ``" — "``,
        the document chain comma-joined inside its own segment), so neither is a
        second wording to keep in step:

            FY2025 10-K, Consolidated Statements of Operations, p. 87
            https://example.com/fx — retrieved 2026-07-15
        """
        chain = ", ".join(
            part
            for part in (self.filing, self.statement, f"p. {self.page}" if self.page else None)
            if part
        )
        tail = f"retrieved {self.retrieved}" if self.retrieved else None
        return " — ".join(part for part in (chain, self.url, tail) if part)

    def to_dict(self) -> dict:
        """JSON-ready dict, omitting the fields this source does not carry."""
        out = {
            "filing": self.filing,
            "statement": self.statement,
            "page": self.page,
            "url": self.url,
            "retrieved": self.retrieved,
        }
        return {k: v for k, v in out.items() if v is not None}

    @classmethod
    def from_dict(cls, data: Mapping) -> "FigureSource":
        return cls(
            filing=data.get("filing"),
            statement=data.get("statement"),
            page=data.get("page"),
            url=data.get("url"),
            retrieved=data.get("retrieved"),
        )


def _location_key(location: Any) -> str | None:
    """A comparable key for a ``<tab>!<cell>`` location.

    Absolute markers and sheet quoting are notation, not identity:
    ``captable!F17``, ``captable!$F$17`` and ``'captable'!f17`` are one cell. A
    location with no ``!`` (nothing writes one today) is compared whole.
    """
    text = _clean(location)
    if text is None:
        return None
    text = text.replace("$", "")
    if "!" in text:
        sheet, _, cell = text.rpartition("!")
        return f"{sheet.strip().strip(chr(39)).casefold()}!{cell.strip().casefold()}"
    return text.casefold()


def _figure_key(figure: Any) -> str | None:
    """A comparable key for a figure name — case- and space-insensitive.

    The leniency is deliberate and bounded: a ref is authored by the same stage
    that authored the figure name it points at (or, across stages, by a skill
    reading the other stage's documented label), so the only realistic mismatch
    is capitalisation. Being strict there would report "unresolvable" for a
    working chain, which is the reading this module exists to prevent.
    """
    text = _clean(figure)
    return None if text is None else " ".join(text.split()).casefold()


@dataclass(frozen=True)
class FigureRef:
    """A machine-followable pointer from a derived figure to a component's record.

    One of two keys, and they are not interchangeable:

    - ``location`` — ``"<tab>!<cell>"``, the same string a record puts in
      :attr:`FigureProvenance.location`. The canonical identity for anything in
      the deal workbook: the cap table's Enterprise Value formula reads ``F22``,
      ``F28``, ``F29``, ``F30``, and those cells are what its components *are*,
      whatever a skill chose to call them.
    - ``figure`` — the component's figure name, for a link that crosses stages and
      workbook tabs (the `financial-summary` LTM cell pointing at the
      `ltm-metrics` bridge total it links to). ``stage`` narrows it when the
      caller knows which stage recorded it; left unset the ref resolves against
      whichever stage did, which is what makes a fragment written before its
      component's fragment still resolve **in the merge**.

    When both are given, ``location`` decides and ``figure`` is a human label.
    """

    figure: str | None = None
    stage: str | None = None
    location: str | None = None

    def __post_init__(self) -> None:
        for name in ("figure", "stage", "location"):
            object.__setattr__(self, name, _clean(getattr(self, name)))
        if not self.figure and not self.location:
            raise ProvenanceError(
                "a component reference needs a figure name or a location; got "
                f"{self!r}. Pass location='<tab>!<cell>' for a workbook cell, or "
                f"figure='<figure name>' (optionally with stage=…) across stages."
            )

    def matches(self, record: "FigureProvenance") -> bool:
        """Whether `record` is the figure this ref points at."""
        if self.stage and record.stage and self.stage.casefold() != record.stage.casefold():
            return False
        if self.location:
            return _location_key(self.location) == _location_key(record.location)
        return _figure_key(self.figure) == _figure_key(record.figure)

    def render(self) -> str:
        """The ref as one short string — what a report prints."""
        if self.location:
            return self.location
        return f"{self.stage}.{self.figure}" if self.stage else str(self.figure)

    def to_dict(self) -> dict:
        out = {"figure": self.figure, "stage": self.stage, "location": self.location}
        return {k: v for k, v in out.items() if v is not None}

    @classmethod
    def from_dict(cls, data: "Mapping | str") -> "FigureRef":
        """Read a ref from JSON. A bare string is a figure name — see `_as_refs`."""
        if isinstance(data, str):
            return cls(figure=data)
        return cls(
            figure=data.get("figure"),
            stage=data.get("stage"),
            location=data.get("location"),
        )


@dataclass(frozen=True)
class DeckPlacement:
    """Where a figure LANDS on the deck — the identity `deckcheck` joins on.

    - ``slide`` — the 1-based number PowerPoint shows, which for both deck plans
      is the slide plan's ``order + 1`` (the assemblers emit one deck slide per
      plan entry, in order). Resolve it from the plan the stage was handed, never
      by counting slides in your head.
    - ``field`` — for a typed content bundle, the field the figure sits in:
      ``"executive_summary_bullets[1]"``, ``"market_entry_targets[3].cells[10]"``.
      The bundle *is* what the assembler writes, so the field is the most precise
      thing a content stage can say — and it survives a library re-layout, which
      a shape name does not.
    - ``shape`` — the shape name, for a recorder that genuinely knows it.

    At least one must be given: a placement that names nowhere is not a placement.
    """

    slide: int | None = None
    field: str | None = None
    shape: str | None = None

    def __post_init__(self) -> None:
        for name in ("field", "shape"):
            object.__setattr__(self, name, _clean(getattr(self, name)))
        if self.slide is not None:
            try:
                slide = int(self.slide)
            except (TypeError, ValueError):
                raise ProvenanceError(
                    f"a placement's slide must be the 1-based slide number, got {self.slide!r}"
                ) from None
            if slide < 1:
                raise ProvenanceError(
                    f"a placement's slide is 1-based, as PowerPoint numbers it; got {slide}"
                )
            object.__setattr__(self, "slide", slide)
        if self.slide is None and not self.field and not self.shape:
            raise ProvenanceError(
                "a deck placement must name a slide, a content field or a shape — "
                "an empty placement claims an identity it does not have."
            )

    def render(self) -> str:
        parts = [f"slide {self.slide}" if self.slide is not None else None, self.field, self.shape]
        return " · ".join(p for p in parts if p)

    def to_dict(self) -> dict:
        out = {"slide": self.slide, "field": self.field, "shape": self.shape}
        return {k: v for k, v in out.items() if v is not None}

    @classmethod
    def from_dict(cls, data: Mapping) -> "DeckPlacement":
        return cls(slide=data.get("slide"), field=data.get("field"), shape=data.get("shape"))


def _as_refs(refs: Any) -> tuple[FigureRef, ...]:
    """Coerce ``derived_from`` into refs, accepting a bare figure name.

    A string is allowed here, unlike a source: :class:`FigureSource` collapsed
    three fields into one sentence when handed a string, which is why that raises.
    A ref has one identifying field, so ``"LTM Revenue"`` loses nothing —
    it is the figure name, resolved against whichever stage recorded it.
    """
    if refs is None:
        return ()
    if isinstance(refs, (FigureRef, str, Mapping)):
        refs = [refs]
    out: list[FigureRef] = []
    for item in refs:
        if isinstance(item, FigureRef):
            out.append(item)
        elif isinstance(item, str):
            out.append(FigureRef(figure=item))
        elif isinstance(item, Mapping):
            out.append(FigureRef.from_dict(item))
        else:
            raise ProvenanceError(
                f"expected FigureRef, a figure name or a mapping in derived_from, got "
                f"{type(item).__name__} ({item!r})."
            )
    return tuple(out)


def _as_placement(placement: Any) -> "DeckPlacement | None":
    if placement is None or isinstance(placement, DeckPlacement):
        return placement
    if isinstance(placement, Mapping):
        return DeckPlacement.from_dict(placement)
    raise ProvenanceError(
        f"expected a DeckPlacement, got {type(placement).__name__} ({placement!r}). "
        f"Build DeckPlacement(slide=…, field=…) — the slide is 1-based."
    )


@dataclass(frozen=True)
class FigureProvenance:
    """One figure, its value, where it was written, and where it came from.

    - ``figure`` — what the figure IS, in the words the artefact uses:
      ``"Revenue FY2023"``, ``"LTM Revenue"``, ``"Share price"``. The join key
      `deckcheck` reads, so it should name the figure a reader would look for.
    - ``value`` — the number written (or the Excel formula, for a combined
      figure). `deckcheck` matches deck text against this.
    - ``units`` — ``"US$MM"``, ``"%"``, ``"shares"``. Scale matters: the deck may
      render the same figure as ``$1.2B``.
    - ``location`` — where it landed, as ``"<tab>!<cell>"``. A cell reference is
      enough because the deal owns one workbook.
    - ``sources`` — one or more :class:`FigureSource`. Required *unless*
      ``derivation`` explains that this figure was computed from others.
    - ``derivation`` — how a computed figure was derived, in the artefact's own
      terms (``"FY2025 + Q3 2026 YTD − Q3 2025 YTD"``). A derived figure's real
      provenance is its components' records; this names the chain **for a reader**.
    - ``derived_from`` — the same chain as :class:`FigureRef` pointers, for a
      machine. Pass both: the prose says what the arithmetic *is*, the refs are
      what `deckcheck` follows to the filings underneath
      (:meth:`ProvenanceLedger.trace`).
    - ``placement`` — the :class:`DeckPlacement` this figure lands at, when the
      recording stage knows. The join key for a figure on a slide; without it a
      deck figure can only be value-matched, and a value match is a lead.
    - ``stage`` — the plan stage that recorded it. Stamped on merge from the
      fragment when a record does not carry its own.
    """

    figure: str
    sources: tuple[FigureSource, ...] = ()
    value: float | str | None = None
    units: str | None = None
    location: str | None = None
    derivation: str | None = None
    derived_from: tuple[FigureRef, ...] = ()
    placement: "DeckPlacement | None" = None
    stage: str | None = None

    def __post_init__(self) -> None:
        object.__setattr__(self, "figure", _clean(self.figure) or "")
        object.__setattr__(self, "sources", tuple(self.sources))
        object.__setattr__(self, "derived_from", _as_refs(self.derived_from))
        object.__setattr__(self, "placement", _as_placement(self.placement))
        if not self.figure:
            raise ProvenanceError("a provenance record needs a figure name")
        if not self.sources and not self.derivation and not self.derived_from:
            raise ProvenanceError(
                f"figure {self.figure!r} has no source and no derivation. Every "
                f"extracted figure carries the filing it came from; a computed one "
                f"carries a derivation naming how it was built from figures that do "
                f"(prose in `derivation`, and the components in `derived_from` so the "
                f"chain can be followed)."
            )

    @property
    def citation_lines(self) -> tuple[str, ...]:
        """The rendered citation text, one line per source — what a comment shows."""
        return tuple(source.render() for source in self.sources)

    @property
    def derived(self) -> bool:
        """Whether this figure was computed from others rather than read off a page."""
        return bool(self.derivation or self.derived_from)

    @property
    def derivation_line(self) -> str | None:
        """The derivation as one line — the prose, else the refs it points at.

        A derived record with refs and no prose is legal (the refs are the
        stronger claim), so the artefact's ``Derived: …`` comment and every report
        row render from here rather than from ``derivation`` directly.
        """
        if self.derivation:
            return self.derivation
        if self.derived_from:
            return "derived from " + ", ".join(ref.render() for ref in self.derived_from)
        return None

    def to_dict(self) -> dict:
        out = {
            "figure": self.figure,
            "value": self.value,
            "units": self.units,
            "location": self.location,
            "derivation": self.derivation,
            "derived_from": [r.to_dict() for r in self.derived_from],
            "placement": None if self.placement is None else self.placement.to_dict(),
            "stage": self.stage,
            "sources": [s.to_dict() for s in self.sources],
        }
        return {k: v for k, v in out.items() if v is not None and v != []}

    @classmethod
    def from_dict(cls, data: Mapping) -> "FigureProvenance":
        placement = data.get("placement")
        return cls(
            figure=data.get("figure", ""),
            sources=tuple(FigureSource.from_dict(s) for s in data.get("sources", ())),
            value=data.get("value"),
            units=data.get("units"),
            location=data.get("location"),
            derivation=data.get("derivation"),
            derived_from=tuple(FigureRef.from_dict(r) for r in data.get("derived_from", ())),
            placement=None if placement is None else DeckPlacement.from_dict(placement),
            stage=data.get("stage"),
        )


def _as_sources(sources: "FigureSource | Sequence[FigureSource] | None") -> tuple[FigureSource, ...]:
    if sources is None:
        return ()
    if isinstance(sources, FigureSource):
        return (sources,)
    for item in sources:
        if not isinstance(item, FigureSource):
            raise ProvenanceError(
                f"expected FigureSource instances, got {type(item).__name__} "
                f"({item!r}). A citation string is no longer a source — build the "
                f"record: FigureSource(filing=…, statement=…, page=…)."
            )
    return tuple(sources)


@dataclass(frozen=True)
class DerivationTrace:
    """What a derived figure's structured refs resolved to — and what they didn't.

    ``components`` are the direct refs that found a record; ``unresolved`` are the
    refs that found nothing, anywhere in the walk. An unresolved ref is the
    interesting case, and it is reported rather than dropped: it means a stage
    claimed a figure was built from something that has no record — a gap that used
    to be indistinguishable from a healthy derivation, because prose cannot be
    checked.

    ``root_sources`` is the transitive answer: every :class:`FigureSource` reachable
    by following the refs down. That is what makes a derived figure on a slide
    traceable to a filing page instead of to another sentence.
    """

    figure: FigureProvenance
    components: tuple[FigureProvenance, ...] = ()
    unresolved: tuple[FigureRef, ...] = ()
    root_sources: tuple[FigureSource, ...] = ()

    @property
    def structured(self) -> bool:
        """Whether the record declared machine-followable components at all."""
        return bool(self.figure.derived_from)

    @property
    def resolved(self) -> bool:
        """Structured, and every ref in the walk found its record."""
        return self.structured and not self.unresolved

    def render(self) -> str:
        """One line for a report row: what it resolved to, and what it did not."""
        if not self.structured:
            return self.figure.derivation_line or "—"
        parts = []
        if self.components:
            parts.append("← " + ", ".join(c.figure for c in self.components))
        if self.root_sources:
            parts.append("; ".join(dict.fromkeys(s.render() for s in self.root_sources)))
        if self.unresolved:
            parts.append(
                "UNRESOLVABLE: " + ", ".join(dict.fromkeys(r.render() for r in self.unresolved))
            )
        return " — ".join(parts) or "—"


@dataclass
class ProvenanceLedger:
    """Every figure one stage recorded. Written as that stage's fragment.

    Mutable by design: a builder fills it as it writes cells, in the order it
    writes them, so the ledger reads like the tab does.
    """

    stage: str | None = None
    figures: list[FigureProvenance] = field(default_factory=list)

    def record(
        self,
        figure: str,
        *,
        sources: "FigureSource | Sequence[FigureSource] | None" = None,
        value: float | str | None = None,
        units: str | None = None,
        location: str | None = None,
        derivation: str | None = None,
        derived_from: Any = None,
        placement: "DeckPlacement | Mapping | None" = None,
    ) -> FigureProvenance:
        """Append one figure's provenance and return it.

        Raises :class:`ProvenanceError` when the figure has neither a source nor a
        derivation — the "sources are REQUIRED" rule, made structural rather than
        left to a SKILL.md sentence.
        """
        entry = FigureProvenance(
            figure=figure,
            sources=_as_sources(sources),
            value=value,
            units=units,
            location=location,
            derivation=derivation,
            derived_from=_as_refs(derived_from),
            placement=_as_placement(placement),
            stage=self.stage,
        )
        self.figures.append(entry)
        return entry

    def extend(self, other: "ProvenanceLedger") -> None:
        """Absorb another ledger's figures, stamping its stage on any that lack one."""
        for entry in other.figures:
            if entry.stage is None and other.stage is not None:
                entry = FigureProvenance(
                    figure=entry.figure,
                    sources=entry.sources,
                    value=entry.value,
                    units=entry.units,
                    location=entry.location,
                    derivation=entry.derivation,
                    derived_from=entry.derived_from,
                    placement=entry.placement,
                    stage=other.stage,
                )
            self.figures.append(entry)

    def find(self, ref: "FigureRef | str") -> "FigureProvenance | None":
        """The record a ref points at, or None.

        A stage-less ref prefers a record from the ref's own stage when the ledger
        holds several matches — in the merged run record two stages can both hold a
        figure called "Revenue FY2025", and the nearer one is the intended one.
        """
        if isinstance(ref, str):
            ref = FigureRef(figure=ref)
        matches = [entry for entry in self.figures if ref.matches(entry)]
        if not matches:
            return None
        if ref.stage is None and self.stage is not None:
            for entry in matches:
                if entry.stage == self.stage:
                    return entry
        return matches[0]

    def trace(self, figure: FigureProvenance) -> DerivationTrace:
        """Follow a record's structured refs to its components and their sources.

        Depth-first over ``derived_from``, cycle-safe (a workbook can be
        self-referential by accident, and a traversal that hangs would take the
        whole review with it). Direct refs that resolve become ``components``;
        every ref that resolves to nothing — at any depth — is reported.
        """
        components: list[FigureProvenance] = []
        unresolved: list[FigureRef] = []
        root_sources: list[FigureSource] = []
        seen: set[tuple[str | None, str | None, str | None]] = set()

        def visit(record: FigureProvenance, *, root: bool) -> None:
            key = (record.stage, _figure_key(record.figure), _location_key(record.location))
            if key in seen:
                return
            seen.add(key)
            # The root's own sources are its own, not something it was derived from.
            if not root:
                root_sources.extend(record.sources)
            for ref in record.derived_from:
                target = self.find(ref)
                if target is None:
                    unresolved.append(ref)
                    continue
                if root:
                    components.append(target)
                visit(target, root=False)

        visit(figure, root=True)
        return DerivationTrace(
            figure=figure,
            components=tuple(components),
            unresolved=tuple(unresolved),
            root_sources=tuple(root_sources),
        )

    def __len__(self) -> int:
        return len(self.figures)

    @property
    def stages(self) -> tuple[str, ...]:
        """Every stage represented, in first-seen order."""
        return tuple(dict.fromkeys(f.stage for f in self.figures if f.stage))

    def to_dict(self) -> dict:
        out: dict[str, Any] = {"figures": [f.to_dict() for f in self.figures]}
        if self.stage:
            out["stage"] = self.stage
        return out

    @classmethod
    def from_dict(cls, data: Mapping) -> "ProvenanceLedger":
        return cls(
            stage=data.get("stage"),
            figures=[FigureProvenance.from_dict(f) for f in data.get("figures", ())],
        )

    def to_json(self) -> str:
        # `ensure_ascii=False`: the record is read by people as well as machines,
        # and a bridge derivation carries the typographic minus (U+2212) the
        # workbook uses. `−` in the file would be legible to neither.
        return (
            json.dumps(self.to_dict(), indent=2, sort_keys=True, default=str, ensure_ascii=False)
            + "\n"
        )

    def write(self, directory: Path | str) -> Path:
        """Write this ledger as `<directory>/provenance.json`. Returns the path.

        Called by a dispatched stage with `io.stage_dir`, so the fragment lands
        beside that stage's `inputs.json` / `outputs.json` and no two stages ever
        write the same file.
        """
        path = Path(directory) / PROVENANCE_FILENAME
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(self.to_json(), encoding="utf-8")
        return path

    @classmethod
    def read(cls, path: Path | str) -> "ProvenanceLedger":
        """Read a fragment (or a consolidated record) from a file or its directory."""
        p = Path(path)
        if p.is_dir():
            p = p / PROVENANCE_FILENAME
        return cls.from_dict(json.loads(p.read_text(encoding="utf-8")))


def stage_provenance_path(stage_dir: Path | str) -> Path:
    """`<stage_dir>/provenance.json` — a stage's own fragment, written or not."""
    return Path(stage_dir) / PROVENANCE_FILENAME


def iter_stage_fragments(run_dir: Path | str) -> Iterable[Path]:
    """Every `stages/*/provenance.json` under a run directory, in stage-id order.

    Missing fragments are simply absent: a stage that extracts no figures (the
    wireframe, the assembler) has nothing to record, and that is not an error.
    """
    stages = Path(run_dir) / "stages"
    if not stages.is_dir():
        return []
    return sorted(
        (d / PROVENANCE_FILENAME for d in stages.iterdir() if d.is_dir()),
        key=lambda p: p.parent.name,
    )


def read_run_provenance(run_dir: Path | str) -> ProvenanceLedger:
    """Merge every stage fragment in `run_dir` into one ledger.

    A fragment whose stage key is unset is stamped with its stage directory's
    name, so a record always says which stage produced it even if the stage
    forgot to label its ledger.
    """
    merged = ProvenanceLedger()
    for fragment in iter_stage_fragments(run_dir):
        if not fragment.is_file():
            continue
        ledger = ProvenanceLedger.read(fragment)
        if ledger.stage is None:
            ledger.stage = fragment.parent.name
        merged.extend(ledger)
    return merged


def write_run_provenance(run_dir: Path | str) -> Path:
    """Consolidate the fragments into `<run_dir>/provenance.json` and return it.

    The per-run record Phase G asks for. Written by the `deckcheck` stage, which
    is the first thing in the run that needs the whole picture; it is a pure
    merge, so re-running it is idempotent.
    """
    path = Path(run_dir) / PROVENANCE_FILENAME
    path.write_text(read_run_provenance(run_dir).to_json(), encoding="utf-8")
    return path
