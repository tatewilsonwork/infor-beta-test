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
      provenance is its components' records; this names the chain.
    - ``stage`` — the plan stage that recorded it. Stamped on merge from the
      fragment when a record does not carry its own.
    """

    figure: str
    sources: tuple[FigureSource, ...] = ()
    value: float | str | None = None
    units: str | None = None
    location: str | None = None
    derivation: str | None = None
    stage: str | None = None

    def __post_init__(self) -> None:
        object.__setattr__(self, "figure", _clean(self.figure) or "")
        object.__setattr__(self, "sources", tuple(self.sources))
        if not self.figure:
            raise ProvenanceError("a provenance record needs a figure name")
        if not self.sources and not self.derivation:
            raise ProvenanceError(
                f"figure {self.figure!r} has no source and no derivation. Every "
                f"extracted figure carries the filing it came from; a computed one "
                f"carries a derivation naming how it was built from figures that do."
            )

    @property
    def citation_lines(self) -> tuple[str, ...]:
        """The rendered citation text, one line per source — what a comment shows."""
        return tuple(source.render() for source in self.sources)

    def to_dict(self) -> dict:
        out = {
            "figure": self.figure,
            "value": self.value,
            "units": self.units,
            "location": self.location,
            "derivation": self.derivation,
            "stage": self.stage,
            "sources": [s.to_dict() for s in self.sources],
        }
        return {k: v for k, v in out.items() if v is not None and v != []}

    @classmethod
    def from_dict(cls, data: Mapping) -> "FigureProvenance":
        return cls(
            figure=data.get("figure", ""),
            sources=tuple(FigureSource.from_dict(s) for s in data.get("sources", ())),
            value=data.get("value"),
            units=data.get("units"),
            location=data.get("location"),
            derivation=data.get("derivation"),
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
                    stage=other.stage,
                )
            self.figures.append(entry)

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
