"""JSON Schema exporter.

Run as a module to regenerate the language-agnostic JSON Schemas under
`infor-beta/scripts/schemas/json/`:

    python -m schemas.export

(With `infor-beta/scripts` on `sys.path` — the test harness's
`pyproject.toml` adds it. From the repo root you can also run
`PYTHONPATH=infor-beta/scripts python -m schemas.export`.)

Idempotent — re-running produces identical output, byte for byte, given the
same schema definitions.
"""

from __future__ import annotations

import json
from pathlib import Path

from . import (
    Company,
    DealContext,
    Filing,
    Plan,
    SlidePlan,
    EarningsUpdateContent,
    PitchDeckContent,
)

# (filename_stem, model class) pairs.
_MODELS = [
    ("company", Company),
    ("filing", Filing),
    ("slide_plan", SlidePlan),
    ("earnings_update_content", EarningsUpdateContent),
    ("pitch_deck_content", PitchDeckContent),
    ("deal_context", DealContext),
    ("plan", Plan),
]


def _out_dir() -> Path:
    return Path(__file__).resolve().parent / "json"


def export_all(out_dir: Path | None = None) -> list[Path]:
    """Write one `<stem>.schema.json` per model. Returns the list of written paths."""
    target = out_dir if out_dir is not None else _out_dir()
    target.mkdir(parents=True, exist_ok=True)

    written: list[Path] = []
    for stem, model in _MODELS:
        schema = model.model_json_schema()
        # Sort keys + final newline → byte-identical output across runs.
        payload = json.dumps(schema, indent=2, sort_keys=True) + "\n"
        path = target / f"{stem}.schema.json"
        path.write_text(payload, encoding="utf-8")
        written.append(path)
    return written


if __name__ == "__main__":
    paths = export_all()
    for p in paths:
        print(p)
