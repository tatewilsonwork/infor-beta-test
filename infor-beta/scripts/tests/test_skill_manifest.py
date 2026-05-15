"""Unit tests for SkillManifest + its spec sub-models."""

import pytest
from pydantic import ValidationError

from schemas import (
    InputSpec,
    OutputSpec,
    SideEffectSpec,
    SkillManifest,
)


def _base_manifest(**overrides):
    kwargs = dict(
        name="comps-infor",
        description="Use this skill when the analyst asks for trading comps.",
        version="0.1.0",
        inputs=[InputSpec(name="tickers", type="list[str]")],
        outputs=[OutputSpec(name="comps_table", type="Path")],
        side_effects=[SideEffectSpec(kind="file_write", target="comps.xlsx")],
        allowed_tools=["Read", "Bash", "Write"],
    )
    kwargs.update(overrides)
    return SkillManifest(**kwargs)


def test_minimum_seven_fields():
    """All seven required fields present → manifest constructs cleanly (G8)."""
    m = _base_manifest()
    assert m.name == "comps-infor"
    assert m.version == "0.1.0"
    assert m.inputs[0].name == "tickers"
    assert m.outputs[0].type == "Path"
    assert m.side_effects[0].kind == "file_write"
    assert "Read" in m.allowed_tools


@pytest.mark.parametrize("missing_field", ["name", "description", "version"])
def test_missing_required_string_fields_rejected(missing_field):
    kwargs = dict(
        name="x",
        description="y",
        version="0.1.0",
    )
    kwargs.pop(missing_field)
    with pytest.raises(ValidationError):
        SkillManifest(**kwargs)  # type: ignore[arg-type]


@pytest.mark.parametrize(
    "good_version", ["0.0.1", "0.1.0", "1.0.0", "12.34.56"],
)
def test_version_semver_accepted(good_version):
    m = _base_manifest(version=good_version)
    assert m.version == good_version


@pytest.mark.parametrize(
    "bad_version",
    ["0.1", "0.1.0-rc1", "v0.1.0", "1", "0.1.0.0", "not-a-version", ""],
)
def test_version_non_semver_rejected(bad_version):
    with pytest.raises(ValidationError):
        _base_manifest(version=bad_version)


def test_list_fields_default_to_empty():
    m = SkillManifest(
        name="x",
        description="y",
        version="0.1.0",
    )
    assert m.inputs == []
    assert m.outputs == []
    assert m.side_effects == []
    assert m.allowed_tools == []


def test_input_spec_required_default_true():
    spec = InputSpec(name="x", type="str")
    assert spec.required is True


def test_extra_fields_rejected_on_manifest():
    with pytest.raises(ValidationError):
        SkillManifest(
            name="x",
            description="y",
            version="0.1.0",
            cost_class="cheap",  # type: ignore[call-arg]
        )
