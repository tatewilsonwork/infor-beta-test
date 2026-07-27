"""Tests for the cap-table Excel→PowerPoint renderer's LibreOffice recalc path,
plus the shared ``find_soffice`` locator every LibreOffice caller resolves through.

The Windows COM path forces ``excel.CalculateFull()``; the non-Windows
(Cowork/Linux) path must force LibreOffice to recalculate the openpyxl-authored,
manual-calc workbook on load, or every formula cell prints as 0/blank.
"""

import io
import re
import sys
from pathlib import Path

import pytest
from openpyxl import Workbook, load_workbook

import excel_to_powerpoint as e2p
from excel_to_powerpoint import _soffice_convert, _write_lo_recalc_profile, find_soffice
from tests.fake_com import FakeExcelApp, install_fake_com

_SCRIPTS_DIR = Path(__file__).resolve().parents[1]


def test_soffice_convert_timeout_raises_runtime_error(tmp_path: Path, monkeypatch):
    """v0.5.21: a wedged LibreOffice (TimeoutExpired) must surface as RuntimeError
    like every other soffice failure, so callers' graceful-degradation nets
    (``except RuntimeError``) engage instead of the stage aborting raw."""
    import subprocess

    import excel_to_powerpoint as e2p

    def _hang(*_args, **_kwargs):
        raise subprocess.TimeoutExpired(cmd="soffice", timeout=180)

    monkeypatch.setattr(e2p.subprocess, "run", _hang)
    with pytest.raises(RuntimeError, match="timed out"):
        _soffice_convert("soffice", tmp_path / "x.xlsx", "pdf", tmp_path)


# ─── The shared LibreOffice locator (v0.5.35 flip / v0.5.36 consolidation) ────


def test_find_soffice_returns_path_hit_before_install_locations(tmp_path: Path, monkeypatch):
    """PATH first — that is how Cowork / Linux prod resolves the binary."""
    monkeypatch.setattr("shutil.which", lambda name: "/usr/bin/soffice" if name == "soffice" else None)
    monkeypatch.setattr(e2p, "_SOFFICE_FALLBACK_PATHS", (str(tmp_path),))

    assert find_soffice() == "/usr/bin/soffice"


def test_find_soffice_falls_back_to_standard_install_location(tmp_path: Path, monkeypatch):
    """The Windows MSI puts nothing on PATH, so PATH-miss must not mean absent."""
    monkeypatch.setattr("shutil.which", lambda *_a, **_k: None)
    installed = tmp_path / "soffice.exe"
    installed.write_bytes(b"")
    monkeypatch.setattr(
        e2p, "_SOFFICE_FALLBACK_PATHS", (str(tmp_path / "not-here.exe"), str(installed))
    )

    assert find_soffice() == str(installed)


def test_find_soffice_is_none_when_libreoffice_absent(monkeypatch):
    monkeypatch.setattr("shutil.which", lambda *_a, **_k: None)
    monkeypatch.setattr(e2p, "_SOFFICE_FALLBACK_PATHS", ())

    assert find_soffice() is None


def test_no_bare_libreoffice_path_lookups_outside_the_locator():
    """Every LibreOffice caller resolves through ``find_soffice``.

    v0.5.35 made LibreOffice the default renderer on every platform but wired the
    locator into ``slide_render`` only; five other call sites still probed PATH
    directly, which the Windows MSI never satisfies — so they failed or silently
    degraded on the very dev box the flip existed to keep honest. This locks the
    consolidation: the PATH probe lives in exactly one module.
    """
    probe = re.compile(r"""which\(\s*["'](soffice|libreoffice)""")
    offenders = [
        f"{py.relative_to(_SCRIPTS_DIR).as_posix()}:{n}"
        for py in sorted(_SCRIPTS_DIR.rglob("*.py"))
        if py.name != "excel_to_powerpoint.py"  # the locator's own module
        for n, line in enumerate(py.read_text(encoding="utf-8").splitlines(), 1)
        if probe.search(line)
    ]

    assert not offenders, (
        "resolve LibreOffice through excel_to_powerpoint.find_soffice() rather than a "
        f"bare PATH lookup (it also probes the standard install locations): {offenders}"
    )


def test_write_lo_recalc_profile_forces_always_recalc(tmp_path: Path):
    """The throwaway profile must set OOXML recalc-on-load to 0 (Always)."""
    uri = _write_lo_recalc_profile(tmp_path)

    xcu = tmp_path / "lo_profile" / "user" / "registrymodifications.xcu"
    assert xcu.exists(), "registrymodifications.xcu must live under <profile>/user/"
    text = xcu.read_text(encoding="utf-8")
    assert "OOXMLRecalcMode" in text
    # 0 = Always; 1 = Never and 2 = Prompt would both skip the recalc headless.
    assert "<value>0</value>" in text
    assert uri.startswith("file:") and uri.endswith("lo_profile")


def _build_cacheless_cap_table(path: Path) -> None:
    """A manual-calc cap table with formula cells and NO cached values, plus a
    CapIQ ``_xll.*`` cell LibreOffice cannot resolve."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Cap with Links"
    ws["F16"] = 30.0  # share price (hardcoded)
    ws["F17"] = 100.0  # basic shares (hardcoded)
    ws["F18"] = "=F16*F17"  # basic market cap -> 3000
    ws["F28"] = 500.0  # net debt (hardcoded)
    ws["F31"] = "=F18+F28"  # Enterprise Value -> 3500
    ws["D47"] = 1200.0  # LTM revenue, hardcoded by ltm-metrics
    ws["D48"] = 300.0  # LTM Adj. EBITDA, hardcoded
    ws["D34"] = '=IFERROR(D47,"n/a ")'  # -> 1200
    ws["D35"] = '=IFERROR(D48,"n/a ")'  # -> 300
    # A CapIQ add-in cell: unknown to LibreOffice -> #NAME?, must degrade not crash.
    ws["E47"] = '=_xll.SNL.Clients.Office.Excel.Functions.SPG("X","IQ_REV")'
    ws["E34"] = '=IFERROR(E47,"n/a ")'  # -> "n/a "
    wb.calculation.calcMode = "manual"
    wb.calculation.fullCalcOnLoad = True
    wb.save(path)


@pytest.mark.skipif(sys.platform == "win32", reason="Windows uses the Excel COM path")
def test_libreoffice_recalc_populates_inworkbook_formulas(tmp_path: Path):
    soffice = find_soffice()
    if soffice is None:
        pytest.skip("LibreOffice (soffice/libreoffice) not installed")

    src = tmp_path / "cap_cacheless.xlsx"
    _build_cacheless_cap_table(src)
    # Precondition: openpyxl wrote no cached value, so without recalc this is blank.
    assert load_workbook(src, data_only=True)["Cap with Links"]["F31"].value is None

    out_dir = tmp_path / "out"
    out_dir.mkdir()
    _soffice_convert(soffice, src, "xlsx:Calc MS Excel 2007 XML", out_dir)

    ws = load_workbook(out_dir / "cap_cacheless.xlsx", data_only=True)["Cap with Links"]
    # In-workbook arithmetic now computes through Enterprise Value.
    assert ws["F18"].value == pytest.approx(3000.0)
    assert ws["F31"].value == pytest.approx(3500.0)
    # The hardcoded LTM column (fed by ltm-metrics) resolves through IFERROR.
    assert ws["D34"].value == pytest.approx(1200.0)
    assert ws["D35"].value == pytest.approx(300.0)
    # CapIQ cell could not resolve -> degraded to n/a; the recalc did not crash.
    assert "n/a" in str(ws["E34"].value)


# ─── The shared Excel-COM instance owner (v0.5.38 orphan fix) ─────────────────


def test_excel_com_app_quits_then_uninitializes(monkeypatch):
    log: list[str] = []
    app = FakeExcelApp(log)
    install_fake_com(monkeypatch, log, app)

    with e2p.excel_com_app(purpose="unit test", visible=False) as excel:
        assert excel is app
        log.append("body")

    assert log == ["CoInitialize", "DispatchEx(Excel.Application)", "body", "Quit", "CoUninitialize"]
    assert app.settings["Visible"] is False
    assert app.settings["DisplayAlerts"] is False


def test_excel_com_app_uninitializes_even_when_quit_raises(monkeypatch, capsys):
    """The v0.5.37 divergence: all four Excel sites called ``excel.Quit()`` bare in
    a ``finally``, so a Quit that raises (a modal add-in dialog does exactly this)
    escaped and ``CoUninitialize()`` never ran — leaving the thread
    COM-initialized for the rest of the process. ``slide_render`` already guarded
    its Quit; these now match."""
    log: list[str] = []
    app = FakeExcelApp(log, quit_raises=True)
    install_fake_com(monkeypatch, log, app)

    with e2p.excel_com_app(purpose="unit test", visible=False):
        pass  # no error escapes the block

    assert log[-2:] == ["Quit", "CoUninitialize"], "the apartment teardown is unconditional"
    assert "Quit() failed" in capsys.readouterr().err


def test_excel_com_app_uninitializes_when_the_body_raises(monkeypatch):
    log: list[str] = []
    app = FakeExcelApp(log)
    install_fake_com(monkeypatch, log, app)

    with pytest.raises(ValueError):
        with e2p.excel_com_app(purpose="unit test", visible=False):
            raise ValueError("boom")

    assert log[-2:] == ["Quit", "CoUninitialize"]


def test_excel_com_app_normalizes_a_failed_dispatch_and_warns(monkeypatch, capsys):
    """The MEASURED orphan source (2026-07-27): ``CoCreateInstanceEx`` fails with
    "Server execution failed" *after* Excel has already launched, so a
    fully-started instance is left with no client and no pointer to Quit. It must
    normalize to RuntimeError (so callers degrade to their non-COM backend) and
    say so loudly — silently degrading is how 13 orphans accumulated unnoticed."""
    log: list[str] = []
    install_fake_com(monkeypatch, log, None)

    with pytest.raises(RuntimeError, match="Excel COM unavailable"):
        with e2p.excel_com_app(purpose="unit test", visible=False):
            pass

    err = capsys.readouterr().err
    assert "Excel COM startup failed" in err
    assert "orphan signature" in err
    assert "Do not force-kill" in err
    # The apartment is still torn down even though no app was ever obtained.
    assert log == ["CoInitialize", "DispatchEx(Excel.Application)", "CoUninitialize"]


def test_excel_com_app_parks_a_visible_instance_offscreen(monkeypatch):
    log: list[str] = []
    app = FakeExcelApp(log)
    install_fake_com(monkeypatch, log, app)

    with e2p.excel_com_app(
        purpose="unit test", visible=True, park_offscreen=True, hide_comment_indicators=True
    ):
        pass

    assert app.settings["Visible"] is True
    assert (app.settings["Top"], app.settings["Left"]) == (4000, 6000)
    # xlNoIndicator: CopyPicture(xlScreen) would otherwise bake the red
    # cell-comment triangles into the picture.
    assert app.settings["DisplayCommentIndicator"] == 0


def test_excel_com_app_missing_pywin32_raises_runtime_error(monkeypatch):
    """No pywin32 must read as RuntimeError so each caller's non-COM fallback engages."""
    import builtins

    real_import = builtins.__import__

    def _no_pywin32(name, *args, **kwargs):
        if name in ("pythoncom", "win32com.client", "win32com"):
            raise ImportError(f"no module named {name}")
        return real_import(name, *args, **kwargs)

    monkeypatch.setattr(builtins, "__import__", _no_pywin32)
    with pytest.raises(RuntimeError, match="pywin32 is required for unit test"):
        with e2p.excel_com_app(purpose="unit test", visible=False):
            pass


def test_no_bare_excel_dispatch_outside_the_instance_owner():
    """Every Excel instance is created by ``excel_com_app``.

    The drift lock's sibling for COM, mirroring the ``find_soffice`` one above: the
    Quit-guard inconsistency that Phase I item 3 fixes existed precisely *because*
    four copies of this block were maintained independently. ``slide_render`` is
    exempt — it dispatches PowerPoint, whose COM server is a singleton and needs
    the opposite handling (a ``Presentations.Count`` guard so it never closes the
    analyst's open decks).
    """
    probe = re.compile(r"""DispatchEx\(\s*["']Excel\.Application""")
    offenders = [
        f"{py.relative_to(_SCRIPTS_DIR).as_posix()}:{n}"
        for py in sorted(_SCRIPTS_DIR.rglob("*.py"))
        if py.name != "excel_to_powerpoint.py"  # the instance owner's own module
        and "tests" not in py.relative_to(_SCRIPTS_DIR).parts
        for n, line in enumerate(py.read_text(encoding="utf-8").splitlines(), 1)
        if probe.search(line)
    ]

    assert not offenders, (
        "create Excel through excel_to_powerpoint.excel_com_app() rather than a bare "
        f"DispatchEx (it guards Quit and always tears the apartment down): {offenders}"
    )


def test_com_sites_release_their_workbook_before_quit(monkeypatch):
    """Each site must close/release its COM children *inside* the apartment that
    created them, before ``excel_com_app`` quits the instance."""
    log: list[str] = []

    class _Border:
        LineStyle = 0

    class _ChartArea:
        Border = _Border()

    class _Shapes:
        Count = 1

    class _Chart:
        ChartArea = _ChartArea()
        Shapes = _Shapes()

        def Paste(self) -> None:
            log.append("Paste")

        def Export(self, Filename: str, FilterName: str) -> None:  # noqa: N803
            log.append("Export")
            Path(Filename).write_bytes(b"\x89PNG\r\n\x1a\n" + b"0" * 32)

    class _ChartObject:
        Chart = _Chart()

        def Delete(self) -> None:
            log.append("chart_obj.Delete")

    class _ChartObjects:
        def Add(self, **_kwargs) -> _ChartObject:
            return _ChartObject()

    class _Range:
        Width = 400
        Height = 300

        def CopyPicture(self, **_kwargs) -> None:
            log.append("CopyPicture")

    class _Worksheet:
        def Range(self, _ref: str) -> _Range:
            return _Range()

        def ChartObjects(self) -> _ChartObjects:
            return _ChartObjects()

    class _Workbook:
        def Worksheets(self, _name: str) -> _Worksheet:
            return _Worksheet()

        def Close(self, SaveChanges: bool) -> None:  # noqa: N803
            log.append("wb.Close")

    class _Workbooks:
        def Open(self, _path: str, **_kwargs) -> _Workbook:
            log.append("wb.Open")
            return _Workbook()

    class _App(FakeExcelApp):
        Workbooks = _Workbooks()

        def CalculateFull(self) -> None:
            log.append("CalculateFull")

    install_fake_com(monkeypatch, log, _App(log))

    buf = e2p._excel_com_range_to_png(Path("cap.xlsx"), "Cap with Links", "B15:F40")

    assert buf.getvalue().startswith(b"\x89PNG")
    # The chart object, then the workbook, then Quit, then the apartment.
    assert log.index("chart_obj.Delete") < log.index("wb.Close") < log.index("Quit")
    assert log.index("Quit") < log.index("CoUninitialize")


# ─── Phase I item 1: an exhausted clipboard retry degrades, not aborts ────────


def test_exhausted_clipboard_retry_degrades_to_libreoffice(monkeypatch, capsys):
    """``_ClipboardPasteError`` stays a non-RuntimeError on purpose (a transient
    clipboard race must not be mislabeled "Excel unavailable"), but once the
    retries are SPENT the render is not happening on the COM path — and
    LibreOffice renders the same range fine. Before this, the error escaped
    ``_render_range_to_png`` and aborted the stage."""
    sentinel = io.BytesIO(b"libreoffice-png")

    def _exhausted(*_args, **_kwargs):
        raise e2p._ClipboardPasteError("Chart.Paste pasted nothing for B15:F40")

    monkeypatch.setattr(e2p.sys, "platform", "win32")
    monkeypatch.setattr(e2p, "_excel_com_range_to_png", _exhausted)
    monkeypatch.setattr(e2p, "_libreoffice_range_to_png", lambda *a, **k: sentinel)

    assert e2p._render_range_to_png(Path("cap.xlsx"), "Cap with Links", "B15:F40") is sentinel
    assert "clipboard retries exhausted" in capsys.readouterr().err


def test_com_unavailable_still_degrades_to_libreoffice(monkeypatch):
    """The pre-existing RuntimeError fall-through is unchanged."""
    sentinel = io.BytesIO(b"libreoffice-png")

    def _unavailable(*_args, **_kwargs):
        raise RuntimeError("Excel COM unavailable: nope")

    monkeypatch.setattr(e2p.sys, "platform", "win32")
    monkeypatch.setattr(e2p, "_excel_com_range_to_png", _unavailable)
    monkeypatch.setattr(e2p, "_libreoffice_range_to_png", lambda *a, **k: sentinel)

    assert e2p._render_range_to_png(Path("cap.xlsx"), "Cap with Links", "B15:F40") is sentinel


# ─── Phase I item 2: the Excel-COM tests are gated off by default ─────────────

_GATED_COM_TESTS = {
    "test_slide_library_poc.py": (
        "test_pitch_deck_inserts_cap_table_into_slide7",
        "test_pitch_deck_inserts_ownership_into_slide",
        "test_pitch_deck_inserts_institutions_with_bloomberg",
    ),
    "test_earnings_update_assembler.py": (
        "test_assemble_earnings_update_deck_inserts_cap_table_from_workbook",
    ),
}


def test_excel_spawning_tests_carry_the_opt_in_marker():
    """The four tests that spawn a real visible Excel must stay behind
    ``@pytest.mark.excel_com``. Unmarked, they run by default and each can park a
    modal Cap IQ add-in dialog in front of the suite."""
    tests_dir = Path(__file__).resolve().parent
    unmarked = []
    for filename, names in _GATED_COM_TESTS.items():
        lines = (tests_dir / filename).read_text(encoding="utf-8").splitlines()
        for name in names:
            idx = next(
                (i for i, line in enumerate(lines) if line.startswith(f"def {name}(")), None
            )
            assert idx is not None, f"{filename}::{name} not found — did it get renamed?"
            if "@pytest.mark.excel_com" not in lines[idx - 1]:
                unmarked.append(f"{filename}::{name}")

    assert not unmarked, f"missing @pytest.mark.excel_com: {unmarked}"


def test_libreoffice_captable_variant_is_not_gated():
    """The LibreOffice sibling is a DIFFERENT surface — it is what should keep
    running by default, so it must NOT pick up the Excel-COM gate."""
    lines = (
        (Path(__file__).resolve().parent / "test_earnings_update_assembler.py")
        .read_text(encoding="utf-8")
        .splitlines()
    )
    idx = next(
        i
        for i, line in enumerate(lines)
        if line.startswith("def test_assemble_earnings_update_deck_inserts_cap_table_via_libreoffice(")
    )
    assert "@pytest.mark.excel_com" not in lines[idx - 1]


def _conftest_module():
    """Load ``conftest.py`` by path — ``tests/`` is a package, so it is not
    importable as a top-level ``conftest`` module."""
    import importlib.util

    path = Path(__file__).resolve().parent / "conftest.py"
    spec = importlib.util.spec_from_file_location("_infor_beta_conftest", path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)  # type: ignore[union-attr]
    return module


@pytest.mark.parametrize(
    "value, opted_in",
    [(None, False), ("", False), ("0", False), ("false", False), ("no", False),
     ("1", True), ("true", True), ("yes", True)],
)
def test_excel_com_gate_env_var_parsing(monkeypatch, value, opted_in):
    conftest = _conftest_module()

    if value is None:
        monkeypatch.delenv(conftest._EXCEL_COM_ENV_VAR, raising=False)
    else:
        monkeypatch.setenv(conftest._EXCEL_COM_ENV_VAR, value)

    assert conftest._excel_com_opted_in() is opted_in


def test_excel_com_gate_skips_marked_items_only(monkeypatch):
    conftest = _conftest_module()

    class _Item:
        def __init__(self, keywords):
            self.keywords = keywords
            self.markers = []

        def add_marker(self, marker):
            self.markers.append(marker)

    marked, plain = _Item({"excel_com": True}), _Item({})

    monkeypatch.delenv(conftest._EXCEL_COM_ENV_VAR, raising=False)
    conftest.pytest_collection_modifyitems(None, [marked, plain])
    assert len(marked.markers) == 1 and marked.markers[0].name == "skip"
    assert plain.markers == []

    marked2, plain2 = _Item({"excel_com": True}), _Item({})
    monkeypatch.setenv(conftest._EXCEL_COM_ENV_VAR, "1")
    conftest.pytest_collection_modifyitems(None, [marked2, plain2])
    assert marked2.markers == [] and plain2.markers == []
