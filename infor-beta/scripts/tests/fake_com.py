"""A fake Excel-COM stack, so the four COM sites' cleanup ordering is locked on
every platform **without spawning Excel** — spawning it is exactly what leaks.

The real COM backends had no pytest coverage at all before v0.5.38: the chart and
aggregator tests stub `_build_charts_com` / `_combine_via_com` wholesale to
exercise their openpyxl fallbacks, and the range renderer's only coverage was the
four real-Excel tests that are now gated off by default. That left the v0.5.38
restructuring of all four bodies unverified by the suite, which is what these
fakes fix.

Every fake records into one shared ordered `log`, so a test can assert the
invariant that matters: **COM children are released before `Quit`, and the
apartment is torn down last.**
"""

from __future__ import annotations

import types


class FakePythoncom:
    def __init__(self, log: list[str]) -> None:
        self._log = log

    def CoInitialize(self) -> None:
        self._log.append("CoInitialize")

    def CoUninitialize(self) -> None:
        self._log.append("CoUninitialize")


class FakeExcelApp:
    """Records `Quit`; swallows arbitrary app-level property sets into `settings`."""

    def __init__(self, log: list[str], *, quit_raises: bool = False) -> None:
        self._log = log
        self._quit_raises = quit_raises
        self.settings: dict[str, object] = {}

    def __setattr__(self, name: str, value: object) -> None:
        if name.startswith("_") or name == "settings":
            super().__setattr__(name, value)
        else:
            self.settings[name] = value

    def Quit(self) -> None:
        self._log.append("Quit")
        if self._quit_raises:
            raise OSError("modal add-in dialog blocked Quit")


class FakeWin32ComClient:
    def __init__(self, log: list[str], app: object | None) -> None:
        self._log = log
        self._app = app

    def DispatchEx(self, progid: str) -> object:
        self._log.append(f"DispatchEx({progid})")
        if self._app is None:
            # The measured orphan source: "Server execution failed" AFTER the
            # launch already succeeded.
            raise OSError("(-2146959355, 'Server execution failed', None, None)")
        return self._app


def install_fake_com(monkeypatch, log: list[str], app: object | None) -> None:
    """Put the fake `pythoncom` / `win32com.client` into `sys.modules`.

    Replaces the real pywin32 on a Windows dev box too, which is the point: the
    ordering assertions then hold identically on Windows, macOS and Linux.
    """
    import sys

    client = FakeWin32ComClient(log, app)
    pkg = types.ModuleType("win32com")
    pkg.client = client  # type: ignore[attr-defined]
    monkeypatch.setitem(sys.modules, "pythoncom", FakePythoncom(log))
    monkeypatch.setitem(sys.modules, "win32com", pkg)
    monkeypatch.setitem(sys.modules, "win32com.client", client)


def assert_released_before_quit(log: list[str], *children: str) -> None:
    """Every named child release must precede `Quit`, which precedes the teardown."""
    assert "Quit" in log and "CoUninitialize" in log, log
    for child in children:
        assert child in log, f"{child} never happened: {log}"
        assert log.index(child) < log.index("Quit"), (
            f"{child} must be released before Quit: {log}"
        )
    assert log.index("Quit") < log.index("CoUninitialize"), log
