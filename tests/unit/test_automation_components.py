from __future__ import annotations

import subprocess

import automation_components


class _DummyRect:
    def __init__(self, width: int, height: int) -> None:
        self._width = width
        self._height = height

    def width(self) -> int:
        return self._width

    def height(self) -> int:
        return self._height


class _DummyElementInfo:
    def __init__(self, process_id: int) -> None:
        self.process_id = process_id


class _DummyWindow:
    def __init__(self, title: str, process_id: int, *, width: int = 1200, height: int = 800) -> None:
        self._title = title
        self._rect = _DummyRect(width, height)
        self.element_info = _DummyElementInfo(process_id)

    def window_text(self) -> str:
        return self._title

    def rectangle(self) -> _DummyRect:
        return self._rect


class _DummyDesktop:
    def __init__(self, windows: list[_DummyWindow]) -> None:
        self._windows = windows

    def __call__(self, backend: str = "uia") -> "_DummyDesktop":
        assert backend == "uia"
        return self

    def windows(self) -> list[_DummyWindow]:
        return list(self._windows)


def test_connect_software_window_prefers_matching_process_name(monkeypatch) -> None:
    vscode_window = _DummyWindow(
        "run_shotcut_validation.py - code for autotest - Visual Studio Code",
        1001,
        width=1800,
        height=1100,
    )
    shotcut_window = _DummyWindow("Shotcut", 2002, width=1280, height=900)

    monkeypatch.setattr(
        automation_components,
        "_import_pywinauto_desktop",
        lambda: _DummyDesktop([vscode_window, shotcut_window]),
    )

    class _DummyProcess:
        def __init__(self, pid: int) -> None:
            self.pid = pid

        def name(self) -> str:
            return {
                1001: "Code.exe",
                2002: "shotcut.exe",
            }[self.pid]

    monkeypatch.setattr(automation_components.psutil, "Process", _DummyProcess)
    monkeypatch.setattr(automation_components.psutil, "process_iter", lambda _attrs: [])

    resolved = automation_components.connect_software_window("shotcut", timeout=0.2)
    assert resolved is shotcut_window


def test_wait_for_main_window_uses_startup_timeout_and_backend_fallback(monkeypatch) -> None:
    recorded: dict[str, object] = {}

    def _connect_software_window(software: str, timeout: float = 15.0, *, backends: tuple[str, ...] = ("uia",)):
        recorded["software"] = software
        recorded["timeout"] = timeout
        recorded["backends"] = backends
        return object()

    monkeypatch.setattr(automation_components, "connect_software_window", _connect_software_window)

    launcher = automation_components.SoftwareLauncher()
    launcher.wait_for_main_window("shotcut", timeout=30.0)

    assert recorded == {
        "software": "shotcut",
        "timeout": 90.0,
        "backends": ("uia", "win32"),
    }


def test_wait_for_main_window_uses_kdenlive_startup_timeout(monkeypatch) -> None:
    recorded: dict[str, object] = {}

    def _connect_software_window(software: str, timeout: float = 15.0, *, backends: tuple[str, ...] = ("uia",)):
        recorded["software"] = software
        recorded["timeout"] = timeout
        recorded["backends"] = backends
        return object()

    monkeypatch.setattr(automation_components, "connect_software_window", _connect_software_window)

    launcher = automation_components.SoftwareLauncher()
    launcher.wait_for_main_window("kdenlive", timeout=30.0)

    assert recorded == {
        "software": "kdenlive",
        "timeout": 90.0,
        "backends": ("uia", "win32"),
    }


def test_run_main_decodes_binary_output_without_locale_crash(monkeypatch, tmp_path) -> None:
    def _fake_run(*args, **kwargs):
        return subprocess.CompletedProcess(
            args=["python", "main.py"],
            returncode=0,
            stdout=b"session_id=test-session\nstatus=completed\n",
            stderr=b"\xa7\xffbinary-warning\n",
        )

    monkeypatch.setattr(automation_components.subprocess, "run", _fake_run)

    bridge = automation_components.MonitorBridge(repo_root=tmp_path)
    completed = bridge._run_main(["status", "--session-id", "test-session"])

    assert completed.stdout.startswith("session_id=test-session")
    assert isinstance(completed.stderr, str)
    assert completed.stderr


def test_completed_output_text_tolerates_empty_streams() -> None:
    completed = subprocess.CompletedProcess(args=["python"], returncode=1, stdout="", stderr="")
    assert automation_components._completed_output_text(completed) == ""
