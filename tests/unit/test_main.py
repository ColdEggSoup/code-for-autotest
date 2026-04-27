from __future__ import annotations

import json
from pathlib import Path
import subprocess
from types import SimpleNamespace

import main
from monitoring_common import save_json


def test_default_blender_executable_prefers_existing_repo_install(monkeypatch, tmp_path: Path) -> None:
    repo_candidate = tmp_path / "software" / "blender" / "blender.exe"
    repo_candidate.parent.mkdir(parents=True, exist_ok=True)
    repo_candidate.write_bytes(b"exe")
    program_files_candidate = tmp_path / "program_files" / "blender.exe"
    program_files_candidate.parent.mkdir(parents=True, exist_ok=True)
    program_files_candidate.write_bytes(b"exe")

    monkeypatch.setattr(
        main,
        "default_blender_executable_candidates",
        lambda: (repo_candidate, program_files_candidate),
    )

    assert main.default_blender_executable() == repo_candidate


def test_default_blender_executable_returns_repo_candidate_when_none_exist(monkeypatch, tmp_path: Path) -> None:
    repo_candidate = tmp_path / "software" / "blender" / "blender.exe"
    program_files_candidate = tmp_path / "program_files" / "blender.exe"

    monkeypatch.setattr(
        main,
        "default_blender_executable_candidates",
        lambda: (repo_candidate, program_files_candidate),
    )

    assert main.default_blender_executable() == repo_candidate


def test_build_blender_command_visible_mode_omits_direct_render_flags_and_registers_listener_before_render(
    monkeypatch,
    tmp_path: Path,
) -> None:
    blender_exe = tmp_path / "blender.exe"
    blender_exe.write_bytes(b"exe")
    blend_file = tmp_path / "scene.blend"
    blend_file.write_text("blend", encoding="utf-8")
    listener_script = tmp_path / "blender_listener.py"
    listener_script.write_text("print('listener')", encoding="utf-8")

    monkeypatch.setattr(main, "default_blender_listener_script", lambda: listener_script)

    args = SimpleNamespace(
        blender_exe=str(blender_exe),
        blend_file=str(blend_file),
        blender_ui_mode="visible",
        render_mode="frame",
        frame=1,
    )
    record = {"output_path": tmp_path / "out.csv", "session_id": "session-001"}

    command = main.build_blender_command(args, record)

    assert "-b" not in command
    assert "-f" not in command
    assert "-a" not in command
    assert command.index("--python") < command.index("--")
    assert command[-4:] == ["", "visible", "frame", "1"]


def test_build_blender_command_headless_mode_keeps_background_flag(monkeypatch, tmp_path: Path) -> None:
    blender_exe = tmp_path / "blender.exe"
    blender_exe.write_bytes(b"exe")
    blend_file = tmp_path / "scene.blend"
    blend_file.write_text("blend", encoding="utf-8")
    listener_script = tmp_path / "blender_listener.py"
    listener_script.write_text("print('listener')", encoding="utf-8")

    monkeypatch.setattr(main, "default_blender_listener_script", lambda: listener_script)

    args = SimpleNamespace(
        blender_exe=str(blender_exe),
        blend_file=str(blend_file),
        blender_ui_mode="headless",
        render_mode="frame",
        frame=1,
    )
    record = {"output_path": tmp_path / "out.csv", "session_id": "session-001"}

    command = main.build_blender_command(args, record)

    assert "-b" in command
    assert command.index("--python") < command.index("-f")
    assert command[-4:] == ["", "headless", "frame", "1"]


def test_minimize_visible_blender_window_minimizes_all_visible_windows_for_the_process(monkeypatch) -> None:
    class _FakeRect:
        def __init__(self, width: int, height: int) -> None:
            self._width = width
            self._height = height

        def width(self) -> int:
            return self._width

        def height(self) -> int:
            return self._height

    class _FakeWindow:
        def __init__(self, handle: int, title: str, *, process_id: int, visible: bool = True, area: int = 10000) -> None:
            self.handle = handle
            self._title = title
            self._visible = visible
            self._rect = _FakeRect(area, 1)
            self.element_info = SimpleNamespace(process_id=process_id)

        def is_visible(self) -> bool:
            return self._visible

        def window_text(self) -> str:
            return self._title

        def rectangle(self):
            return self._rect

    class _FakeDesktop:
        def __init__(self, backend: str = "uia") -> None:
            assert backend == "uia"

        def windows(self):
            return [
                _FakeWindow(1001, "Blender", process_id=4321, area=50000),
                _FakeWindow(1002, "Blender Render", process_id=4321, area=25000),
                _FakeWindow(1003, "Other App", process_id=9999, area=75000),
                _FakeWindow(1004, "Hidden Blender", process_id=4321, visible=False, area=20000),
            ]

    minimized_handles: list[int] = []

    monkeypatch.setattr(main, "_import_pywinauto_desktop", lambda: _FakeDesktop)
    monkeypatch.setattr(main, "minimize_window", lambda window: minimized_handles.append(window.handle))

    minimized_count = main.minimize_visible_blender_window(4321, timeout=0.1)

    assert minimized_count == 2
    assert minimized_handles == [1001, 1002]


def test_run_blender_session_visible_mode_terminates_lingering_process_after_csv_capture(
    monkeypatch,
    tmp_path: Path,
) -> None:
    blender_exe = tmp_path / "blender.exe"
    blender_exe.write_bytes(b"exe")
    blend_file = tmp_path / "scene.blend"
    blend_file.write_text("blend", encoding="utf-8")
    state_path = tmp_path / "session.json"
    save_json(state_path, {})
    output_path = tmp_path / "out.csv"
    record = {
        "output_path": str(output_path),
        "session_id": "session-001",
        "software": "blender",
        "test_name": "demo",
    }
    args = SimpleNamespace(
        blender_exe=str(blender_exe),
        blend_file=str(blend_file),
        blender_ui_mode="visible",
        render_mode="frame",
        frame=1,
    )

    class _FakeProcess:
        def __init__(self) -> None:
            self.pid = 4321
            self.returncode = None

        def poll(self):
            return self.returncode

        def wait(self, timeout=None):
            if timeout is not None and self.returncode is None:
                raise subprocess.TimeoutExpired(cmd="blender", timeout=timeout)
            return self.returncode if self.returncode is not None else 0

        def kill(self) -> None:
            self.returncode = 0

    fake_process = _FakeProcess()
    terminated: list[int] = []
    row_counts = iter([1, 1])

    monkeypatch.setattr(main, "build_blender_command", lambda current_args, current_record: ["blender"])
    monkeypatch.setattr(main.subprocess, "Popen", lambda *args, **kwargs: fake_process)
    monkeypatch.setattr(main, "count_csv_rows", lambda path: next(row_counts, 1))
    monkeypatch.setattr(
        main,
        "minimize_visible_blender_window",
        lambda process_id, timeout=main.BLENDER_VISIBLE_WINDOW_MINIMIZE_TIMEOUT_SECONDS, minimized_handles=None: 1,
    )
    monkeypatch.setattr(main, "_collect_visible_windows_for_pid", lambda process_id: ([], []))
    monkeypatch.setattr(
        main,
        "terminate_process_tree",
        lambda process, wait_seconds=main.BLENDER_VISIBLE_FORCE_EXIT_WAIT_SECONDS: terminated.append(process.pid) or 0,
    )
    monkeypatch.setattr(main.time, "sleep", lambda _seconds: None)

    exit_code = main.run_blender_session(args, state_path, record)

    assert exit_code == 0
    assert terminated == [4321]


def test_run_blender_session_visible_mode_minimizes_newly_visible_windows_while_running(monkeypatch, tmp_path: Path) -> None:
    blender_exe = tmp_path / "blender.exe"
    blender_exe.write_bytes(b"exe")
    blend_file = tmp_path / "scene.blend"
    blend_file.write_text("blend", encoding="utf-8")
    state_path = tmp_path / "session.json"
    save_json(state_path, {})
    output_path = tmp_path / "out.csv"
    record = {
        "output_path": str(output_path),
        "session_id": "session-001",
        "software": "blender",
        "test_name": "demo",
    }
    args = SimpleNamespace(
        blender_exe=str(blender_exe),
        blend_file=str(blend_file),
        blender_ui_mode="visible",
        render_mode="frame",
        frame=1,
    )

    class _FakeProcess:
        def __init__(self) -> None:
            self.pid = 2468
            self._poll_values = iter([None, 0])

        def poll(self):
            return next(self._poll_values, 0)

        def wait(self, timeout=None):
            return 0

        def kill(self) -> None:
            return None

    additional_minimize_calls: list[int] = []

    monkeypatch.setattr(main, "build_blender_command", lambda current_args, current_record: ["blender"])
    monkeypatch.setattr(main.subprocess, "Popen", lambda *args, **kwargs: _FakeProcess())
    monkeypatch.setattr(main, "count_csv_rows", lambda path: 0)
    monkeypatch.setattr(
        main,
        "minimize_visible_blender_window",
        lambda process_id, timeout=main.BLENDER_VISIBLE_WINDOW_MINIMIZE_TIMEOUT_SECONDS, minimized_handles=None: 1,
    )
    monkeypatch.setattr(
        main,
        "_collect_visible_windows_for_pid",
        lambda process_id: ([SimpleNamespace(handle=9001)], ["Blender Render"]),
    )
    monkeypatch.setattr(
        main,
        "_minimize_visible_windows",
        lambda windows, minimized_handles=None: additional_minimize_calls.append(len(list(windows))) or 1,
    )
    monkeypatch.setattr(main.time, "sleep", lambda _seconds: None)

    exit_code = main.run_blender_session(args, state_path, record)

    assert exit_code == 2
    assert additional_minimize_calls == [1]


def test_run_blender_session_visible_mode_attempts_to_minimize_window(monkeypatch, tmp_path: Path) -> None:
    blender_exe = tmp_path / "blender.exe"
    blender_exe.write_bytes(b"exe")
    blend_file = tmp_path / "scene.blend"
    blend_file.write_text("blend", encoding="utf-8")
    state_path = tmp_path / "session.json"
    save_json(state_path, {})
    output_path = tmp_path / "out.csv"
    record = {
        "output_path": str(output_path),
        "session_id": "session-001",
        "software": "blender",
        "test_name": "demo",
    }
    args = SimpleNamespace(
        blender_exe=str(blender_exe),
        blend_file=str(blend_file),
        blender_ui_mode="visible",
        render_mode="frame",
        frame=1,
    )

    class _FakeProcess:
        def __init__(self) -> None:
            self.pid = 9876
            self.returncode = 0

        def poll(self):
            return self.returncode

        def wait(self, timeout=None):
            return self.returncode

        def kill(self) -> None:
            self.returncode = 0

    minimized: list[int] = []

    monkeypatch.setattr(main, "build_blender_command", lambda current_args, current_record: ["blender"])
    monkeypatch.setattr(main.subprocess, "Popen", lambda *args, **kwargs: _FakeProcess())
    monkeypatch.setattr(
        main,
        "minimize_visible_blender_window",
        lambda process_id, timeout=main.BLENDER_VISIBLE_WINDOW_MINIMIZE_TIMEOUT_SECONDS, minimized_handles=None: minimized.append(process_id) or 1,
    )
    monkeypatch.setattr(main, "_collect_visible_windows_for_pid", lambda process_id: ([], []))

    exit_code = main.run_blender_session(args, state_path, record)

    assert exit_code == 2
    assert minimized == [9876]


def test_load_session_record_retries_transient_json_decode_error(monkeypatch, tmp_path: Path) -> None:
    runtime_root = tmp_path / "runtime"
    state_path = main.resolve_session_paths(runtime_root, "session-001")["state_path"]
    state_path.parent.mkdir(parents=True, exist_ok=True)
    state_path.write_text("{}", encoding="utf-8")
    expected_record = {"status": "running", "session_id": "session-001"}
    call_count = {"value": 0}

    def _fake_load_json(path: Path):
        assert path == state_path
        call_count["value"] += 1
        if call_count["value"] == 1:
            raise json.JSONDecodeError("Expecting value", "", 0)
        return expected_record

    monkeypatch.setattr(main, "load_json", _fake_load_json)
    monkeypatch.setattr(main.time, "sleep", lambda _seconds: None)

    resolved_state_path, resolved_record = main.load_session_record(runtime_root, "session-001")

    assert resolved_state_path == state_path
    assert resolved_record == expected_record
    assert call_count["value"] == 2
