from __future__ import annotations

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


def test_build_blender_command_visible_mode_omits_background_flag_and_registers_listener_before_render(
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
    assert command.index("--python") < command.index("-f")
    assert command.index("-f") < command.index("--")
    assert command[-2:] == ["", "visible"]


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
    assert command[-2:] == ["", "headless"]


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
    monkeypatch.setattr(main, "minimize_visible_blender_window", lambda process_id, timeout=main.BLENDER_VISIBLE_WINDOW_MINIMIZE_TIMEOUT_SECONDS: None)
    monkeypatch.setattr(
        main,
        "terminate_process_tree",
        lambda process, wait_seconds=main.BLENDER_VISIBLE_FORCE_EXIT_WAIT_SECONDS: terminated.append(process.pid) or 0,
    )
    monkeypatch.setattr(main.time, "sleep", lambda _seconds: None)

    exit_code = main.run_blender_session(args, state_path, record)

    assert exit_code == 0
    assert terminated == [4321]


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
        lambda process_id, timeout=main.BLENDER_VISIBLE_WINDOW_MINIMIZE_TIMEOUT_SECONDS: minimized.append(process_id),
    )

    exit_code = main.run_blender_session(args, state_path, record)

    assert exit_code == 2
    assert minimized == [9876]
