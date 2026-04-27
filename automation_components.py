from __future__ import annotations

from dataclasses import dataclass
import locale
import logging
import os
from pathlib import Path
import re
import subprocess
import sys
import time

import psutil

from monitoring_common import FINAL_SESSION_STATUSES
from ui_automation import (
    UiAutomationError,
    WindowTopmostKeeper,
    clear_window_topmost,
    connect_window,
    request_window_close,
    set_performance_boost_state,
)


REPO_ROOT = Path(__file__).resolve().parent
WHITELIST_APP_DIR = REPO_ROOT / "whitelist app"
TEST_DATA_DIR = REPO_ROOT / "test_data"
DEFAULT_AI_TURBO_TITLE_RE = r"^Lenovo AI Turbo Engine$"
DEFAULT_PIPELINE_WORKLOAD_NAME = "4k_big_video_processing_speed"
logger = logging.getLogger(__name__)
TERMINAL_SESSION_STATUSES = FINAL_SESSION_STATUSES | {"stale"}


@dataclass(frozen=True)
class SoftwareRuntimeSpec:
    software: str
    launch_path: Path | None
    main_window_title_re: str
    ai_turbo_app_name: str
    export_suffix: str
    blend_file: Path | None = None
    process_names: tuple[str, ...] = ()
    startup_timeout_seconds: float = 30.0


@dataclass(frozen=True)
class MonitorLaunchResult:
    software: str
    session_id: str
    output_path: Path
    raw_output: str
    state_path: Path | None = None
    session_output_path: Path | None = None
    worker_stdout_path: Path | None = None
    worker_stderr_path: Path | None = None
    status: str = ""


SOFTWARE_SPECS = {
    "shotcut": SoftwareRuntimeSpec(
        software="shotcut",
        launch_path=WHITELIST_APP_DIR / "shotcut.lnk",
        main_window_title_re=r".*Shotcut.*",
        ai_turbo_app_name="shotcut",
        export_suffix=".mp4",
        process_names=("shotcut.exe",),
        startup_timeout_seconds=90.0,
    ),
    "kdenlive": SoftwareRuntimeSpec(
        software="kdenlive",
        launch_path=WHITELIST_APP_DIR / "Kdenlive.lnk",
        main_window_title_re=r".*Kdenlive.*",
        ai_turbo_app_name="kdenlive",
        export_suffix=".mp4",
        process_names=("kdenlive.exe",),
        startup_timeout_seconds=90.0,
    ),
    "shutter_encoder": SoftwareRuntimeSpec(
        software="shutter_encoder",
        launch_path=WHITELIST_APP_DIR / "Shutter Encoder.lnk",
        main_window_title_re=r".*Shutter Encoder.*",
        ai_turbo_app_name="Shutter Encoder",
        export_suffix=".mp4",
        process_names=("javaw.exe", "Shutter Encoder.exe"),
    ),
    "avidemux": SoftwareRuntimeSpec(
        software="avidemux",
        launch_path=WHITELIST_APP_DIR / "avidemux.lnk",
        main_window_title_re=r".*Avidemux.*",
        ai_turbo_app_name="avidemux",
        export_suffix=".mp4",
        process_names=("avidemux.exe",),
    ),
    "handbrake": SoftwareRuntimeSpec(
        software="handbrake",
        launch_path=WHITELIST_APP_DIR / "HandBrake.lnk",
        main_window_title_re=r".*HandBrake.*",
        ai_turbo_app_name="HandBrake",
        export_suffix=".mp4",
        process_names=("HandBrake.exe",),
    ),
    "blender": SoftwareRuntimeSpec(
        software="blender",
        launch_path=None,
        main_window_title_re=r".*Blender.*",
        ai_turbo_app_name="blender",
        export_suffix=".mp4",
        blend_file=WHITELIST_APP_DIR / "blender" / "13263.blend",
        process_names=("blender.exe",),
    ),
}


def _import_pywinauto_desktop():
    try:
        from pywinauto import Desktop
    except ImportError as exc:
        raise UiAutomationError(
            "pywinauto is not installed. Run: pip install -r requirements.txt"
        ) from exc
    return Desktop


def _safe_window_text(window) -> str:
    try:
        return (window.window_text() or "").strip()
    except Exception:
        return ""


def _safe_process_name(pid: int | None) -> str:
    if not pid:
        return ""
    try:
        return (psutil.Process(int(pid)).name() or "").strip()
    except (psutil.Error, OSError, ValueError):
        return ""


def _list_matching_processes(process_names: tuple[str, ...]) -> list[tuple[str, int, str]]:
    expected = {name.casefold() for name in process_names}
    if not expected:
        return []
    matches: list[tuple[str, int, str]] = []
    for process in psutil.process_iter(["pid", "name", "status"]):
        try:
            process_name = (process.info.get("name") or "").strip()
        except (psutil.Error, OSError, ValueError):
            continue
        if not process_name or process_name.casefold() not in expected:
            continue
        pid = int(process.info.get("pid") or 0)
        status = str(process.info.get("status") or "")
        matches.append((process_name, pid, status))
    matches.sort(key=lambda item: (item[0].casefold(), item[1]))
    return matches


def connect_software_window(
    software: str,
    timeout: float = 15.0,
    *,
    backends: tuple[str, ...] = ("uia",),
):
    spec = SOFTWARE_SPECS[software]
    Desktop = _import_pywinauto_desktop()
    title_pattern = re.compile(spec.main_window_title_re, re.IGNORECASE)
    expected_process_names = {name.casefold() for name in spec.process_names}
    deadline = time.monotonic() + timeout
    last_seen_titles: list[str] = []
    last_seen_processes: list[tuple[str, int, str]] = []

    while time.monotonic() < deadline:
        candidates = []
        windows = []
        seen_window_keys: set[tuple[int, int, str]] = set()
        for backend in backends:
            try:
                backend_windows = list(Desktop(backend=backend).windows())
            except Exception:
                continue
            for window in backend_windows:
                title = _safe_window_text(window)
                pid = int(getattr(getattr(window, "element_info", None), "process_id", 0) or 0)
                handle = int(getattr(window, "handle", 0) or 0)
                key = (handle, pid, title.casefold())
                if key in seen_window_keys:
                    continue
                seen_window_keys.add(key)
                windows.append(window)
        last_seen_titles = [_safe_window_text(window) for window in windows if _safe_window_text(window)]
        last_seen_processes = _list_matching_processes(spec.process_names)
        for window in windows:
            title = _safe_window_text(window)
            if not title or not title_pattern.search(title):
                continue
            pid = getattr(window.element_info, "process_id", None)
            process_name = _safe_process_name(pid)
            if expected_process_names and process_name.casefold() not in expected_process_names:
                continue
            try:
                rect = window.rectangle()
                area = max(1, rect.width() * rect.height())
            except Exception:
                area = 1
            candidates.append((-area, title.casefold(), int(pid or 0), window))
        if candidates:
            candidates.sort(key=lambda item: (item[0], item[1], item[2]))
            return candidates[0][3]
        time.sleep(0.25)

    raise UiAutomationError(
        f"Could not connect to a visible {software} window matching '{spec.main_window_title_re}' "
        f"and process names {spec.process_names}. Seen titles: {last_seen_titles[:10]}. "
        f"Matching processes: {last_seen_processes[:10]}"
    )


def parse_key_value_output(raw_output: str) -> dict[str, str]:
    result: dict[str, str] = {}
    for line in raw_output.splitlines():
        if "=" not in line:
            continue
        key, value = line.split("=", 1)
        result[key.strip()] = value.strip()
    return result


def _decode_subprocess_stream(payload: bytes | str | None) -> str:
    if payload is None:
        return ""
    if isinstance(payload, str):
        return payload

    encodings: list[str] = []
    for candidate in ("utf-8-sig", "utf-8", locale.getpreferredencoding(False), "gb18030"):
        if candidate and candidate not in encodings:
            encodings.append(candidate)

    for encoding_name in encodings:
        try:
            return payload.decode(encoding_name)
        except UnicodeDecodeError:
            continue
    return payload.decode("utf-8", errors="replace")


def _completed_output_text(completed: subprocess.CompletedProcess[str]) -> str:
    return (completed.stdout or "") + (completed.stderr or "")


def resolve_input_video(workload_name: str) -> Path:
    normalized = workload_name.strip().lower()
    assert normalized, "workload_name must not be empty."
    if "small" in normalized:
        path = TEST_DATA_DIR / "4K_small.mp4"
    elif "big" in normalized or "1080" in normalized:
        path = TEST_DATA_DIR / "4K_big.mp4"
    elif "4k" in normalized:
        path = TEST_DATA_DIR / "4K_small.mp4"
    else:
        path = TEST_DATA_DIR / "4K_big.mp4"
    assert path.exists(), f"Input video not found: {path}"
    return path


def default_monitor_output_path(software: str, test_id: str) -> Path:
    assert software in SOFTWARE_SPECS, f"Unsupported software: {software}"
    normalized_test_id = test_id.strip()
    assert normalized_test_id, "test_id must not be empty."
    return REPO_ROOT / "results" / "csv" / f"{software}_{normalized_test_id}.csv"


class SoftwareLauncher:
    def __init__(self, *, launch_wait_seconds: float = 5.0) -> None:
        assert launch_wait_seconds > 0, "launch_wait_seconds must be positive."
        self.launch_wait_seconds = launch_wait_seconds

    def get_spec(self, software: str) -> SoftwareRuntimeSpec:
        assert software in SOFTWARE_SPECS, f"Unsupported software: {software}"
        return SOFTWARE_SPECS[software]

    def launch(self, software: str) -> SoftwareRuntimeSpec:
        spec = self.get_spec(software)
        if spec.launch_path is None:
            logger.info("No launch shortcut is defined for %s. Skipping launch step.", software)
            return spec
        assert spec.launch_path.exists(), f"Launch target does not exist: {spec.launch_path}"
        logger.info("Launching %s from shortcut: %s", software, spec.launch_path)
        os.startfile(str(spec.launch_path))
        time.sleep(self.launch_wait_seconds)
        return spec

    def wait_for_main_window(self, software: str, timeout: float = 30.0):
        spec = self.get_spec(software)
        effective_timeout = max(timeout, spec.startup_timeout_seconds)
        logger.info(
            "Waiting for %s main window. requested_timeout=%.1fs effective_timeout=%.1fs",
            software,
            timeout,
            effective_timeout,
        )
        return connect_software_window(software, timeout=effective_timeout, backends=("uia", "win32"))


class MonitorBridge:
    def __init__(self, *, repo_root: Path | None = None, python_executable: str | None = None) -> None:
        self.repo_root = (repo_root or REPO_ROOT).resolve(strict=False)
        self.python_executable = python_executable or sys.executable

    def _run_main(self, args: list[str], *, timeout: float | None = None) -> subprocess.CompletedProcess[str]:
        command = [self.python_executable, "main.py", *args]
        completed = subprocess.run(
            command,
            cwd=str(self.repo_root),
            capture_output=True,
            text=False,
            check=False,
            timeout=timeout,
        )
        return subprocess.CompletedProcess(
            completed.args,
            completed.returncode,
            _decode_subprocess_stream(completed.stdout),
            _decode_subprocess_stream(completed.stderr),
        )

    def start_background_monitor(self, software: str, test_name: str, output_path: Path) -> MonitorLaunchResult:
        assert output_path.parent.exists(), f"Output directory does not exist: {output_path.parent}"
        completed = self._run_main(
            [
                "start-bg",
                "--software",
                software,
                "--name",
                test_name,
                "--output",
                str(output_path),
            ]
        )
        assert completed.returncode == 0, _completed_output_text(completed)
        payload = parse_key_value_output(completed.stdout)
        assert "session_id" in payload, completed.stdout
        resolved_output = payload.get("aggregate_output_path") or payload.get("output_path")
        assert resolved_output, completed.stdout
        return MonitorLaunchResult(
            software=software,
            session_id=payload["session_id"],
            output_path=Path(resolved_output).resolve(strict=False),
            state_path=Path(payload["state_path"]).resolve(strict=False) if payload.get("state_path") else None,
            session_output_path=Path(payload["session_output_path"]).resolve(strict=False)
            if payload.get("session_output_path")
            else None,
            worker_stdout_path=Path(payload["worker_stdout_path"]).resolve(strict=False)
            if payload.get("worker_stdout_path")
            else None,
            worker_stderr_path=Path(payload["worker_stderr_path"]).resolve(strict=False)
            if payload.get("worker_stderr_path")
            else None,
            status=str(payload.get("status", "") or ""),
            raw_output=completed.stdout,
        )

    def wait_for_session_completion(
        self,
        session_id: str,
        *,
        timeout_seconds: float = 7200.0,
        poll_interval_seconds: float = 2.0,
    ) -> dict[str, str]:
        assert timeout_seconds > 0, "timeout_seconds must be positive."
        deadline = time.monotonic() + timeout_seconds
        last_payload: dict[str, str] = {}
        while time.monotonic() < deadline:
            completed = self._run_main(["status", "--session-id", session_id])
            assert completed.returncode == 0, _completed_output_text(completed)
            last_payload = parse_key_value_output(completed.stdout)
            status = last_payload.get("status", "")
            if status in TERMINAL_SESSION_STATUSES:
                return last_payload
            time.sleep(poll_interval_seconds)
        raise AssertionError(f"Session '{session_id}' did not reach a final state. Last payload: {last_payload}")

    def stop_session(self, session_id: str, *, wait_seconds: float = 15.0) -> dict[str, str]:
        assert wait_seconds > 0, "wait_seconds must be positive."
        completed = self._run_main(
            [
                "stop",
                "--session-id",
                session_id,
                "--wait-seconds",
                str(wait_seconds),
            ],
            timeout=wait_seconds + 10.0,
        )
        assert completed.returncode == 0, _completed_output_text(completed)
        return parse_key_value_output(completed.stdout)

    def run_blender_monitor(
        self,
        *,
        test_name: str,
        output_path: Path,
        blend_file: Path,
        blender_executable: Path | None = None,
        blender_ui_mode: str = "visible",
        render_mode: str = "frame",
        frame: int = 1,
    ) -> MonitorLaunchResult:
        assert blend_file.exists(), f"Blend file does not exist: {blend_file}"
        assert output_path.parent.exists(), f"Output directory does not exist: {output_path.parent}"
        assert blender_ui_mode in {"visible", "headless"}, f"Unsupported Blender UI mode: {blender_ui_mode}"
        assert render_mode in {"frame", "animation"}, f"Unsupported Blender render mode: {render_mode}"
        assert frame >= 0, "frame must not be negative."
        args = [
            "start",
            "--software",
            "blender",
            "--name",
            test_name,
            "--output",
            str(output_path),
            "--blend-file",
            str(blend_file),
            "--blender-ui-mode",
            blender_ui_mode,
            "--render-mode",
            render_mode,
            "--frame",
            str(frame),
        ]
        if blender_executable is not None:
            args.extend(["--blender-exe", str(blender_executable)])
        completed = self._run_main(
            args,
            timeout=7200.0,
        )
        payload = parse_key_value_output(completed.stdout)
        assert completed.returncode == 0, _completed_output_text(completed)
        resolved_output = payload.get("aggregate_output_path") or payload.get("output_path") or str(output_path)
        return MonitorLaunchResult(
            software="blender",
            session_id=str(payload.get("session_id", "") or ""),
            output_path=Path(resolved_output).resolve(strict=False),
            state_path=Path(payload["state_path"]).resolve(strict=False) if payload.get("state_path") else None,
            session_output_path=Path(payload["session_output_path"]).resolve(strict=False)
            if payload.get("session_output_path")
            else None,
            worker_stdout_path=Path(payload["worker_stdout_path"]).resolve(strict=False)
            if payload.get("worker_stdout_path")
            else None,
            worker_stderr_path=Path(payload["worker_stderr_path"]).resolve(strict=False)
            if payload.get("worker_stderr_path")
            else None,
            status=str(payload.get("status", "") or ""),
            raw_output=completed.stdout,
        )


class AiTurboEngineController:
    def __init__(
        self,
        *,
        window_title_re: str = DEFAULT_AI_TURBO_TITLE_RE,
        shortcut_path: Path | None = None,
        topmost_interval_seconds: float = 1.0,
    ) -> None:
        self.window_title_re = window_title_re
        self.shortcut_path = shortcut_path
        self.topmost_interval_seconds = topmost_interval_seconds
        self._keeper: WindowTopmostKeeper | None = None

    def ensure_running(self) -> None:
        try:
            connect_window(self.window_title_re, timeout=1.0)
            return
        except UiAutomationError:
            pass
        if self.shortcut_path is not None and self.shortcut_path.exists():
            os.startfile(str(self.shortcut_path))
            time.sleep(3.0)
        try:
            connect_window(self.window_title_re, timeout=10.0)
        except UiAutomationError as exc:
            raise AssertionError(
                "AI Turbo Engine is not available. Open it first or add its shortcut under whitelist app."
            ) from exc

    def configure_for_software(self, software: str) -> None:
        self.set_software_boost_enabled(software, enabled=True)

    def set_software_boost_enabled(self, software: str, enabled: bool) -> None:
        spec = SOFTWARE_SPECS[software]
        self.ensure_running()
        results = set_performance_boost_state(
            self.window_title_re,
            spec.ai_turbo_app_name,
            enabled=enabled,
            timeout=15.0,
        )
        action = "configure" if enabled else "disable"
        assert results, f"Failed to {action} AI Turbo Engine for {software}"

    def disable_for_software(self, software: str) -> None:
        self.set_software_boost_enabled(software, enabled=False)

    def start_topmost_guard(self) -> None:
        self.ensure_running()
        if self._keeper is not None:
            return
        self._keeper = WindowTopmostKeeper(
            self.window_title_re,
            interval_seconds=self.topmost_interval_seconds,
        )
        self._keeper.start()

    def stop_topmost_guard(self) -> None:
        if self._keeper is None:
            return
        self._keeper.stop()
        self._keeper = None

    def close(self, *, wait_seconds: float = 10.0) -> None:
        self.stop_topmost_guard()
        try:
            window = connect_window(self.window_title_re, timeout=1.0)
        except UiAutomationError:
            return
        clear_window_topmost(window)
        requested = request_window_close(window)
        if not requested:
            return
        deadline = time.monotonic() + wait_seconds
        while time.monotonic() < deadline:
            try:
                connect_window(self.window_title_re, timeout=0.5)
            except UiAutomationError:
                return
            time.sleep(0.25)
        raise AssertionError(f"AI Turbo Engine did not close within {wait_seconds:.1f} seconds.")

    def run_sequence(self, software: str, *, wait_seconds: float) -> None:
        assert wait_seconds >= 0, "wait_seconds must be non-negative."
        logger.info(
            "Running standalone AI Turbo Engine sequence. software=%s wait_seconds=%.1f",
            software,
            wait_seconds,
        )
        self.start_topmost_guard()
        try:
            self.set_software_boost_enabled(software, enabled=True)
            if wait_seconds > 0:
                time.sleep(wait_seconds)
            self.set_software_boost_enabled(software, enabled=False)
        finally:
            self.close()
