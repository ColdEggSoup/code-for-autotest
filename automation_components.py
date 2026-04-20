from __future__ import annotations

from dataclasses import dataclass
import logging
import os
from pathlib import Path
import subprocess
import sys
import time

from monitoring_common import FINAL_SESSION_STATUSES
from ui_automation import (
    UiAutomationError,
    WindowTopmostKeeper,
    connect_window,
    set_performance_boost_selection,
)


REPO_ROOT = Path(__file__).resolve().parent
WHITELIST_APP_DIR = REPO_ROOT / "whitelist app"
TEST_DATA_DIR = REPO_ROOT / "test_data"
DEFAULT_AI_TURBO_TITLE_RE = r"^AI Turbo Engine$"
logger = logging.getLogger(__name__)


@dataclass(frozen=True)
class SoftwareRuntimeSpec:
    software: str
    launch_path: Path | None
    main_window_title_re: str
    ai_turbo_app_name: str
    export_suffix: str
    blend_file: Path | None = None


@dataclass(frozen=True)
class MonitorLaunchResult:
    software: str
    session_id: str
    output_path: Path
    raw_output: str


SOFTWARE_SPECS = {
    "shotcut": SoftwareRuntimeSpec(
        software="shotcut",
        launch_path=WHITELIST_APP_DIR / "shotcut.lnk",
        main_window_title_re=r".*Shotcut.*",
        ai_turbo_app_name="shotcut",
        export_suffix=".mp4",
    ),
    "kdenlive": SoftwareRuntimeSpec(
        software="kdenlive",
        launch_path=WHITELIST_APP_DIR / "Kdenlive.lnk",
        main_window_title_re=r".*Kdenlive.*",
        ai_turbo_app_name="kdenlive",
        export_suffix=".mp4",
    ),
    "shutter_encoder": SoftwareRuntimeSpec(
        software="shutter_encoder",
        launch_path=WHITELIST_APP_DIR / "Shutter Encoder.lnk",
        main_window_title_re=r".*Shutter Encoder.*",
        ai_turbo_app_name="Shutter Encoder",
        export_suffix=".mp4",
    ),
    "avidemux": SoftwareRuntimeSpec(
        software="avidemux",
        launch_path=WHITELIST_APP_DIR / "avidemux.lnk",
        main_window_title_re=r".*Avidemux.*",
        ai_turbo_app_name="avidemux",
        export_suffix=".mp4",
    ),
    "handbrake": SoftwareRuntimeSpec(
        software="handbrake",
        launch_path=WHITELIST_APP_DIR / "HandBrake.lnk",
        main_window_title_re=r".*HandBrake.*",
        ai_turbo_app_name="HandBrake",
        export_suffix=".mp4",
    ),
    "blender": SoftwareRuntimeSpec(
        software="blender",
        launch_path=None,
        main_window_title_re=r".*Blender.*",
        ai_turbo_app_name="blender",
        export_suffix=".mp4",
        blend_file=WHITELIST_APP_DIR / "blender" / "13263.blend",
    ),
}


def parse_key_value_output(raw_output: str) -> dict[str, str]:
    result: dict[str, str] = {}
    for line in raw_output.splitlines():
        if "=" not in line:
            continue
        key, value = line.split("=", 1)
        result[key.strip()] = value.strip()
    return result


def resolve_input_video(workload_name: str) -> Path:
    normalized = workload_name.strip().lower()
    assert normalized, "workload_name must not be empty."
    if "4k" in normalized:
        path = TEST_DATA_DIR / "4K.mp4"
    else:
        path = TEST_DATA_DIR / "1080.mp4"
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
        logger.info("Waiting for %s main window. timeout=%.1fs", software, timeout)
        return connect_window(spec.main_window_title_re, timeout=timeout)


class MonitorBridge:
    def __init__(self, *, repo_root: Path | None = None, python_executable: str | None = None) -> None:
        self.repo_root = (repo_root or REPO_ROOT).resolve(strict=False)
        self.python_executable = python_executable or sys.executable

    def _run_main(self, args: list[str], *, timeout: float | None = None) -> subprocess.CompletedProcess[str]:
        command = [self.python_executable, "main.py", *args]
        return subprocess.run(
            command,
            cwd=str(self.repo_root),
            capture_output=True,
            text=True,
            check=False,
            timeout=timeout,
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
        assert completed.returncode == 0, completed.stdout + completed.stderr
        payload = parse_key_value_output(completed.stdout)
        assert "session_id" in payload, completed.stdout
        resolved_output = payload.get("aggregate_output_path") or payload.get("output_path")
        assert resolved_output, completed.stdout
        return MonitorLaunchResult(
            software=software,
            session_id=payload["session_id"],
            output_path=Path(resolved_output),
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
            assert completed.returncode == 0, completed.stdout + completed.stderr
            last_payload = parse_key_value_output(completed.stdout)
            status = last_payload.get("status", "")
            if status in FINAL_SESSION_STATUSES:
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
        assert completed.returncode == 0, completed.stdout + completed.stderr
        return parse_key_value_output(completed.stdout)

    def run_blender_monitor(
        self,
        *,
        test_name: str,
        output_path: Path,
        blend_file: Path,
        blender_executable: Path | None = None,
        render_mode: str = "frame",
        frame: int = 1,
    ) -> Path:
        assert blend_file.exists(), f"Blend file does not exist: {blend_file}"
        assert output_path.parent.exists(), f"Output directory does not exist: {output_path.parent}"
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
        assert completed.returncode == 0, completed.stdout + completed.stderr
        payload = parse_key_value_output(completed.stdout)
        resolved_output = payload.get("output_path") or str(output_path)
        return Path(resolved_output)


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
        spec = SOFTWARE_SPECS[software]
        self.ensure_running()
        results = set_performance_boost_selection(self.window_title_re, spec.ai_turbo_app_name, timeout=15.0)
        assert results, f"Failed to configure AI Turbo Engine for {software}"

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
