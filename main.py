from __future__ import annotations

import argparse
import csv
import os
import re
import shutil
import subprocess
import sys
import time
from pathlib import Path

import psutil

from workspace_runtime import configure_workspace_runtime

configure_workspace_runtime()

from log_listener import LOG_PROFILES, run_log_listener
from monitoring_common import (
    FINAL_SESSION_STATUSES,
    append_results,
    default_runtime_root,
    expand_path,
    load_json,
    now_local,
    resolve_counted_output_path,
    resolve_session_paths,
    save_json,
    summarize_result_statuses,
)
from process_listener import PROCESS_PROFILES, run_process_listener
from ui_automation import UiAutomationError, minimize_window, set_performance_boost_selection
from xlsx_report_generator import generate_xlsx_report


ALL_SOFTWARE = sorted(list(PROCESS_PROFILES) + list(LOG_PROFILES) + ["blender"])
RUN_SESSION_PATTERN = re.compile(r"^(?P<base>.+?)_run(?P<index>\d+)$", re.IGNORECASE)
INVALID_NAME_CHARS_PATTERN = re.compile(r'[<>:"/\\|?*\s]+')
SCRIPT_ROOT = Path(__file__).resolve().parent
BLENDER_VISIBLE_EXIT_GRACE_SECONDS = 5.0
BLENDER_VISIBLE_POLL_SECONDS = 0.5
BLENDER_VISIBLE_FORCE_EXIT_WAIT_SECONDS = 10.0
BLENDER_VISIBLE_WINDOW_MINIMIZE_TIMEOUT_SECONDS = 30.0


def get_monitor_type(software: str) -> str:
    if software in PROCESS_PROFILES:
        return "process"
    if software in LOG_PROFILES:
        return "log"
    if software == "blender":
        return "blender"
    raise KeyError(f"Unsupported software: {software}")


def get_runtime_root(raw_value: str) -> Path:
    return expand_path(raw_value) if raw_value else default_runtime_root()


def load_session_record(runtime_root: Path, session_id: str) -> tuple[Path, dict]:
    state_path = resolve_session_paths(runtime_root, session_id)["state_path"]
    if not state_path.exists():
        raise FileNotFoundError(f"Session '{session_id}' was not found under {runtime_root}")
    return state_path, load_json(state_path)


def save_session_status(state_path: Path, record: dict, **updates) -> None:
    record.update(updates)
    save_json(state_path, record)


def bool_text(value: object) -> str:
    return "true" if bool(value) else "false"


def clean_test_name(raw_value: str) -> str:
    return raw_value.strip()


def slugify_test_name(raw_value: str) -> str:
    cleaned = clean_test_name(raw_value)
    if not cleaned:
        return ""
    return INVALID_NAME_CHARS_PATTERN.sub("_", cleaned).strip("._-")


def build_output_stem(software: str, test_name_slug: str) -> str:
    return f"{software}_{test_name_slug}" if test_name_slug else software


def default_output_base_path(software: str, test_name_slug: str) -> Path:
    return expand_path(f"results/csv/{build_output_stem(software, test_name_slug)}.csv")


def default_report_base_path(test_name_slug: str) -> Path:
    stem = f"{test_name_slug}_report" if test_name_slug else "test_report"
    return expand_path(f"results/xlsx/{stem}.xlsx")


def default_csv_root() -> Path:
    return expand_path("results/csv")


def default_blender_executable_candidates() -> tuple[Path, ...]:
    return (
        (SCRIPT_ROOT / "software" / "blender" / "blender.exe").resolve(strict=False),
        expand_path(r"C:/Program Files/Blender Foundation/Blender 4.5/blender.exe"),
        expand_path(r"C:/Program Files/Blender Foundation/Blender/blender.exe"),
    )


def default_blender_executable() -> Path:
    candidates = list(default_blender_executable_candidates())
    for candidate in candidates:
        if candidate.exists():
            return candidate
    return candidates[0]


def default_blender_listener_script() -> Path:
    return (Path(__file__).resolve().parent / "blender_listener.py").resolve(strict=False)


def resolve_blender_executable(raw_value: str) -> Path:
    return expand_path(raw_value) if raw_value else default_blender_executable()


def resolve_blend_file(raw_value: str) -> Path:
    if not raw_value:
        raise ValueError("--blend-file is required when --software blender is used.")
    blend_file = expand_path(raw_value)
    if not blend_file.exists():
        raise FileNotFoundError(f"Blend file was not found: {blend_file}")
    return blend_file


def build_blender_command(args: argparse.Namespace, record: dict) -> list[str]:
    blender_exe = resolve_blender_executable(args.blender_exe)
    if not blender_exe.exists():
        raise FileNotFoundError(f"Blender executable was not found: {blender_exe}")
    blend_file = resolve_blend_file(args.blend_file)
    listener_script = default_blender_listener_script()
    if not listener_script.exists():
        raise FileNotFoundError(f"Blender listener script was not found: {listener_script}")
    blender_ui_mode = str(getattr(args, "blender_ui_mode", "visible") or "visible")
    if blender_ui_mode not in {"visible", "headless"}:
        raise ValueError(f"Unsupported Blender UI mode: {blender_ui_mode}")

    command = [str(blender_exe)]
    if blender_ui_mode == "headless":
        command.append("-b")
    command.extend([
        str(blend_file),
        "--python",
        str(listener_script),
    ])
    if args.render_mode == "animation":
        command.append("-a")
    else:
        command.extend(["-f", str(args.frame)])
    command.extend([
        "--",
        str(record["output_path"]),
        str(record["session_id"]),
        "",
        blender_ui_mode,
    ])
    return command


def count_csv_rows(csv_path: Path) -> int:
    if not csv_path.exists() or csv_path.stat().st_size == 0:
        return 0
    with csv_path.open("r", newline="", encoding="utf-8-sig") as handle:
        reader = csv.DictReader(handle)
        return sum(1 for _ in reader)


def _import_pywinauto_desktop():
    try:
        from pywinauto import Desktop
    except ImportError as exc:
        raise UiAutomationError("pywinauto is required to control the Blender window.") from exc
    return Desktop


def find_visible_window_for_pid(process_id: int, *, timeout: float = BLENDER_VISIBLE_WINDOW_MINIMIZE_TIMEOUT_SECONDS):
    assert process_id > 0, "process_id must be positive."
    Desktop = _import_pywinauto_desktop()
    deadline = time.monotonic() + timeout
    last_titles: list[str] = []
    while time.monotonic() < deadline:
        candidates = []
        windows = Desktop(backend="uia").windows()
        for window in windows:
            try:
                pid = int(getattr(getattr(window, "element_info", None), "process_id", 0) or 0)
            except Exception:
                continue
            if pid != process_id:
                continue
            try:
                if not window.is_visible():
                    continue
            except Exception:
                pass
            try:
                title = (window.window_text() or "").strip()
            except Exception:
                title = ""
            if title:
                last_titles.append(title)
            try:
                rect = window.rectangle()
                area = max(0, rect.width() * rect.height())
            except Exception:
                area = 0
            candidates.append((-area, title.casefold(), window))
        if candidates:
            candidates.sort(key=lambda item: (item[0], item[1]))
            return candidates[0][2]
        time.sleep(0.25)
    raise UiAutomationError(
        f"Could not find a visible top-level window for Blender pid={process_id}. Seen titles: {last_titles[:10]}"
    )


def minimize_visible_blender_window(process_id: int, *, timeout: float = BLENDER_VISIBLE_WINDOW_MINIMIZE_TIMEOUT_SECONDS) -> None:
    window = find_visible_window_for_pid(process_id, timeout=timeout)
    minimize_window(window)


def terminate_process_tree(process: subprocess.Popen[object], *, wait_seconds: float = BLENDER_VISIBLE_FORCE_EXIT_WAIT_SECONDS) -> int | None:
    try:
        subprocess.run(
            ["taskkill", "/PID", str(process.pid), "/T", "/F"],
            capture_output=True,
            text=True,
            check=False,
            timeout=wait_seconds,
        )
    except Exception:
        pass

    try:
        process.kill()
    except Exception:
        pass

    deadline = time.monotonic() + wait_seconds
    while time.monotonic() < deadline:
        returncode = process.poll()
        if returncode is not None:
            return returncode
        time.sleep(0.1)
    return process.poll()


def run_blender_session(args: argparse.Namespace, state_path: Path, record: dict) -> int:
    output_path = Path(record["output_path"])
    command = build_blender_command(args, record)
    blender_ui_mode = getattr(args, "blender_ui_mode", "visible") or "visible"
    save_session_status(
        state_path,
        record,
        status="running",
        pid=None,
        started_at=now_local().isoformat(timespec="milliseconds"),
        progress_state="waiting_for_render_activity",
        progress_message="Launching Blender render.",
        progress_updated_at=now_local().isoformat(timespec="milliseconds"),
        current_source_path=args.blend_file,
    )
    print(f"session_id={record['session_id']}")
    print(f"software={record['software']}")
    if record.get("test_name"):
        print(f"test_name={record['test_name']}")
    print(f"output_path={record['output_path']}")
    print(f"blender_exe={resolve_blender_executable(args.blender_exe)}")
    print(f"blend_file={resolve_blend_file(args.blend_file)}")
    print(f"blender_ui_mode={blender_ui_mode}")
    print("Launching Blender...", flush=True)
    process: subprocess.Popen[object] | None = None
    forced_exit_after_capture = False
    returncode: int | None = None
    try:
        process = subprocess.Popen(command, cwd=str(Path.cwd()))
        save_session_status(state_path, record, pid=process.pid)
        if blender_ui_mode == "visible":
            try:
                minimize_visible_blender_window(process.pid)
                print("Minimized the visible Blender window.", flush=True)
            except Exception as exc:
                print(f"Could not minimize the visible Blender window automatically: {exc}", flush=True)
        while True:
            returncode = process.poll()
            row_count = count_csv_rows(output_path)
            if returncode is not None:
                break
            if blender_ui_mode == "visible" and row_count > 0:
                print(
                    "Detected completed Blender render output. Waiting briefly for the visible Blender window to exit.",
                    flush=True,
                )
                try:
                    returncode = process.wait(timeout=BLENDER_VISIBLE_EXIT_GRACE_SECONDS)
                except subprocess.TimeoutExpired:
                    print(
                        "Visible Blender did not exit after render completion. Terminating the Blender process.",
                        flush=True,
                    )
                    forced_exit_after_capture = True
                    returncode = terminate_process_tree(process)
                break
            time.sleep(BLENDER_VISIBLE_POLL_SECONDS)
    except KeyboardInterrupt:
        if process is not None and process.poll() is None:
            terminate_process_tree(process)
        save_session_status(
            state_path,
            record,
            status="failed",
            ended_at=now_local().isoformat(timespec="milliseconds"),
            error="KeyboardInterrupt: Blender monitoring was interrupted by the user.",
            progress_state="failed",
            progress_message="Blender monitoring was interrupted by the user.",
            progress_updated_at=now_local().isoformat(timespec="milliseconds"),
        )
        print("Blender monitoring interrupted.", flush=True)
        return 130

    row_count = count_csv_rows(output_path)
    successful_completion = row_count > 0 and ((returncode or 0) == 0 or forced_exit_after_capture)
    final_status = "completed" if successful_completion else "completed_with_warnings"
    result_statuses = ["ok"] if row_count > 0 else []
    save_session_status(
        state_path,
        record,
        status=final_status,
        ended_at=now_local().isoformat(timespec="milliseconds"),
        row_count=row_count,
        result_statuses=result_statuses,
        error="" if successful_completion else f"Blender exited with code {returncode}",
        xlsx_output_path="",
        xlsx_status="not_generated",
        xlsx_csv_count=0,
        xlsx_row_count=0,
        xlsx_error="",
        progress_state="completed" if row_count > 0 else "completed_with_warnings",
        progress_message=(
            "Blender render finished and wrote csv result files."
            if row_count > 0
            else "Blender render finished, but no completed csv result row was found."
        ),
        progress_updated_at=now_local().isoformat(timespec="milliseconds"),
    )
    print(f"Blender finished. Wrote {row_count} row(s) to {output_path}", flush=True)
    return 0 if successful_completion else 2


def guess_initial_log_source(software: str, options: dict) -> str:
    profile = LOG_PROFILES[software]
    if options.get("log_path"):
        return str(expand_path(options["log_path"]))
    if options.get("log_dir"):
        return str(expand_path(options["log_dir"]))
    if profile.log_file_candidates:
        return str(expand_path(profile.log_file_candidates[0]))
    if profile.log_dir_candidates:
        return str(expand_path(profile.log_dir_candidates[0]))
    return ""


def guess_initial_process_target(software: str) -> str:
    profile = PROCESS_PROFILES[software]
    return ", ".join(profile.worker_process_names)


def default_log_override(software: str) -> tuple[str, str]:
    profile = LOG_PROFILES.get(software)
    if profile is None:
        return "", ""
    if profile.log_file_candidates:
        return str(expand_path(profile.log_file_candidates[0])), ""
    if profile.log_dir_candidates:
        return "", str(expand_path(profile.log_dir_candidates[0]))
    return "", ""


def resolve_log_override(software: str, raw_log_path: str) -> tuple[str, str]:
    profile = LOG_PROFILES.get(software)
    if profile is None:
        return "", ""
    if raw_log_path:
        resolved_value = str(expand_path(raw_log_path))
        if profile.log_file_candidates:
            return resolved_value, ""
        if profile.log_dir_candidates:
            return "", resolved_value
    return default_log_override(software)


def session_exists(runtime_root: Path, session_id: str) -> bool:
    return resolve_session_paths(runtime_root, session_id)["state_path"].exists()


def resolve_session_id(runtime_root: Path, preferred_session_id: str, requested_session_id: str) -> tuple[str, str]:
    if requested_session_id:
        if not session_exists(runtime_root, requested_session_id):
            return requested_session_id, ""
        next_index = 2
        while True:
            candidate = f"{requested_session_id}_session{next_index:02d}"
            if not session_exists(runtime_root, candidate):
                return candidate, requested_session_id
            next_index += 1
    if not session_exists(runtime_root, preferred_session_id):
        return preferred_session_id, ""
    next_index = 2
    while True:
        candidate = f"{preferred_session_id}_session{next_index:02d}"
        if not session_exists(runtime_root, candidate):
            return candidate, ""
        next_index += 1


def is_process_alive(pid: int | None) -> bool:
    if not pid:
        return False
    try:
        process = psutil.Process(pid)
        return process.is_running() and process.status() != psutil.STATUS_ZOMBIE
    except (psutil.Error, OSError):
        return False


def print_session_summary(record: dict) -> None:
    print(f"session_id={record['session_id']}")
    if record.get("requested_session_id"):
        print(f"requested_session_id={record['requested_session_id']}")
    print(f"software={record['software']}")
    if record.get("test_name"):
        print(f"test_name={record['test_name']}")
    print(f"monitor_type={record['monitor_type']}")
    print(f"status={record['status']}")
    print(f"pid={record.get('pid', '')}")
    print(f"state_path={record['state_path']}")
    print(f"stop_path={record['stop_path']}")
    print(f"session_output_path={record['session_output_path']}")
    print(f"aggregate_output_path={record['output_path']}")
    if record.get("output_base_path"):
        print(f"output_base_path={record['output_base_path']}")
    if record.get("output_run_index"):
        print(f"output_run_index={record['output_run_index']}")
    if record.get("xlsx_status"):
        print(f"xlsx_status={record['xlsx_status']}")
    if record.get("xlsx_csv_count") is not None:
        print(f"xlsx_csv_count={record['xlsx_csv_count']}")
    if record.get("xlsx_row_count") is not None:
        print(f"xlsx_row_count={record['xlsx_row_count']}")
    if record.get("xlsx_error"):
        print(f"xlsx_error={record['xlsx_error']}")
    if record.get("worker_stdout_path"):
        print(f"worker_stdout_path={record['worker_stdout_path']}")
    if record.get("worker_stderr_path"):
        print(f"worker_stderr_path={record['worker_stderr_path']}")
    if record.get("progress_state"):
        print(f"progress_state={record['progress_state']}")
    if record.get("progress_message"):
        print(f"progress_message={record['progress_message']}")
    if "capture_detected" in record:
        print(f"capture_detected={bool_text(record.get('capture_detected'))}")
    if "capture_complete" in record:
        print(f"capture_complete={bool_text(record.get('capture_complete'))}")
    if record.get("current_source_path"):
        print(f"current_source_path={record['current_source_path']}")
    if record.get("captured_line_count") is not None:
        print(f"captured_line_count={record['captured_line_count']}")
    if record.get("last_log_activity_at"):
        print(f"last_log_activity_at={record['last_log_activity_at']}")
    if record.get("progress_updated_at"):
        print(f"progress_updated_at={record['progress_updated_at']}")
    if record.get("preview_status"):
        print(f"preview_status={record['preview_status']}")
    if record.get("preview_started_at"):
        print(f"preview_started_at={record['preview_started_at']}")
    if record.get("preview_ended_at"):
        print(f"preview_ended_at={record['preview_ended_at']}")
    if record.get("preview_duration_seconds"):
        print(f"preview_duration_seconds={record['preview_duration_seconds']}")
    if record.get("preview_evidence"):
        print(f"preview_evidence={record['preview_evidence']}")
    if record.get("result_statuses"):
        print(f"result_statuses={','.join(record['result_statuses'])}")
    if record.get("row_count") is not None:
        print(f"row_count={record['row_count']}")
    if record.get("error"):
        print(f"error={record['error']}")


def create_session_record(args: argparse.Namespace) -> tuple[Path, dict, dict[str, Path]]:
    state_path, record = build_session_record(args)
    runtime_root = Path(record["runtime_root"])
    paths = resolve_session_paths(runtime_root, record["session_id"])
    state_file = Path(record["state_path"])
    if state_file.exists():
        raise FileExistsError(f"Session already exists: {record['session_id']}")
    paths["session_dir"].mkdir(parents=True, exist_ok=True)
    save_json(state_path, record)
    return state_path, record, paths


def progress_state_hint(progress_state: str, progress_message: str) -> str:
    hints = {
        "waiting_for_log_activity": "Monitoring. Waiting for log activity.",
        "waiting_for_worker_activity": "Monitoring. Waiting for target process activity.",
        "waiting_for_render_activity": "Monitoring. Waiting for Blender render activity.",
        "log_activity_detected": "Detected log activity. Monitoring continues.",
        "partial_data_captured": "Partial data captured. Waiting for completion.",
        "active_worker_captured": "Target worker process captured. Monitoring continues.",
        "active_render_captured": "Blender render captured. Monitoring continues.",
        "complete_data_captured": "Complete data captured. Finalizing.",
        "finalizing_with_partial_data": "Finalizing from currently captured data.",
        "stop_requested": "Stop requested. Finalizing.",
        "completed": "Monitoring completed.",
        "completed_with_warnings": "Monitoring completed with warnings.",
        "failed": "Monitoring failed.",
        "stale": "Monitor process exited.",
    }
    return hints.get(progress_state, progress_message or progress_state)


def apply_test_name(rows: list[dict[str, str]], test_name: str) -> list[dict[str, str]]:
    normalized_test_name = clean_test_name(test_name)
    if not normalized_test_name:
        return rows
    for row in rows:
        row["test_name"] = normalized_test_name
    return rows


def execute_session(state_path: Path, record: dict) -> tuple[int, list[dict[str, str]], str]:
    config = record["config"]
    monitor_type = record["monitor_type"]
    stop_path = Path(record["stop_path"])
    session_output_path = Path(record["session_output_path"])
    aggregate_output_path = Path(record["output_path"])
    save_session_status(
        state_path,
        record,
        status="running",
        pid=os.getpid(),
        started_at=record["started_at"] or now_local().isoformat(timespec="milliseconds"),
    )
    last_progress_signature: list[tuple[object, ...] | None] = [None]

    def report_progress(payload: dict[str, object]) -> None:
        updates = {
            "progress_state": str(payload.get("progress_state", "") or ""),
            "progress_message": str(payload.get("progress_message", "") or ""),
            "progress_updated_at": now_local().isoformat(timespec="milliseconds"),
            "capture_detected": bool(payload.get("capture_detected", False)),
            "capture_complete": bool(payload.get("capture_complete", False)),
            "current_source_path": str(payload.get("current_source_path", "") or ""),
            "captured_line_count": int(payload.get("captured_line_count", 0) or 0),
            "last_log_activity_at": str(payload.get("last_log_activity_at", "") or ""),
            "preview_status": str(payload.get("preview_status", "") or ""),
            "preview_started_at": str(payload.get("preview_started_at", "") or ""),
            "preview_ended_at": str(payload.get("preview_ended_at", "") or ""),
            "preview_duration_seconds": str(payload.get("preview_duration_seconds", "") or ""),
            "preview_evidence": str(payload.get("preview_evidence", "") or ""),
        }
        save_session_status(state_path, record, **updates)
        log_signature = (
            updates["progress_state"],
            updates["progress_message"],
            updates["capture_detected"],
            updates["capture_complete"],
            updates["preview_status"],
        )
        if log_signature != last_progress_signature[0]:
            print(
                f"[{now_local().isoformat(timespec='seconds')}] "
                f"{progress_state_hint(updates['progress_state'], updates['progress_message'])}",
                flush=True,
            )
            last_progress_signature[0] = log_signature

    try:
        if monitor_type == "process":
            rows = run_process_listener(config, stop_path, progress_callback=report_progress)
        elif monitor_type == "log":
            rows = run_log_listener(config, stop_path, progress_callback=report_progress)
        else:
            raise RuntimeError("Blender worker must run inside Blender, not via standard Python.")
        rows = apply_test_name(rows, str(record.get("test_name", "")))
        append_results(session_output_path, aggregate_output_path, rows)
        result_statuses = summarize_result_statuses(rows)
        final_status = "completed" if result_statuses == ["ok"] else "completed_with_warnings"
        save_session_status(
            state_path,
            record,
            status=final_status,
            ended_at=now_local().isoformat(timespec="milliseconds"),
            row_count=len(rows),
            result_statuses=result_statuses,
            error="",
            xlsx_output_path="",
            xlsx_status="not_generated",
            xlsx_csv_count=0,
            xlsx_row_count=0,
            xlsx_error="",
            progress_state="completed" if final_status == "completed" else "completed_with_warnings",
            progress_message=(
                "Background monitor finished and wrote csv result files."
                if final_status == "completed"
                else "Background monitor finished with warnings after writing the csv result files."
            ),
            progress_updated_at=now_local().isoformat(timespec="milliseconds"),
        )
        return (0 if final_status == "completed" else 2), rows, str(aggregate_output_path)
    except Exception as exc:
        save_session_status(
            state_path,
            record,
            status="failed",
            ended_at=now_local().isoformat(timespec="milliseconds"),
            error=f"{type(exc).__name__}: {exc}",
            progress_state="failed",
            progress_message="Background monitor failed before it could finalize the result rows.",
            progress_updated_at=now_local().isoformat(timespec="milliseconds"),
        )
        raise


def build_session_record(args: argparse.Namespace) -> tuple[Path, dict]:
    runtime_root = get_runtime_root(args.runtime_root)
    test_name = clean_test_name(args.name)
    test_name_slug = slugify_test_name(test_name)
    if args.output:
        requested_output_path = expand_path(args.output)
        output_path, output_run_index = resolve_counted_output_path(requested_output_path)
    else:
        requested_output_path = default_output_base_path(args.software, test_name_slug)
        output_path, output_run_index = resolve_counted_output_path(requested_output_path)
    preferred_session_id = output_path.stem
    session_id, requested_session_id = resolve_session_id(runtime_root, preferred_session_id, args.session_id)
    paths = resolve_session_paths(runtime_root, session_id)
    monitor_type = get_monitor_type(args.software)
    resolved_log_path, resolved_log_dir = resolve_log_override(args.software, args.log_path)
    options = {
        "poll_interval": args.poll_interval,
        "idle_seconds": args.idle_seconds,
        "max_runtime_seconds": args.max_runtime,
        "stop_grace_seconds": args.stop_grace_seconds,
        "include_existing_workers": args.include_existing_workers,
        "allow_detached_worker": args.allow_detached_worker,
        "allow_activity_log": args.allow_activity_log,
        "log_path": resolved_log_path,
        "log_dir": resolved_log_dir,
    }
    config = {
        "software": args.software,
        "session_id": session_id,
        "test_name": test_name,
        "monitor_type": monitor_type,
        "options": options,
    }
    record = {
        "session_id": session_id,
        "requested_session_id": requested_session_id,
        "software": args.software,
        "test_name": test_name,
        "monitor_type": monitor_type,
        "status": "starting",
        "pid": None,
        "created_at": now_local().isoformat(timespec="milliseconds"),
        "started_at": "",
        "ended_at": "",
        "stop_requested_at": "",
        "row_count": 0,
        "result_statuses": [],
        "error": "",
        "runtime_root": str(runtime_root),
        "state_path": str(paths["state_path"]),
        "stop_path": str(paths["stop_path"]),
        "session_output_path": str(paths["session_output_path"]),
        "output_path": str(output_path),
        "xlsx_output_path": "",
        "output_base_path": str(requested_output_path),
        "output_run_index": output_run_index,
        "xlsx_status": "",
        "xlsx_csv_count": 0,
        "xlsx_row_count": 0,
        "xlsx_error": "",
        "worker_stdout_path": str(paths["session_dir"] / "worker.stdout.log"),
        "worker_stderr_path": str(paths["session_dir"] / "worker.stderr.log"),
        "progress_state": "",
        "progress_message": "",
        "progress_updated_at": "",
        "capture_detected": False,
        "capture_complete": False,
        "current_source_path": "",
        "captured_line_count": 0,
        "last_log_activity_at": "",
        "preview_status": "",
        "preview_started_at": "",
        "preview_ended_at": "",
        "preview_duration_seconds": "",
        "preview_evidence": "",
        "config": config,
    }
    return paths["state_path"], record


def build_worker_command(runtime_root: Path, session_id: str) -> list[str]:
    return [
        sys.executable,
        str(Path(__file__).resolve()),
        "worker",
        "--runtime-root",
        str(runtime_root),
        "--session-id",
        session_id,
    ]


def start_session(args: argparse.Namespace) -> int:
    try:
        state_path, record, _ = create_session_record(args)
    except FileExistsError as exc:
        print(str(exc))
        return 2

    if args.software == "blender":
        return run_blender_session(args, state_path, record)

    save_session_status(
        state_path,
        record,
        progress_state=(
            "waiting_for_log_activity"
            if record["monitor_type"] == "log"
            else "waiting_for_worker_activity"
        ),
        progress_message=(
            "Background log listener started. Waiting for new log content."
            if record["monitor_type"] == "log"
            else "Background listener started. Waiting for matching worker activity."
        ),
        progress_updated_at=now_local().isoformat(timespec="milliseconds"),
        current_source_path=(
            guess_initial_log_source(record["software"], record["config"]["options"])
            if record["monitor_type"] == "log"
            else guess_initial_process_target(record["software"])
        ),
    )
    print(f"session_id={record['session_id']}")
    print(f"software={record['software']}")
    if record.get("test_name"):
        print(f"test_name={record['test_name']}")
    print(f"output_path={record['output_path']}")
    print("Monitoring...", flush=True)
    try:
        exit_code, rows, output_path = execute_session(state_path, record)
    except KeyboardInterrupt:
        save_session_status(
            state_path,
            record,
            status="failed",
            ended_at=now_local().isoformat(timespec="milliseconds"),
            error="KeyboardInterrupt: Monitoring was interrupted by the user.",
            progress_state="failed",
            progress_message="Monitoring was interrupted by the user.",
            progress_updated_at=now_local().isoformat(timespec="milliseconds"),
        )
        print("Monitoring interrupted.", flush=True)
        return 130
    print(f"Monitoring completed. Wrote {len(rows)} row(s) to {output_path}", flush=True)
    return exit_code


def start_session_background(args: argparse.Namespace) -> int:
    if args.software == "blender":
        print("Blender background start is not supported from main.py.")
        print("Use: python main.py start --software blender --blend-file ... --blender-exe ...")
        return 2
    try:
        state_path, record, paths = create_session_record(args)
    except FileExistsError as exc:
        print(str(exc))
        return 2
    runtime_root = Path(record["runtime_root"])
    stdout_path = paths["session_dir"] / "worker.stdout.log"
    stderr_path = paths["session_dir"] / "worker.stderr.log"
    creationflags = 0
    creationflags |= getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0)
    creationflags |= getattr(subprocess, "DETACHED_PROCESS", 0)
    creationflags |= getattr(subprocess, "CREATE_NO_WINDOW", 0)
    with stdout_path.open("a", encoding="utf-8") as stdout_handle, stderr_path.open("a", encoding="utf-8") as stderr_handle:
        worker = subprocess.Popen(
            build_worker_command(runtime_root, record["session_id"]),
            cwd=str(Path.cwd()),
            stdout=stdout_handle,
            stderr=stderr_handle,
            creationflags=creationflags,
        )
    save_session_status(
        state_path,
        record,
        status="running",
        pid=worker.pid,
        started_at=now_local().isoformat(timespec="milliseconds"),
        progress_state=(
            "waiting_for_log_activity"
            if record["monitor_type"] == "log"
            else "waiting_for_worker_activity"
        ),
        progress_message=(
            "Background log listener started. Waiting for new log content."
            if record["monitor_type"] == "log"
            else "Background listener started. Waiting for matching worker activity."
        ),
        progress_updated_at=now_local().isoformat(timespec="milliseconds"),
        current_source_path=(
            guess_initial_log_source(record["software"], record["config"]["options"])
            if record["monitor_type"] == "log"
            else guess_initial_process_target(record["software"])
        ),
    )
    print_session_summary(record)
    print(f"status_hint=python main.py status --session-id {record['session_id']}")
    return 0


def status_session(args: argparse.Namespace) -> int:
    runtime_root = get_runtime_root(args.runtime_root)
    state_path, record = load_session_record(runtime_root, args.session_id)
    if record["status"] not in FINAL_SESSION_STATUSES and not is_process_alive(record.get("pid")):
        save_session_status(
            state_path,
            record,
            status="stale",
            error="Worker process is no longer running.",
            progress_state="stale",
            progress_message="Worker process is no longer running.",
            progress_updated_at=now_local().isoformat(timespec="milliseconds"),
        )
    print_session_summary(record)
    return 0


def stop_session(args: argparse.Namespace) -> int:
    runtime_root = get_runtime_root(args.runtime_root)
    state_path, record = load_session_record(runtime_root, args.session_id)
    stop_path = Path(record["stop_path"])
    stop_path.parent.mkdir(parents=True, exist_ok=True)
    stop_path.touch(exist_ok=True)
    if not record.get("stop_requested_at"):
        save_session_status(state_path, record, stop_requested_at=now_local().isoformat(timespec="milliseconds"))
    deadline = time.monotonic() + args.wait_seconds
    while time.monotonic() < deadline:
        state_path, record = load_session_record(runtime_root, args.session_id)
        if record["status"] in FINAL_SESSION_STATUSES:
            break
        if not is_process_alive(record.get("pid")) and record["status"] not in FINAL_SESSION_STATUSES:
            save_session_status(
                state_path,
                record,
                status="stale",
                error="Worker process stopped before finalizing.",
                progress_state="stale",
                progress_message="Worker process stopped before finalizing.",
                progress_updated_at=now_local().isoformat(timespec="milliseconds"),
            )
            break
        time.sleep(0.5)
    _, record = load_session_record(runtime_root, args.session_id)
    print_session_summary(record)
    return 0 if record["status"] in FINAL_SESSION_STATUSES else 2


def list_software(_: argparse.Namespace) -> int:
    print("Available software:")
    for software in ALL_SOFTWARE:
        print(f"  - {software} ({get_monitor_type(software)})")
    return 0

def set_ui_boost(args: argparse.Namespace) -> int:
    try:
        results = set_performance_boost_selection(
            window_title_re=args.window_title,
            app_name=args.app_name,
            timeout=args.timeout,
        )
    except UiAutomationError as exc:
        print(f"ui_automation_error={exc}")
        return 2
    except Exception as exc:
        print(f"ui_automation_error={type(exc).__name__}: {exc}")
        return 2
    for result in results:
        print(f"target={result.target}")
        print(f"state={result.state}")
        print(f"changed={bool_text(result.changed)}")
    return 0



def run_worker(args: argparse.Namespace) -> int:
    runtime_root = get_runtime_root(args.runtime_root)
    state_path, record = load_session_record(runtime_root, args.session_id)
    exit_code, _, _ = execute_session(state_path, record)
    return exit_code


def csv_matches_test_name(csv_path: Path, test_name_slug: str) -> bool:
    if not test_name_slug:
        return True
    match = RUN_SESSION_PATTERN.fullmatch(csv_path.stem)
    base_stem = match.group("base") if match else csv_path.stem
    return base_stem.endswith(f"_{test_name_slug}")


def csv_contains_test_name(csv_path: Path, test_name: str) -> bool:
    if not test_name:
        return True
    try:
        with csv_path.open("r", newline="", encoding="utf-8-sig") as handle:
            reader = csv.DictReader(handle)
            for row in reader:
                if clean_test_name(row.get("test_name", "")) == test_name:
                    return True
    except OSError:
        return False
    return False


def collect_pending_csv_paths(csv_root: Path, test_name: str, test_name_slug: str) -> list[Path]:
    if not csv_root.exists():
        return []
    csv_paths = sorted(path for path in csv_root.glob("*.csv") if path.is_file())
    filtered_paths: list[Path] = []
    for path in csv_paths:
        if csv_matches_test_name(path, test_name_slug):
            filtered_paths.append(path)
            continue
        if test_name and csv_contains_test_name(path, test_name):
            filtered_paths.append(path)
    return filtered_paths


def unique_destination_path(target_dir: Path, file_name: str) -> Path:
    candidate = target_dir / file_name
    if not candidate.exists():
        return candidate
    stem = Path(file_name).stem
    suffix = Path(file_name).suffix
    next_index = 2
    while True:
        candidate = target_dir / f"{stem}_{next_index:02d}{suffix}"
        if not candidate.exists():
            return candidate
        next_index += 1


def archive_csv_paths(csv_paths: list[Path], report_output_path: Path) -> tuple[Path, list[Path]]:
    archive_dir = report_output_path.parent / report_output_path.stem
    archive_dir.mkdir(parents=True, exist_ok=True)
    archived_paths: list[Path] = []
    for csv_path in csv_paths:
        destination = unique_destination_path(archive_dir, csv_path.name)
        shutil.move(str(csv_path), str(destination))
        archived_paths.append(destination)
    return archive_dir, archived_paths


def report_results(args: argparse.Namespace) -> int:
    test_name = clean_test_name(args.name)
    test_name_slug = slugify_test_name(test_name)
    csv_root = expand_path(args.input_dir) if args.input_dir else default_csv_root()
    csv_paths = collect_pending_csv_paths(csv_root, test_name, test_name_slug)
    if not csv_paths:
        filter_text = f" for name '{test_name}'" if test_name else ""
        print(f"No pending csv files were found in {csv_root}{filter_text}.")
        return 2
    if args.output:
        requested_output_path = expand_path(args.output)
        output_path, _ = resolve_counted_output_path(requested_output_path)
    else:
        requested_output_path = default_report_base_path(test_name_slug)
        output_path, _ = resolve_counted_output_path(requested_output_path)
    csv_count, row_count, generated_output_path = generate_xlsx_report(
        csv_paths,
        output_path,
        default_test_case=test_name,
    )
    archive_dir, archived_paths = archive_csv_paths(csv_paths, generated_output_path)
    print(f"xlsx_output_path={generated_output_path}")
    print(f"archived_csv_dir={archive_dir}")
    print(f"archived_csv_count={len(archived_paths)}")
    print(f"included_csv_count={csv_count}")
    print(f"included_row_count={row_count}")
    return 0


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Background monitor controller for video software.")
    subparsers = parser.add_subparsers(dest="command", required=True)
    list_parser = subparsers.add_parser("list-software", help="List supported software.")
    list_parser.set_defaults(handler=list_software)
    shared_runtime = argparse.ArgumentParser(add_help=False)
    shared_runtime.add_argument("--runtime-root", default="", help=argparse.SUPPRESS)
    start_parser = subparsers.add_parser("start", parents=[shared_runtime], help="Start a foreground monitor session.")
    start_parser.add_argument("--software", required=True, choices=ALL_SOFTWARE, help="Target software.")
    start_parser.add_argument("--name", default="", help="Optional test name. Used in default csv naming and later xlsx test_case.")
    start_parser.add_argument("--session-id", default="", help="Optional session id.")
    start_parser.add_argument("--output", default="", help="Optional csv output path. If omitted, uses results/csv/<software>_<name>_runNNN.csv.")
    start_parser.add_argument("--log-path", default="", help="Optional log path override for avidemux or handbrake.")
    start_parser.add_argument("--blender-exe", default="", help="Optional Blender executable path. Only used when --software blender.")
    start_parser.add_argument("--blend-file", default="", help="Blend file path. Required when --software blender.")
    start_parser.add_argument(
        "--blender-ui-mode",
        choices=("visible", "headless"),
        default="visible",
        help="Whether Blender should render with a visible UI window or in headless background mode.",
    )
    start_parser.add_argument("--render-mode", choices=("frame", "animation"), default="frame", help="Blender render mode. Only used when --software blender.")
    start_parser.add_argument("--frame", type=int, default=1, help="Blender frame number when --render-mode frame is used.")
    start_parser.add_argument("--poll-interval", type=float, default=0.5, help=argparse.SUPPRESS)
    start_parser.add_argument("--idle-seconds", type=float, default=3.0, help=argparse.SUPPRESS)
    start_parser.add_argument("--max-runtime", type=float, default=7200, help=argparse.SUPPRESS)
    start_parser.add_argument("--stop-grace-seconds", type=float, default=10.0, help=argparse.SUPPRESS)
    start_parser.add_argument("--include-existing-workers", action="store_true", help=argparse.SUPPRESS)
    start_parser.add_argument("--allow-detached-worker", action="store_true", help=argparse.SUPPRESS)
    start_parser.add_argument("--allow-activity-log", action="store_true", help=argparse.SUPPRESS)
    start_parser.set_defaults(handler=start_session)
    start_bg_parser = subparsers.add_parser("start-bg", parents=[shared_runtime], help="Start a background monitor session.")
    start_bg_parser.add_argument("--software", required=True, choices=ALL_SOFTWARE, help="Target software.")
    start_bg_parser.add_argument("--name", default="", help="Optional test name. Used in default csv naming and later xlsx test_case.")
    start_bg_parser.add_argument("--session-id", default="", help="Optional session id.")
    start_bg_parser.add_argument("--output", default="", help="Optional csv output path. If omitted, uses results/csv/<software>_<name>_runNNN.csv.")
    start_bg_parser.add_argument("--log-path", default="", help="Optional log path override for avidemux or handbrake.")
    start_bg_parser.add_argument("--blender-exe", default="", help=argparse.SUPPRESS)
    start_bg_parser.add_argument("--blend-file", default="", help=argparse.SUPPRESS)
    start_bg_parser.add_argument("--blender-ui-mode", choices=("visible", "headless"), default="visible", help=argparse.SUPPRESS)
    start_bg_parser.add_argument("--render-mode", choices=("frame", "animation"), default="frame", help=argparse.SUPPRESS)
    start_bg_parser.add_argument("--frame", type=int, default=1, help=argparse.SUPPRESS)
    start_bg_parser.add_argument("--poll-interval", type=float, default=0.5, help=argparse.SUPPRESS)
    start_bg_parser.add_argument("--idle-seconds", type=float, default=3.0, help=argparse.SUPPRESS)
    start_bg_parser.add_argument("--max-runtime", type=float, default=7200, help=argparse.SUPPRESS)
    start_bg_parser.add_argument("--stop-grace-seconds", type=float, default=10.0, help=argparse.SUPPRESS)
    start_bg_parser.add_argument("--include-existing-workers", action="store_true", help=argparse.SUPPRESS)
    start_bg_parser.add_argument("--allow-detached-worker", action="store_true", help=argparse.SUPPRESS)
    start_bg_parser.add_argument("--allow-activity-log", action="store_true", help=argparse.SUPPRESS)
    start_bg_parser.set_defaults(handler=start_session_background)
    report_parser = subparsers.add_parser("report", help="Generate one xlsx report from pending csv files and archive those csv files.")
    report_parser.add_argument("--name", default="", help="Optional test name filter. When provided, only matching csv files are included.")
    report_parser.add_argument("--input-dir", default="results/csv", help="Directory containing pending csv files. Default: results/csv")
    report_parser.add_argument("--output", default="", help="Optional xlsx output path. If omitted, uses results/xlsx/<name>_report_runNNN.xlsx.")
    report_parser.set_defaults(handler=report_results)
    status_parser = subparsers.add_parser("status", parents=[shared_runtime], help="Show one session status.")
    status_parser.add_argument("--session-id", required=True, help="Session id.")
    status_parser.set_defaults(handler=status_session)
    stop_parser = subparsers.add_parser("stop", parents=[shared_runtime], help="Stop one session and wait for completion.")
    stop_parser.add_argument("--session-id", required=True, help="Session id.")
    stop_parser.add_argument("--wait-seconds", type=float, default=15.0, help="How long to wait for finalization.")
    stop_parser.set_defaults(handler=stop_session)
    ui_boost_parser = subparsers.add_parser("ui-boost", help="Enable Performance Boost and check one app in the target Windows UI.")
    ui_boost_parser.add_argument("--window-title", default="^AI Turbo Engine$", help="Regular expression used to connect to the target top-level window.")
    ui_boost_parser.add_argument("--app-name", required=True, help="App label shown in the Performance Boost list, for example avidemux.")
    ui_boost_parser.add_argument("--timeout", type=float, default=15.0, help="Seconds to wait for the target window.")
    ui_boost_parser.set_defaults(handler=set_ui_boost)
    worker_parser = subparsers.add_parser("worker", parents=[shared_runtime], help=argparse.SUPPRESS)
    worker_parser.add_argument("--session-id", required=True, help="Session id.")
    worker_parser.set_defaults(handler=run_worker)
    return parser


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    return args.handler(args)


if __name__ == "__main__":
    raise SystemExit(main())

