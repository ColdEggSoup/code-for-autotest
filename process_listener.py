from __future__ import annotations

import time
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Callable, Iterable, Optional

import psutil

from monitoring_common import (
    LOCAL_TZ,
    calculate_duration_seconds,
    format_seconds,
    isoformat_or_empty,
    make_result_row,
    now_local,
)


@dataclass(frozen=True)
class ProcessProfile:
    software: str
    main_process_names: tuple[str, ...]
    worker_process_names: tuple[str, ...]
    require_parent_match: bool = True
    aggregation_mode: str = "session_span"
    timing_mode: str = "process_lifetime"
    require_activity_for_capture: bool = False
    min_capture_seconds: float = 0.0
    notes: str = ""


ProgressCallback = Optional[Callable[[dict[str, object]], None]]


PROCESS_PROFILES = {
    "shotcut": ProcessProfile(
        software="shotcut",
        main_process_names=("shotcut.exe",),
        worker_process_names=("melt.exe",),
        aggregation_mode="session_span",
        timing_mode="process_lifetime",
        require_activity_for_capture=False,
        min_capture_seconds=0.0,
        notes="Measure only melt.exe instances started by Shotcut to avoid unrelated MLT jobs.",
    ),
    "kdenlive": ProcessProfile(
        software="kdenlive",
        main_process_names=("kdenlive.exe",),
        worker_process_names=("melt.exe",),
        aggregation_mode="session_span",
        timing_mode="process_lifetime",
        require_activity_for_capture=False,
        min_capture_seconds=0.0,
        notes="Measure only melt.exe instances started by Kdenlive to avoid unrelated MLT jobs.",
    ),
    "shutter_encoder": ProcessProfile(
        software="shutter_encoder",
        main_process_names=("shutter encoder.exe", "shutterencoder.exe", "javaw.exe", "java.exe"),
        worker_process_names=("ffmpeg.exe",),
        aggregation_mode="longest_worker",
        timing_mode="write_activity_span",
        require_activity_for_capture=True,
        min_capture_seconds=3.0,
        notes="Measure only ffmpeg.exe instances started by Shutter Encoder (often launched from javaw.exe) to avoid unrelated transcodes.",
    ),
}


def emit_progress(progress_callback: ProgressCallback, **payload: object) -> None:
    if progress_callback is None:
        return
    progress_callback(payload)


def lower_names(values: Iterable[str]) -> set[str]:
    return {value.lower() for value in values}


def process_key(process: psutil.Process) -> tuple[int, float]:
    return process.pid, process.create_time()


def is_same_process(process: psutil.Process, created_at: float) -> bool:
    try:
        return process.is_running() and abs(process.create_time() - created_at) < 0.001
    except (psutil.Error, OSError):
        return False


def safe_process_name(process: Optional[psutil.Process]) -> str:
    if process is None:
        return ""
    try:
        return process.name() or ""
    except (psutil.Error, OSError):
        return ""


def safe_process_pid(process: Optional[psutil.Process]) -> str:
    if process is None:
        return ""
    return str(process.pid)


def safe_process_write_snapshot(process: psutil.Process) -> tuple[Optional[int], Optional[int]]:
    try:
        counters = process.io_counters()
    except (psutil.Error, OSError, AttributeError):
        return None, None
    return (
        getattr(counters, "write_bytes", None),
        getattr(counters, "write_count", None),
    )


def iter_matching_processes(process_names: Iterable[str]) -> list[psutil.Process]:
    wanted = lower_names(process_names)
    matched: list[psutil.Process] = []
    for process in psutil.process_iter(attrs=("name",)):
        try:
            process_name = (process.info.get("name") or "").lower()
            if process_name in wanted:
                matched.append(process)
        except (psutil.Error, OSError):
            continue
    return matched


def find_matching_ancestor(process: psutil.Process, main_process_names: Iterable[str]) -> Optional[psutil.Process]:
    wanted = lower_names(main_process_names)
    try:
        for parent in process.parents():
            if (parent.name() or "").lower() in wanted:
                return parent
    except (psutil.Error, OSError):
        return None
    return None


def pick_running_main_process(main_process_names: Iterable[str]) -> Optional[psutil.Process]:
    matched = iter_matching_processes(main_process_names)
    if not matched:
        return None
    try:
        matched.sort(key=lambda proc: proc.create_time())
    except (psutil.Error, OSError):
        return matched[0]
    return matched[0]


def build_process_watch_target(profile: ProcessProfile) -> str:
    return ", ".join(profile.worker_process_names)


def parse_iso_datetime(raw_value: str) -> Optional[datetime]:
    if not raw_value:
        return None
    try:
        return datetime.fromisoformat(raw_value)
    except ValueError:
        return None


def ordered_unique(values: Iterable[str]) -> list[str]:
    seen: set[str] = set()
    result: list[str] = []
    for value in values:
        normalized = value.strip()
        if not normalized or normalized in seen:
            continue
        seen.add(normalized)
        result.append(normalized)
    return result


def parse_float(raw_value: str) -> Optional[float]:
    if not raw_value:
        return None
    try:
        return float(raw_value)
    except ValueError:
        return None


def choose_longest_worker_row(rows: list[dict[str, str]]) -> dict[str, str]:
    def sort_key(row: dict[str, str]) -> tuple[float, float]:
        duration_value = parse_float(row.get("duration_seconds", "")) or 0.0
        started_at = parse_iso_datetime(row.get("started_at", ""))
        started_at_ts = started_at.timestamp() if started_at is not None else float("inf")
        return duration_value, -started_at_ts

    return max(rows, key=sort_key)


def resolve_worker_started_at(profile: ProcessProfile, worker_info: dict[str, object]) -> Optional[datetime]:
    if profile.timing_mode == "write_activity_span":
        activity_started_at = worker_info.get("activity_started_at")
        if isinstance(activity_started_at, datetime):
            return activity_started_at
        first_seen_at = worker_info.get("first_seen_at")
        if isinstance(first_seen_at, datetime):
            return first_seen_at
    started_at = worker_info.get("started_at")
    return started_at if isinstance(started_at, datetime) else None


def resolve_worker_ended_at(
    profile: ProcessProfile,
    worker_info: dict[str, object],
    fallback_ended_at: datetime,
) -> datetime:
    if profile.timing_mode == "write_activity_span":
        last_activity_at = worker_info.get("last_activity_at")
        if isinstance(last_activity_at, datetime):
            return last_activity_at
        last_seen_at = worker_info.get("last_seen_at")
        if isinstance(last_seen_at, datetime):
            return last_seen_at
    return fallback_ended_at


def resolve_worker_evidence(profile: ProcessProfile, worker_info: dict[str, object], row_status: str) -> str:
    if profile.timing_mode == "write_activity_span" and row_status == "ok":
        if isinstance(worker_info.get("activity_started_at"), datetime):
            return "worker_process_write_activity"
    return "worker_process_runtime"


def resolve_worker_notes(profile: ProcessProfile, worker_info: dict[str, object], row_status: str) -> str:
    notes = profile.notes.strip()
    if profile.timing_mode == "write_activity_span" and row_status == "ok":
        if isinstance(worker_info.get("activity_started_at"), datetime):
            notes = (
                f"{notes} Timed by ffmpeg write activity instead of raw process lifetime "
                "to better match the Shutter Encoder UI timer."
            ).strip()
    return notes


def should_capture_worker(profile: ProcessProfile, worker_info: dict[str, object], started_at: Optional[datetime], ended_at: Optional[datetime]) -> bool:
    if profile.require_activity_for_capture and not isinstance(worker_info.get("activity_started_at"), datetime):
        return False
    duration_seconds = calculate_duration_seconds(started_at, ended_at)
    if duration_seconds is None:
        return False
    return duration_seconds >= profile.min_capture_seconds


def refresh_worker_activity(
    profile: ProcessProfile,
    process: psutil.Process,
    worker_info: dict[str, object],
    sampled_at: datetime,
) -> None:
    worker_info["last_seen_at"] = sampled_at
    if profile.timing_mode != "write_activity_span":
        return

    write_bytes, write_count = safe_process_write_snapshot(process)
    previous_write_bytes = worker_info.get("last_write_bytes")
    previous_write_count = worker_info.get("last_write_count")

    write_progress_detected = False
    if isinstance(previous_write_bytes, int) and isinstance(write_bytes, int):
        write_progress_detected = write_bytes > previous_write_bytes
    elif isinstance(previous_write_count, int) and isinstance(write_count, int):
        write_progress_detected = write_count > previous_write_count

    if write_progress_detected:
        if not isinstance(worker_info.get("activity_started_at"), datetime):
            worker_info["activity_started_at"] = sampled_at
        worker_info["last_activity_at"] = sampled_at

    worker_info["last_write_bytes"] = write_bytes
    worker_info["last_write_count"] = write_count


def aggregate_process_rows(profile: ProcessProfile, rows: list[dict[str, str]]) -> dict[str, str]:
    if not rows:
        raise ValueError("rows must not be empty")

    if profile.aggregation_mode == "longest_worker":
        primary_row = choose_longest_worker_row(rows)
        notes = profile.notes.strip()
        if len(rows) > 1:
            notes = (
                f"{notes} Selected the longest matching worker runtime out of {len(rows)} captured workers "
                "to avoid auxiliary helper ffmpeg runs."
            ).strip()
        return {
            **primary_row,
            "evidence": "worker_process_runtime_longest",
            "notes": notes,
        }

    started_at_candidates = [
        parsed
        for parsed in (parse_iso_datetime(row.get("started_at", "")) for row in rows)
        if parsed is not None
    ]
    ended_at_candidates = [
        parsed
        for parsed in (parse_iso_datetime(row.get("ended_at", "")) for row in rows)
        if parsed is not None
    ]

    aggregate_started_at = min(started_at_candidates) if started_at_candidates else None
    aggregate_ended_at = max(ended_at_candidates) if ended_at_candidates else None

    statuses = [row.get("status", "") for row in rows]
    if any(status == "timeout" for status in statuses):
        aggregate_status = "timeout"
    elif any(status == "stopped" for status in statuses):
        aggregate_status = "stopped"
    elif any(status == "ok" for status in statuses):
        aggregate_status = "ok"
    else:
        aggregate_status = statuses[-1] if statuses else "no_capture"

    main_process_names = ordered_unique(row.get("main_process_name", "") for row in rows)
    main_process_pids = ordered_unique(row.get("main_process_pid", "") for row in rows)
    worker_process_names = ordered_unique(row.get("worker_process_name", "") for row in rows)
    worker_process_pids = ordered_unique(row.get("worker_process_pid", "") for row in rows)

    notes = profile.notes.strip()
    if len(rows) > 1:
        notes = f"{notes} Aggregated {len(rows)} matching worker process runs into one session result.".strip()

    evidence = (
        "worker_process_runtime_group"
        if len(rows) > 1
        else rows[-1].get("evidence", "worker_process_runtime")
    )

    return make_result_row(
        software=profile.software,
        mode="process",
        session_id=rows[-1].get("session_id", ""),
        status=aggregate_status,
        started_at=aggregate_started_at,
        ended_at=aggregate_ended_at,
        main_process_name=", ".join(main_process_names),
        main_process_pid=", ".join(main_process_pids),
        worker_process_name=", ".join(worker_process_names),
        worker_process_pid=", ".join(worker_process_pids),
        evidence=evidence,
        notes=notes,
    )


def build_process_preview(
    profile: ProcessProfile,
    active_workers: dict[tuple[int, float], dict[str, object]],
    finished_rows: list[dict[str, str]],
) -> dict[str, object]:
    if active_workers:
        finished_started = [
            parsed
            for parsed in (parse_iso_datetime(row.get("started_at", "")) for row in finished_rows)
            if parsed is not None
        ]
        latest_active = max(
            active_workers.values(),
            key=lambda worker_info: float(resolve_worker_started_at(profile, worker_info).timestamp()),
        )
        active_started = [
            resolved_started_at
            for resolved_started_at in (resolve_worker_started_at(profile, worker_info) for worker_info in active_workers.values())
            if resolved_started_at is not None
        ]
        started_at = min(
            active_started + finished_started
            if finished_started
            else active_started
        )
        worker_name = str(latest_active["worker_name"])
        worker_pid = str(latest_active["worker_pid"])
        return {
            "progress_state": "active_worker_captured",
            "progress_message": (
                f"Detected a matching worker process ({worker_name}, pid={worker_pid}) and is monitoring until it exits."
            ),
            "capture_detected": True,
            "capture_complete": False,
            "current_source_path": f"{worker_name} pid={worker_pid}",
            "captured_line_count": len(active_workers) + len(finished_rows),
            "last_log_activity_at": "",
            "preview_status": "partial",
            "preview_started_at": isoformat_or_empty(started_at),
            "preview_ended_at": "",
            "preview_duration_seconds": format_seconds(calculate_duration_seconds(started_at, now_local())),
            "preview_evidence": "worker_process_runtime",
        }

    if finished_rows:
        aggregate_row = aggregate_process_rows(profile, finished_rows)
        worker_name = aggregate_row.get("worker_process_name", "")
        worker_pid = aggregate_row.get("worker_process_pid", "")
        return {
            "progress_state": "complete_data_captured",
            "progress_message": (
                f"Captured {len(finished_rows)} completed worker runtime result(s) and aggregated them into one session result."
            ),
            "capture_detected": True,
            "capture_complete": True,
            "current_source_path": f"{worker_name} pid={worker_pid}".strip(),
            "captured_line_count": len(finished_rows),
            "last_log_activity_at": "",
            "preview_status": "complete",
            "preview_started_at": aggregate_row.get("started_at", ""),
            "preview_ended_at": aggregate_row.get("ended_at", ""),
            "preview_duration_seconds": aggregate_row.get("duration_seconds", ""),
            "preview_evidence": aggregate_row.get("evidence", ""),
        }

    return {
        "progress_state": "waiting_for_worker_activity",
        "progress_message": "Waiting for a new matching worker process.",
        "capture_detected": False,
        "capture_complete": False,
        "current_source_path": build_process_watch_target(profile),
        "captured_line_count": 0,
        "last_log_activity_at": "",
        "preview_status": "none",
        "preview_started_at": "",
        "preview_ended_at": "",
        "preview_duration_seconds": "",
        "preview_evidence": "",
    }


def run_process_listener(
    config: dict,
    stop_file: Path,
    progress_callback: ProgressCallback = None,
) -> list[dict[str, str]]:
    software = config["software"]
    session_id = config["session_id"]
    options = config["options"]
    profile = PROCESS_PROFILES[software]

    deadline = time.monotonic() + float(options.get("max_runtime_seconds", 7200))
    poll_interval = float(options.get("poll_interval", 0.5))
    idle_seconds = float(options.get("idle_seconds", 3.0))
    stop_grace_seconds = float(options.get("stop_grace_seconds", 10.0))
    include_existing_workers = bool(options.get("include_existing_workers", False))
    allow_detached_worker = bool(options.get("allow_detached_worker", False))

    baseline_workers = set()
    if not include_existing_workers:
        baseline_workers = {process_key(process) for process in iter_matching_processes(profile.worker_process_names)}

    active_workers: dict[tuple[int, float], dict[str, object]] = {}
    finished_rows: list[dict[str, str]] = []
    stop_requested_at: Optional[float] = None
    last_worker_exit_at: Optional[float] = None
    last_preview_signature: Optional[tuple[object, ...]] = None

    emit_progress(
        progress_callback,
        progress_state="waiting_for_worker_activity",
        progress_message="Watching for a new matching worker process.",
        capture_detected=False,
        capture_complete=False,
        current_source_path=build_process_watch_target(profile),
        captured_line_count=0,
        last_log_activity_at="",
        preview_status="none",
        preview_started_at="",
        preview_ended_at="",
        preview_duration_seconds="",
        preview_evidence="",
    )

    while time.monotonic() < deadline:
        if stop_file.exists() and stop_requested_at is None:
            stop_requested_at = time.monotonic()
            live_preview = build_process_preview(profile, active_workers, finished_rows)
            emit_progress(
                progress_callback,
                progress_state="stop_requested",
                progress_message=(
                    "Stop was requested; waiting for active worker processes to exit within the grace period."
                    if active_workers
                    else "Stop was requested before any matching worker process was captured."
                ),
                capture_detected=bool(live_preview.get("capture_detected", False)),
                capture_complete=bool(live_preview.get("capture_complete", False)),
                current_source_path=str(live_preview.get("current_source_path", "")),
                captured_line_count=int(live_preview.get("captured_line_count", 0) or 0),
                last_log_activity_at="",
                preview_status=str(live_preview.get("preview_status", "")),
                preview_started_at=str(live_preview.get("preview_started_at", "")),
                preview_ended_at=str(live_preview.get("preview_ended_at", "")),
                preview_duration_seconds=str(live_preview.get("preview_duration_seconds", "")),
                preview_evidence=str(live_preview.get("preview_evidence", "")),
            )

        current_workers: dict[tuple[int, float], psutil.Process] = {}
        detected_new_worker = False
        for process in iter_matching_processes(profile.worker_process_names):
            try:
                key = process_key(process)
            except (psutil.Error, OSError):
                continue
            current_workers[key] = process

            if key in baseline_workers or key in active_workers:
                continue

            main_process = find_matching_ancestor(process, profile.main_process_names)
            if main_process is None and profile.require_parent_match and not allow_detached_worker:
                continue
            if main_process is None:
                main_process = pick_running_main_process(profile.main_process_names)

            captured_at = now_local()
            write_bytes, write_count = safe_process_write_snapshot(process)
            active_workers[key] = {
                "worker_name": safe_process_name(process),
                "worker_pid": str(process.pid),
                "started_at": datetime.fromtimestamp(key[1], tz=LOCAL_TZ),
                "first_seen_at": captured_at,
                "last_seen_at": captured_at,
                "activity_started_at": None,
                "last_activity_at": None,
                "last_write_bytes": write_bytes,
                "last_write_count": write_count,
                "main_process_name": safe_process_name(main_process),
                "main_process_pid": safe_process_pid(main_process),
            }
            detected_new_worker = True
            last_worker_exit_at = None

        finished_this_round = False
        for key, worker_info in list(active_workers.items()):
            running_process = current_workers.get(key)
            if running_process is not None and is_same_process(running_process, key[1]):
                refresh_worker_activity(profile, running_process, worker_info, now_local())
                continue

            row_started_at = resolve_worker_started_at(profile, worker_info)
            row_ended_at = resolve_worker_ended_at(profile, worker_info, now_local())
            if should_capture_worker(profile, worker_info, row_started_at, row_ended_at):
                row_status = "ok"
                finished_rows.append(
                    make_result_row(
                        software=profile.software,
                        mode="process",
                        session_id=session_id,
                        status=row_status,
                        started_at=row_started_at,
                        ended_at=row_ended_at,
                        main_process_name=str(worker_info["main_process_name"]),
                        main_process_pid=str(worker_info["main_process_pid"]),
                        worker_process_name=str(worker_info["worker_name"]),
                        worker_process_pid=str(worker_info["worker_pid"]),
                        evidence=resolve_worker_evidence(profile, worker_info, row_status),
                        notes=resolve_worker_notes(profile, worker_info, row_status),
                    )
                )
                finished_this_round = True
                last_worker_exit_at = time.monotonic()
            active_workers.pop(key, None)

        if detected_new_worker or finished_this_round or (active_workers and stop_requested_at is not None):
            live_preview = build_process_preview(profile, active_workers, finished_rows)
            preview_signature = (
                live_preview.get("progress_state"),
                live_preview.get("capture_detected"),
                live_preview.get("capture_complete"),
                live_preview.get("preview_status"),
                live_preview.get("current_source_path"),
                live_preview.get("captured_line_count"),
            )
            if preview_signature != last_preview_signature:
                emit_progress(progress_callback, **live_preview)
                last_preview_signature = preview_signature

        if (
            stop_requested_at is None
            and finished_rows
            and not active_workers
            and last_worker_exit_at is not None
            and time.monotonic() - last_worker_exit_at >= idle_seconds
        ):
            break

        if stop_requested_at is not None:
            if not active_workers:
                break
            if time.monotonic() - stop_requested_at >= stop_grace_seconds:
                break

        time.sleep(poll_interval)

    if active_workers:
        ended_at = now_local()
        row_status = "stopped" if stop_requested_at is not None else "timeout"
        suffix = (
            "Stop was requested before the worker process exited."
            if stop_requested_at is not None
            else "Monitor reached max runtime before the worker process exited."
        )
        for worker_info in active_workers.values():
            row_started_at = resolve_worker_started_at(profile, worker_info)
            row_ended_at = resolve_worker_ended_at(profile, worker_info, ended_at)
            finished_rows.append(
                make_result_row(
                    software=profile.software,
                    mode="process",
                    session_id=session_id,
                    status=row_status,
                    started_at=row_started_at,
                    ended_at=row_ended_at,
                    main_process_name=str(worker_info["main_process_name"]),
                    main_process_pid=str(worker_info["main_process_pid"]),
                    worker_process_name=str(worker_info["worker_name"]),
                    worker_process_pid=str(worker_info["worker_pid"]),
                    evidence=resolve_worker_evidence(profile, worker_info, row_status),
                    notes=f"{resolve_worker_notes(profile, worker_info, row_status)} {suffix}".strip(),
                )
            )

    if finished_rows:
        aggregate_row = aggregate_process_rows(profile, finished_rows)
        emit_progress(progress_callback, **build_process_preview(profile, {}, finished_rows))
        return [aggregate_row]

    note = (
        f"{profile.notes} Stop was requested before any matching worker process was captured."
        if stop_requested_at is not None
        else f"{profile.notes} No new matching worker process was captured."
    )
    return [
        make_result_row(
            software=profile.software,
            mode="process",
            session_id=session_id,
            status="no_capture",
            started_at=None,
            ended_at=None,
            evidence="no_matching_worker",
            notes=note.strip(),
        )
    ]
