#!/usr/bin/env python3
"""Collect background job runtimes from several Windows video editors."""

from __future__ import annotations

import argparse
import csv
import os
import time
import uuid
from dataclasses import dataclass
from datetime import date, datetime, time as dt_time, timedelta
from pathlib import Path
from typing import Iterable, Optional

import psutil


LOCAL_TZ = datetime.now().astimezone().tzinfo
TIME_ONLY_PATTERN = r"(?<!\d)(\d{2}:\d{2}:\d{2})(?:[-:.,](\d{1,3}))?(?!\d)"
CSV_HEADERS = [
    "software",
    "mode",
    "session_id",
    "status",
    "started_at",
    "ended_at",
    "duration_seconds",
    "duration_ms",
    "main_process_name",
    "main_process_pid",
    "worker_process_name",
    "worker_process_pid",
    "source_path",
    "evidence",
    "notes",
]


@dataclass(frozen=True)
class ProcessProfile:
    software: str
    main_process_names: tuple[str, ...]
    worker_process_names: tuple[str, ...]
    require_parent_match: bool = True
    notes: str = ""


@dataclass(frozen=True)
class LogProfile:
    software: str
    main_process_names: tuple[str, ...]
    log_file_candidates: tuple[str, ...] = ()
    log_dir_candidates: tuple[str, ...] = ()
    include_keywords: tuple[str, ...] = ()
    ignored_filename_prefixes: tuple[str, ...] = ()
    ignored_filenames: tuple[str, ...] = ()
    notes: str = ""


PROCESS_PROFILES = {
    "shotcut": ProcessProfile(
        software="shotcut",
        main_process_names=("shotcut.exe",),
        worker_process_names=("melt.exe",),
        require_parent_match=True,
        notes="Measure only melt.exe instances started by Shotcut to avoid unrelated MLT jobs.",
    ),
    "kdenlive": ProcessProfile(
        software="kdenlive",
        main_process_names=("kdenlive.exe",),
        worker_process_names=("melt.exe",),
        require_parent_match=True,
        notes="Measure only melt.exe instances started by Kdenlive to avoid unrelated MLT jobs.",
    ),
    "shutter_encoder": ProcessProfile(
        software="shutter_encoder",
        main_process_names=("shutter encoder.exe", "shutterencoder.exe"),
        worker_process_names=("ffmpeg.exe",),
        require_parent_match=True,
        notes="Measure only ffmpeg.exe instances started by Shutter Encoder to avoid unrelated transcodes.",
    ),
}

LOG_PROFILES = {
    "avidemux": LogProfile(
        software="avidemux",
        main_process_names=("avidemux.exe",),
        log_file_candidates=(
            r"%LOCALAPPDATA%\avidemux\admlog.txt",
            r"%APPDATA%\Avidemux\admlog.txt",
            r"%APPDATA%\avidemux\admlog.txt",
        ),
        include_keywords=("save", "mux", "wrote", "encoding", "export"),
        notes="Prefer %LOCALAPPDATA% on Windows; %APPDATA%/Local is not a valid Local AppData path.",
    ),
    "handbrake": LogProfile(
        software="handbrake",
        main_process_names=("handbrake.exe", "handbrakecli.exe"),
        log_dir_candidates=(
            r"%APPDATA%\HandBrake\logs",
            r"%LOCALAPPDATA%\HandBrake\logs",
        ),
        ignored_filename_prefixes=("activity_log",),
        ignored_filenames=("handbrake-activitylog.txt",),
        notes="Prefer per-job log files under the logs directory and ignore generic activity logs by default.",
    ),
}


def now_local() -> datetime:
    return datetime.now().astimezone()


def expand_path(raw_path: str) -> Path:
    expanded = os.path.expandvars(os.path.expanduser(raw_path))
    return Path(expanded).resolve(strict=False)


def lower_names(values: Iterable[str]) -> set[str]:
    return {value.lower() for value in values}


def isoformat_or_empty(value: Optional[datetime]) -> str:
    if value is None:
        return ""
    return value.isoformat(timespec="milliseconds")


def format_duration(started_at: Optional[datetime], ended_at: Optional[datetime]) -> tuple[str, str]:
    if started_at is None or ended_at is None:
        return "", ""
    duration_seconds = max((ended_at - started_at).total_seconds(), 0.0)
    duration_ms = int(round(duration_seconds * 1000))
    return f"{duration_seconds:.3f}", str(duration_ms)


def ensure_parent_directory(path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)


def write_rows(csv_path: Path, rows: list[dict[str, str]]) -> None:
    ensure_parent_directory(csv_path)
    should_write_header = not csv_path.exists() or csv_path.stat().st_size == 0
    with csv_path.open("a", newline="", encoding="utf-8-sig") as handle:
        writer = csv.DictWriter(handle, fieldnames=CSV_HEADERS)
        if should_write_header:
            writer.writeheader()
        writer.writerows(rows)


def make_result_row(
    *,
    software: str,
    mode: str,
    session_id: str,
    status: str,
    started_at: Optional[datetime],
    ended_at: Optional[datetime],
    main_process_name: str = "",
    main_process_pid: str = "",
    worker_process_name: str = "",
    worker_process_pid: str = "",
    source_path: str = "",
    evidence: str = "",
    notes: str = "",
) -> dict[str, str]:
    duration_seconds, duration_ms = format_duration(started_at, ended_at)
    return {
        "software": software,
        "mode": mode,
        "session_id": session_id,
        "status": status,
        "started_at": isoformat_or_empty(started_at),
        "ended_at": isoformat_or_empty(ended_at),
        "duration_seconds": duration_seconds,
        "duration_ms": duration_ms,
        "main_process_name": main_process_name,
        "main_process_pid": main_process_pid,
        "worker_process_name": worker_process_name,
        "worker_process_pid": worker_process_pid,
        "source_path": source_path,
        "evidence": evidence,
        "notes": notes,
    }


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


def read_file_segment(file_path: Path, offset: int) -> tuple[str, int]:
    with file_path.open("rb") as handle:
        handle.seek(offset)
        payload = handle.read()
        new_offset = handle.tell()
    return payload.decode("utf-8", errors="ignore"), new_offset


def choose_existing_path(candidates: Iterable[Path]) -> Optional[Path]:
    for candidate in candidates:
        if candidate.exists():
            return candidate
    return None


def is_ignored_log_name(file_name: str, profile: LogProfile, allow_ignored_logs: bool) -> bool:
    if allow_ignored_logs:
        return False
    lowered = file_name.lower()
    if lowered in profile.ignored_filenames:
        return True
    return any(lowered.startswith(prefix) for prefix in profile.ignored_filename_prefixes)


def extract_timestamps(lines: list[str], observation_started_at: datetime) -> list[datetime]:
    import re

    timestamps: list[datetime] = []
    current_date: date = observation_started_at.date()
    previous_timestamp: Optional[datetime] = None
    for line in lines:
        for match in re.finditer(TIME_ONLY_PATTERN, line):
            clock_value = match.group(1)
            millisecond_value = match.group(2) or "0"
            hour, minute, second = (int(value) for value in clock_value.split(":"))
            milliseconds = int(millisecond_value.ljust(3, "0")[:3])
            candidate = datetime.combine(
                current_date,
                dt_time(hour=hour, minute=minute, second=second, microsecond=milliseconds * 1000),
                tzinfo=LOCAL_TZ,
            )
            if previous_timestamp and candidate < previous_timestamp - timedelta(hours=12):
                current_date += timedelta(days=1)
                candidate = datetime.combine(
                    current_date,
                    dt_time(hour=hour, minute=minute, second=second, microsecond=milliseconds * 1000),
                    tzinfo=LOCAL_TZ,
                )
            while candidate < observation_started_at - timedelta(hours=12):
                candidate += timedelta(days=1)
            timestamps.append(candidate)
            previous_timestamp = candidate
    return timestamps


def filter_relevant_lines(lines: list[str], include_keywords: tuple[str, ...]) -> tuple[list[str], str]:
    if not include_keywords:
        return lines, "all_lines"
    filtered = [line for line in lines if any(keyword in line.lower() for keyword in include_keywords)]
    if filtered:
        return filtered, "keyword_filtered"
    return lines, "all_lines_fallback"


def build_log_result(
    *,
    profile: LogProfile,
    session_id: str,
    source_path: str,
    lines: list[str],
    observation_started_at: datetime,
    first_change_at: Optional[datetime],
    last_change_at: Optional[datetime],
    main_process: Optional[psutil.Process],
    base_notes: str,
) -> dict[str, str]:
    relevant_lines, evidence_suffix = filter_relevant_lines(lines, profile.include_keywords)
    timestamps = extract_timestamps(relevant_lines, observation_started_at)

    if len(timestamps) >= 2:
        started_at = timestamps[0]
        ended_at = timestamps[-1]
        evidence = f"log_timestamps:{evidence_suffix}"
        notes = base_notes
    elif len(timestamps) == 1:
        started_at = timestamps[0]
        ended_at = last_change_at or timestamps[0]
        evidence = f"single_log_timestamp:{evidence_suffix}"
        notes = f"{base_notes} Fallback to file change window because only one timestamp was found.".strip()
    elif first_change_at and last_change_at:
        started_at = first_change_at
        ended_at = last_change_at
        evidence = "file_change_window"
        notes = f"{base_notes} Fallback to file write times because no parsable timestamps were found.".strip()
    else:
        return make_result_row(
            software=profile.software,
            mode="log",
            session_id=session_id,
            status="no_capture",
            started_at=None,
            ended_at=None,
            main_process_name=safe_process_name(main_process),
            main_process_pid=safe_process_pid(main_process),
            source_path=source_path,
            evidence="no_log_activity",
            notes=f"{base_notes} No new log content was captured.".strip(),
        )

    return make_result_row(
        software=profile.software,
        mode="log",
        session_id=session_id,
        status="ok",
        started_at=started_at,
        ended_at=ended_at,
        main_process_name=safe_process_name(main_process),
        main_process_pid=safe_process_pid(main_process),
        source_path=source_path,
        evidence=evidence,
        notes=notes.strip(),
    )


def monitor_process_profile(
    profile: ProcessProfile,
    *,
    session_id: str,
    timeout_seconds: float,
    poll_interval: float,
    idle_seconds: float,
    include_existing_workers: bool,
    allow_detached_worker: bool,
) -> list[dict[str, str]]:
    deadline = time.monotonic() + timeout_seconds
    baseline_workers = set()
    if not include_existing_workers:
        baseline_workers = {process_key(process) for process in iter_matching_processes(profile.worker_process_names)}

    active_workers: dict[tuple[int, float], dict[str, object]] = {}
    finished_rows: list[dict[str, str]] = []
    last_worker_finished_at: Optional[float] = None
    saw_any_worker = False

    while time.monotonic() < deadline:
        current_workers: dict[tuple[int, float], psutil.Process] = {}
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

            active_workers[key] = {
                "worker_name": safe_process_name(process),
                "worker_pid": str(process.pid),
                "started_at": datetime.fromtimestamp(key[1], tz=LOCAL_TZ),
                "main_process_name": safe_process_name(main_process),
                "main_process_pid": safe_process_pid(main_process),
            }
            saw_any_worker = True

        for key, worker_info in list(active_workers.items()):
            running_process = current_workers.get(key)
            if running_process is not None and is_same_process(running_process, key[1]):
                continue

            finished_rows.append(
                make_result_row(
                    software=profile.software,
                    mode="process",
                    session_id=session_id,
                    status="ok",
                    started_at=worker_info["started_at"],
                    ended_at=now_local(),
                    main_process_name=str(worker_info["main_process_name"]),
                    main_process_pid=str(worker_info["main_process_pid"]),
                    worker_process_name=str(worker_info["worker_name"]),
                    worker_process_pid=str(worker_info["worker_pid"]),
                    evidence="worker_process_runtime",
                    notes=profile.notes,
                )
            )
            active_workers.pop(key, None)
            last_worker_finished_at = time.monotonic()

        if saw_any_worker and not active_workers and last_worker_finished_at is not None:
            if time.monotonic() - last_worker_finished_at >= idle_seconds:
                break

        time.sleep(poll_interval)

    if active_workers:
        timeout_at = now_local()
        for worker_info in active_workers.values():
            finished_rows.append(
                make_result_row(
                    software=profile.software,
                    mode="process",
                    session_id=session_id,
                    status="timeout",
                    started_at=worker_info["started_at"],
                    ended_at=timeout_at,
                    main_process_name=str(worker_info["main_process_name"]),
                    main_process_pid=str(worker_info["main_process_pid"]),
                    worker_process_name=str(worker_info["worker_name"]),
                    worker_process_pid=str(worker_info["worker_pid"]),
                    evidence="worker_process_runtime",
                    notes=f"{profile.notes} Monitor timed out before the worker process exited.".strip(),
                )
            )

    if finished_rows:
        return finished_rows

    return [
        make_result_row(
            software=profile.software,
            mode="process",
            session_id=session_id,
            status="no_capture",
            started_at=None,
            ended_at=None,
            evidence="no_matching_worker",
            notes=f"{profile.notes} No new matching worker process was captured.".strip(),
        )
    ]


def monitor_single_log_file(
    profile: LogProfile,
    *,
    session_id: str,
    timeout_seconds: float,
    poll_interval: float,
    idle_seconds: float,
    log_path_override: Optional[Path],
) -> list[dict[str, str]]:
    observation_started_at = now_local()
    deadline = time.monotonic() + timeout_seconds
    candidate_paths = [log_path_override] if log_path_override else [expand_path(path) for path in profile.log_file_candidates]
    chosen_path = choose_existing_path(candidate_paths) or candidate_paths[0]
    source_path = str(chosen_path)
    baseline_offset = chosen_path.stat().st_size if chosen_path.exists() else 0
    captured_lines: list[str] = []
    first_change_at: Optional[datetime] = None
    last_change_at: Optional[datetime] = None

    while time.monotonic() < deadline:
        existing_path = choose_existing_path(candidate_paths)
        if existing_path is not None and existing_path != chosen_path:
            chosen_path = existing_path
            source_path = str(chosen_path)
            baseline_offset = 0

        if chosen_path.exists():
            current_size = chosen_path.stat().st_size
            if current_size < baseline_offset:
                baseline_offset = 0
                captured_lines.clear()
                first_change_at = None
                last_change_at = None

            if current_size > baseline_offset:
                chunk, baseline_offset = read_file_segment(chosen_path, baseline_offset)
                lines = [line for line in chunk.splitlines() if line.strip()]
                if lines:
                    change_time = now_local()
                    first_change_at = first_change_at or change_time
                    last_change_at = change_time
                    captured_lines.extend(lines)
                    source_path = str(chosen_path)

        if captured_lines and last_change_at is not None:
            if (now_local() - last_change_at).total_seconds() >= idle_seconds:
                main_process = pick_running_main_process(profile.main_process_names)
                return [
                    build_log_result(
                        profile=profile,
                        session_id=session_id,
                        source_path=source_path,
                        lines=captured_lines,
                        observation_started_at=observation_started_at,
                        first_change_at=first_change_at,
                        last_change_at=last_change_at,
                        main_process=main_process,
                        base_notes=profile.notes,
                    )
                ]

        time.sleep(poll_interval)

    if captured_lines:
        main_process = pick_running_main_process(profile.main_process_names)
        return [
            build_log_result(
                profile=profile,
                session_id=session_id,
                source_path=source_path,
                lines=captured_lines,
                observation_started_at=observation_started_at,
                first_change_at=first_change_at,
                last_change_at=last_change_at,
                main_process=main_process,
                base_notes=f"{profile.notes} Monitor timed out before the log went idle.".strip(),
            )
        ]

    return [
        make_result_row(
            software=profile.software,
            mode="log",
            session_id=session_id,
            status="no_capture",
            started_at=None,
            ended_at=None,
            source_path=source_path,
            evidence="no_log_activity",
            notes=f"{profile.notes} No new log content was captured.".strip(),
        )
    ]


def list_log_files(directory: Path) -> list[Path]:
    if not directory.exists():
        return []
    return [path for path in directory.iterdir() if path.is_file() and path.suffix.lower() in {".txt", ".log"}]


def choose_log_file_for_directory_monitor(
    *,
    files: list[Path],
    baseline_sizes: dict[Path, int],
    profile: LogProfile,
    allow_ignored_logs: bool,
    observation_started_at: datetime,
) -> Optional[Path]:
    candidates: list[tuple[float, Path]] = []
    for file_path in files:
        if is_ignored_log_name(file_path.name, profile, allow_ignored_logs):
            continue
        try:
            stat_result = file_path.stat()
        except OSError:
            continue
        baseline_size = baseline_sizes.get(file_path, 0)
        changed_after_observation = stat_result.st_mtime >= observation_started_at.timestamp() - 1
        appended_after_observation = stat_result.st_size > baseline_size
        if changed_after_observation or appended_after_observation or file_path not in baseline_sizes:
            candidates.append((stat_result.st_mtime, file_path))
    if not candidates:
        return None
    candidates.sort(key=lambda item: item[0], reverse=True)
    return candidates[0][1]


def monitor_log_directory(
    profile: LogProfile,
    *,
    session_id: str,
    timeout_seconds: float,
    poll_interval: float,
    idle_seconds: float,
    log_dir_override: Optional[Path],
    allow_ignored_logs: bool,
) -> list[dict[str, str]]:
    observation_started_at = now_local()
    deadline = time.monotonic() + timeout_seconds
    candidate_dirs = [log_dir_override] if log_dir_override else [expand_path(path) for path in profile.log_dir_candidates]
    chosen_dir = choose_existing_path(candidate_dirs) or candidate_dirs[0]
    baseline_sizes: dict[Path, int] = {}
    for existing_file in list_log_files(chosen_dir):
        try:
            baseline_sizes[existing_file] = existing_file.stat().st_size
        except OSError:
            continue

    active_file: Optional[Path] = None
    active_offset = 0
    captured_lines: list[str] = []
    first_change_at: Optional[datetime] = None
    last_change_at: Optional[datetime] = None

    while time.monotonic() < deadline:
        existing_dir = choose_existing_path(candidate_dirs)
        if existing_dir is not None:
            chosen_dir = existing_dir

        files = list_log_files(chosen_dir)
        if active_file is None:
            active_file = choose_log_file_for_directory_monitor(
                files=files,
                baseline_sizes=baseline_sizes,
                profile=profile,
                allow_ignored_logs=allow_ignored_logs,
                observation_started_at=observation_started_at,
            )
            if active_file is not None:
                active_offset = baseline_sizes.get(active_file, 0)

        if active_file is not None and active_file.exists():
            current_size = active_file.stat().st_size
            if current_size < active_offset:
                active_offset = 0
                captured_lines.clear()
                first_change_at = None
                last_change_at = None

            if current_size > active_offset:
                chunk, active_offset = read_file_segment(active_file, active_offset)
                lines = [line for line in chunk.splitlines() if line.strip()]
                if lines:
                    change_time = now_local()
                    first_change_at = first_change_at or change_time
                    last_change_at = change_time
                    captured_lines.extend(lines)

        for file_path in files:
            try:
                baseline_sizes[file_path] = file_path.stat().st_size
            except OSError:
                continue

        if captured_lines and last_change_at is not None:
            if (now_local() - last_change_at).total_seconds() >= idle_seconds:
                main_process = pick_running_main_process(profile.main_process_names)
                return [
                    build_log_result(
                        profile=profile,
                        session_id=session_id,
                        source_path=str(active_file),
                        lines=captured_lines,
                        observation_started_at=observation_started_at,
                        first_change_at=first_change_at,
                        last_change_at=last_change_at,
                        main_process=main_process,
                        base_notes=profile.notes,
                    )
                ]

        time.sleep(poll_interval)

    if captured_lines and active_file is not None:
        main_process = pick_running_main_process(profile.main_process_names)
        return [
            build_log_result(
                profile=profile,
                session_id=session_id,
                source_path=str(active_file),
                lines=captured_lines,
                observation_started_at=observation_started_at,
                first_change_at=first_change_at,
                last_change_at=last_change_at,
                main_process=main_process,
                base_notes=f"{profile.notes} Monitor timed out before the log went idle.".strip(),
            )
        ]

    return [
        make_result_row(
            software=profile.software,
            mode="log",
            session_id=session_id,
            status="no_capture",
            started_at=None,
            ended_at=None,
            source_path=str(chosen_dir),
            evidence="no_log_activity",
            notes=f"{profile.notes} No new per-job log file was captured.".strip(),
        )
    ]


def print_profiles() -> int:
    print("Process profiles:")
    for profile in PROCESS_PROFILES.values():
        print(
            f"  - {profile.software}: main={','.join(profile.main_process_names)} "
            f"worker={','.join(profile.worker_process_names)}"
        )
    print("Log profiles:")
    for profile in LOG_PROFILES.values():
        location = profile.log_file_candidates or profile.log_dir_candidates
        print(f"  - {profile.software}: source={','.join(location)}")
    print("Special profile:")
    print("  - blender: use blender_hook.py inside Blender")
    return 0


def run_monitor_command(args: argparse.Namespace) -> int:
    session_id = args.session_id or str(uuid.uuid4())
    output_path = expand_path(args.output)

    if args.software == "blender":
        row = make_result_row(
            software="blender",
            mode="hook",
            session_id=session_id,
            status="no_capture",
            started_at=None,
            ended_at=None,
            evidence="external_hook_required",
            notes="Use blender_hook.py inside Blender because bpy is unavailable in standard Python.",
        )
        write_rows(output_path, [row])
        print(f"Wrote reminder row to {output_path}")
        return 2

    if args.software in PROCESS_PROFILES:
        rows = monitor_process_profile(
            PROCESS_PROFILES[args.software],
            session_id=session_id,
            timeout_seconds=args.timeout,
            poll_interval=args.poll_interval,
            idle_seconds=args.idle_seconds,
            include_existing_workers=args.include_existing_workers,
            allow_detached_worker=args.allow_detached_worker,
        )
    else:
        profile = LOG_PROFILES[args.software]
        log_path_override = expand_path(args.log_path) if args.log_path else None
        log_dir_override = expand_path(args.log_dir) if args.log_dir else None
        if profile.log_file_candidates:
            rows = monitor_single_log_file(
                profile,
                session_id=session_id,
                timeout_seconds=args.timeout,
                poll_interval=args.poll_interval,
                idle_seconds=args.idle_seconds,
                log_path_override=log_path_override,
            )
        else:
            rows = monitor_log_directory(
                profile,
                session_id=session_id,
                timeout_seconds=args.timeout,
                poll_interval=args.poll_interval,
                idle_seconds=args.idle_seconds,
                log_dir_override=log_dir_override,
                allow_ignored_logs=args.allow_activity_log,
            )

    write_rows(output_path, rows)
    status_summary = ",".join(sorted({row["status"] for row in rows}))
    print(f"Wrote {len(rows)} row(s) to {output_path} with status={status_summary}")
    if any(row["status"] != "ok" for row in rows):
        return 2
    return 0


def build_argument_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Measure background processing durations for several Windows video editors."
    )
    subparsers = parser.add_subparsers(dest="command", required=True)

    list_parser = subparsers.add_parser("list-profiles", help="Show all built-in software profiles.")
    list_parser.set_defaults(handler=lambda args: print_profiles())

    monitor_parser = subparsers.add_parser("monitor", help="Run one software monitor and append results to CSV.")
    monitor_parser.add_argument(
        "--software",
        required=True,
        choices=sorted(list(PROCESS_PROFILES) + list(LOG_PROFILES) + ["blender"]),
        help="Target software profile.",
    )
    monitor_parser.add_argument(
        "--output",
        default="results/video_job_durations.csv",
        help="CSV output file path. Default: results/video_job_durations.csv",
    )
    monitor_parser.add_argument("--timeout", type=float, default=3600, help="Maximum monitor duration in seconds.")
    monitor_parser.add_argument("--poll-interval", type=float, default=0.5, help="Polling interval in seconds.")
    monitor_parser.add_argument(
        "--idle-seconds",
        type=float,
        default=3.0,
        help="Finish after this many idle seconds once data has started arriving.",
    )
    monitor_parser.add_argument("--session-id", default="", help="Optional session id. A UUID is generated when omitted.")
    monitor_parser.add_argument("--log-path", default="", help="Optional override for single log file profiles.")
    monitor_parser.add_argument("--log-dir", default="", help="Optional override for log directory profiles.")
    monitor_parser.add_argument(
        "--include-existing-workers",
        action="store_true",
        help="Track already-running worker processes instead of only new ones.",
    )
    monitor_parser.add_argument(
        "--allow-detached-worker",
        action="store_true",
        help="Allow worker capture even when no matching parent GUI process is found.",
    )
    monitor_parser.add_argument(
        "--allow-activity-log",
        action="store_true",
        help="Allow generic activity logs when monitoring HandBrake.",
    )
    monitor_parser.set_defaults(handler=run_monitor_command)
    return parser


def main(argv: Optional[list[str]] = None) -> int:
    parser = build_argument_parser()
    args = parser.parse_args(argv)
    return args.handler(args)


if __name__ == "__main__":
    raise SystemExit(main())
