from __future__ import annotations

import re
import time
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Callable, Optional

from monitoring_common import (
    calculate_duration_seconds,
    expand_path,
    format_seconds,
    isoformat_or_empty,
    make_result_row,
    now_local,
)


WORK_MARKER_PATTERN = re.compile(
    r"^\[(?P<clock>\d{2}:\d{2}:\d{2})\]\s+"
    r"(?P<event>Starting work at|Finished work at):\s+"
    r"(?P<full_text>.+?)\s*$"
)
AVIDEMUX_ELAPSED_PATTERN = re.compile(
    r"^\s*(?P<percent>\d+)% done\s+frames:\s+(?P<frames>\d+)\s+elapsed:\s+(?P<elapsed>\d{2}:\d{2}:\d{2},\d{3})\s*$"
)
AVIDEMUX_TIMESTAMP_PATTERN = re.compile(r"(?P<clock>\d{2}:\d{2}:\d{2})-(?P<millis>\d{3})")
AVIDEMUX_PRIMARY_START_MARKERS = ("[A_Save] Saving..",)
AVIDEMUX_FALLBACK_START_MARKERS = ("[FF] Saving",)
AVIDEMUX_END_MARKERS = ("End of flush",)
AVIDEMUX_COMPLETION_SETTLE_SECONDS = 1.0
HANDBRAKE_ENCODE_FILENAME_PATTERN = re.compile(r"_encode_", re.IGNORECASE)
HANDBRAKE_ENCODE_MARKERS = (
    "# Starting Encode ...",
    "Starting work at:",
    "Finished work at:",
    "# Job Completed!",
)
HANDBRAKE_SCAN_MARKERS = (
    "# Starting Scan ...",
    "# Scan Finished ...",
    "hb_scan:",
    "scan: decoding previews",
    "libhb: scan thread found",
)
AVIDEMUX_DURATION_WARNING_SECONDS = 1.0
MONTHS = {
    "Jan": 1,
    "Feb": 2,
    "Mar": 3,
    "Apr": 4,
    "May": 5,
    "Jun": 6,
    "Jul": 7,
    "Aug": 8,
    "Sep": 9,
    "Oct": 10,
    "Nov": 11,
    "Dec": 12,
}


@dataclass(frozen=True)
class LogProfile:
    software: str
    main_process_names: tuple[str, ...]
    parser_kind: str
    log_file_candidates: tuple[str, ...] = ()
    log_dir_candidates: tuple[str, ...] = ()
    ignored_filename_prefixes: tuple[str, ...] = ()
    ignored_filenames: tuple[str, ...] = ()
    notes: str = ""


@dataclass(frozen=True)
class AvidemuxParseResult:
    duration_seconds: Optional[float]
    started_at: Optional[datetime]
    ended_at: Optional[datetime]
    computed_duration_seconds: Optional[float]
    duration_delta_seconds: Optional[float]
    evidence: str
    is_complete: bool
    notes: tuple[str, ...] = ()


ProgressCallback = Optional[Callable[[dict[str, object]], None]]


LOG_PROFILES = {
    "avidemux": LogProfile(
        software="avidemux",
        main_process_names=("avidemux.exe",),
        parser_kind="avidemux_elapsed",
        log_file_candidates=(
            r"%LOCALAPPDATA%\avidemux\admlog.txt",
            r"%APPDATA%\Avidemux\admlog.txt",
            r"%APPDATA%\avidemux\admlog.txt",
        ),
        notes="Prefer %LOCALAPPDATA% on Windows; %APPDATA%/Local is not a valid Local AppData path.",
    ),
    "handbrake": LogProfile(
        software="handbrake",
        main_process_names=("handbrake.exe", "handbrakecli.exe"),
        parser_kind="work_markers",
        log_dir_candidates=(
            r"%APPDATA%\HandBrake\logs",
            r"%LOCALAPPDATA%\HandBrake\logs",
        ),
        ignored_filename_prefixes=("activity_log",),
        ignored_filenames=("handbrake-activitylog.txt",),
        notes="Prefer per-job log files under the logs directory and ignore generic activity logs by default.",
    ),
}


def read_file_segment(file_path: Path, offset: int) -> tuple[str, int]:
    with file_path.open("rb") as handle:
        handle.seek(offset)
        payload = handle.read()
        new_offset = handle.tell()
    return payload.decode("utf-8", errors="ignore"), new_offset


def split_complete_lines(pending_text: str, chunk: str) -> tuple[list[str], str]:
    merged = pending_text + chunk
    if not merged:
        return [], ""

    raw_lines = merged.splitlines(keepends=True)
    if merged.endswith("\n") or merged.endswith("\r"):
        complete_lines = [line.rstrip("\r\n") for line in raw_lines]
        return complete_lines, ""

    if len(raw_lines) == 1:
        return [], raw_lines[0]

    complete_lines = [line.rstrip("\r\n") for line in raw_lines[:-1]]
    return complete_lines, raw_lines[-1]


def finalize_captured_lines(captured_lines: list[str], pending_text: str) -> list[str]:
    if pending_text.strip():
        return [*captured_lines, pending_text.strip()]
    return captured_lines


def choose_existing_path(candidates: list[Path]) -> Optional[Path]:
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


def normalize_log_line(line: str) -> str:
    return line.lstrip("\ufeff").strip()


def parse_work_datetime(raw_value: str) -> Optional[datetime]:
    parts = raw_value.strip().split()
    if len(parts) != 5:
        return None

    _, month_name, day_value, clock_value, year_value = parts
    month_value = MONTHS.get(month_name)
    if month_value is None:
        return None

    try:
        hour, minute, second = (int(value) for value in clock_value.split(":"))
        return datetime(
            year=int(year_value),
            month=month_value,
            day=int(day_value),
            hour=hour,
            minute=minute,
            second=second,
            tzinfo=now_local().tzinfo,
        )
    except ValueError:
        return None


def extract_latest_work_window(lines: list[str]) -> tuple[Optional[datetime], Optional[datetime]]:
    active_start: Optional[datetime] = None
    completed_pairs: list[tuple[datetime, datetime]] = []

    for line in lines:
        match = WORK_MARKER_PATTERN.search(normalize_log_line(line))
        if not match:
            continue

        marker_time = parse_work_datetime(match.group("full_text"))
        if marker_time is None:
            continue

        event_name = match.group("event")
        if event_name == "Starting work at":
            active_start = marker_time
            continue

        if active_start is not None and marker_time >= active_start:
            completed_pairs.append((active_start, marker_time))
            active_start = None

    if completed_pairs:
        return completed_pairs[-1]
    return active_start, None


def parse_elapsed_seconds(raw_value: str) -> Optional[float]:
    try:
        hour_value = int(raw_value[0:2])
        minute_value = int(raw_value[3:5])
        second_value = int(raw_value[6:8])
        millisecond_value = int(raw_value[9:12])
    except (ValueError, IndexError):
        return None
    return hour_value * 3600 + minute_value * 60 + second_value + millisecond_value / 1000.0


def combine_notes(*parts: str) -> str:
    return " ".join(part.strip() for part in parts if part and part.strip())


def emit_progress(progress_callback: ProgressCallback, **payload: object) -> None:
    if progress_callback is None:
        return
    progress_callback(payload)


def read_text_preview(file_path: Path, *, max_bytes: int = 8192) -> str:
    try:
        with file_path.open("rb") as handle:
            payload = handle.read(max_bytes)
    except OSError:
        return ""
    return payload.decode("utf-8", errors="ignore")


def classify_handbrake_log_file(file_path: Path) -> str:
    if HANDBRAKE_ENCODE_FILENAME_PATTERN.search(file_path.name):
        return "encode"

    preview_text = read_text_preview(file_path)
    if not preview_text:
        return "unknown"

    if any(marker in preview_text for marker in HANDBRAKE_ENCODE_MARKERS):
        return "encode"
    if any(marker in preview_text for marker in HANDBRAKE_SCAN_MARKERS):
        return "scan"
    return "unknown"


def should_wait_for_completion(profile: LogProfile, live_preview: dict[str, object]) -> bool:
    if profile.parser_kind == "avidemux_elapsed":
        return not bool(live_preview.get("capture_complete", False))
    if profile.software == "handbrake":
        return bool(live_preview.get("capture_detected", False)) and not bool(
            live_preview.get("capture_complete", False)
        )
    return False


def should_finalize_avidemux_completed_capture(
    profile: LogProfile,
    live_preview: dict[str, object],
    *,
    completion_detected_at: Optional[float],
    now_monotonic: float,
) -> bool:
    if profile.parser_kind != "avidemux_elapsed":
        return False
    if not bool(live_preview.get("capture_complete", False)):
        return False
    if completion_detected_at is None:
        return False
    return now_monotonic - completion_detected_at >= AVIDEMUX_COMPLETION_SETTLE_SECONDS


def extract_avidemux_elapsed_seconds(lines: list[str]) -> tuple[Optional[float], str]:
    latest_any: Optional[float] = None
    latest_complete: Optional[float] = None

    for line in lines:
        match = AVIDEMUX_ELAPSED_PATTERN.search(normalize_log_line(line))
        if not match:
            continue

        elapsed_seconds = parse_elapsed_seconds(match.group("elapsed"))
        if elapsed_seconds is None:
            continue

        latest_any = elapsed_seconds
        if int(match.group("percent")) == 100:
            # Avidemux appends to a single log file, so the newest completed
            # run is the last 100% line, not the longest one.
            latest_complete = elapsed_seconds

    if latest_complete is not None:
        return latest_complete, "avidemux_elapsed_100_percent"
    if latest_any is not None:
        return latest_any, "avidemux_elapsed_partial"
    return None, "no_elapsed_marker"


def line_contains_marker(line: str, markers: tuple[str, ...]) -> bool:
    return any(marker in line for marker in markers)


def avidemux_has_export_activity(lines: list[str]) -> bool:
    for raw_line in lines:
        line = normalize_log_line(raw_line)
        if not line:
            continue
        if line_contains_marker(line, AVIDEMUX_PRIMARY_START_MARKERS):
            return True
        if line_contains_marker(line, AVIDEMUX_FALLBACK_START_MARKERS):
            return True
        if line_contains_marker(line, AVIDEMUX_END_MARKERS):
            return True
        if AVIDEMUX_ELAPSED_PATTERN.search(line):
            return True
    return False


def infer_log_date(source_path: str) -> date:
    reference = now_local()
    source = Path(source_path)
    try:
        if source.exists():
            reference = datetime.fromtimestamp(source.stat().st_mtime, tz=reference.tzinfo)
    except OSError:
        pass
    return reference.date()


def parse_avidemux_timestamp(raw_line: str, log_date: date) -> Optional[datetime]:
    match = AVIDEMUX_TIMESTAMP_PATTERN.search(normalize_log_line(raw_line))
    if not match:
        return None

    try:
        hour_value, minute_value, second_value = (int(value) for value in match.group("clock").split(":"))
        millisecond_value = int(match.group("millis"))
    except ValueError:
        return None

    return datetime(
        year=log_date.year,
        month=log_date.month,
        day=log_date.day,
        hour=hour_value,
        minute=minute_value,
        second=second_value,
        microsecond=millisecond_value * 1000,
        tzinfo=now_local().tzinfo,
    )


def find_avidemux_run_start_indexes(lines: list[str]) -> list[int]:
    primary = [index for index, line in enumerate(lines) if line_contains_marker(line, AVIDEMUX_PRIMARY_START_MARKERS)]
    if primary:
        return primary

    fallback = [index for index, line in enumerate(lines) if line_contains_marker(line, AVIDEMUX_FALLBACK_START_MARKERS)]
    if fallback:
        return fallback

    if any(AVIDEMUX_ELAPSED_PATTERN.search(line) for line in lines):
        return [0]
    return []


def find_latest_avidemux_timestamp(
    lines: list[str],
    *,
    start_index: int,
    end_index: int,
    log_date: date,
) -> Optional[datetime]:
    for index in range(end_index - 1, start_index - 1, -1):
        timestamp = parse_avidemux_timestamp(lines[index], log_date)
        if timestamp is not None:
            return timestamp
    return None


def find_first_avidemux_timestamp(
    lines: list[str],
    *,
    start_index: int,
    end_index: int,
    log_date: date,
) -> Optional[datetime]:
    for index in range(start_index, end_index):
        timestamp = parse_avidemux_timestamp(lines[index], log_date)
        if timestamp is not None:
            return timestamp
    return None


def parse_avidemux_run(
    lines: list[str],
    *,
    start_index: int,
    end_index: int,
    log_date: date,
) -> Optional[AvidemuxParseResult]:
    latest_any: Optional[tuple[int, float]] = None
    latest_complete: Optional[tuple[int, float]] = None
    for index in range(start_index, end_index):
        match = AVIDEMUX_ELAPSED_PATTERN.search(lines[index])
        if not match:
            continue

        elapsed_seconds = parse_elapsed_seconds(match.group("elapsed"))
        if elapsed_seconds is None:
            continue

        latest_any = (index, elapsed_seconds)
        if int(match.group("percent")) == 100:
            latest_complete = (index, elapsed_seconds)

    selected_progress = latest_complete or latest_any
    if selected_progress is None:
        return None

    progress_index, elapsed_seconds = selected_progress
    started_at = parse_avidemux_timestamp(lines[start_index], log_date)
    if started_at is None:
        started_at = find_first_avidemux_timestamp(
            lines,
            start_index=start_index,
            end_index=end_index,
            log_date=log_date,
        )

    save_loop_index: Optional[int] = None
    for index in range(start_index, progress_index + 1):
        if line_contains_marker(lines[index], AVIDEMUX_FALLBACK_START_MARKERS):
            save_loop_index = index

    if save_loop_index is not None:
        effective_started_at = find_latest_avidemux_timestamp(
            lines,
            start_index=start_index,
            end_index=save_loop_index + 1,
            log_date=log_date,
        )
        if effective_started_at is not None:
            started_at = effective_started_at

    ended_at: Optional[datetime] = None
    explicit_end_index: Optional[int] = None
    for index in range(progress_index, start_index - 1, -1):
        if line_contains_marker(lines[index], AVIDEMUX_END_MARKERS):
            explicit_end_index = index
            break

    if explicit_end_index is not None:
        ended_at = find_latest_avidemux_timestamp(
            lines,
            start_index=start_index,
            end_index=explicit_end_index + 1,
            log_date=log_date,
        )

    if ended_at is None:
        ended_at = find_latest_avidemux_timestamp(
            lines,
            start_index=start_index,
            end_index=progress_index + 1,
            log_date=log_date,
        )

    if ended_at is None:
        ended_at = find_first_avidemux_timestamp(
            lines,
            start_index=progress_index + 1,
            end_index=end_index,
            log_date=log_date,
        )

    computed_duration_seconds = calculate_duration_seconds(started_at, ended_at)
    duration_delta_seconds = None
    if computed_duration_seconds is not None:
        duration_delta_seconds = computed_duration_seconds - elapsed_seconds

    is_complete = latest_complete is not None
    evidence = "avidemux_elapsed_100_percent" if is_complete else "avidemux_elapsed_partial"
    if started_at is not None and ended_at is not None:
        evidence = f"{evidence}_with_time_window"
    elif started_at is not None or ended_at is not None:
        evidence = f"{evidence}_with_partial_time_window"

    notes: list[str] = []
    if not is_complete:
        notes.append("Used the latest partial elapsed progress because no completed 100% line was found.")
    if started_at is None:
        notes.append("Could not infer the Avidemux start timestamp from the captured log segment.")
    if ended_at is None:
        notes.append("Could not infer the Avidemux end timestamp after the selected progress line.")
    if duration_delta_seconds is not None and abs(duration_delta_seconds) > AVIDEMUX_DURATION_WARNING_SECONDS:
        notes.append(
            "Computed wall-clock duration differs from the Avidemux elapsed value by "
            f"{abs(duration_delta_seconds):.3f} seconds."
        )

    return AvidemuxParseResult(
        duration_seconds=elapsed_seconds,
        started_at=started_at,
        ended_at=ended_at,
        computed_duration_seconds=computed_duration_seconds,
        duration_delta_seconds=duration_delta_seconds,
        evidence=evidence,
        is_complete=is_complete,
        notes=tuple(notes),
    )


def extract_avidemux_result(lines: list[str], source_path: str) -> AvidemuxParseResult:
    normalized_lines = [normalize_log_line(line) for line in lines if normalize_log_line(line)]
    start_indexes = find_avidemux_run_start_indexes(normalized_lines)
    log_date = infer_log_date(source_path)

    parsed_runs: list[AvidemuxParseResult] = []
    for offset, start_index in enumerate(start_indexes):
        end_index = start_indexes[offset + 1] if offset + 1 < len(start_indexes) else len(normalized_lines)
        parsed_run = parse_avidemux_run(
            normalized_lines,
            start_index=start_index,
            end_index=end_index,
            log_date=log_date,
        )
        if parsed_run is not None:
            parsed_runs.append(parsed_run)

    if parsed_runs:
        completed_runs = [run for run in parsed_runs if run.is_complete]
        if completed_runs:
            return completed_runs[-1]
        return parsed_runs[-1]

    elapsed_seconds, evidence = extract_avidemux_elapsed_seconds(normalized_lines)
    if elapsed_seconds is None:
        return AvidemuxParseResult(
            duration_seconds=None,
            started_at=None,
            ended_at=None,
            computed_duration_seconds=None,
            duration_delta_seconds=None,
            evidence=evidence,
            is_complete=False,
            notes=("No 'elapsed:' progress line was captured from Avidemux.",),
        )

    fallback_notes = ["Could not infer Avidemux run boundaries from the captured log segment."]
    if evidence == "avidemux_elapsed_partial":
        fallback_notes.append("Used the latest partial elapsed progress because no completed 100% line was found.")
    return AvidemuxParseResult(
        duration_seconds=elapsed_seconds,
        started_at=None,
        ended_at=None,
        computed_duration_seconds=None,
        duration_delta_seconds=None,
        evidence=evidence,
        is_complete=evidence == "avidemux_elapsed_100_percent",
        notes=tuple(fallback_notes),
    )


def build_live_log_preview(profile: LogProfile, source_path: str, lines: list[str]) -> dict[str, object]:
    if profile.parser_kind == "avidemux_elapsed":
        avidemux_result = extract_avidemux_result(lines, source_path)
        if avidemux_result.duration_seconds is None:
            if avidemux_has_export_activity(lines):
                return {
                    "progress_state": "log_activity_detected",
                    "progress_message": "Detected Avidemux export activity, but no usable elapsed marker has been captured yet.",
                    "capture_detected": False,
                    "capture_complete": False,
                    "preview_status": "none",
                    "preview_started_at": "",
                    "preview_ended_at": "",
                    "preview_duration_seconds": "",
                    "preview_evidence": "avidemux_export_activity_without_elapsed",
                }
            return {
                "progress_state": "log_activity_detected",
                "progress_message": "Detected new log content, but Avidemux export has not started yet.",
                "capture_detected": False,
                "capture_complete": False,
                "preview_status": "none",
                "preview_started_at": "",
                "preview_ended_at": "",
                "preview_duration_seconds": "",
                "preview_evidence": "",
            }

        return {
            "progress_state": "complete_data_captured" if avidemux_result.is_complete else "partial_data_captured",
            "progress_message": (
                "Captured a complete Avidemux result from the log."
                if avidemux_result.is_complete
                else "Captured partial Avidemux progress from the log; waiting for the final completion marker."
            ),
            "capture_detected": True,
            "capture_complete": avidemux_result.is_complete,
            "preview_status": "complete" if avidemux_result.is_complete else "partial",
            "preview_started_at": isoformat_or_empty(avidemux_result.started_at),
            "preview_ended_at": isoformat_or_empty(avidemux_result.ended_at),
            "preview_duration_seconds": format_seconds(avidemux_result.duration_seconds),
            "preview_evidence": avidemux_result.evidence,
        }

    started_at, ended_at = extract_latest_work_window(lines)
    if started_at and ended_at:
        return {
            "progress_state": "complete_data_captured",
            "progress_message": "Captured complete log markers from the log.",
            "capture_detected": True,
            "capture_complete": True,
            "preview_status": "complete",
            "preview_started_at": isoformat_or_empty(started_at),
            "preview_ended_at": isoformat_or_empty(ended_at),
            "preview_duration_seconds": format_seconds(calculate_duration_seconds(started_at, ended_at)),
            "preview_evidence": "work_markers",
        }
    if started_at:
        return {
            "progress_state": "partial_data_captured",
            "progress_message": "Captured a start marker from the log; waiting for the finish marker.",
            "capture_detected": True,
            "capture_complete": False,
            "preview_status": "partial",
            "preview_started_at": isoformat_or_empty(started_at),
            "preview_ended_at": "",
            "preview_duration_seconds": "",
            "preview_evidence": "work_start_marker_only",
        }
    return {
        "progress_state": "log_activity_detected",
        "progress_message": "Detected new log content, but no usable work markers have been captured yet.",
        "capture_detected": False,
        "capture_complete": False,
        "preview_status": "none",
        "preview_started_at": "",
        "preview_ended_at": "",
        "preview_duration_seconds": "",
        "preview_evidence": "",
    }


def build_log_result(
    *,
    profile: LogProfile,
    session_id: str,
    source_path: str,
    lines: list[str],
    row_status: str,
    base_notes: str,
) -> dict[str, str]:
    if profile.parser_kind == "avidemux_elapsed":
        avidemux_result = extract_avidemux_result(lines, source_path)
        if avidemux_result.duration_seconds is None:
            return make_result_row(
                software=profile.software,
                mode="log",
                session_id=session_id,
                status="no_capture",
                started_at=None,
                ended_at=None,
                source_path=source_path,
                evidence=avidemux_result.evidence,
                notes=combine_notes(base_notes, *avidemux_result.notes),
            )

        return make_result_row(
            software=profile.software,
            mode="log",
            session_id=session_id,
            status=row_status,
            started_at=avidemux_result.started_at,
            ended_at=avidemux_result.ended_at,
            source_path=source_path,
            evidence=avidemux_result.evidence,
            notes=combine_notes(base_notes, *avidemux_result.notes),
            duration_seconds_override=avidemux_result.duration_seconds,
            computed_duration_seconds_override=avidemux_result.computed_duration_seconds,
            duration_delta_seconds_override=avidemux_result.duration_delta_seconds,
        )

    started_at, ended_at = extract_latest_work_window(lines)
    if started_at and ended_at:
        evidence = "work_markers"
        notes = base_notes
    elif started_at and row_status in {"stopped", "timeout"}:
        evidence = "work_start_marker_only"
        notes = combine_notes(base_notes, "Found 'Starting work at' but no matching 'Finished work at'.")
        ended_at = None
    else:
        return make_result_row(
            software=profile.software,
            mode="log",
            session_id=session_id,
            status="no_capture",
            started_at=None,
            ended_at=None,
            source_path=source_path,
            evidence="no_log_activity",
            notes=combine_notes(base_notes, "No complete 'Starting work at'/'Finished work at' markers were captured."),
        )

    return make_result_row(
        software=profile.software,
        mode="log",
        session_id=session_id,
        status=row_status,
        started_at=started_at,
        ended_at=ended_at,
        source_path=source_path,
        evidence=evidence,
        notes=notes.strip(),
    )


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
        if profile.software == "handbrake":
            log_kind = classify_handbrake_log_file(file_path)
            if log_kind != "encode":
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


def monitor_single_log_file(
    config: dict,
    profile: LogProfile,
    stop_file: Path,
    progress_callback: ProgressCallback = None,
) -> list[dict[str, str]]:
    session_id = config["session_id"]
    options = config["options"]
    observation_started_at = now_local()
    deadline = time.monotonic() + float(options.get("max_runtime_seconds", 7200))
    poll_interval = float(options.get("poll_interval", 0.5))
    idle_seconds = float(options.get("idle_seconds", 3.0))

    log_path_override = options.get("log_path") or ""
    candidate_paths = [expand_path(log_path_override)] if log_path_override else [
        expand_path(path) for path in profile.log_file_candidates
    ]
    chosen_path = choose_existing_path(candidate_paths) or candidate_paths[0]
    source_path = str(chosen_path)
    baseline_offset = chosen_path.stat().st_size if chosen_path.exists() else 0
    captured_lines: list[str] = []
    pending_text = ""
    first_change_at: Optional[datetime] = None
    last_change_at: Optional[datetime] = None
    finalized_row: Optional[dict[str, str]] = None
    last_preview_signature: Optional[tuple[object, ...]] = None
    completion_detected_at: Optional[float] = None

    emit_progress(
        progress_callback,
        progress_state="waiting_for_log_activity",
        progress_message="Watching the log file and waiting for new log content.",
        capture_detected=False,
        capture_complete=False,
        current_source_path=source_path,
        captured_line_count=0,
        last_log_activity_at="",
        preview_status="none",
        preview_started_at="",
        preview_ended_at="",
        preview_duration_seconds="",
        preview_evidence="",
    )

    while time.monotonic() < deadline:
        if finalized_row is None:
            existing_path = choose_existing_path(candidate_paths)
            if existing_path is not None and existing_path != chosen_path:
                chosen_path = existing_path
                source_path = str(chosen_path)
                baseline_offset = 0
                emit_progress(
                    progress_callback,
                    progress_state="waiting_for_log_activity",
                    progress_message="Switched to an existing log file candidate and is waiting for new log content.",
                    capture_detected=False,
                    capture_complete=False,
                    current_source_path=source_path,
                    captured_line_count=len(captured_lines),
                    last_log_activity_at=isoformat_or_empty(last_change_at),
                    preview_status="none",
                    preview_started_at="",
                    preview_ended_at="",
                    preview_duration_seconds="",
                    preview_evidence="",
                )

            if chosen_path.exists():
                current_size = chosen_path.stat().st_size
                if current_size < baseline_offset:
                    baseline_offset = 0
                    captured_lines.clear()
                    pending_text = ""
                    first_change_at = None
                    last_change_at = None
                    completion_detected_at = None

                if current_size > baseline_offset:
                    chunk, baseline_offset = read_file_segment(chosen_path, baseline_offset)
                    lines, pending_text = split_complete_lines(pending_text, chunk)
                    lines = [line for line in lines if line.strip()]
                    if lines:
                        change_time = now_local()
                        first_change_at = first_change_at or change_time
                        last_change_at = change_time
                        captured_lines.extend(lines)
                        source_path = str(chosen_path)
                        live_preview = build_live_log_preview(
                            profile,
                            source_path,
                            finalize_captured_lines(captured_lines, pending_text),
                        )
                        preview_signature = (
                            live_preview.get("progress_state"),
                            live_preview.get("capture_detected"),
                            live_preview.get("capture_complete"),
                            live_preview.get("preview_status"),
                            live_preview.get("preview_evidence"),
                        )
                        if preview_signature != last_preview_signature:
                            emit_progress(
                                progress_callback,
                                current_source_path=source_path,
                                captured_line_count=len(captured_lines),
                                last_log_activity_at=isoformat_or_empty(change_time),
                                **live_preview,
                            )
                            last_preview_signature = preview_signature
                        if profile.parser_kind == "avidemux_elapsed":
                            if bool(live_preview.get("capture_complete", False)):
                                completion_detected_at = completion_detected_at or time.monotonic()
                            else:
                                completion_detected_at = None

            if captured_lines and profile.parser_kind == "avidemux_elapsed":
                final_lines = finalize_captured_lines(captured_lines, pending_text)
                live_preview = build_live_log_preview(profile, source_path, final_lines)
                now_monotonic = time.monotonic()
                if bool(live_preview.get("capture_complete", False)) and completion_detected_at is None:
                    completion_detected_at = now_monotonic
                if should_finalize_avidemux_completed_capture(
                    profile,
                    live_preview,
                    completion_detected_at=completion_detected_at,
                    now_monotonic=now_monotonic,
                ):
                    emit_progress(
                        progress_callback,
                        progress_state="complete_data_captured",
                        progress_message="Captured the final Avidemux elapsed marker; finalizing without waiting for trailing cleanup logs.",
                        current_source_path=source_path,
                        captured_line_count=len(captured_lines),
                        last_log_activity_at=isoformat_or_empty(last_change_at),
                        capture_detected=bool(live_preview.get("capture_detected", False)),
                        capture_complete=True,
                        preview_status=str(live_preview.get("preview_status", "")),
                        preview_started_at=str(live_preview.get("preview_started_at", "")),
                        preview_ended_at=str(live_preview.get("preview_ended_at", "")),
                        preview_duration_seconds=str(live_preview.get("preview_duration_seconds", "")),
                        preview_evidence=str(live_preview.get("preview_evidence", "")),
                    )
                    finalized_row = build_log_result(
                        profile=profile,
                        session_id=session_id,
                        source_path=source_path,
                        lines=final_lines,
                        row_status="ok",
                        base_notes=profile.notes,
                    )
                    break

            if captured_lines and last_change_at is not None:
                if (now_local() - last_change_at).total_seconds() >= idle_seconds:
                    final_lines = finalize_captured_lines(captured_lines, pending_text)
                    live_preview = build_live_log_preview(profile, source_path, final_lines)
                    if profile.parser_kind == "avidemux_elapsed" and not avidemux_has_export_activity(final_lines):
                        emit_progress(
                            progress_callback,
                            progress_state="log_activity_detected",
                            progress_message="Detected non-export Avidemux log activity; still waiting for export markers.",
                            current_source_path=source_path,
                            captured_line_count=len(captured_lines),
                            last_log_activity_at=isoformat_or_empty(last_change_at),
                            capture_detected=False,
                            capture_complete=False,
                            preview_status="none",
                            preview_started_at="",
                            preview_ended_at="",
                            preview_duration_seconds="",
                            preview_evidence="",
                        )
                        time.sleep(poll_interval)
                        continue
                    if should_wait_for_completion(profile, live_preview):
                        emit_progress(
                            progress_callback,
                            progress_state=str(live_preview.get("progress_state", "log_activity_detected")),
                            progress_message=(
                                "Avidemux export has started, but the final elapsed marker has not been flushed to the log yet."
                                if profile.software == "avidemux"
                                else "HandBrake encode log has started, but the final completion marker has not been flushed to the log yet."
                            ),
                            current_source_path=source_path,
                            captured_line_count=len(captured_lines),
                            last_log_activity_at=isoformat_or_empty(last_change_at),
                            capture_detected=bool(live_preview.get("capture_detected", False)),
                            capture_complete=False,
                            preview_status=str(live_preview.get("preview_status", "")),
                            preview_started_at=str(live_preview.get("preview_started_at", "")),
                            preview_ended_at=str(live_preview.get("preview_ended_at", "")),
                            preview_duration_seconds=str(live_preview.get("preview_duration_seconds", "")),
                            preview_evidence=str(live_preview.get("preview_evidence", "")),
                        )
                        time.sleep(poll_interval)
                        continue
                    emit_progress(
                        progress_callback,
                        progress_state=(
                            "complete_data_captured"
                            if live_preview.get("capture_complete")
                            else "finalizing_with_partial_data"
                        ),
                        progress_message=(
                            "Log activity is idle; finalizing the captured result."
                            if live_preview.get("capture_complete")
                            else "Log activity is idle; finalizing the currently captured partial data."
                        ),
                        current_source_path=source_path,
                        captured_line_count=len(captured_lines),
                        last_log_activity_at=isoformat_or_empty(last_change_at),
                        capture_detected=bool(live_preview.get("capture_detected", False)),
                        capture_complete=bool(live_preview.get("capture_complete", False)),
                        preview_status=str(live_preview.get("preview_status", "")),
                        preview_started_at=str(live_preview.get("preview_started_at", "")),
                        preview_ended_at=str(live_preview.get("preview_ended_at", "")),
                        preview_duration_seconds=str(live_preview.get("preview_duration_seconds", "")),
                        preview_evidence=str(live_preview.get("preview_evidence", "")),
                    )
                    finalized_row = build_log_result(
                        profile=profile,
                        session_id=session_id,
                        source_path=source_path,
                        lines=final_lines,
                        row_status="ok",
                        base_notes=profile.notes,
                    )
                    break

        if stop_file.exists():
            if finalized_row is not None:
                break
            settled_lines = finalize_captured_lines(captured_lines, pending_text)
            if settled_lines:
                live_preview = build_live_log_preview(profile, source_path, settled_lines)
                emit_progress(
                    progress_callback,
                    progress_state="stop_requested",
                    progress_message="Stop was requested; finalizing the currently captured log data.",
                    current_source_path=source_path,
                    captured_line_count=len(captured_lines),
                    last_log_activity_at=isoformat_or_empty(last_change_at),
                    capture_detected=bool(live_preview.get("capture_detected", False)),
                    capture_complete=bool(live_preview.get("capture_complete", False)),
                    preview_status=str(live_preview.get("preview_status", "")),
                    preview_started_at=str(live_preview.get("preview_started_at", "")),
                    preview_ended_at=str(live_preview.get("preview_ended_at", "")),
                    preview_duration_seconds=str(live_preview.get("preview_duration_seconds", "")),
                    preview_evidence=str(live_preview.get("preview_evidence", "")),
                )
                finalized_row = build_log_result(
                    profile=profile,
                    session_id=session_id,
                    source_path=source_path,
                    lines=settled_lines,
                    row_status="stopped",
                    base_notes=f"{profile.notes} Stop was requested before the log capture went idle.".strip(),
                )
            break

        time.sleep(poll_interval)

    if finalized_row is not None:
        return [finalized_row]

    settled_lines = finalize_captured_lines(captured_lines, pending_text)
    if settled_lines:
        return [
            build_log_result(
                profile=profile,
                session_id=session_id,
                source_path=source_path,
                lines=settled_lines,
                row_status="timeout",
                base_notes=f"{profile.notes} Monitor reached max runtime before the log capture went idle.".strip(),
            )
        ]

    note = (
        f"{profile.notes} Stop was requested before any new log content was captured."
        if stop_file.exists()
        else f"{profile.notes} No new log content was captured."
    )
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
            notes=note.strip(),
        )
    ]


def monitor_log_directory(
    config: dict,
    profile: LogProfile,
    stop_file: Path,
    progress_callback: ProgressCallback = None,
) -> list[dict[str, str]]:
    session_id = config["session_id"]
    options = config["options"]
    observation_started_at = now_local()
    deadline = time.monotonic() + float(options.get("max_runtime_seconds", 7200))
    poll_interval = float(options.get("poll_interval", 0.5))
    idle_seconds = float(options.get("idle_seconds", 3.0))
    allow_ignored_logs = bool(options.get("allow_activity_log", False))

    log_dir_override = options.get("log_dir") or ""
    candidate_dirs = [expand_path(log_dir_override)] if log_dir_override else [
        expand_path(path) for path in profile.log_dir_candidates
    ]
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
    pending_text = ""
    first_change_at: Optional[datetime] = None
    last_change_at: Optional[datetime] = None
    finalized_row: Optional[dict[str, str]] = None
    last_preview_signature: Optional[tuple[object, ...]] = None

    emit_progress(
        progress_callback,
        progress_state="waiting_for_log_activity",
        progress_message="Watching the log directory and waiting for a new job log file.",
        capture_detected=False,
        capture_complete=False,
        current_source_path=str(chosen_dir),
        captured_line_count=0,
        last_log_activity_at="",
        preview_status="none",
        preview_started_at="",
        preview_ended_at="",
        preview_duration_seconds="",
        preview_evidence="",
    )

    while time.monotonic() < deadline:
        if finalized_row is None:
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
                    emit_progress(
                        progress_callback,
                        progress_state="log_file_selected",
                        progress_message="Detected a candidate job log file and is waiting for new content.",
                        capture_detected=False,
                        capture_complete=False,
                        current_source_path=str(active_file),
                        captured_line_count=len(captured_lines),
                        last_log_activity_at=isoformat_or_empty(last_change_at),
                        preview_status="none",
                        preview_started_at="",
                        preview_ended_at="",
                        preview_duration_seconds="",
                        preview_evidence="",
                    )

            if active_file is not None and active_file.exists():
                current_size = active_file.stat().st_size
                if current_size < active_offset:
                    active_offset = 0
                    captured_lines.clear()
                    pending_text = ""
                    first_change_at = None
                    last_change_at = None

                if current_size > active_offset:
                    chunk, active_offset = read_file_segment(active_file, active_offset)
                    lines, pending_text = split_complete_lines(pending_text, chunk)
                    lines = [line for line in lines if line.strip()]
                    if lines:
                        change_time = now_local()
                        first_change_at = first_change_at or change_time
                        last_change_at = change_time
                        captured_lines.extend(lines)
                        live_preview = build_live_log_preview(
                            profile,
                            str(active_file),
                            finalize_captured_lines(captured_lines, pending_text),
                        )
                        preview_signature = (
                            live_preview.get("progress_state"),
                            live_preview.get("capture_detected"),
                            live_preview.get("capture_complete"),
                            live_preview.get("preview_status"),
                            live_preview.get("preview_evidence"),
                        )
                        if preview_signature != last_preview_signature:
                            emit_progress(
                                progress_callback,
                                current_source_path=str(active_file),
                                captured_line_count=len(captured_lines),
                                last_log_activity_at=isoformat_or_empty(change_time),
                                **live_preview,
                            )
                            last_preview_signature = preview_signature

            for file_path in files:
                try:
                    baseline_sizes[file_path] = file_path.stat().st_size
                except OSError:
                    continue

            if captured_lines and last_change_at is not None:
                if (now_local() - last_change_at).total_seconds() >= idle_seconds:
                    final_lines = finalize_captured_lines(captured_lines, pending_text)
                    live_preview = build_live_log_preview(profile, str(active_file), final_lines)
                    if profile.parser_kind == "avidemux_elapsed" and not avidemux_has_export_activity(final_lines):
                        emit_progress(
                            progress_callback,
                            progress_state="log_activity_detected",
                            progress_message="Detected non-export Avidemux log activity; still waiting for export markers.",
                            current_source_path=str(active_file),
                            captured_line_count=len(captured_lines),
                            last_log_activity_at=isoformat_or_empty(last_change_at),
                            capture_detected=False,
                            capture_complete=False,
                            preview_status="none",
                            preview_started_at="",
                            preview_ended_at="",
                            preview_duration_seconds="",
                            preview_evidence="",
                        )
                        time.sleep(poll_interval)
                        continue
                    if should_wait_for_completion(profile, live_preview):
                        emit_progress(
                            progress_callback,
                            progress_state=str(live_preview.get("progress_state", "log_activity_detected")),
                            progress_message=(
                                "Avidemux export has started, but the final elapsed marker has not been flushed to the log yet."
                                if profile.software == "avidemux"
                                else "HandBrake encode log has started, but the final completion marker has not been flushed to the log yet."
                            ),
                            current_source_path=str(active_file),
                            captured_line_count=len(captured_lines),
                            last_log_activity_at=isoformat_or_empty(last_change_at),
                            capture_detected=bool(live_preview.get("capture_detected", False)),
                            capture_complete=False,
                            preview_status=str(live_preview.get("preview_status", "")),
                            preview_started_at=str(live_preview.get("preview_started_at", "")),
                            preview_ended_at=str(live_preview.get("preview_ended_at", "")),
                            preview_duration_seconds=str(live_preview.get("preview_duration_seconds", "")),
                            preview_evidence=str(live_preview.get("preview_evidence", "")),
                        )
                        time.sleep(poll_interval)
                        continue
                    emit_progress(
                        progress_callback,
                        progress_state=(
                            "complete_data_captured"
                            if live_preview.get("capture_complete")
                            else "finalizing_with_partial_data"
                        ),
                        progress_message=(
                            "Log activity is idle; finalizing the captured result."
                            if live_preview.get("capture_complete")
                            else "Log activity is idle; finalizing the currently captured partial data."
                        ),
                        current_source_path=str(active_file),
                        captured_line_count=len(captured_lines),
                        last_log_activity_at=isoformat_or_empty(last_change_at),
                        capture_detected=bool(live_preview.get("capture_detected", False)),
                        capture_complete=bool(live_preview.get("capture_complete", False)),
                        preview_status=str(live_preview.get("preview_status", "")),
                        preview_started_at=str(live_preview.get("preview_started_at", "")),
                        preview_ended_at=str(live_preview.get("preview_ended_at", "")),
                        preview_duration_seconds=str(live_preview.get("preview_duration_seconds", "")),
                        preview_evidence=str(live_preview.get("preview_evidence", "")),
                    )
                    finalized_row = build_log_result(
                        profile=profile,
                        session_id=session_id,
                        source_path=str(active_file),
                        lines=final_lines,
                        row_status="ok",
                        base_notes=profile.notes,
                    )
                    break

        if stop_file.exists():
            if finalized_row is not None:
                break
            settled_lines = finalize_captured_lines(captured_lines, pending_text)
            if settled_lines and active_file is not None:
                live_preview = build_live_log_preview(profile, str(active_file), settled_lines)
                emit_progress(
                    progress_callback,
                    progress_state="stop_requested",
                    progress_message="Stop was requested; finalizing the currently captured log data.",
                    current_source_path=str(active_file),
                    captured_line_count=len(captured_lines),
                    last_log_activity_at=isoformat_or_empty(last_change_at),
                    capture_detected=bool(live_preview.get("capture_detected", False)),
                    capture_complete=bool(live_preview.get("capture_complete", False)),
                    preview_status=str(live_preview.get("preview_status", "")),
                    preview_started_at=str(live_preview.get("preview_started_at", "")),
                    preview_ended_at=str(live_preview.get("preview_ended_at", "")),
                    preview_duration_seconds=str(live_preview.get("preview_duration_seconds", "")),
                    preview_evidence=str(live_preview.get("preview_evidence", "")),
                )
                finalized_row = build_log_result(
                    profile=profile,
                    session_id=session_id,
                    source_path=str(active_file),
                    lines=settled_lines,
                    row_status="stopped",
                    base_notes=f"{profile.notes} Stop was requested before the log capture went idle.".strip(),
                )
            break

        time.sleep(poll_interval)

    if finalized_row is not None:
        return [finalized_row]

    settled_lines = finalize_captured_lines(captured_lines, pending_text)
    if settled_lines and active_file is not None:
        return [
            build_log_result(
                profile=profile,
                session_id=session_id,
                source_path=str(active_file),
                lines=settled_lines,
                row_status="timeout",
                base_notes=f"{profile.notes} Monitor reached max runtime before the log capture went idle.".strip(),
            )
        ]

    note = (
        f"{profile.notes} Stop was requested before any new per-job log file was captured."
        if stop_file.exists()
        else f"{profile.notes} No new per-job log file was captured."
    )
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
            notes=note.strip(),
        )
    ]


def run_log_listener(
    config: dict,
    stop_file: Path,
    progress_callback: ProgressCallback = None,
) -> list[dict[str, str]]:
    software = config["software"]
    profile = LOG_PROFILES[software]
    if profile.log_file_candidates:
        return monitor_single_log_file(config, profile, stop_file, progress_callback=progress_callback)
    return monitor_log_directory(config, profile, stop_file, progress_callback=progress_callback)
