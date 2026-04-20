from __future__ import annotations

import csv
import json
import os
import re
from datetime import datetime
from pathlib import Path
from typing import Optional


LOCAL_TZ = datetime.now().astimezone().tzinfo
COUNTED_OUTPUT_PATTERN = re.compile(r"^(?P<base>.+?)_run(?P<index>\d+)$", re.IGNORECASE)
CSV_HEADERS = [
    "software",
    "mode",
    "session_id",
    "test_name",
    "status",
    "started_at",
    "ended_at",
    "duration_seconds",
    "computed_duration_seconds",
    "duration_delta_seconds",
    "main_process_name",
    "main_process_pid",
    "worker_process_name",
    "worker_process_pid",
    "source_path",
    "evidence",
    "notes",
]
FINAL_SESSION_STATUSES = {"completed", "completed_with_warnings", "failed"}


def now_local() -> datetime:
    return datetime.now().astimezone()


POWERSHELL_ENV_PATTERN = re.compile(r"\$env:(?P<name>[A-Za-z_][A-Za-z0-9_]*)", re.IGNORECASE)


def expand_powershell_env_vars(raw_path: str) -> str:
    def replace(match: re.Match[str]) -> str:
        env_name = match.group("name")
        return os.environ.get(env_name, match.group(0))

    return POWERSHELL_ENV_PATTERN.sub(replace, raw_path)


def expand_path(raw_path: str) -> Path:
    normalized = expand_powershell_env_vars(raw_path)
    expanded = os.path.expandvars(os.path.expanduser(normalized))
    return Path(expanded).resolve(strict=False)


def ensure_parent_directory(path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)


def split_counted_output_stem(stem: str) -> tuple[str, Optional[int]]:
    match = COUNTED_OUTPUT_PATTERN.fullmatch(stem)
    if not match:
        return stem, None
    return match.group("base"), int(match.group("index"))


def extract_run_label_from_path(path: Path) -> str:
    _, run_index = split_counted_output_stem(path.stem)
    if run_index is None:
        return ""
    return f"run{run_index:03d}"


def resolve_counted_output_path(path: Path) -> tuple[Path, int]:
    base_stem, explicit_index = split_counted_output_stem(path.stem)
    if explicit_index is not None:
        ensure_parent_directory(path)
        return path, explicit_index

    ensure_parent_directory(path)
    max_index = 0
    try:
        for candidate in path.parent.glob(f"{base_stem}_run*{path.suffix}"):
            if not candidate.is_file():
                continue
            candidate_base_stem, candidate_index = split_counted_output_stem(candidate.stem)
            if candidate_base_stem == base_stem and candidate_index is not None:
                max_index = max(max_index, candidate_index)
    except OSError:
        pass

    next_index = max_index + 1
    counted_name = f"{base_stem}_run{next_index:03d}{path.suffix}"
    return path.with_name(counted_name), next_index


def isoformat_or_empty(value: Optional[datetime]) -> str:
    if value is None:
        return ""
    return value.isoformat(timespec="milliseconds")


def calculate_duration_seconds(started_at: Optional[datetime], ended_at: Optional[datetime]) -> Optional[float]:
    if started_at is None or ended_at is None:
        return None
    return max((ended_at - started_at).total_seconds(), 0.0)


def format_seconds(value: Optional[float], *, clamp_non_negative: bool = True) -> str:
    if value is None:
        return ""
    if clamp_non_negative:
        value = max(value, 0.0)
    return f"{value:.3f}"


def format_duration(started_at: Optional[datetime], ended_at: Optional[datetime]) -> str:
    return format_seconds(calculate_duration_seconds(started_at, ended_at))


def make_result_row(
    *,
    software: str,
    mode: str,
    session_id: str,
    test_name: str = "",
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
    duration_seconds_override: Optional[float] = None,
    computed_duration_seconds_override: Optional[float] = None,
    duration_delta_seconds_override: Optional[float] = None,
) -> dict[str, str]:
    computed_duration_seconds = (
        computed_duration_seconds_override
        if computed_duration_seconds_override is not None
        else calculate_duration_seconds(started_at, ended_at)
    )
    duration_seconds = (
        duration_seconds_override
        if duration_seconds_override is not None
        else computed_duration_seconds
    )
    duration_delta_seconds = duration_delta_seconds_override
    if (
        duration_delta_seconds is None
        and duration_seconds is not None
        and computed_duration_seconds is not None
    ):
        duration_delta_seconds = computed_duration_seconds - duration_seconds

    return {
        "software": software,
        "mode": mode,
        "session_id": session_id,
        "test_name": test_name,
        "status": status,
        "started_at": isoformat_or_empty(started_at),
        "ended_at": isoformat_or_empty(ended_at),
        "duration_seconds": format_seconds(duration_seconds),
        "computed_duration_seconds": format_seconds(computed_duration_seconds),
        "duration_delta_seconds": format_seconds(duration_delta_seconds, clamp_non_negative=False),
        "main_process_name": main_process_name,
        "main_process_pid": main_process_pid,
        "worker_process_name": worker_process_name,
        "worker_process_pid": worker_process_pid,
        "source_path": source_path,
        "evidence": evidence,
        "notes": notes,
    }


def write_rows(csv_path: Path, rows: list[dict[str, str]]) -> None:
    ensure_parent_directory(csv_path)
    existing_rows: list[dict[str, str]] = []
    should_rewrite = False

    if csv_path.exists() and csv_path.stat().st_size > 0:
        with csv_path.open("r", newline="", encoding="utf-8-sig") as handle:
            reader = csv.DictReader(handle)
            existing_headers = reader.fieldnames or []
            if existing_headers != CSV_HEADERS:
                existing_rows = [
                    {header: row.get(header, "") for header in CSV_HEADERS}
                    for row in reader
                ]
                should_rewrite = True

    mode = "w" if should_rewrite or not csv_path.exists() or csv_path.stat().st_size == 0 else "a"
    with csv_path.open(mode, newline="", encoding="utf-8-sig") as handle:
        writer = csv.DictWriter(handle, fieldnames=CSV_HEADERS)
        if mode == "w":
            writer.writeheader()
            if existing_rows:
                writer.writerows(existing_rows)
        writer.writerows(rows)


def save_json(path: Path, payload: dict) -> None:
    ensure_parent_directory(path)
    with path.open("w", encoding="utf-8") as handle:
        json.dump(payload, handle, ensure_ascii=True, indent=2, sort_keys=True)


def load_json(path: Path) -> dict:
    with path.open("r", encoding="utf-8") as handle:
        return json.load(handle)


def default_runtime_root(base_dir: Optional[Path] = None) -> Path:
    workspace_root = (base_dir or Path(__file__).resolve().parent).resolve(strict=False)
    root = workspace_root / "trash" / ".monitor_sessions"
    return root


def resolve_session_paths(runtime_root: Path, session_id: str) -> dict[str, Path]:
    session_dir = runtime_root / session_id
    return {
        "session_dir": session_dir,
        "state_path": session_dir / "session.json",
        "stop_path": session_dir / "stop.signal",
        "session_output_path": session_dir / "results.csv",
    }


def summarize_result_statuses(rows: list[dict[str, str]]) -> list[str]:
    return sorted({row["status"] for row in rows}) if rows else []


def append_results(session_output_path: Path, aggregate_output_path: Path, rows: list[dict[str, str]]) -> None:
    if not rows:
        return
    write_rows(session_output_path, rows)
    if aggregate_output_path != session_output_path:
        write_rows(aggregate_output_path, rows)
