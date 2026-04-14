"""Register Blender render handlers and append timing results to a CSV file.

Examples:
    blender -b project.blend --python blender_listener.py -- results\blender_results.csv blender_001
    blender project.blend --python blender_listener.py -- results\blender_results.csv blender_001

Extra args after `--`:
    1. csv output path
    2. session id
    3. optional json status path
"""

from __future__ import annotations

import csv
import json
import os
import re
import sys
import uuid
from datetime import datetime
from pathlib import Path
from typing import Optional

import bpy
from bpy.app.handlers import persistent


CSV_HEADERS = [
    "software",
    "mode",
    "session_id",
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
COUNTED_OUTPUT_PATTERN = re.compile(r"^(?P<base>.+?)_run(?P<index>\d+)$", re.IGNORECASE)

STATE = {
    "output_path": None,
    "session_id": None,
    "started_at": None,
    "status_path": None,
    "row_count": 0,
    "last_row_status": "",
}


def now_local() -> datetime:
    return datetime.now().astimezone()


def ensure_parent_directory(path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)


def split_counted_output_stem(stem: str) -> tuple[str, Optional[int]]:
    match = COUNTED_OUTPUT_PATTERN.fullmatch(stem)
    if not match:
        return stem, None
    return match.group("base"), int(match.group("index"))


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
        json.dump(payload, handle, ensure_ascii=False, indent=2, sort_keys=True)


def parse_extra_args() -> list[str]:
    if "--" in sys.argv:
        return sys.argv[sys.argv.index("--") + 1 :]
    return []


def candidate_script_paths() -> list[Path]:
    candidates: list[Path] = []

    raw_file = globals().get("__file__")
    if isinstance(raw_file, str) and raw_file.strip():
        candidates.append(Path(raw_file.strip()))

    try:
        space_data = getattr(bpy.context, "space_data", None)
        active_text = getattr(space_data, "text", None)
        active_text_path = getattr(active_text, "filepath", "")
        if active_text_path:
            candidates.append(Path(active_text_path))
    except Exception:
        pass

    try:
        for text_block in bpy.data.texts:
            text_path = getattr(text_block, "filepath", "")
            if text_path:
                candidates.append(Path(text_path))
    except Exception:
        pass

    return candidates


def normalize_runtime_base(path: Path) -> Path:
    raw_value = str(path.expanduser())
    embedded_file_match = re.match(r"(?is)^(.*?\.(?:pyw?|blend))(?:[\\/].*)?$", raw_value)
    if embedded_file_match:
        resolved_file_path = Path(embedded_file_match.group(1)).resolve(strict=False)
        return resolved_file_path.parent

    resolved = Path(raw_value).resolve(strict=False)
    if resolved.suffix.lower() in {".py", ".pyw", ".blend"}:
        return resolved.parent
    return resolved


def base_runtime_dir() -> Path:
    for candidate in candidate_script_paths():
        normalized = normalize_runtime_base(candidate)
        if normalized.exists():
            return normalized
    for candidate in candidate_script_paths():
        return normalize_runtime_base(candidate)
    if bpy.data.filepath:
        return normalize_runtime_base(Path(bpy.data.filepath))
    return normalize_runtime_base(Path.cwd())


def resolve_runtime_path(raw_value: str) -> Path:
    raw_path = Path(raw_value).expanduser()
    if raw_path.is_absolute():
        return raw_path.resolve(strict=False)
    return (base_runtime_dir() / raw_path).resolve(strict=False)


def get_output_path() -> Path:
    if STATE["output_path"] is not None:
        return STATE["output_path"]

    extra_args = parse_extra_args()
    output_value = extra_args[0] if extra_args else "results/csv/blender.csv"
    requested_output_path = resolve_runtime_path(output_value)
    output_path, _ = resolve_counted_output_path(requested_output_path)
    STATE["output_path"] = output_path
    return output_path


def get_session_id() -> str:
    if STATE["session_id"] is not None:
        return STATE["session_id"]

    extra_args = parse_extra_args()
    STATE["session_id"] = extra_args[1] if len(extra_args) >= 2 else str(uuid.uuid4())
    return STATE["session_id"]


def get_status_path() -> Path:
    if STATE["status_path"] is not None:
        return STATE["status_path"]

    extra_args = parse_extra_args()
    if len(extra_args) >= 3 and extra_args[2]:
        status_path = resolve_runtime_path(extra_args[2])
    else:
        output_path = get_output_path()
        session_id = get_session_id()
        status_path = output_path.parent / f"{output_path.stem}.{session_id}.status.json"
    ensure_parent_directory(status_path)
    STATE["status_path"] = status_path
    return status_path


def blend_source_path() -> str:
    return bpy.data.filepath or "<unsaved>"


def append_row(row: dict[str, str]) -> None:
    write_rows(get_output_path(), [row])
    STATE["row_count"] = int(STATE["row_count"]) + 1
    STATE["last_row_status"] = row.get("status", "")


def reset_started_at() -> None:
    STATE["started_at"] = None


def process_pid() -> str:
    try:
        return str(os.getpid())
    except OSError:
        return ""


def build_row(status: str, started_at: Optional[datetime], ended_at: Optional[datetime], notes: str) -> dict[str, str]:
    return make_result_row(
        software="blender",
        mode="hook",
        session_id=get_session_id(),
        status=status,
        started_at=started_at,
        ended_at=ended_at,
        main_process_name="blender.exe",
        main_process_pid=process_pid(),
        worker_process_name="bpy.app.handlers",
        worker_process_pid="",
        source_path=blend_source_path(),
        evidence="render_handlers",
        notes=notes,
    )


def emit_progress(
    *,
    progress_state: str,
    progress_message: str,
    status: str,
    capture_detected: bool,
    capture_complete: bool,
    preview_status: str,
    preview_started_at: str = "",
    preview_ended_at: str = "",
    preview_duration_seconds: str = "",
    preview_evidence: str = "",
) -> None:
    payload = {
        "session_id": get_session_id(),
        "software": "blender",
        "monitor_type": "blender",
        "status": status,
        "created_at": isoformat_or_empty(now_local()),
        "progress_state": progress_state,
        "progress_message": progress_message,
        "progress_updated_at": isoformat_or_empty(now_local()),
        "capture_detected": capture_detected,
        "capture_complete": capture_complete,
        "current_source_path": blend_source_path(),
        "captured_line_count": int(STATE["row_count"]),
        "preview_status": preview_status,
        "preview_started_at": preview_started_at,
        "preview_ended_at": preview_ended_at,
        "preview_duration_seconds": preview_duration_seconds,
        "preview_evidence": preview_evidence,
        "session_output_path": str(get_output_path()),
        "status_output_path": str(get_status_path()),
        "last_result_status": str(STATE["last_row_status"]),
    }
    save_json(get_status_path(), payload)
    print(
        f"[{now_local().isoformat(timespec='seconds')}] {progress_state}: {progress_message}",
        flush=True,
    )


@persistent
def on_render_init(scene):
    del scene
    get_session_id()
    started_at = now_local()
    STATE["started_at"] = started_at
    emit_progress(
        progress_state="active_render_captured",
        progress_message="Detected Blender render start and is measuring until completion or cancel.",
        status="running",
        capture_detected=True,
        capture_complete=False,
        preview_status="partial",
        preview_started_at=isoformat_or_empty(started_at),
        preview_ended_at="",
        preview_duration_seconds="",
        preview_evidence="render_handlers",
    )


@persistent
def on_render_complete(scene):
    del scene
    ended_at = now_local()
    started_at = STATE["started_at"] or ended_at
    row = build_row(
        status="ok",
        started_at=started_at,
        ended_at=ended_at,
        notes="Captured from bpy.app.handlers.render_init/render_complete.",
    )
    append_row(row)
    emit_progress(
        progress_state="complete_data_captured",
        progress_message="Captured a completed Blender render result and wrote it to the csv output.",
        status="completed",
        capture_detected=True,
        capture_complete=True,
        preview_status="complete",
        preview_started_at=row["started_at"],
        preview_ended_at=row["ended_at"],
        preview_duration_seconds=row["duration_seconds"],
        preview_evidence=row["evidence"],
    )
    reset_started_at()


@persistent
def on_render_cancel(scene):
    del scene
    ended_at = now_local()
    started_at = STATE["started_at"] or ended_at
    row = build_row(
        status="cancelled",
        started_at=started_at,
        ended_at=ended_at,
        notes="Captured from bpy.app.handlers.render_init/render_cancel.",
    )
    append_row(row)
    emit_progress(
        progress_state="cancelled_render_captured",
        progress_message="Captured a cancelled Blender render result and wrote it to the csv output.",
        status="cancelled",
        capture_detected=True,
        capture_complete=True,
        preview_status="complete",
        preview_started_at=row["started_at"],
        preview_ended_at=row["ended_at"],
        preview_duration_seconds=row["duration_seconds"],
        preview_evidence=row["evidence"],
    )
    reset_started_at()


def remove_handler(handler_list, handler) -> None:
    while handler in handler_list:
        handler_list.remove(handler)


def register() -> None:
    get_output_path()
    get_session_id()
    get_status_path()
    remove_handler(bpy.app.handlers.render_init, on_render_init)
    remove_handler(bpy.app.handlers.render_complete, on_render_complete)
    remove_handler(bpy.app.handlers.render_cancel, on_render_cancel)
    bpy.app.handlers.render_init.append(on_render_init)
    bpy.app.handlers.render_complete.append(on_render_complete)
    bpy.app.handlers.render_cancel.append(on_render_cancel)
    emit_progress(
        progress_state="waiting_for_render_activity",
        progress_message="Blender listener registered. Waiting for render activity.",
        status="running",
        capture_detected=False,
        capture_complete=False,
        preview_status="none",
    )
    print(f"runtime_base={base_runtime_dir()}", flush=True)
    print(f"csv_output={get_output_path()}", flush=True)
    print(f"status_output={get_status_path()}", flush=True)


def unregister() -> None:
    remove_handler(bpy.app.handlers.render_init, on_render_init)
    remove_handler(bpy.app.handlers.render_complete, on_render_complete)
    remove_handler(bpy.app.handlers.render_cancel, on_render_cancel)
    emit_progress(
        progress_state="listener_unregistered",
        progress_message="Blender listener was unregistered.",
        status="stopped",
        capture_detected=bool(STATE["row_count"]),
        capture_complete=bool(STATE["row_count"]),
        preview_status="complete" if STATE["row_count"] else "none",
    )


if __name__ == "__main__":
    register()
