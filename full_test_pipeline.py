from __future__ import annotations

import argparse
from dataclasses import dataclass
from datetime import datetime
import json
import logging
from pathlib import Path
import re
import time
import traceback

from workspace_runtime import configure_workspace_runtime

configure_workspace_runtime()

from automation_components import (
    DEFAULT_AI_TURBO_TITLE_RE,
    DEFAULT_PIPELINE_WORKLOAD_NAME,
    REPO_ROOT,
    SOFTWARE_SPECS,
    AiTurboEngineController,
    MonitorBridge,
    SoftwareLauncher,
    resolve_input_video,
)
from software_operations import OPERATION_PROFILES, build_operator
from xlsx_report_generator import generate_xlsx_report


logger = logging.getLogger(__name__)


DEFAULT_SOFTWARE_ORDER = ("shotcut", "kdenlive", "shutter_encoder", "avidemux", "handbrake", "blender")
MONITOR_STOP_AFTER_AUTOMATION_SOFTWARE = frozenset({"avidemux", "kdenlive"})
PIPELINE_STATE_SCHEMA_VERSION = 1
SHOTCUT_SEQUENCE_NAME = "shotcut_multi_export"
SHOTCUT_MULTI_CASE_SMALL_NAME = "4k(1大1小)"
SHOTCUT_MULTI_CASE_DOUBLE_NAME = "4k(2大2小)"
SHOTCUT_MULTI_CASE_SMALL_SLUG = "4k_1big_1small"
SHOTCUT_MULTI_CASE_DOUBLE_SLUG = "4k_2big_2small"
SHOTCUT_MULTI_CASE_SMALL_NAME = "4k(1\u59271\u5c0f)"
SHOTCUT_MULTI_CASE_DOUBLE_NAME = "4k(2\u59272\u5c0f)"
KDENLIVE_SEQUENCE_NAME = "kdenlive_multi_render"
MULTI_CLIP_SEQUENCE_SOFTWARE = frozenset({"shotcut", "kdenlive"})


@dataclass(frozen=True)
class PipelinePaths:
    root_dir: Path
    csv_dir: Path
    export_dir: Path
    report_dir: Path
    cases_dir: Path
    manifest_path: Path


@dataclass(frozen=True)
class PipelineCase:
    case_id: str
    software: str
    variant: str
    ai_turbo_enabled: bool
    case_name: str = ""
    case_slug: str = ""
    sequence_name: str = ""
    sequence_order: int = 0


@dataclass(frozen=True)
class PipelineCasePaths:
    case_dir: Path
    state_path: Path
    log_path: Path
    traceback_path: Path


@dataclass(frozen=True)
class CaseExecutionResult:
    csv_path: Path
    requested_csv_path: Path
    export_output_path: Path | None = None
    session_id: str = ""
    monitor_status: str = ""
    monitor_state_path: Path | None = None
    session_output_path: Path | None = None
    worker_stdout_path: Path | None = None
    worker_stderr_path: Path | None = None


def _now_timestamp() -> str:
    return datetime.now().isoformat(timespec="seconds")


def build_case_id(software: str, variant: str, *, case_slug: str = "") -> str:
    pieces = [slugify(software)]
    if case_slug:
        pieces.append(slugify(case_slug))
    pieces.append(slugify(variant))
    return "__".join(pieces)


def resolve_selected_software(args: argparse.Namespace) -> tuple[str, ...]:
    selected_software = tuple(args.software or DEFAULT_SOFTWARE_ORDER)
    excluded_software = set(getattr(args, "exclude_software", []) or [])
    if excluded_software:
        selected_software = tuple(software for software in selected_software if software not in excluded_software)
    assert selected_software, "At least one software entry is required."
    for software in selected_software:
        assert software in SOFTWARE_SPECS, f"Unsupported software: {software}"
    return selected_software


def resolve_ai_turbo_sequence_software(args: argparse.Namespace) -> str:
    requested_software = str(getattr(args, "ai_turbo_sequence_software", "") or "").strip()
    if requested_software:
        assert requested_software in SOFTWARE_SPECS, f"Unsupported AI Turbo sequence software: {requested_software}"
        return requested_software
    selected_software = resolve_selected_software(args)
    assert len(selected_software) == 1, (
        "Standalone AI Turbo sequence needs exactly one target software. "
        "Use --ai-turbo-sequence-software or narrow --software to one entry."
    )
    return selected_software[0]


def should_expand_shotcut_cases(input_video_path: Path) -> bool:
    return input_video_path.name.casefold() == "4k_big.mp4" and input_video_path.with_name("4K_small.mp4").exists()


def should_expand_multi_clip_cases(input_video_path: Path) -> bool:
    return should_expand_shotcut_cases(input_video_path)


def build_pipeline_cases(
    selected_software: tuple[str, ...],
    passes: list[tuple[str, bool]],
    *,
    workload_name: str,
    input_video_path: Path,
) -> tuple[PipelineCase, ...]:
    cases: list[PipelineCase] = []
    expand_multi_clip_cases = should_expand_multi_clip_cases(input_video_path)
    for variant, enable_turbo in passes:
        for software in selected_software:
            if software in MULTI_CLIP_SEQUENCE_SOFTWARE and expand_multi_clip_cases:
                sequence_name = SHOTCUT_SEQUENCE_NAME if software == "shotcut" else KDENLIVE_SEQUENCE_NAME
                cases.extend(
                    [
                        PipelineCase(
                            case_id=build_case_id(software, variant),
                            software=software,
                            variant=variant,
                            ai_turbo_enabled=enable_turbo,
                            case_name=workload_name,
                            case_slug="",
                            sequence_name=sequence_name,
                            sequence_order=1,
                        ),
                        PipelineCase(
                            case_id=build_case_id(software, variant, case_slug=SHOTCUT_MULTI_CASE_SMALL_SLUG),
                            software=software,
                            variant=variant,
                            ai_turbo_enabled=enable_turbo,
                            case_name=SHOTCUT_MULTI_CASE_SMALL_NAME,
                            case_slug=SHOTCUT_MULTI_CASE_SMALL_SLUG,
                            sequence_name=sequence_name,
                            sequence_order=2,
                        ),
                        PipelineCase(
                            case_id=build_case_id(software, variant, case_slug=SHOTCUT_MULTI_CASE_DOUBLE_SLUG),
                            software=software,
                            variant=variant,
                            ai_turbo_enabled=enable_turbo,
                            case_name=SHOTCUT_MULTI_CASE_DOUBLE_NAME,
                            case_slug=SHOTCUT_MULTI_CASE_DOUBLE_SLUG,
                            sequence_name=sequence_name,
                            sequence_order=3,
                        ),
                    ]
                )
                continue
            cases.append(
                PipelineCase(
                    case_id=build_case_id(software, variant),
                    software=software,
                    variant=variant,
                    ai_turbo_enabled=enable_turbo,
                    case_name=workload_name,
                )
            )
    return tuple(cases)


def build_case_paths(pipeline_paths: PipelinePaths, case: PipelineCase) -> PipelineCasePaths:
    case_dir = pipeline_paths.cases_dir / case.case_id
    return PipelineCasePaths(
        case_dir=case_dir,
        state_path=case_dir / "case_state.json",
        log_path=case_dir / "case_events.jsonl",
        traceback_path=case_dir / "traceback.txt",
    )


def _write_json(path: Path, payload: dict[str, object]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, indent=2, ensure_ascii=True), encoding="utf-8")


def _read_json(path: Path) -> dict[str, object]:
    return json.loads(path.read_text(encoding="utf-8"))


def build_requested_csv_path(
    pipeline_paths: PipelinePaths,
    *,
    software: str,
    workload_name: str,
    variant: str,
    case_slug: str = "",
) -> Path:
    name_token = case_slug or slugify(workload_name)
    return (pipeline_paths.csv_dir / f"{software}_{name_token}_{variant}.csv").resolve(strict=False)


def build_export_output_path(
    pipeline_paths: PipelinePaths,
    *,
    software: str,
    workload_name: str,
    variant: str,
    case_slug: str = "",
) -> Path:
    export_profile = OPERATION_PROFILES[software]
    name_token = case_slug or slugify(workload_name)
    return (
        pipeline_paths.export_dir / f"{software}_{name_token}_{variant}{export_profile.output_suffix}"
    ).resolve(strict=False)


def load_pipeline_manifest(run_root: Path) -> dict[str, object]:
    manifest_path = run_root / "pipeline_manifest.json"
    if not manifest_path.exists():
        return {}
    try:
        return _read_json(manifest_path)
    except (json.JSONDecodeError, OSError):
        return {}


def write_pipeline_manifest(
    pipeline_paths: PipelinePaths,
    *,
    workload_name: str,
    input_video_path: Path,
    selected_software: tuple[str, ...],
    passes: list[tuple[str, bool]],
    cases: tuple[PipelineCase, ...],
) -> None:
    existing_manifest = load_pipeline_manifest(pipeline_paths.root_dir)
    created_at = str(existing_manifest.get("created_at") or _now_timestamp())
    manifest = {
        "schema_version": PIPELINE_STATE_SCHEMA_VERSION,
        "created_at": created_at,
        "updated_at": _now_timestamp(),
        "workload_name": workload_name,
        "run_root": str(pipeline_paths.root_dir),
        "input_video_path": str(input_video_path),
        "selected_software": list(selected_software),
        "passes": [
            {
                "variant": variant,
                "ai_turbo_enabled": enable_turbo,
            }
            for variant, enable_turbo in passes
        ],
        "cases": [
            {
                "case_id": case.case_id,
                "software": case.software,
                "variant": case.variant,
                "ai_turbo_enabled": case.ai_turbo_enabled,
                "case_name": case.case_name,
                "case_slug": case.case_slug,
                "sequence_name": case.sequence_name,
                "sequence_order": case.sequence_order,
                "case_state_path": str(build_case_paths(pipeline_paths, case).state_path),
            }
            for case in cases
        ],
    }
    _write_json(pipeline_paths.manifest_path, manifest)


class PipelineCaseTracker:
    def __init__(self, pipeline_paths: PipelinePaths, case: PipelineCase) -> None:
        self.pipeline_paths = pipeline_paths
        self.case = case
        self.paths = build_case_paths(pipeline_paths, case)
        self.state = self._load_or_initialize_state()

    def _load_or_initialize_state(self) -> dict[str, object]:
        if self.paths.state_path.exists():
            try:
                state = _read_json(self.paths.state_path)
            except (json.JSONDecodeError, OSError):
                state = {}
        else:
            state = {}
        if not state:
            state = self._default_state()
            self.state = state
            self._save_state()
        return state

    def _default_state(self) -> dict[str, object]:
        return {
            "schema_version": PIPELINE_STATE_SCHEMA_VERSION,
            "case_id": self.case.case_id,
            "software": self.case.software,
            "variant": self.case.variant,
            "case_name": self.case.case_name,
            "case_slug": self.case.case_slug,
            "sequence_name": self.case.sequence_name,
            "sequence_order": self.case.sequence_order,
            "status": "pending",
            "attempt_count": 0,
            "created_at": _now_timestamp(),
            "updated_at": _now_timestamp(),
            "started_at": "",
            "ended_at": "",
            "current_stage": "",
            "current_stage_status": "",
            "failure_stage": "",
            "input_path": "",
            "requested_output_path": "",
            "output_path": "",
            "export_output_path": "",
            "session_id": "",
            "monitor_status": "",
            "monitor_state_path": "",
            "session_output_path": "",
            "worker_stdout_path": "",
            "worker_stderr_path": "",
            "log_path": str(self.paths.log_path),
            "traceback_path": "",
            "case_state_path": str(self.paths.state_path),
            "error": "",
            "exception_type": "",
        }

    def _save_state(self) -> None:
        self.paths.case_dir.mkdir(parents=True, exist_ok=True)
        self.state["updated_at"] = _now_timestamp()
        _write_json(self.paths.state_path, self.state)

    def ensure_registered(self) -> None:
        self.paths.case_dir.mkdir(parents=True, exist_ok=True)
        self.state.setdefault("log_path", str(self.paths.log_path))
        self.state.setdefault("case_state_path", str(self.paths.state_path))
        self._save_state()

    def _append_event(self, stage: str, status: str, **fields: object) -> None:
        self.paths.case_dir.mkdir(parents=True, exist_ok=True)
        event = {
            "timestamp": _now_timestamp(),
            "case_id": self.case.case_id,
            "stage": stage,
            "status": status,
        }
        event.update(fields)
        with self.paths.log_path.open("a", encoding="utf-8") as handle:
            handle.write(json.dumps(event, ensure_ascii=True) + "\n")

    def record_stage(self, stage: str, status: str, **fields: object) -> None:
        self.state["current_stage"] = stage
        self.state["current_stage_status"] = status
        self._save_state()
        self._append_event(stage, status, **fields)

    def start_attempt(
        self,
        *,
        input_path: Path,
        requested_output_path: Path,
        export_output_path: Path | None = None,
    ) -> None:
        self.paths.case_dir.mkdir(parents=True, exist_ok=True)
        if self.paths.traceback_path.exists():
            self.paths.traceback_path.unlink()
        self.state.update(
            {
                "status": "running",
                "attempt_count": int(self.state.get("attempt_count", 0) or 0) + 1,
                "started_at": _now_timestamp(),
                "ended_at": "",
                "failure_stage": "",
                "input_path": str(input_path),
                "requested_output_path": str(requested_output_path),
                "output_path": str(requested_output_path),
                "export_output_path": str(export_output_path) if export_output_path is not None else "",
                "session_id": "",
                "monitor_status": "",
                "monitor_state_path": "",
                "session_output_path": "",
                "worker_stdout_path": "",
                "worker_stderr_path": "",
                "error": "",
                "exception_type": "",
                "traceback_path": "",
            }
        )
        self._save_state()
        self._append_event(
            "case",
            "running",
            input_path=str(input_path),
            requested_output_path=str(requested_output_path),
            export_output_path=str(export_output_path) if export_output_path is not None else "",
        )

    def update_monitor_metadata(
        self,
        *,
        session_id: str = "",
        output_path: Path | None = None,
        monitor_status: str = "",
        monitor_state_path: Path | None = None,
        session_output_path: Path | None = None,
        worker_stdout_path: Path | None = None,
        worker_stderr_path: Path | None = None,
    ) -> None:
        if session_id:
            self.state["session_id"] = session_id
        if output_path is not None:
            self.state["output_path"] = str(output_path)
        if monitor_status:
            self.state["monitor_status"] = monitor_status
        if monitor_state_path is not None:
            self.state["monitor_state_path"] = str(monitor_state_path)
        if session_output_path is not None:
            self.state["session_output_path"] = str(session_output_path)
        if worker_stdout_path is not None:
            self.state["worker_stdout_path"] = str(worker_stdout_path)
        if worker_stderr_path is not None:
            self.state["worker_stderr_path"] = str(worker_stderr_path)
        self._save_state()

    def mark_completed(self, result: CaseExecutionResult) -> None:
        if self.paths.traceback_path.exists():
            self.paths.traceback_path.unlink()
        self.state.update(
            {
                "status": "completed",
                "ended_at": _now_timestamp(),
                "output_path": str(result.csv_path),
                "export_output_path": str(result.export_output_path) if result.export_output_path is not None else "",
                "session_id": result.session_id,
                "monitor_status": result.monitor_status,
                "monitor_state_path": str(result.monitor_state_path) if result.monitor_state_path is not None else "",
                "session_output_path": str(result.session_output_path) if result.session_output_path is not None else "",
                "worker_stdout_path": str(result.worker_stdout_path) if result.worker_stdout_path is not None else "",
                "worker_stderr_path": str(result.worker_stderr_path) if result.worker_stderr_path is not None else "",
                "error": "",
                "exception_type": "",
                "traceback_path": "",
                "failure_stage": "",
            }
        )
        self._save_state()
        self._append_event(
            "case",
            "completed",
            csv_path=str(result.csv_path),
            session_id=result.session_id,
            monitor_status=result.monitor_status,
        )

    def mark_failed(self, exc: Exception) -> Path:
        traceback_text = traceback.format_exc()
        self.paths.case_dir.mkdir(parents=True, exist_ok=True)
        self.paths.traceback_path.write_text(traceback_text, encoding="utf-8")
        self.state.update(
            {
                "status": "failed",
                "ended_at": _now_timestamp(),
                "failure_stage": str(self.state.get("current_stage", "") or ""),
                "error": str(exc),
                "exception_type": type(exc).__name__,
                "traceback_path": str(self.paths.traceback_path),
            }
        )
        self._save_state()
        self._append_event(
            "case",
            "failed",
            error=str(exc),
            exception_type=type(exc).__name__,
            traceback_path=str(self.paths.traceback_path),
        )
        return self.paths.traceback_path

    def record_reused(self) -> None:
        self._append_event(
            "case",
            "reused",
            csv_path=str(self.state.get("output_path", "") or ""),
            session_id=str(self.state.get("session_id", "") or ""),
        )

    def completed_csv_path(self) -> Path | None:
        if str(self.state.get("status", "")) != "completed":
            return None
        output_path = str(self.state.get("output_path", "") or "")
        if not output_path:
            return None
        resolved = Path(output_path).resolve(strict=False)
        if not resolved.exists():
            return None
        return resolved


def _case_rows_from_state_files(run_root: Path) -> list[dict[str, str]]:
    manifest = load_pipeline_manifest(run_root)
    ordered_state_paths: list[Path] = []
    for case_payload in manifest.get("cases", []):
        state_path = Path(str(case_payload.get("case_state_path", "") or "")).resolve(strict=False)
        if state_path.exists():
            ordered_state_paths.append(state_path)

    if not ordered_state_paths:
        ordered_state_paths = sorted((run_root / "cases").glob("*/case_state.json"))

    rows: list[dict[str, str]] = []
    for state_path in ordered_state_paths:
        try:
            payload = _read_json(state_path)
        except (json.JSONDecodeError, OSError):
            continue
        rows.append(
            {
                "case_id": str(payload.get("case_id", "") or ""),
                "software": str(payload.get("software", "") or ""),
                "variant": str(payload.get("variant", "") or ""),
                "status": str(payload.get("status", "") or ""),
                "output_path": str(payload.get("output_path", "") or ""),
                "case_state_path": str(payload.get("case_state_path", "") or str(state_path)),
                "log_path": str(payload.get("log_path", "") or ""),
                "traceback_path": str(payload.get("traceback_path", "") or ""),
                "session_id": str(payload.get("session_id", "") or ""),
                "error": str(payload.get("error", "") or ""),
            }
        )
    return rows


def collect_completed_case_csv_paths(
    pipeline_paths: PipelinePaths,
    cases: tuple[PipelineCase, ...],
) -> list[Path]:
    csv_paths: list[Path] = []
    seen: set[Path] = set()
    for case in cases:
        tracker = PipelineCaseTracker(pipeline_paths, case)
        completed_csv_path = tracker.completed_csv_path()
        if completed_csv_path is None or completed_csv_path in seen:
            continue
        seen.add(completed_csv_path)
        csv_paths.append(completed_csv_path)
    return csv_paths


class PipelineRunRecorder:
    def __init__(self, run_root: Path) -> None:
        self.run_root = run_root
        self.events_path = run_root / "pipeline_progress.jsonl"
        self.summary_path = run_root / "pipeline_summary.md"
        self._events: list[dict[str, object]] = []
        if self.events_path.exists():
            for raw_line in self.events_path.read_text(encoding="utf-8").splitlines():
                if not raw_line.strip():
                    continue
                try:
                    self._events.append(json.loads(raw_line))
                except json.JSONDecodeError:
                    continue

    def record(self, stage: str, status: str, **fields: object) -> None:
        event = {
            "timestamp": datetime.now().isoformat(timespec="seconds"),
            "stage": stage,
            "status": status,
        }
        event.update(fields)
        self._events.append(event)
        with self.events_path.open("a", encoding="utf-8") as handle:
            handle.write(json.dumps(event, ensure_ascii=True) + "\n")
        self._write_summary()
        print(self._format_console_event(event), flush=True)

    def _format_console_event(self, event: dict[str, object]) -> str:
        pieces = [f"[pipeline][{event['status']}]", str(event["stage"])]
        if event.get("case_id"):
            pieces.append(f"case_id={event['case_id']}")
        if event.get("variant"):
            pieces.append(f"variant={event['variant']}")
        if event.get("software"):
            pieces.append(f"software={event['software']}")
        if event.get("report_path"):
            pieces.append(f"report={event['report_path']}")
        if event.get("csv_path"):
            pieces.append(f"csv={event['csv_path']}")
        if event.get("traceback_path"):
            pieces.append(f"traceback={event['traceback_path']}")
        if event.get("reused"):
            pieces.append(f"reused={event['reused']}")
        if event.get("error"):
            pieces.append(f"error={event['error']}")
        return " ".join(pieces)

    def _latest_pipeline_status(self) -> str:
        for event in reversed(self._events):
            if event["stage"] == "pipeline":
                return str(event["status"])
        return "running"

    def _write_summary(self) -> None:
        case_rows: dict[tuple[str, str], dict[str, object]] = {}
        timeline: list[str] = []
        for event in self._events:
            timestamp = str(event["timestamp"])
            stage = str(event["stage"])
            status = str(event["status"])
            software = str(event.get("software", "") or "")
            variant = str(event.get("variant", "") or "")
            if stage == "case" and software and variant:
                case_rows[(software, variant)] = {
                    "status": status,
                    "csv_path": str(event.get("csv_path", "") or ""),
                    "error": str(event.get("error", "") or ""),
                }
            summary_parts = [timestamp, stage, status]
            if variant:
                summary_parts.append(f"variant={variant}")
            if software:
                summary_parts.append(f"software={software}")
            if event.get("csv_path"):
                summary_parts.append(f"csv={event['csv_path']}")
            if event.get("report_path"):
                summary_parts.append(f"report={event['report_path']}")
            if event.get("traceback_path"):
                summary_parts.append(f"traceback={event['traceback_path']}")
            if event.get("reused"):
                summary_parts.append(f"reused={event['reused']}")
            if event.get("error"):
                summary_parts.append(f"error={event['error']}")
            timeline.append("- " + " | ".join(summary_parts))

        manifest_path = self.run_root / "pipeline_manifest.json"
        state_rows = _case_rows_from_state_files(self.run_root)
        lines = [
            "# Pipeline Summary",
            "",
            f"- status: {self._latest_pipeline_status()}",
            f"- run_root: {self.run_root}",
            f"- progress_log: {self.events_path}",
            f"- manifest_path: {manifest_path}",
            "",
            "## Case Status",
            "",
            "| case_id | software | variant | status | csv | state | log | traceback | session_id | error |",
            "| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |",
        ]
        if state_rows:
            for row in state_rows:
                output_path = row["output_path"].replace("|", "\\|")
                case_state_path = row["case_state_path"].replace("|", "\\|")
                log_path = row["log_path"].replace("|", "\\|")
                traceback_path = row["traceback_path"].replace("|", "\\|")
                session_id = row["session_id"].replace("|", "\\|")
                error = row["error"].replace("|", "\\|")
                lines.append(
                    f"| {row['case_id']} | {row['software']} | {row['variant']} | {row['status']} | "
                    f"{output_path} | {case_state_path} | {log_path} | {traceback_path} | {session_id} | {error} |"
                )
        elif case_rows:
            for software, variant in sorted(case_rows):
                row = case_rows[(software, variant)]
                csv_path = str(row["csv_path"]).replace("|", "\\|")
                error = str(row["error"]).replace("|", "\\|")
                lines.append(f"| {software}_{variant} | {software} | {variant} | {row['status']} | {csv_path} |  |  |  |  | {error} |")
        else:
            lines.append("|  |  |  | not_started |  |  |  |  |  |  |")
        lines.extend([
            "",
            "## Timeline",
            "",
            *timeline,
            "",
        ])
        self.summary_path.write_text("\n".join(lines), encoding="utf-8")


def slugify(value: str) -> str:
    normalized = re.sub(r"[^\w.-]+", "_", value.strip().lower())
    return normalized.strip("._") or "pipeline_run"


def build_pipeline_paths(
    results_root: Path,
    workload_name: str,
    *,
    resume_run_root: Path | None = None,
) -> PipelinePaths:
    results_root = results_root.resolve(strict=False)
    if resume_run_root is not None:
        run_root = resume_run_root.resolve(strict=False)
    else:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        run_root = results_root / f"{slugify(workload_name)}_{timestamp}"
    csv_dir = run_root / "csv"
    export_dir = run_root / "exports"
    report_dir = run_root / "report"
    cases_dir = run_root / "cases"
    csv_dir.mkdir(parents=True, exist_ok=True)
    export_dir.mkdir(parents=True, exist_ok=True)
    report_dir.mkdir(parents=True, exist_ok=True)
    cases_dir.mkdir(parents=True, exist_ok=True)
    assert csv_dir.exists() and export_dir.exists() and report_dir.exists() and cases_dir.exists()
    return PipelinePaths(
        root_dir=run_root,
        csv_dir=csv_dir,
        export_dir=export_dir,
        report_dir=report_dir,
        cases_dir=cases_dir,
        manifest_path=run_root / "pipeline_manifest.json",
    )


def run_non_blender_case(
    *,
    software: str,
    variant: str,
    workload_name: str,
    input_video_path: Path,
    requested_csv_path: Path,
    output_video_path: Path,
    pipeline_paths: PipelinePaths,
    launcher: SoftwareLauncher,
    monitor_bridge: MonitorBridge,
    ai_turbo_controller: AiTurboEngineController | None,
    recorder: PipelineRunRecorder,
    case_tracker: PipelineCaseTracker | None = None,
) -> CaseExecutionResult:
    run_name = f"{slugify(workload_name)}_{variant}"
    operator = build_operator(software)

    logger.info(
        "Preparing %s %s run. monitor_csv=%s output_video=%s",
        software,
        variant,
        requested_csv_path,
        output_video_path,
    )
    logger.info("Closing any leftover %s window before starting a clean run.", software)
    operator.close()

    if ai_turbo_controller is not None:
        logger.info("Configuring AI Turbo Engine Performance Boost for %s.", software)
        ai_turbo_controller.configure_for_software(software)

    logger.info("Starting the %s background timing monitor before launching the application.", software)
    launch_result = monitor_bridge.start_background_monitor(software, run_name, requested_csv_path)
    if case_tracker is not None:
        case_tracker.update_monitor_metadata(
            session_id=launch_result.session_id,
            output_path=launch_result.output_path,
            monitor_state_path=launch_result.state_path,
            session_output_path=launch_result.session_output_path,
            worker_stdout_path=launch_result.worker_stdout_path,
            worker_stderr_path=launch_result.worker_stderr_path,
        )
    recorder.record("monitor", "started", software=software, variant=variant, csv_path=str(requested_csv_path))
    if case_tracker is not None:
        case_tracker.record_stage(
            "monitor",
            "started",
            session_id=launch_result.session_id,
            csv_path=str(launch_result.output_path),
        )

    try:
        logger.info("Launching %s for the %s run.", software, variant)
        launcher.launch(software)
        launcher.wait_for_main_window(software, timeout=30.0)
        recorder.record("launch", "completed", software=software, variant=variant)
        if case_tracker is not None:
            case_tracker.record_stage("launch", "completed")
        logger.info("Running the %s UI automation flow for the %s run.", software, variant)
        recorder.record("automation", "started", software=software, variant=variant)
        if case_tracker is not None:
            case_tracker.record_stage("automation", "started")
        operator.perform(input_video_path, output_video_path)
        recorder.record("automation", "completed", software=software, variant=variant, output_video=str(output_video_path))
        if case_tracker is not None:
            case_tracker.record_stage("automation", "completed", output_video=str(output_video_path))
        if software in MONITOR_STOP_AFTER_AUTOMATION_SOFTWARE:
            logger.info(
                "Automation for %s has completed. Requesting the timing monitor to finalize immediately.",
                software,
            )
            status_payload = monitor_bridge.stop_session(launch_result.session_id)
        else:
            logger.info("Waiting for the %s timing monitor to finalize.", software)
            status_payload = monitor_bridge.wait_for_session_completion(launch_result.session_id)
    except Exception as exc:
        logger.exception("%s %s run failed. Requesting the monitor session to stop.", software, variant)
        recorder.record("automation", "failed", software=software, variant=variant, error=str(exc))
        if case_tracker is not None:
            case_tracker.record_stage("automation", "failed", error=str(exc))
        try:
            monitor_bridge.stop_session(launch_result.session_id)
        except Exception as stop_exc:
            logger.warning(
                "Could not stop the %s monitor session cleanly while unwinding the failed case. "
                "session_id=%s error=%s",
                software,
                launch_result.session_id,
                stop_exc,
            )
        raise
    finally:
        recorder.record("cleanup", "started", software=software, variant=variant)
        if case_tracker is not None:
            case_tracker.record_stage("cleanup", "started")
        try:
            operator.close()
        except Exception as exc:
            recorder.record("cleanup", "failed", software=software, variant=variant, error=str(exc))
            if case_tracker is not None:
                case_tracker.record_stage("cleanup", "failed", error=str(exc))
            raise
        recorder.record("cleanup", "completed", software=software, variant=variant)
        if case_tracker is not None:
            case_tracker.record_stage("cleanup", "completed")

    assert status_payload["status"] in {"completed", "completed_with_warnings"}, status_payload
    resolved_output_path = Path(status_payload.get("aggregate_output_path") or str(launch_result.output_path)).resolve(strict=False)
    assert resolved_output_path.exists(), f"Expected csv output was not created: {resolved_output_path}"
    return CaseExecutionResult(
        csv_path=resolved_output_path,
        requested_csv_path=requested_csv_path,
        export_output_path=output_video_path,
        session_id=launch_result.session_id,
        monitor_status=str(status_payload["status"]),
        monitor_state_path=Path(status_payload["state_path"]).resolve(strict=False)
        if status_payload.get("state_path")
        else launch_result.state_path,
        session_output_path=Path(status_payload["session_output_path"]).resolve(strict=False)
        if status_payload.get("session_output_path")
        else launch_result.session_output_path,
        worker_stdout_path=Path(status_payload["worker_stdout_path"]).resolve(strict=False)
        if status_payload.get("worker_stdout_path")
        else launch_result.worker_stdout_path,
        worker_stderr_path=Path(status_payload["worker_stderr_path"]).resolve(strict=False)
        if status_payload.get("worker_stderr_path")
        else launch_result.worker_stderr_path,
    )


def build_case_monitor_name(case: PipelineCase, *, workload_name: str) -> str:
    case_name = case.case_name or workload_name
    return f"{case_name}_{case.variant}"


def build_case_requested_csv_path(
    pipeline_paths: PipelinePaths,
    *,
    case: PipelineCase,
    workload_name: str,
) -> Path:
    return build_requested_csv_path(
        pipeline_paths,
        software=case.software,
        workload_name=workload_name,
        variant=case.variant,
        case_slug=case.case_slug,
    )


def build_case_export_output_path(
    pipeline_paths: PipelinePaths,
    *,
    case: PipelineCase,
    workload_name: str,
) -> Path | None:
    if case.software == "blender":
        return None
    return build_export_output_path(
        pipeline_paths,
        software=case.software,
        workload_name=workload_name,
        variant=case.variant,
        case_slug=case.case_slug,
    )


def run_shotcut_case_sequence(
    *,
    cases: tuple[PipelineCase, ...],
    workload_name: str,
    input_video_path: Path,
    pipeline_paths: PipelinePaths,
    launcher: SoftwareLauncher,
    monitor_bridge: MonitorBridge,
    ai_turbo_controller: AiTurboEngineController | None,
    recorder: PipelineRunRecorder,
) -> list[tuple[PipelineCase, CaseExecutionResult | None]]:
    assert cases, "Shotcut case sequence must not be empty."
    assert all(case.software == "shotcut" for case in cases), cases
    small_input_video_path = input_video_path.with_name("4K_small.mp4")
    assert small_input_video_path.exists(), f"Expected companion Shotcut input video is missing: {small_input_video_path}"

    operator = build_operator("shotcut")
    csv_results: list[tuple[PipelineCase, CaseExecutionResult | None]] = []
    pending_cases = [
        case
        for case in cases
        if PipelineCaseTracker(pipeline_paths, case).completed_csv_path() is None
    ]
    if not pending_cases:
        for case in cases:
            csv_results.append((case, None))
        return csv_results

    logger.info("Closing any leftover shotcut window before starting a clean run.")
    operator.close()
    if ai_turbo_controller is not None:
        logger.info("Configuring AI Turbo Engine Performance Boost for shotcut.")
        ai_turbo_controller.configure_for_software("shotcut")

    window = None
    launched = False
    try:
        for case in cases:
            case_tracker = PipelineCaseTracker(pipeline_paths, case)
            case_tracker.ensure_registered()
            completed_csv_path = case_tracker.completed_csv_path()
            if completed_csv_path is not None:
                logger.info(
                    "Skipping pipeline case because a completed case_state already exists. software=%s variant=%s case_name=%s csv=%s",
                    case.software,
                    case.variant,
                    case.case_name or workload_name,
                    completed_csv_path,
                )
                case_tracker.record_reused()
                recorder.record(
                    "case",
                    "completed",
                    case_id=case.case_id,
                    software=case.software,
                    variant=case.variant,
                    csv_path=str(completed_csv_path),
                    session_id=str(case_tracker.state.get("session_id", "") or ""),
                    reused="true",
                    case_state_path=str(case_tracker.paths.state_path),
                    log_path=str(case_tracker.paths.log_path),
                )
                csv_results.append((case, None))
                continue

            run_name = build_case_monitor_name(case, workload_name=workload_name)
            launch_result = None
            try:
                if not launched:
                    logger.info("Launching shotcut for the %s run.", case.variant)
                    launcher.launch("shotcut")
                    launcher.wait_for_main_window("shotcut", timeout=30.0)
                    launched = True
                    window = operator._connect_main_window(
                        timeout=max(30.0, SOFTWARE_SPECS["shotcut"].startup_timeout_seconds)
                    )

                requested_csv_path = build_case_requested_csv_path(
                    pipeline_paths,
                    case=case,
                    workload_name=workload_name,
                )
                output_video_path = build_case_export_output_path(
                    pipeline_paths,
                    case=case,
                    workload_name=workload_name,
                )
                assert output_video_path is not None

                current_input_path = small_input_video_path if case.sequence_order == 2 else input_video_path
                case_tracker.start_attempt(
                    input_path=current_input_path.resolve(strict=False),
                    requested_output_path=requested_csv_path,
                    export_output_path=output_video_path,
                )
                recorder.record(
                    "case",
                    "started",
                    case_id=case.case_id,
                    software=case.software,
                    variant=case.variant,
                    case_state_path=str(case_tracker.paths.state_path),
                    log_path=str(case_tracker.paths.log_path),
                )
                logger.info(
                    "Starting the shotcut background timing monitor for case '%s'. csv=%s",
                    run_name,
                    requested_csv_path,
                )
                launch_result = monitor_bridge.start_background_monitor("shotcut", run_name, requested_csv_path)
                case_tracker.update_monitor_metadata(
                    session_id=launch_result.session_id,
                    output_path=launch_result.output_path,
                    monitor_state_path=launch_result.state_path,
                    session_output_path=launch_result.session_output_path,
                    worker_stdout_path=launch_result.worker_stdout_path,
                    worker_stderr_path=launch_result.worker_stderr_path,
                )
                recorder.record(
                    "monitor",
                    "started",
                    software="shotcut",
                    variant=case.variant,
                    csv_path=str(requested_csv_path),
                )
                case_tracker.record_stage(
                    "monitor",
                    "started",
                    session_id=launch_result.session_id,
                    csv_path=str(launch_result.output_path),
                )
                recorder.record("launch", "completed", software="shotcut", variant=case.variant)
                case_tracker.record_stage("launch", "completed")
                recorder.record("automation", "started", software="shotcut", variant=case.variant)
                case_tracker.record_stage("automation", "started")

                logger.info("Running the shotcut UI automation flow for case '%s'.", run_name)
                if case.sequence_order == 1:
                    window = operator._open_input_clip(window, input_video_path)
                    operator._append_selected_clip_to_timeline(window)
                elif case.sequence_order == 2:
                    window = operator._open_input_clip(window, small_input_video_path)
                    operator._append_selected_clip_to_timeline(window)
                elif case.sequence_order == 3:
                    operator._append_selected_clip_to_timeline(window)
                    window = operator._open_input_clip(window, input_video_path)
                    operator._append_selected_clip_to_timeline(window)
                else:
                    raise AssertionError(f"Unsupported Shotcut sequence order: {case.sequence_order}")
                window = operator._export_current_timeline(window, output_video_path)
                recorder.record(
                    "automation",
                    "completed",
                    software="shotcut",
                    variant=case.variant,
                    output_video=str(output_video_path),
                )
                case_tracker.record_stage("automation", "completed", output_video=str(output_video_path))
                logger.info(
                    "Automation for shotcut case '%s' has completed. Finalizing its timing monitor immediately.",
                    run_name,
                )
                status_payload = monitor_bridge.stop_session(launch_result.session_id)
            except Exception as exc:
                logger.exception("shotcut %s case '%s' failed. Requesting the monitor session to stop.", case.variant, run_name)
                recorder.record("automation", "failed", software="shotcut", variant=case.variant, error=str(exc))
                case_tracker.record_stage("automation", "failed", error=str(exc))
                try:
                    if launch_result is not None:
                        monitor_bridge.stop_session(launch_result.session_id)
                except Exception as stop_exc:
                    logger.warning(
                        "Could not stop the shotcut monitor session cleanly while unwinding the failed case. "
                        "session_id=%s error=%s",
                        launch_result.session_id if launch_result is not None else "",
                        stop_exc,
                    )
                traceback_path = case_tracker.mark_failed(exc)
                recorder.record(
                    "case",
                    "failed",
                    case_id=case.case_id,
                    software=case.software,
                    variant=case.variant,
                    error=str(exc),
                    case_state_path=str(case_tracker.paths.state_path),
                    log_path=str(case_tracker.paths.log_path),
                    traceback_path=str(traceback_path),
                )
                raise

            recorder.record("cleanup", "started", software="shotcut", variant=case.variant)
            case_tracker.record_stage("cleanup", "started")
            recorder.record("cleanup", "completed", software="shotcut", variant=case.variant)
            case_tracker.record_stage("cleanup", "completed")

            assert status_payload["status"] in {"completed", "completed_with_warnings"}, status_payload
            resolved_output_path = Path(status_payload.get("aggregate_output_path") or str(launch_result.output_path)).resolve(strict=False)
            assert resolved_output_path.exists(), f"Expected csv output was not created: {resolved_output_path}"
            result = CaseExecutionResult(
                csv_path=resolved_output_path,
                requested_csv_path=requested_csv_path,
                export_output_path=output_video_path,
                session_id=launch_result.session_id,
                monitor_status=str(status_payload["status"]),
                monitor_state_path=Path(status_payload["state_path"]).resolve(strict=False)
                if status_payload.get("state_path")
                else launch_result.state_path,
                session_output_path=Path(status_payload["session_output_path"]).resolve(strict=False)
                if status_payload.get("session_output_path")
                else launch_result.session_output_path,
                worker_stdout_path=Path(status_payload["worker_stdout_path"]).resolve(strict=False)
                if status_payload.get("worker_stdout_path")
                else launch_result.worker_stdout_path,
                worker_stderr_path=Path(status_payload["worker_stderr_path"]).resolve(strict=False)
                if status_payload.get("worker_stderr_path")
                else launch_result.worker_stderr_path,
            )
            case_tracker.mark_completed(result)
            recorder.record(
                "case",
                "completed",
                case_id=case.case_id,
                software=case.software,
                variant=case.variant,
                csv_path=str(result.csv_path),
                session_id=result.session_id,
                case_state_path=str(case_tracker.paths.state_path),
                log_path=str(case_tracker.paths.log_path),
            )
            csv_results.append((case, result))
    finally:
        if launched:
            try:
                operator.close()
            except Exception:
                logger.exception("Shotcut close failed while finalizing the multi-case sequence.")
                raise
    return csv_results


def run_kdenlive_case_sequence(
    *,
    cases: tuple[PipelineCase, ...],
    workload_name: str,
    input_video_path: Path,
    pipeline_paths: PipelinePaths,
    launcher: SoftwareLauncher,
    monitor_bridge: MonitorBridge,
    ai_turbo_controller: AiTurboEngineController | None,
    recorder: PipelineRunRecorder,
) -> list[tuple[PipelineCase, CaseExecutionResult | None]]:
    assert cases, "Kdenlive case sequence must not be empty."
    assert all(case.software == "kdenlive" for case in cases), cases
    small_input_video_path = input_video_path.with_name("4K_small.mp4")
    assert small_input_video_path.exists(), f"Expected companion Kdenlive input video is missing: {small_input_video_path}"

    operator = build_operator("kdenlive")
    csv_results: list[tuple[PipelineCase, CaseExecutionResult | None]] = []
    pending_cases = [
        case
        for case in cases
        if PipelineCaseTracker(pipeline_paths, case).completed_csv_path() is None
    ]
    if not pending_cases:
        for case in cases:
            csv_results.append((case, None))
        return csv_results

    logger.info("Closing any leftover kdenlive window before starting a clean run.")
    operator.close()
    if ai_turbo_controller is not None:
        logger.info("Configuring AI Turbo Engine Performance Boost for kdenlive.")
        ai_turbo_controller.configure_for_software("kdenlive")

    window = None
    launched = False
    try:
        for case in cases:
            case_tracker = PipelineCaseTracker(pipeline_paths, case)
            case_tracker.ensure_registered()
            completed_csv_path = case_tracker.completed_csv_path()
            if completed_csv_path is not None:
                logger.info(
                    "Skipping pipeline case because a completed case_state already exists. software=%s variant=%s case_name=%s csv=%s",
                    case.software,
                    case.variant,
                    case.case_name or workload_name,
                    completed_csv_path,
                )
                case_tracker.record_reused()
                recorder.record(
                    "case",
                    "completed",
                    case_id=case.case_id,
                    software=case.software,
                    variant=case.variant,
                    csv_path=str(completed_csv_path),
                    session_id=str(case_tracker.state.get("session_id", "") or ""),
                    reused="true",
                    case_state_path=str(case_tracker.paths.state_path),
                    log_path=str(case_tracker.paths.log_path),
                )
                csv_results.append((case, None))
                continue

            run_name = build_case_monitor_name(case, workload_name=workload_name)
            launch_result = None
            try:
                if not launched:
                    logger.info("Launching kdenlive for the %s run.", case.variant)
                    launcher.launch("kdenlive")
                    launcher.wait_for_main_window("kdenlive", timeout=30.0)
                    launched = True
                    window = operator._connect_main_window(
                        timeout=max(30.0, SOFTWARE_SPECS["kdenlive"].startup_timeout_seconds)
                    )

                requested_csv_path = build_case_requested_csv_path(
                    pipeline_paths,
                    case=case,
                    workload_name=workload_name,
                )
                output_video_path = build_case_export_output_path(
                    pipeline_paths,
                    case=case,
                    workload_name=workload_name,
                )
                assert output_video_path is not None

                current_input_path = small_input_video_path if case.sequence_order == 2 else input_video_path
                case_tracker.start_attempt(
                    input_path=current_input_path.resolve(strict=False),
                    requested_output_path=requested_csv_path,
                    export_output_path=output_video_path,
                )
                recorder.record(
                    "case",
                    "started",
                    case_id=case.case_id,
                    software=case.software,
                    variant=case.variant,
                    case_state_path=str(case_tracker.paths.state_path),
                    log_path=str(case_tracker.paths.log_path),
                )
                logger.info(
                    "Starting the kdenlive background timing monitor for case '%s'. csv=%s",
                    run_name,
                    requested_csv_path,
                )
                launch_result = monitor_bridge.start_background_monitor("kdenlive", run_name, requested_csv_path)
                case_tracker.update_monitor_metadata(
                    session_id=launch_result.session_id,
                    output_path=launch_result.output_path,
                    monitor_state_path=launch_result.state_path,
                    session_output_path=launch_result.session_output_path,
                    worker_stdout_path=launch_result.worker_stdout_path,
                    worker_stderr_path=launch_result.worker_stderr_path,
                )
                recorder.record(
                    "monitor",
                    "started",
                    software="kdenlive",
                    variant=case.variant,
                    csv_path=str(requested_csv_path),
                )
                case_tracker.record_stage(
                    "monitor",
                    "started",
                    session_id=launch_result.session_id,
                    csv_path=str(launch_result.output_path),
                )
                recorder.record("launch", "completed", software="kdenlive", variant=case.variant)
                case_tracker.record_stage("launch", "completed")
                recorder.record("automation", "started", software="kdenlive", variant=case.variant)
                case_tracker.record_stage("automation", "started")

                logger.info("Running the kdenlive UI automation flow for case '%s'.", run_name)
                if case.sequence_order == 1:
                    window = operator._open_input_clip(window, input_video_path)
                    operator._insert_clip_to_timeline(window)
                elif case.sequence_order == 2:
                    window = operator._open_input_clip(window, small_input_video_path)
                    operator._insert_clip_to_timeline(window)
                elif case.sequence_order == 3:
                    operator._close_render_dialog_only()
                    operator._insert_clip_to_timeline(window)
                    window = operator._open_input_clip(window, input_video_path)
                    operator._insert_clip_to_timeline(window)
                else:
                    raise AssertionError(f"Unsupported Kdenlive sequence order: {case.sequence_order}")
                window = operator._render_current_timeline(window, output_video_path)
                recorder.record(
                    "automation",
                    "completed",
                    software="kdenlive",
                    variant=case.variant,
                    output_video=str(output_video_path),
                )
                case_tracker.record_stage("automation", "completed", output_video=str(output_video_path))
                logger.info(
                    "Automation for kdenlive case '%s' has completed. Finalizing its timing monitor immediately.",
                    run_name,
                )
                status_payload = monitor_bridge.stop_session(launch_result.session_id)
            except Exception as exc:
                logger.exception("kdenlive %s case '%s' failed. Requesting the monitor session to stop.", case.variant, run_name)
                recorder.record("automation", "failed", software="kdenlive", variant=case.variant, error=str(exc))
                case_tracker.record_stage("automation", "failed", error=str(exc))
                try:
                    if launch_result is not None:
                        monitor_bridge.stop_session(launch_result.session_id)
                except Exception as stop_exc:
                    logger.warning(
                        "Could not stop the kdenlive monitor session cleanly while unwinding the failed case. "
                        "session_id=%s error=%s",
                        launch_result.session_id if launch_result is not None else "",
                        stop_exc,
                    )
                traceback_path = case_tracker.mark_failed(exc)
                recorder.record(
                    "case",
                    "failed",
                    case_id=case.case_id,
                    software=case.software,
                    variant=case.variant,
                    error=str(exc),
                    case_state_path=str(case_tracker.paths.state_path),
                    log_path=str(case_tracker.paths.log_path),
                    traceback_path=str(traceback_path),
                )
                raise

            recorder.record("cleanup", "started", software="kdenlive", variant=case.variant)
            case_tracker.record_stage("cleanup", "started")
            recorder.record("cleanup", "completed", software="kdenlive", variant=case.variant)
            case_tracker.record_stage("cleanup", "completed")

            assert status_payload["status"] in {"completed", "completed_with_warnings"}, status_payload
            resolved_output_path = Path(status_payload.get("aggregate_output_path") or str(launch_result.output_path)).resolve(strict=False)
            assert resolved_output_path.exists(), f"Expected csv output was not created: {resolved_output_path}"
            result = CaseExecutionResult(
                csv_path=resolved_output_path,
                requested_csv_path=requested_csv_path,
                export_output_path=output_video_path,
                session_id=launch_result.session_id,
                monitor_status=str(status_payload["status"]),
                monitor_state_path=Path(status_payload["state_path"]).resolve(strict=False)
                if status_payload.get("state_path")
                else launch_result.state_path,
                session_output_path=Path(status_payload["session_output_path"]).resolve(strict=False)
                if status_payload.get("session_output_path")
                else launch_result.session_output_path,
                worker_stdout_path=Path(status_payload["worker_stdout_path"]).resolve(strict=False)
                if status_payload.get("worker_stdout_path")
                else launch_result.worker_stdout_path,
                worker_stderr_path=Path(status_payload["worker_stderr_path"]).resolve(strict=False)
                if status_payload.get("worker_stderr_path")
                else launch_result.worker_stderr_path,
            )
            case_tracker.mark_completed(result)
            recorder.record(
                "case",
                "completed",
                case_id=case.case_id,
                software=case.software,
                variant=case.variant,
                csv_path=str(result.csv_path),
                session_id=result.session_id,
                case_state_path=str(case_tracker.paths.state_path),
                log_path=str(case_tracker.paths.log_path),
            )
            csv_results.append((case, result))
    finally:
        if launched:
            try:
                operator.close()
            except Exception:
                logger.exception("Kdenlive close failed while finalizing the multi-case sequence.")
                raise
    return csv_results


def run_blender_case(
    *,
    variant: str,
    workload_name: str,
    pipeline_paths: PipelinePaths,
    monitor_bridge: MonitorBridge,
    ai_turbo_controller: AiTurboEngineController | None,
    requested_csv_path: Path,
    blend_file: Path | None = None,
    blender_executable: Path | None = None,
    blender_ui_mode: str = "visible",
    render_mode: str = "frame",
    frame: int = 1,
    recorder: PipelineRunRecorder | None = None,
    case_tracker: PipelineCaseTracker | None = None,
) -> CaseExecutionResult:
    spec = SOFTWARE_SPECS["blender"]
    effective_blend_file = blend_file or spec.blend_file
    assert effective_blend_file is not None and effective_blend_file.exists(), "Blender blend file is required."
    run_name = f"{slugify(workload_name)}_{variant}"
    logger.info(
        "Preparing blender %s run. blend_file=%s monitor_csv=%s blender_ui_mode=%s render_mode=%s frame=%s",
        variant,
        effective_blend_file,
        requested_csv_path,
        blender_ui_mode,
        render_mode,
        frame,
    )

    if ai_turbo_controller is not None:
        logger.info("Configuring AI Turbo Engine Performance Boost for blender.")
        ai_turbo_controller.configure_for_software("blender")
    if recorder is not None:
        recorder.record("monitor", "started", software="blender", variant=variant, csv_path=str(requested_csv_path))
    if case_tracker is not None:
        case_tracker.record_stage("monitor", "started", csv_path=str(requested_csv_path))
    try:
        run_result = monitor_bridge.run_blender_monitor(
            test_name=run_name,
            output_path=requested_csv_path,
            blend_file=effective_blend_file,
            blender_executable=blender_executable,
            blender_ui_mode=blender_ui_mode,
            render_mode=render_mode,
            frame=frame,
        )
    except Exception as exc:
        logger.exception("Blender %s run failed.", variant)
        if recorder is not None:
            recorder.record("automation", "failed", software="blender", variant=variant, error=str(exc))
        if case_tracker is not None:
            case_tracker.record_stage("automation", "failed", error=str(exc))
        raise
    if case_tracker is not None:
        case_tracker.update_monitor_metadata(
            session_id=run_result.session_id,
            output_path=run_result.output_path,
            monitor_status=run_result.status,
            monitor_state_path=run_result.state_path,
            session_output_path=run_result.session_output_path,
            worker_stdout_path=run_result.worker_stdout_path,
            worker_stderr_path=run_result.worker_stderr_path,
        )
    if recorder is not None:
        recorder.record("automation", "completed", software="blender", variant=variant)
    if case_tracker is not None:
        case_tracker.record_stage("automation", "completed")
    assert run_result.output_path.exists(), f"Expected blender csv output was not created: {run_result.output_path}"
    return CaseExecutionResult(
        csv_path=run_result.output_path,
        requested_csv_path=requested_csv_path,
        session_id=run_result.session_id,
        monitor_status=run_result.status,
        monitor_state_path=run_result.state_path,
        session_output_path=run_result.session_output_path,
        worker_stdout_path=run_result.worker_stdout_path,
        worker_stderr_path=run_result.worker_stderr_path,
    )


def run_pipeline_pass(
    *,
    variant: str,
    enable_turbo: bool,
    selected_software: tuple[str, ...],
    workload_name: str,
    input_video_path: Path,
    pipeline_paths: PipelinePaths,
    launcher: SoftwareLauncher,
    monitor_bridge: MonitorBridge,
    ai_turbo_controller: AiTurboEngineController | None,
    resolved_blend_file: Path | None,
    resolved_blender_executable: Path | None,
    blender_ui_mode: str,
    render_mode: str,
    frame: int,
    recorder: PipelineRunRecorder,
) -> list[Path]:
    logger.info("Starting pipeline pass '%s'. ai_turbo_enabled=%s", variant, enable_turbo)
    recorder.record("pass", "started", variant=variant, ai_turbo_enabled=enable_turbo)
    planned_cases = build_pipeline_cases(
        selected_software,
        [(variant, enable_turbo)],
        workload_name=workload_name,
        input_video_path=input_video_path,
    )
    controller = ai_turbo_controller if enable_turbo else None
    csv_paths: list[Path] = []
    pending_cases = [
        case
        for case in planned_cases
        if PipelineCaseTracker(pipeline_paths, case).completed_csv_path() is None
    ]
    if controller is not None and pending_cases:
        logger.info("Starting AI Turbo Engine topmost guard for the turbo pass.")
        controller.start_topmost_guard()
    try:
        case_index = 0
        while case_index < len(planned_cases):
            case = planned_cases[case_index]
            software = case.software
            if (
                software == "shotcut"
                and case.sequence_name == SHOTCUT_SEQUENCE_NAME
                and case.sequence_order == 1
            ):
                grouped_cases: list[PipelineCase] = [case]
                next_index = case_index + 1
                while (
                    next_index < len(planned_cases)
                    and planned_cases[next_index].software == software
                    and planned_cases[next_index].variant == case.variant
                    and planned_cases[next_index].sequence_name == case.sequence_name
                ):
                    grouped_cases.append(planned_cases[next_index])
                    next_index += 1
                sequence_results = run_shotcut_case_sequence(
                    cases=tuple(grouped_cases),
                    workload_name=workload_name,
                    input_video_path=input_video_path,
                    pipeline_paths=pipeline_paths,
                    launcher=launcher,
                    monitor_bridge=monitor_bridge,
                    ai_turbo_controller=controller,
                    recorder=recorder,
                )
                for grouped_case, result in sequence_results:
                    if result is not None:
                        csv_paths.append(result.csv_path)
                        logger.info(
                            "Pipeline case finished. software=%s variant=%s csv=%s",
                            grouped_case.software,
                            grouped_case.variant,
                            result.csv_path,
                        )
                        continue
                    completed_csv_path = PipelineCaseTracker(pipeline_paths, grouped_case).completed_csv_path()
                    if completed_csv_path is not None:
                        csv_paths.append(completed_csv_path)
                case_index = next_index
                continue
            if (
                software == "kdenlive"
                and case.sequence_name == KDENLIVE_SEQUENCE_NAME
                and case.sequence_order == 1
            ):
                grouped_cases: list[PipelineCase] = [case]
                next_index = case_index + 1
                while (
                    next_index < len(planned_cases)
                    and planned_cases[next_index].software == software
                    and planned_cases[next_index].variant == case.variant
                    and planned_cases[next_index].sequence_name == case.sequence_name
                ):
                    grouped_cases.append(planned_cases[next_index])
                    next_index += 1
                sequence_results = run_kdenlive_case_sequence(
                    cases=tuple(grouped_cases),
                    workload_name=workload_name,
                    input_video_path=input_video_path,
                    pipeline_paths=pipeline_paths,
                    launcher=launcher,
                    monitor_bridge=monitor_bridge,
                    ai_turbo_controller=controller,
                    recorder=recorder,
                )
                for grouped_case, result in sequence_results:
                    if result is not None:
                        csv_paths.append(result.csv_path)
                        logger.info(
                            "Pipeline case finished. software=%s variant=%s csv=%s",
                            grouped_case.software,
                            grouped_case.variant,
                            result.csv_path,
                        )
                        continue
                    completed_csv_path = PipelineCaseTracker(pipeline_paths, grouped_case).completed_csv_path()
                    if completed_csv_path is not None:
                        csv_paths.append(completed_csv_path)
                case_index = next_index
                continue

            case_tracker = PipelineCaseTracker(pipeline_paths, case)
            case_tracker.ensure_registered()
            completed_csv_path = case_tracker.completed_csv_path()
            if completed_csv_path is not None:
                logger.info(
                    "Skipping pipeline case because a completed case_state already exists. software=%s variant=%s csv=%s",
                    software,
                    variant,
                    completed_csv_path,
                )
                case_tracker.record_reused()
                recorder.record(
                    "case",
                    "completed",
                    case_id=case.case_id,
                    software=software,
                    variant=variant,
                    csv_path=str(completed_csv_path),
                    session_id=str(case_tracker.state.get("session_id", "") or ""),
                    reused="true",
                    case_state_path=str(case_tracker.paths.state_path),
                    log_path=str(case_tracker.paths.log_path),
                )
                csv_paths.append(completed_csv_path)
                case_index += 1
                continue

            logger.info("Starting pipeline case. software=%s variant=%s", software, variant)
            requested_csv_path = build_requested_csv_path(
                pipeline_paths,
                software=software,
                workload_name=workload_name,
                variant=variant,
            )
            output_video_path = (
                build_export_output_path(
                    pipeline_paths,
                    software=software,
                    workload_name=workload_name,
                    variant=variant,
                )
                if software != "blender"
                else None
            )
            input_path_for_state = (
                (resolved_blend_file or SOFTWARE_SPECS["blender"].blend_file)
                if software == "blender"
                else input_video_path
            )
            assert input_path_for_state is not None
            case_tracker.start_attempt(
                input_path=input_path_for_state.resolve(strict=False),
                requested_output_path=requested_csv_path,
                export_output_path=output_video_path,
            )
            recorder.record(
                "case",
                "started",
                case_id=case.case_id,
                software=software,
                variant=variant,
                case_state_path=str(case_tracker.paths.state_path),
                log_path=str(case_tracker.paths.log_path),
            )
            if software == "blender":
                try:
                    result = run_blender_case(
                        variant=variant,
                        workload_name=workload_name,
                        pipeline_paths=pipeline_paths,
                        monitor_bridge=monitor_bridge,
                        ai_turbo_controller=controller,
                        requested_csv_path=requested_csv_path,
                        blend_file=resolved_blend_file,
                        blender_executable=resolved_blender_executable,
                        blender_ui_mode=blender_ui_mode,
                        render_mode=render_mode,
                        frame=frame,
                        recorder=recorder,
                        case_tracker=case_tracker,
                    )
                except Exception as exc:
                    traceback_path = case_tracker.mark_failed(exc)
                    recorder.record(
                        "case",
                        "failed",
                        case_id=case.case_id,
                        software=software,
                        variant=variant,
                        error=str(exc),
                        case_state_path=str(case_tracker.paths.state_path),
                        log_path=str(case_tracker.paths.log_path),
                        traceback_path=str(traceback_path),
                    )
                    raise
            else:
                try:
                    assert output_video_path is not None
                    result = run_non_blender_case(
                        software=software,
                        variant=variant,
                        workload_name=workload_name,
                        input_video_path=input_video_path,
                        requested_csv_path=requested_csv_path,
                        output_video_path=output_video_path,
                        pipeline_paths=pipeline_paths,
                        launcher=launcher,
                        monitor_bridge=monitor_bridge,
                        ai_turbo_controller=controller,
                        recorder=recorder,
                        case_tracker=case_tracker,
                    )
                except Exception as exc:
                    traceback_path = case_tracker.mark_failed(exc)
                    recorder.record(
                        "case",
                        "failed",
                        case_id=case.case_id,
                        software=software,
                        variant=variant,
                        error=str(exc),
                        case_state_path=str(case_tracker.paths.state_path),
                        log_path=str(case_tracker.paths.log_path),
                        traceback_path=str(traceback_path),
                    )
                    raise
            case_tracker.mark_completed(result)
            csv_paths.append(result.csv_path)
            logger.info("Pipeline case finished. software=%s variant=%s csv=%s", software, variant, result.csv_path)
            recorder.record(
                "case",
                "completed",
                case_id=case.case_id,
                software=software,
                variant=variant,
                csv_path=str(result.csv_path),
                session_id=result.session_id,
                case_state_path=str(case_tracker.paths.state_path),
                log_path=str(case_tracker.paths.log_path),
            )
            case_index += 1
    except Exception as exc:
        recorder.record("pass", "failed", variant=variant, ai_turbo_enabled=enable_turbo, error=str(exc))
        raise
    finally:
        if controller is not None and pending_cases:
            logger.info("Stopping AI Turbo Engine topmost guard for the turbo pass.")
            controller.stop_topmost_guard()
    recorder.record("pass", "completed", variant=variant, ai_turbo_enabled=enable_turbo)
    return csv_paths


def run_pipeline(args: argparse.Namespace) -> Path:
    selected_software = resolve_selected_software(args)
    launcher = SoftwareLauncher()
    monitor_bridge = MonitorBridge()
    ai_turbo_shortcut = Path(args.ai_turbo_shortcut).resolve(strict=False) if args.ai_turbo_shortcut else None
    ai_turbo_controller = AiTurboEngineController(
        window_title_re=args.ai_turbo_window_title,
        shortcut_path=ai_turbo_shortcut,
    )
    resolved_blend_file = Path(args.blend_file).resolve(strict=False) if args.blend_file else None
    resolved_blender_executable = Path(args.blender_exe).resolve(strict=False) if args.blender_exe else None
    resolved_blender_ui_mode = str(getattr(args, "blender_ui_mode", "visible") or "visible")

    passes = [("baseline", False), ("turbo", True)]
    if args.skip_turbo:
        passes = [("baseline", False)]
    if args.skip_baseline:
        passes = [("turbo", True)]
    pipeline_paths = build_pipeline_paths(
        Path(args.results_root),
        args.workload_name,
        resume_run_root=Path(args.resume_run_root) if getattr(args, "resume_run_root", "") else None,
    )
    recorder = PipelineRunRecorder(pipeline_paths.root_dir)
    input_video_path = Path(args.input_video).resolve(strict=False) if args.input_video else resolve_input_video(args.workload_name)
    assert input_video_path.exists(), f"Input video does not exist: {input_video_path}"
    planned_cases = build_pipeline_cases(
        selected_software,
        passes,
        workload_name=args.workload_name,
        input_video_path=input_video_path,
    )
    existing_manifest = load_pipeline_manifest(pipeline_paths.root_dir)
    if existing_manifest:
        existing_workload_name = str(existing_manifest.get("workload_name", "") or "")
        assert existing_workload_name in {"", args.workload_name}, (
            existing_workload_name,
            args.workload_name,
        )
    write_pipeline_manifest(
        pipeline_paths,
        workload_name=args.workload_name,
        input_video_path=input_video_path,
        selected_software=selected_software,
        passes=passes,
        cases=planned_cases,
    )
    for case in planned_cases:
        PipelineCaseTracker(pipeline_paths, case).ensure_registered()

    recorder.record(
        "pipeline",
        "started",
        workload_name=args.workload_name,
        run_root=str(pipeline_paths.root_dir),
        selected_software=",".join(selected_software),
        manifest_path=str(pipeline_paths.manifest_path),
    )
    try:
        for variant, enable_turbo in passes:
            run_pipeline_pass(
                variant=variant,
                enable_turbo=enable_turbo,
                selected_software=selected_software,
                workload_name=args.workload_name,
                input_video_path=input_video_path,
                pipeline_paths=pipeline_paths,
                launcher=launcher,
                monitor_bridge=monitor_bridge,
                ai_turbo_controller=ai_turbo_controller,
                resolved_blend_file=resolved_blend_file,
                resolved_blender_executable=resolved_blender_executable,
                blender_ui_mode=resolved_blender_ui_mode,
                render_mode=args.render_mode,
                frame=args.frame,
                recorder=recorder,
            )
        csv_paths = collect_completed_case_csv_paths(pipeline_paths, planned_cases)
        assert csv_paths, f"No completed case csv files were found under {pipeline_paths.root_dir}"
        report_path = pipeline_paths.report_dir / f"{slugify(args.workload_name)}_summary.xlsx"
        logger.info("Generating xlsx report with baseline/turbo comparison formulas. output=%s", report_path)
        recorder.record("report", "started", report_path=str(report_path))
        csv_count, row_count, generated_report_path = generate_xlsx_report(csv_paths, report_path, default_test_case=args.workload_name)
        assert csv_count == len(csv_paths), (csv_count, len(csv_paths))
        assert row_count >= len(csv_paths), (row_count, len(csv_paths))
        assert generated_report_path.exists(), f"Expected xlsx report was not created: {generated_report_path}"
        recorder.record("report", "completed", report_path=str(generated_report_path), csv_count=csv_count, row_count=row_count)
        recorder.record("pipeline", "completed", report_path=str(generated_report_path))
        return generated_report_path
    except Exception as exc:
        recorder.record("pipeline", "failed", error=str(exc))
        raise


def run_ai_turbo_sequence(args: argparse.Namespace) -> str:
    ai_turbo_shortcut = Path(args.ai_turbo_shortcut).resolve(strict=False) if args.ai_turbo_shortcut else None
    controller = AiTurboEngineController(
        window_title_re=args.ai_turbo_window_title,
        shortcut_path=ai_turbo_shortcut,
    )
    software = resolve_ai_turbo_sequence_software(args)
    wait_seconds = float(getattr(args, "ai_turbo_sequence_wait_seconds", 30.0))
    controller.run_sequence(software, wait_seconds=wait_seconds)
    return software


def run_ai_turbo_open_check(args: argparse.Namespace) -> None:
    ai_turbo_shortcut = Path(args.ai_turbo_shortcut).resolve(strict=False) if args.ai_turbo_shortcut else None
    controller = AiTurboEngineController(
        window_title_re=args.ai_turbo_window_title,
        shortcut_path=ai_turbo_shortcut,
    )
    wait_seconds = float(getattr(args, "ai_turbo_open_check_wait_seconds", 5.0))
    assert wait_seconds >= 0, "ai_turbo_open_check_wait_seconds must be non-negative."
    controller.ensure_running()
    try:
        if wait_seconds > 0:
            time.sleep(wait_seconds)
    finally:
        controller.close()


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Run the full baseline vs AI Turbo video-processing automation workflow.")
    parser.add_argument("--workload-name", default=DEFAULT_PIPELINE_WORKLOAD_NAME, help="Logical workload label used in csv/xlsx outputs.")
    parser.add_argument("--input-video", default="", help="Optional override video path. Defaults to test_data mapping based on workload name.")
    parser.add_argument(
        "--software",
        nargs="+",
        choices=DEFAULT_SOFTWARE_ORDER,
        default=list(DEFAULT_SOFTWARE_ORDER),
        help="Software list to run. Default: all 6 software entries.",
    )
    parser.add_argument(
        "--exclude-software",
        nargs="+",
        choices=DEFAULT_SOFTWARE_ORDER,
        default=[],
        help="Software list to skip after selection. Useful when one app has already been validated.",
    )
    parser.add_argument("--results-root", default="results/pipeline_runs", help="Root directory for pipeline outputs.")
    parser.add_argument("--ai-turbo-window-title", default=DEFAULT_AI_TURBO_TITLE_RE, help="AI Turbo Engine main window title regex.")
    parser.add_argument("--ai-turbo-shortcut", default=str(REPO_ROOT / "whitelist app" / "AI Turbo Engine.lnk"), help="Optional AI Turbo Engine shortcut path.")
    parser.add_argument(
        "--run-ai-turbo-sequence",
        "--run_ai_turbo_sequence",
        dest="run_ai_turbo_sequence",
        action="store_true",
        help="Run only the standalone AI Turbo flow: open AI Turbo, enable boost, wait, disable boost, then close AI Turbo.",
    )
    parser.add_argument(
        "--run-ai-turbo-open-check",
        "--run_ai_turbo_open_check",
        dest="run_ai_turbo_open_check",
        action="store_true",
        help="Run only an AI Turbo Engine smoke test: open the app, verify the window appears, optionally wait, then close it.",
    )
    parser.add_argument(
        "--ai-turbo-sequence-software",
        default="",
        choices=("", *DEFAULT_SOFTWARE_ORDER),
        help="Target software for the standalone AI Turbo flow. If omitted, --software must resolve to exactly one entry.",
    )
    parser.add_argument(
        "--ai-turbo-sequence-wait-seconds",
        type=float,
        default=30.0,
        help="How long the standalone AI Turbo flow should keep boost enabled before turning it off.",
    )
    parser.add_argument(
        "--ai-turbo-open-check-wait-seconds",
        type=float,
        default=5.0,
        help="How long the standalone AI Turbo open-check should keep the app open before closing it.",
    )
    parser.add_argument("--blend-file", default="", help="Optional override blend file for blender runs. Defaults to whitelist app/blender/*.blend from the software spec.")
    parser.add_argument("--blender-exe", default="", help="Optional Blender executable path for the background Blender render pass.")
    parser.add_argument(
        "--blender-ui-mode",
        choices=("visible", "headless"),
        default="visible",
        help="Whether Blender should render with a visible UI window or in headless background mode.",
    )
    parser.add_argument("--render-mode", choices=("frame", "animation"), default="frame", help="Blender render mode for the pipeline run.")
    parser.add_argument("--frame", type=int, default=1, help="Blender frame number when --render-mode frame is used.")
    parser.add_argument(
        "--resume-run-root",
        default="",
        help="Optional existing pipeline run directory. When provided, completed cases are reused and only missing/failed cases are rerun.",
    )
    parser.add_argument(
        "--skip-baseline",
        "--skip_baseline",
        dest="skip_baseline",
        action="store_true",
        help="Skip the baseline run and execute only the AI Turbo pass.",
    )
    parser.add_argument(
        "--skip-turbo",
        "--skip_turbo",
        dest="skip_turbo",
        action="store_true",
        help="Skip the AI Turbo pass and execute only the baseline run.",
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s | %(levelname)s | %(message)s",
    )
    parser = build_parser()
    args = parser.parse_args(argv)
    assert not (args.run_ai_turbo_sequence and args.run_ai_turbo_open_check), (
        "Choose only one standalone AI Turbo mode: --run-ai-turbo-sequence or --run-ai-turbo-open-check."
    )
    if args.run_ai_turbo_open_check:
        logger.info("Starting standalone AI Turbo open check.")
        run_ai_turbo_open_check(args)
        print(f"ai_turbo_open_check_wait_seconds={float(args.ai_turbo_open_check_wait_seconds):.1f}")
        print("ai_turbo_open_check_status=completed")
        return 0
    if args.run_ai_turbo_sequence:
        logger.info("Starting standalone AI Turbo sequence.")
        software = run_ai_turbo_sequence(args)
        print(f"ai_turbo_sequence_software={software}")
        print(f"ai_turbo_sequence_wait_seconds={float(args.ai_turbo_sequence_wait_seconds):.1f}")
        print("ai_turbo_sequence_status=completed")
        return 0
    assert not (args.skip_baseline and args.skip_turbo), "Cannot skip both baseline and turbo passes."
    logger.info("Starting full automation pipeline.")
    report_path = run_pipeline(args)
    print(f"pipeline_report={report_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
