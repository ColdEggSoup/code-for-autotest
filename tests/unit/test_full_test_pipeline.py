from __future__ import annotations

import csv
from argparse import Namespace
from pathlib import Path
from types import SimpleNamespace

from openpyxl import load_workbook

import full_test_pipeline
import xlsx_report_generator


def _write_monitor_csv(csv_path: Path, *, software: str, test_name: str, duration_seconds: float) -> None:
    csv_path.parent.mkdir(parents=True, exist_ok=True)
    with csv_path.open("w", newline="", encoding="utf-8-sig") as handle:
        writer = csv.DictWriter(
            handle,
            fieldnames=("software", "duration_seconds", "started_at", "ended_at", "test_name"),
        )
        writer.writeheader()
        writer.writerow(
            {
                "software": software,
                "duration_seconds": str(duration_seconds),
                "started_at": "2026-04-21T10:00:00.000+08:00",
                "ended_at": "2026-04-21T10:00:10.000+08:00",
                "test_name": test_name,
            }
        )


class DummyLauncher:
    def __init__(self) -> None:
        self.launched: list[str] = []
        self.waited: list[tuple[str, float]] = []

    def launch(self, software: str) -> None:
        self.launched.append(software)

    def wait_for_main_window(self, software: str, timeout: float = 30.0) -> None:
        self.waited.append((software, timeout))


class DummyOperator:
    def __init__(self, software: str, event_log: list[tuple[str, str]]) -> None:
        self.software = software
        self.event_log = event_log

    def perform(self, input_video_path: Path, output_video_path: Path) -> None:
        assert input_video_path.exists()
        output_video_path.parent.mkdir(parents=True, exist_ok=True)
        output_video_path.write_bytes(b"fake-video")
        self.event_log.append(("perform", self.software))

    def close(self) -> None:
        self.event_log.append(("close", self.software))


class DummyMonitorBridge:
    def __init__(self) -> None:
        self.sessions: dict[str, tuple[str, str, Path]] = {}
        self.waited_sessions: list[str] = []
        self.stopped_sessions: list[str] = []

    def start_background_monitor(self, software: str, test_name: str, output_path: Path):
        output_path.parent.mkdir(parents=True, exist_ok=True)
        session_id = f"{software}-{test_name}"
        self.sessions[session_id] = (software, test_name, output_path)
        return SimpleNamespace(session_id=session_id, output_path=output_path)

    def wait_for_session_completion(self, session_id: str) -> dict[str, str]:
        self.waited_sessions.append(session_id)
        software, test_name, output_path = self.sessions[session_id]
        duration_seconds = 5.0 if test_name.endswith("turbo") else 10.0
        _write_monitor_csv(
            output_path,
            software=software,
            test_name=test_name,
            duration_seconds=duration_seconds,
        )
        return {"status": "completed"}

    def stop_session(self, session_id: str, *, wait_seconds: float = 15.0) -> dict[str, str]:
        self.stopped_sessions.append(session_id)
        software, test_name, output_path = self.sessions[session_id]
        duration_seconds = 4.0 if test_name.endswith("turbo") else 9.0
        _write_monitor_csv(
            output_path,
            software=software,
            test_name=test_name,
            duration_seconds=duration_seconds,
        )
        return {"status": "completed_with_warnings"}

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
    ) -> Path:
        assert blend_file.exists()
        assert blender_ui_mode in {"visible", "headless"}
        duration_seconds = 5.0 if test_name.endswith("turbo") else 10.0
        _write_monitor_csv(
            output_path,
            software="blender",
            test_name=test_name,
            duration_seconds=duration_seconds,
        )
        return output_path


class DummyAiTurboController:
    def __init__(self) -> None:
        self.configured: list[str] = []
        self.started = 0
        self.stopped = 0

    def start_topmost_guard(self) -> None:
        self.started += 1

    def stop_topmost_guard(self) -> None:
        self.stopped += 1

    def configure_for_software(self, software: str) -> None:
        self.configured.append(software)


def test_run_pipeline_with_test_doubles_generates_comparison_report(monkeypatch, tmp_path: Path) -> None:
    input_video = tmp_path / "input.mp4"
    input_video.write_bytes(b"input-video")
    blend_file = tmp_path / "scene.blend"
    blend_file.write_text("blend", encoding="utf-8")

    launcher = DummyLauncher()
    monitor_bridge = DummyMonitorBridge()
    ai_turbo_controller = DummyAiTurboController()
    operator_log: list[tuple[str, str]] = []

    monkeypatch.setattr(full_test_pipeline, "SoftwareLauncher", lambda: launcher)
    monkeypatch.setattr(full_test_pipeline, "MonitorBridge", lambda: monitor_bridge)
    monkeypatch.setattr(full_test_pipeline, "AiTurboEngineController", lambda **kwargs: ai_turbo_controller)
    monkeypatch.setattr(full_test_pipeline, "resolve_input_video", lambda workload_name: input_video)
    monkeypatch.setattr(
        full_test_pipeline,
        "build_operator",
        lambda software: DummyOperator(software, operator_log),
    )
    monkeypatch.setattr(
        xlsx_report_generator,
        "collect_device_info",
        lambda: [("device_name", "pytest-host"), ("windows_system_version", "pytest-os")],
    )

    args = Namespace(
        workload_name="demo",
        input_video="",
        software=["shotcut", "blender"],
        results_root=str(tmp_path / "results"),
        ai_turbo_window_title="^AI Turbo Engine$",
        ai_turbo_shortcut="",
        blend_file=str(blend_file),
        blender_exe="",
        blender_ui_mode="visible",
        render_mode="frame",
        frame=1,
        skip_baseline=False,
        skip_turbo=False,
    )

    report_path = full_test_pipeline.run_pipeline(args)
    assert report_path.exists()
    run_root = report_path.parent.parent
    progress_log = run_root / "pipeline_progress.jsonl"
    summary_path = run_root / "pipeline_summary.md"
    assert progress_log.exists()
    assert summary_path.exists()
    summary_text = summary_path.read_text(encoding="utf-8")
    assert "status: completed" in summary_text
    assert "| shotcut | baseline | completed |" in summary_text
    assert "| blender | turbo | completed |" in summary_text

    workbook = load_workbook(report_path, data_only=False)
    try:
        comparison_sheet = workbook["Comparison"]
        assert comparison_sheet.max_row == 3
        assert {comparison_sheet["A2"].value, comparison_sheet["A3"].value} == {"shotcut", "blender"}
        assert comparison_sheet["E2"].value == 10.0
        assert comparison_sheet["F2"].value == 5.0
        assert comparison_sheet["E3"].value == 10.0
        assert comparison_sheet["F3"].value == 5.0
        assert comparison_sheet["G2"].value == '=IF(OR(E2="",F2="",E2<=0),"",(E2-F2)/E2)'
        assert comparison_sheet["G3"].value == '=IF(OR(E3="",F3="",E3<=0),"",(E3-F3)/E3)'
    finally:
        workbook.close()

    assert ai_turbo_controller.started == 1
    assert ai_turbo_controller.stopped == 1
    assert ai_turbo_controller.configured == ["shotcut", "blender"]
    assert ("perform", "shotcut") in operator_log


def test_build_pipeline_paths_resolves_relative_results_root(monkeypatch, tmp_path: Path) -> None:
    monkeypatch.chdir(tmp_path)

    pipeline_paths = full_test_pipeline.build_pipeline_paths(Path("results/pipeline_runs"), "demo")

    expected_root = (tmp_path / "results" / "pipeline_runs").resolve(strict=False)
    assert pipeline_paths.root_dir.is_absolute()
    assert pipeline_paths.csv_dir.is_absolute()
    assert pipeline_paths.export_dir.is_absolute()
    assert pipeline_paths.report_dir.is_absolute()
    assert pipeline_paths.root_dir.parent == expected_root
    assert pipeline_paths.csv_dir.parent == pipeline_paths.root_dir
    assert pipeline_paths.export_dir.parent == pipeline_paths.root_dir
    assert pipeline_paths.report_dir.parent == pipeline_paths.root_dir


def test_run_pipeline_honors_excluded_software(monkeypatch, tmp_path: Path) -> None:
    input_video = tmp_path / "input.mp4"
    input_video.write_bytes(b"input-video")
    selected_software_per_pass: list[tuple[str, ...]] = []

    def _fake_run_pipeline_pass(**kwargs):
        selected_software_per_pass.append(kwargs["selected_software"])
        return []

    def _fake_generate_xlsx_report(csv_paths, report_path, default_test_case):
        report_path.parent.mkdir(parents=True, exist_ok=True)
        report_path.write_bytes(b"fake-report")
        return 0, 0, report_path

    monkeypatch.setattr(full_test_pipeline, "resolve_input_video", lambda workload_name: input_video)
    monkeypatch.setattr(full_test_pipeline, "run_pipeline_pass", _fake_run_pipeline_pass)
    monkeypatch.setattr(full_test_pipeline, "SoftwareLauncher", lambda: object())
    monkeypatch.setattr(full_test_pipeline, "MonitorBridge", lambda: object())
    monkeypatch.setattr(full_test_pipeline, "AiTurboEngineController", lambda **kwargs: object())
    monkeypatch.setattr(full_test_pipeline, "generate_xlsx_report", _fake_generate_xlsx_report)

    args = Namespace(
        workload_name="demo",
        input_video="",
        software=["shotcut", "kdenlive", "blender"],
        exclude_software=["shotcut"],
        results_root=str(tmp_path / "results"),
        ai_turbo_window_title="^AI Turbo Engine$",
        ai_turbo_shortcut="",
        blend_file="",
        blender_exe="",
        blender_ui_mode="visible",
        render_mode="frame",
        frame=1,
        skip_baseline=False,
        skip_turbo=True,
    )

    report_path = full_test_pipeline.run_pipeline(args)

    assert report_path.exists()
    assert selected_software_per_pass == [("kdenlive", "blender")]


def test_run_non_blender_case_stops_avidemux_monitor_after_automation(monkeypatch, tmp_path: Path) -> None:
    input_video = tmp_path / "input.mp4"
    input_video.write_bytes(b"input-video")
    pipeline_paths = full_test_pipeline.build_pipeline_paths(tmp_path / "results", "demo")
    launcher = DummyLauncher()
    monitor_bridge = DummyMonitorBridge()
    operator_log: list[tuple[str, str]] = []
    recorder = full_test_pipeline.PipelineRunRecorder(pipeline_paths.root_dir)

    monkeypatch.setattr(
        full_test_pipeline,
        "build_operator",
        lambda software: DummyOperator(software, operator_log),
    )

    csv_path = full_test_pipeline.run_non_blender_case(
        software="avidemux",
        variant="baseline",
        workload_name="demo",
        input_video_path=input_video,
        pipeline_paths=pipeline_paths,
        launcher=launcher,
        monitor_bridge=monitor_bridge,
        ai_turbo_controller=None,
        recorder=recorder,
    )

    assert csv_path.exists()
    assert monitor_bridge.stopped_sessions == ["avidemux-demo_baseline"]
    assert monitor_bridge.waited_sessions == []
