from __future__ import annotations

import argparse
from dataclasses import dataclass
from datetime import datetime
import logging
from pathlib import Path
import re

from workspace_runtime import configure_workspace_runtime

configure_workspace_runtime()

from automation_components import (
    DEFAULT_AI_TURBO_TITLE_RE,
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


@dataclass(frozen=True)
class PipelinePaths:
    root_dir: Path
    csv_dir: Path
    export_dir: Path
    report_dir: Path


def slugify(value: str) -> str:
    normalized = re.sub(r"[^\w.-]+", "_", value.strip().lower())
    return normalized.strip("._") or "pipeline_run"


def build_pipeline_paths(results_root: Path, workload_name: str) -> PipelinePaths:
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    run_root = results_root / f"{slugify(workload_name)}_{timestamp}"
    csv_dir = run_root / "csv"
    export_dir = run_root / "exports"
    report_dir = run_root / "report"
    csv_dir.mkdir(parents=True, exist_ok=True)
    export_dir.mkdir(parents=True, exist_ok=True)
    report_dir.mkdir(parents=True, exist_ok=True)
    assert csv_dir.exists() and export_dir.exists() and report_dir.exists()
    return PipelinePaths(root_dir=run_root, csv_dir=csv_dir, export_dir=export_dir, report_dir=report_dir)


def run_non_blender_case(
    *,
    software: str,
    variant: str,
    workload_name: str,
    input_video_path: Path,
    pipeline_paths: PipelinePaths,
    launcher: SoftwareLauncher,
    monitor_bridge: MonitorBridge,
    ai_turbo_controller: AiTurboEngineController | None,
) -> Path:
    export_profile = OPERATION_PROFILES[software]
    output_video_path = pipeline_paths.export_dir / f"{software}_{slugify(workload_name)}_{variant}{export_profile.output_suffix}"
    requested_csv_path = pipeline_paths.csv_dir / f"{software}_{slugify(workload_name)}_{variant}.csv"
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

    try:
        logger.info("Launching %s for the %s run.", software, variant)
        launcher.launch(software)
        launcher.wait_for_main_window(software, timeout=30.0)
        logger.info("Running the %s UI automation flow for the %s run.", software, variant)
        operator.perform(input_video_path, output_video_path)
        logger.info("Waiting for the %s timing monitor to finalize.", software)
        status_payload = monitor_bridge.wait_for_session_completion(launch_result.session_id)
    except Exception:
        logger.exception("%s %s run failed. Requesting the monitor session to stop.", software, variant)
        try:
            monitor_bridge.stop_session(launch_result.session_id)
        except Exception:
            logger.exception("Could not stop the %s monitor session cleanly.", software)
        raise
    finally:
        operator.close()

    assert status_payload["status"] in {"completed", "completed_with_warnings"}, status_payload
    assert Path(launch_result.output_path).exists(), f"Expected csv output was not created: {launch_result.output_path}"
    return Path(launch_result.output_path)


def run_blender_case(
    *,
    variant: str,
    workload_name: str,
    pipeline_paths: PipelinePaths,
    monitor_bridge: MonitorBridge,
    ai_turbo_controller: AiTurboEngineController | None,
    blend_file: Path | None = None,
    blender_executable: Path | None = None,
    render_mode: str = "frame",
    frame: int = 1,
) -> Path:
    spec = SOFTWARE_SPECS["blender"]
    effective_blend_file = blend_file or spec.blend_file
    assert effective_blend_file is not None and effective_blend_file.exists(), "Blender blend file is required."
    requested_csv_path = pipeline_paths.csv_dir / f"blender_{slugify(workload_name)}_{variant}.csv"
    run_name = f"{slugify(workload_name)}_{variant}"
    logger.info(
        "Preparing blender %s run. blend_file=%s monitor_csv=%s render_mode=%s frame=%s",
        variant,
        effective_blend_file,
        requested_csv_path,
        render_mode,
        frame,
    )

    if ai_turbo_controller is not None:
        logger.info("Configuring AI Turbo Engine Performance Boost for blender.")
        ai_turbo_controller.configure_for_software("blender")
    try:
        csv_path = monitor_bridge.run_blender_monitor(
            test_name=run_name,
            output_path=requested_csv_path,
            blend_file=effective_blend_file,
            blender_executable=blender_executable,
            render_mode=render_mode,
            frame=frame,
        )
    except Exception:
        logger.exception("Blender %s run failed.", variant)
        raise
    assert csv_path.exists(), f"Expected blender csv output was not created: {csv_path}"
    return csv_path


def run_pipeline(args: argparse.Namespace) -> Path:
    selected_software = tuple(args.software or DEFAULT_SOFTWARE_ORDER)
    assert selected_software, "At least one software entry is required."
    for software in selected_software:
        assert software in SOFTWARE_SPECS, f"Unsupported software: {software}"

    pipeline_paths = build_pipeline_paths(Path(args.results_root), args.workload_name)
    input_video_path = Path(args.input_video).resolve(strict=False) if args.input_video else resolve_input_video(args.workload_name)
    assert input_video_path.exists(), f"Input video does not exist: {input_video_path}"

    launcher = SoftwareLauncher()
    monitor_bridge = MonitorBridge()
    ai_turbo_shortcut = Path(args.ai_turbo_shortcut).resolve(strict=False) if args.ai_turbo_shortcut else None
    ai_turbo_controller = AiTurboEngineController(
        window_title_re=args.ai_turbo_window_title,
        shortcut_path=ai_turbo_shortcut,
    )
    resolved_blend_file = Path(args.blend_file).resolve(strict=False) if args.blend_file else None
    resolved_blender_executable = Path(args.blender_exe).resolve(strict=False) if args.blender_exe else None

    csv_paths: list[Path] = []
    passes = [("baseline", False), ("turbo", True)]
    if args.skip_turbo:
        passes = [("baseline", False)]
    if args.skip_baseline:
        passes = [("turbo", True)]

    for variant, enable_turbo in passes:
        logger.info("Starting pipeline pass '%s'. ai_turbo_enabled=%s", variant, enable_turbo)
        for software in selected_software:
            logger.info("Starting pipeline case. software=%s variant=%s", software, variant)
            controller = ai_turbo_controller if enable_turbo else None
            if software == "blender":
                csv_path = run_blender_case(
                    variant=variant,
                    workload_name=args.workload_name,
                    pipeline_paths=pipeline_paths,
                    monitor_bridge=monitor_bridge,
                    ai_turbo_controller=controller,
                    blend_file=resolved_blend_file,
                    blender_executable=resolved_blender_executable,
                    render_mode=args.render_mode,
                    frame=args.frame,
                )
            else:
                csv_path = run_non_blender_case(
                    software=software,
                    variant=variant,
                    workload_name=args.workload_name,
                    input_video_path=input_video_path,
                    pipeline_paths=pipeline_paths,
                    launcher=launcher,
                    monitor_bridge=monitor_bridge,
                    ai_turbo_controller=controller,
                )
            csv_paths.append(csv_path)
            logger.info("Pipeline case finished. software=%s variant=%s csv=%s", software, variant, csv_path)

    report_path = pipeline_paths.report_dir / f"{slugify(args.workload_name)}_summary.xlsx"
    logger.info("Generating xlsx report with baseline/turbo comparison formulas. output=%s", report_path)
    csv_count, row_count, generated_report_path = generate_xlsx_report(csv_paths, report_path, default_test_case=args.workload_name)
    assert csv_count == len(csv_paths), (csv_count, len(csv_paths))
    assert row_count >= len(csv_paths), (row_count, len(csv_paths))
    assert generated_report_path.exists(), f"Expected xlsx report was not created: {generated_report_path}"
    return generated_report_path


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Run the full baseline vs AI Turbo video-processing automation workflow.")
    parser.add_argument("--workload-name", default="1080p_video_processing_speed", help="Logical workload label used in csv/xlsx outputs.")
    parser.add_argument("--input-video", default="", help="Optional override video path. Defaults to test_data mapping based on workload name.")
    parser.add_argument("--software", nargs="*", default=list(DEFAULT_SOFTWARE_ORDER), help="Software list to run. Default: all 6 software entries.")
    parser.add_argument("--results-root", default="results/pipeline_runs", help="Root directory for pipeline outputs.")
    parser.add_argument("--ai-turbo-window-title", default=DEFAULT_AI_TURBO_TITLE_RE, help="AI Turbo Engine main window title regex.")
    parser.add_argument("--ai-turbo-shortcut", default=str(REPO_ROOT / "whitelist app" / "AI Turbo Engine.lnk"), help="Optional AI Turbo Engine shortcut path.")
    parser.add_argument("--blend-file", default="", help="Optional override blend file for blender runs. Defaults to whitelist app/blender/*.blend from the software spec.")
    parser.add_argument("--blender-exe", default="", help="Optional Blender executable path for the background Blender render pass.")
    parser.add_argument("--render-mode", choices=("frame", "animation"), default="frame", help="Blender render mode for the pipeline run.")
    parser.add_argument("--frame", type=int, default=1, help="Blender frame number when --render-mode frame is used.")
    parser.add_argument("--skip-baseline", action="store_true", help="Skip the baseline run and execute only the AI Turbo pass.")
    parser.add_argument("--skip-turbo", action="store_true", help="Skip the AI Turbo pass and execute only the baseline run.")
    return parser


def main(argv: list[str] | None = None) -> int:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s | %(levelname)s | %(message)s",
    )
    parser = build_parser()
    args = parser.parse_args(argv)
    assert not (args.skip_baseline and args.skip_turbo), "Cannot skip both baseline and turbo passes."
    logger.info("Starting full automation pipeline.")
    report_path = run_pipeline(args)
    print(f"pipeline_report={report_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
