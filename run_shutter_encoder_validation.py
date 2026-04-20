from __future__ import annotations

import argparse
import logging
from pathlib import Path

from workspace_runtime import configure_workspace_runtime

configure_workspace_runtime()

from automation_components import MonitorBridge, REPO_ROOT, SoftwareLauncher, default_monitor_output_path, resolve_input_video
from software_operations import run_shutter_encoder_operation
from ui_automation import UiAutomationError


logger = logging.getLogger(__name__)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Run only the Shutter Encoder UI automation flow for quick validation.")
    parser.add_argument(
        "--workload",
        choices=("1080", "4k"),
        default="1080",
        help="Pick the test_data video preset.",
    )
    parser.add_argument(
        "--input-video",
        default="",
        help="Optional explicit input video path. Overrides --workload.",
    )
    parser.add_argument(
        "--output-dir",
        default=str(REPO_ROOT / "results" / "shutter_encoder_validation"),
        help="Directory for the copied validation video.",
    )
    parser.add_argument(
        "--test-id",
        required=True,
        help="Video output file stem used for validation, for example shutter_encoder_smoke_001.",
    )
    parser.add_argument(
        "--monitor-output",
        default="",
        help="Optional csv path for the background timing monitor. Defaults to results/csv/<software>_<test-id>.csv.",
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s | %(levelname)s | %(message)s",
    )
    parser = build_parser()
    args = parser.parse_args(argv)

    logger.info("Starting Shutter Encoder validation entrypoint.")
    if args.input_video:
        input_video_path = Path(args.input_video).resolve(strict=False)
    else:
        workload_name = "4k_video_processing_speed" if args.workload == "4k" else "1080p_video_processing_speed"
        input_video_path = resolve_input_video(workload_name)
    logger.info("Validating selected input video path: %s", input_video_path)
    assert input_video_path.exists(), f"Input video does not exist: {input_video_path}"

    output_dir = Path(args.output_dir).resolve(strict=False)
    logger.info("Ensuring Shutter Encoder validation output directory exists: %s", output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    output_video_path = output_dir / f"{args.test_id}.mp4"
    logger.info("Resolved Shutter Encoder validation output file: %s", output_video_path)

    monitor_bridge = MonitorBridge()
    monitor_output_path = (
        Path(args.monitor_output).resolve(strict=False)
        if args.monitor_output
        else default_monitor_output_path("shutter_encoder", args.test_id)
    )
    logger.info("Ensuring Shutter Encoder timing output directory exists: %s", monitor_output_path.parent)
    monitor_output_path.parent.mkdir(parents=True, exist_ok=True)
    logger.info("Starting the Shutter Encoder background timing monitor.")
    launch_result = monitor_bridge.start_background_monitor("shutter_encoder", args.test_id, monitor_output_path)

    launcher = SoftwareLauncher()
    try:
        logger.info("Checking whether Shutter Encoder is already running.")
        launcher.wait_for_main_window("shutter_encoder", timeout=2.0)
    except UiAutomationError:
        logger.info("Shutter Encoder is not running. Launching it now.")
        launcher.launch("shutter_encoder")
        launcher.wait_for_main_window("shutter_encoder", timeout=30.0)

    try:
        logger.info("Running Shutter Encoder operation.")
        run_shutter_encoder_operation(input_video_path, output_video_path)
    except Exception:
        logger.exception("Shutter Encoder validation failed. Requesting the background timing monitor to stop.")
        try:
            monitor_bridge.stop_session(launch_result.session_id)
        except Exception:
            logger.exception("Could not stop the Shutter Encoder background timing monitor cleanly.")
        raise

    logger.info("Waiting for the Shutter Encoder background timing monitor to finalize.")
    status_payload = monitor_bridge.wait_for_session_completion(launch_result.session_id)
    logger.info("Shutter Encoder timing monitor finished with status: %s", status_payload.get("status", "<unknown>"))
    assert status_payload["status"] in {"completed", "completed_with_warnings"}, status_payload
    assert Path(launch_result.output_path).exists(), f"Shutter Encoder timing csv was not created: {launch_result.output_path}"

    logger.info("Validating Shutter Encoder output file exists: %s", output_video_path)
    assert output_video_path.exists(), f"Shutter Encoder validation did not create the expected output: {output_video_path}"

    print(f"input_video={input_video_path}")
    print(f"monitor_csv={launch_result.output_path}")
    print(f"monitor_status={status_payload['status']}")
    print(f"output_video={output_video_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
