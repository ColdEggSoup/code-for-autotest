from __future__ import annotations

from pathlib import Path

import log_listener


def test_should_finalize_avidemux_completed_capture_after_settle_window() -> None:
    profile = log_listener.LOG_PROFILES["avidemux"]
    live_preview = {
        "capture_complete": True,
        "preview_status": "complete",
    }

    assert (
        log_listener.should_finalize_avidemux_completed_capture(
            profile,
            live_preview,
            completion_detected_at=10.0,
            now_monotonic=10.5,
        )
        is False
    )
    assert (
        log_listener.should_finalize_avidemux_completed_capture(
            profile,
            live_preview,
            completion_detected_at=10.0,
            now_monotonic=11.1,
        )
        is True
    )


def test_monitor_single_log_file_finalizes_avidemux_without_waiting_for_cleanup_idle(
    monkeypatch,
    tmp_path: Path,
) -> None:
    log_path = tmp_path / "admlog.txt"
    log_path.write_text("", encoding="utf-8")

    config = {
        "software": "avidemux",
        "session_id": "avidemux_test_session",
        "options": {
            "log_path": str(log_path),
            "poll_interval": 0.5,
            "idle_seconds": 3.0,
            "max_runtime_seconds": 30.0,
        },
    }
    profile = log_listener.LOG_PROFILES["avidemux"]
    fake_clock = {"value": 0.0}
    write_once = {"done": False}

    completed_log = "\n".join(
        [
            "[A_Save] Saving..",
            "[muxerFFmpeg::saveLoop] 06:21:41-577 [FF] Saving",
            "[x264Encoder::encode] 06:27:58-036 End of flush",
            "100% done\tframes: 4828\telapsed: 00:06:16,941",
        ]
    ) + "\n"

    monkeypatch.setattr(log_listener.time, "monotonic", lambda: fake_clock["value"])

    def _sleep(seconds: float) -> None:
        if not write_once["done"]:
            log_path.write_text(completed_log, encoding="utf-8")
            write_once["done"] = True
        fake_clock["value"] += seconds

    monkeypatch.setattr(log_listener.time, "sleep", _sleep)

    rows = log_listener.monitor_single_log_file(config, profile, tmp_path / "stop.signal")

    assert rows[0]["status"] == "ok"
    assert rows[0]["evidence"] == "avidemux_elapsed_100_percent_with_time_window"
    assert float(rows[0]["duration_seconds"]) > 300.0
    assert fake_clock["value"] < 3.0
