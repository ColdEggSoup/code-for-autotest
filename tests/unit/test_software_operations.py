from __future__ import annotations

from pathlib import Path

import software_operations


class _DummyWindow:
    def __init__(self, title: str = "", *, process_id: int = 100) -> None:
        self._title = title
        self.handle = 100
        self.element_info = type("ElementInfo", (), {"process_id": process_id})()
        self._children = []

    def window_text(self) -> str:
        return self._title

    def close(self) -> None:
        return None

    def descendants(self):
        return list(self._children)

    def is_visible(self) -> bool:
        return True


class _DummyButton:
    def __init__(self) -> None:
        self.clicked = False

    def click_input(self) -> None:
        self.clicked = True


class _DummyDialog(_DummyWindow):
    def __init__(self, title: str = "Info", *, process_id: int = 100, texts: list[str] | None = None) -> None:
        super().__init__(title, process_id=process_id)
        self._children = [_DummyWindow(text, process_id=process_id) for text in (texts or [])]


class _VisibleControl:
    def is_visible(self) -> bool:
        return True


def test_avidemux_open_input_uses_standard_file_dialog(monkeypatch, tmp_path: Path) -> None:
    operator = software_operations.AvidemuxOperator(software_operations.OPERATION_PROFILES["avidemux"])
    input_video_path = tmp_path / "input.mp4"
    input_video_path.write_bytes(b"video")
    recorded: dict[str, object] = {}

    monkeypatch.setattr(software_operations, "bring_window_to_front", lambda *args, **kwargs: None)
    monkeypatch.setattr(software_operations, "send_hotkey", lambda keys: recorded.setdefault("hotkey", keys))
    monkeypatch.setattr(operator, "_dismiss_thanks_popup_if_present", lambda *args, **kwargs: False)
    monkeypatch.setattr(operator, "_dismiss_open_failure_dialog_if_present", lambda *args, **kwargs: "")
    monkeypatch.setattr(operator, "_wait_for_loaded_content", lambda timeout=15.0: True)
    monkeypatch.setattr(
        software_operations,
        "fill_file_dialog",
        lambda path, **kwargs: recorded.update({"path": path, **kwargs}),
    )

    operator._open_input_clip(_DummyWindow("Avidemux"), input_video_path)

    assert recorded["hotkey"] == "^o"
    assert recorded["path"] == input_video_path
    assert recorded["dialog_patterns"] == software_operations.AVIDEMUX_OPEN_DIALOG_PATTERNS
    assert recorded["must_exist"] is True


def test_avidemux_open_input_reports_missing_demuxer(monkeypatch, tmp_path: Path) -> None:
    operator = software_operations.AvidemuxOperator(software_operations.OPERATION_PROFILES["avidemux"])
    input_video_path = tmp_path / "input.mp4"
    input_video_path.write_bytes(b"video")

    monkeypatch.setattr(software_operations, "bring_window_to_front", lambda *args, **kwargs: None)
    monkeypatch.setattr(software_operations, "send_hotkey", lambda keys: None)
    monkeypatch.setattr(operator, "_dismiss_thanks_popup_if_present", lambda *args, **kwargs: False)
    monkeypatch.setattr(
        operator,
        "_dismiss_open_failure_dialog_if_present",
        lambda *args, **kwargs: "Cannot find a demuxer for C:/tmp/input.mp4",
    )
    monkeypatch.setattr(
        software_operations,
        "fill_file_dialog",
        lambda path, **kwargs: None,
    )

    try:
        operator._open_input_clip(_DummyWindow("Avidemux"), input_video_path)
    except AssertionError as exc:
        assert "demuxer plugins are unavailable" in str(exc)
    else:
        raise AssertionError("Expected _open_input_clip to fail when Avidemux reports a demuxer error.")


def test_avidemux_dismiss_export_success_dialog(monkeypatch, tmp_path: Path) -> None:
    operator = software_operations.AvidemuxOperator(software_operations.OPERATION_PROFILES["avidemux"])
    output_video_path = tmp_path / "out.mp4"
    dialog = _DummyDialog(texts=["Done", f"File {output_video_path.name} has been successfully saved."])
    ok_button = _DummyButton()

    monkeypatch.setattr(operator, "_find_process_top_level_window_by_content", lambda *args, **kwargs: dialog)
    monkeypatch.setattr(software_operations, "bring_window_to_front", lambda *args, **kwargs: None)
    monkeypatch.setattr(software_operations, "find_text_control", lambda *args, **kwargs: ok_button)
    monkeypatch.setattr(software_operations.time, "sleep", lambda _seconds: None)

    assert operator._dismiss_export_success_dialog_if_present(_DummyWindow("Avidemux"), output_video_path) is True
    assert ok_button.clicked is True


def test_avidemux_wait_for_export_completion_accepts_done_dialog(monkeypatch, tmp_path: Path) -> None:
    operator = software_operations.AvidemuxOperator(software_operations.OPERATION_PROFILES["avidemux"])
    output_video_path = tmp_path / "out.mp4"
    output_video_path.write_bytes(b"video")

    monkeypatch.setattr(operator, "_connect_main_window", lambda timeout=0.5: _DummyWindow("Avidemux"))
    monkeypatch.setattr(operator, "_dismiss_export_success_dialog_if_present", lambda *args, **kwargs: True)
    monkeypatch.setattr(software_operations.time, "sleep", lambda _seconds: None)

    operator._wait_for_export_completion(output_video_path)


def test_avidemux_close_dismisses_success_dialog_before_closing(monkeypatch) -> None:
    operator = software_operations.AvidemuxOperator(software_operations.OPERATION_PROFILES["avidemux"])
    calls: list[str] = []

    class _CloseTrackingWindow(_DummyWindow):
        def close(self) -> None:
            calls.append("close")

    window = _CloseTrackingWindow("Avidemux")

    monkeypatch.setattr(operator, "_connect_main_window", lambda timeout=2.0: window)
    monkeypatch.setattr(
        operator,
        "_dismiss_export_success_dialog_if_present",
        lambda *args, **kwargs: calls.append("dismiss") or True,
    )
    monkeypatch.setattr(operator, "_iter_process_top_level_windows", lambda process_id: [])
    monkeypatch.setattr(software_operations, "bring_window_to_front", lambda *args, **kwargs: None)
    monkeypatch.setattr(software_operations, "dismiss_close_prompts", lambda *args, **kwargs: None)
    monkeypatch.setattr(software_operations.time, "sleep", lambda _seconds: None)

    operator.close()

    assert calls[:2] == ["dismiss", "close"]


def test_avidemux_close_terminates_process_if_windows_still_visible(monkeypatch) -> None:
    operator = software_operations.AvidemuxOperator(software_operations.OPERATION_PROFILES["avidemux"])

    class _CloseTrackingWindow(_DummyWindow):
        def __init__(self, title: str = "", *, process_id: int = 100) -> None:
            super().__init__(title, process_id=process_id)
            self.close_calls = 0

        def close(self) -> None:
            self.close_calls += 1

    window = _CloseTrackingWindow("Avidemux", process_id=4321)
    visible_windows = [window]
    terminated: list[list[str]] = []

    monkeypatch.setattr(operator, "_connect_main_window", lambda timeout=2.0: window)
    monkeypatch.setattr(operator, "_dismiss_export_success_dialog_if_present", lambda *args, **kwargs: False)

    def _iter_windows(_process_id):
        return list(visible_windows)

    monkeypatch.setattr(operator, "_iter_process_top_level_windows", _iter_windows)
    monkeypatch.setattr(software_operations, "bring_window_to_front", lambda *args, **kwargs: None)
    monkeypatch.setattr(software_operations, "dismiss_close_prompts", lambda *args, **kwargs: None)
    monkeypatch.setattr(
        software_operations.subprocess,
        "run",
        lambda command, **kwargs: terminated.append(command) or type("Completed", (), {"returncode": 0})(),
    )

    fake_clock = {"value": 0.0}

    monkeypatch.setattr(software_operations.time, "monotonic", lambda: fake_clock["value"])

    def _sleep(seconds: float) -> None:
        fake_clock["value"] += seconds
        if terminated:
            visible_windows.clear()

    monkeypatch.setattr(software_operations.time, "sleep", _sleep)

    operator.close()

    assert window.close_calls == software_operations.AVIDEMUX_CLOSE_RETRY_COUNT
    assert terminated == [["taskkill", "/PID", "4321", "/T", "/F"]]


def test_handbrake_open_input_uses_standard_file_dialog(monkeypatch, tmp_path: Path) -> None:
    operator = software_operations.HandBrakeOperator(software_operations.OPERATION_PROFILES["handbrake"])
    input_video_path = tmp_path / "input.mp4"
    input_video_path.write_bytes(b"video")
    button = _DummyButton()
    recorded: dict[str, object] = {}

    monkeypatch.setattr(operator, "_find_button", lambda *args, **kwargs: button)
    monkeypatch.setattr(
        software_operations,
        "fill_file_dialog",
        lambda path, **kwargs: recorded.update({"path": path, **kwargs}),
    )
    monkeypatch.setattr(operator, "_wait_for_source_loaded", lambda path: recorded.setdefault("loaded", path))
    monkeypatch.setattr(software_operations.time, "sleep", lambda _seconds: None)

    operator._open_input_clip(_DummyWindow("HandBrake"), input_video_path)

    assert button.clicked is True
    assert recorded["path"] == input_video_path.resolve(strict=False)
    assert recorded["dialog_patterns"] == software_operations.OPEN_DIALOG_PATTERNS
    assert recorded["confirm_patterns"] == software_operations.HAND_BRAKE_DIALOG_OPEN_PATTERNS
    assert recorded["must_exist"] is True
    assert recorded["loaded"] == input_video_path.resolve(strict=False)


def test_handbrake_wait_for_export_completion_accepts_existing_window(monkeypatch, tmp_path: Path) -> None:
    operator = software_operations.HandBrakeOperator(software_operations.OPERATION_PROFILES["handbrake"])
    output_video_path = tmp_path / "out.mp4"
    output_video_path.write_bytes(b"x" * 2048)
    window = _DummyWindow("HandBrake")

    monkeypatch.setattr(operator, "_connect_main_window", lambda timeout=2.0: window)
    monkeypatch.setattr(operator, "_status_line", lambda current_window, patterns: "")
    monkeypatch.setattr(software_operations.time, "sleep", lambda _seconds: None)

    operator._wait_for_export_completion(output_video_path, window)


def test_kdenlive_set_render_output_path_uses_absolute_path(monkeypatch, tmp_path: Path) -> None:
    monkeypatch.chdir(tmp_path)
    operator = software_operations.KdenliveOperator(software_operations.OPERATION_PROFILES["kdenlive"])
    pasted: dict[str, str] = {}

    monkeypatch.setattr(operator, "_focus_output_file_entry", lambda render_dialog: None)
    monkeypatch.setattr(
        software_operations,
        "paste_text_via_clipboard",
        lambda text, replace_existing=True: pasted.setdefault("text", text),
    )
    monkeypatch.setattr(
        software_operations,
        "wait_for_text_control",
        lambda *args, **kwargs: _VisibleControl(),
    )

    output_video_path = Path("results/pipeline_runs/demo/exports/out.mp4")
    operator._set_render_output_path(object(), output_video_path)

    assert pasted["text"] == str(output_video_path.resolve(strict=False))
    assert Path(pasted["text"]).is_absolute()


def test_kdenlive_wait_for_render_completion_accepts_finished_status(monkeypatch, tmp_path: Path) -> None:
    operator = software_operations.KdenliveOperator(software_operations.OPERATION_PROFILES["kdenlive"])
    output_video_path = tmp_path / "finished.mp4"
    output_video_path.write_bytes(b"x" * (software_operations.KDENLIVE_MIN_OUTPUT_BYTES + 1))

    monkeypatch.setattr(operator, "_show_job_queue_tab", lambda render_dialog: None)
    monkeypatch.setattr(operator, "_minimize_window_during_wait", lambda render_dialog, phase="render": None)
    monkeypatch.setattr(operator, "_render_still_active", lambda render_dialog: (False, False, False, False, None))
    monkeypatch.setattr(operator, "_render_marked_finished", lambda render_dialog: True)

    operator._wait_for_render_completion(object(), output_video_path)


def test_kdenlive_wait_for_render_completion_accepts_stable_output_fallback(monkeypatch, tmp_path: Path) -> None:
    operator = software_operations.KdenliveOperator(software_operations.OPERATION_PROFILES["kdenlive"])
    output_video_path = tmp_path / "finished.mp4"
    output_video_path.write_bytes(b"x" * (software_operations.KDENLIVE_MIN_OUTPUT_BYTES + 1))

    monkeypatch.setattr(operator, "_show_job_queue_tab", lambda render_dialog: None)
    monkeypatch.setattr(operator, "_minimize_window_during_wait", lambda render_dialog, phase="render": None)
    monkeypatch.setattr(operator, "_render_still_active", lambda render_dialog: (False, False, False, False, None))
    monkeypatch.setattr(operator, "_render_marked_finished", lambda render_dialog: False)

    operator._wait_for_render_completion(object(), output_video_path)


def test_kdenlive_close_dismisses_save_prompt_after_render_window_close(monkeypatch) -> None:
    operator = software_operations.KdenliveOperator(software_operations.OPERATION_PROFILES["kdenlive"])
    operator._main_window = _DummyWindow("Kdenlive")
    operator._render_dialog = _DummyWindow("Rendering - Kdenlive")
    main_window = operator._main_window
    dismiss_calls: list[tuple[str, float, object]] = []
    request_close_calls: list[object] = []

    monkeypatch.setattr(software_operations, "bring_window_to_front", lambda *args, **kwargs: None)
    monkeypatch.setattr(software_operations, "find_text_control", lambda *args, **kwargs: _DummyButton())
    monkeypatch.setattr(software_operations.time, "sleep", lambda _seconds: None)
    monkeypatch.setattr(software_operations, "request_window_close", lambda window: request_close_calls.append(window))
    monkeypatch.setattr(
        software_operations,
        "dismiss_close_prompts",
        lambda timeout=2.0, owner_window=None: dismiss_calls.append(("generic", timeout, owner_window)),
    )
    monkeypatch.setattr(operator, "_dismiss_profile_switch_prompt_if_present", lambda window: False)
    monkeypatch.setattr(
        operator,
        "_dismiss_kdenlive_save_dialog_if_present",
        lambda owner_window, timeout=3.0: dismiss_calls.append(("targeted", timeout, owner_window)) or True,
    )
    monkeypatch.setattr(operator, "_wait_for_main_window_to_close", lambda timeout=3.0: True)

    operator.close()

    assert dismiss_calls
    assert dismiss_calls[0][0] == "generic"
    assert dismiss_calls[0][1] == 1.5
    assert any(call[0] == "targeted" and call[1] == 3.0 for call in dismiss_calls)
    assert request_close_calls == [main_window]


def test_kdenlive_dismiss_save_dialog_clicks_dont_save(monkeypatch) -> None:
    operator = software_operations.KdenliveOperator(software_operations.OPERATION_PROFILES["kdenlive"])
    owner_window = _DummyWindow("Kdenlive", process_id=321)
    owner_window.handle = 100
    dialog = _DummyDialog("警告 - Kdenlive", process_id=321, texts=["Save changes to document?"])
    dialog.handle = 200
    button = _DummyButton()

    monkeypatch.setattr(operator, "_iter_process_top_level_windows", lambda process_id: [dialog])
    monkeypatch.setattr(operator, "_dialog_matches", lambda *args, **kwargs: True)
    monkeypatch.setattr(
        operator,
        "_first_matching_control",
        lambda root, patterns, control_types=(): button if root is dialog else None,
    )
    monkeypatch.setattr(software_operations, "bring_window_to_front", lambda *args, **kwargs: None)
    monkeypatch.setattr(software_operations.time, "sleep", lambda _seconds: None)

    assert operator._dismiss_kdenlive_save_dialog_if_present(owner_window, timeout=0.2) is True
    assert button.clicked is True
