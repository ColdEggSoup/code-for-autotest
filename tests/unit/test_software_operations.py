from __future__ import annotations

from pathlib import Path

import software_operations


class _DummyRect:
    def __init__(self, left: int, top: int, right: int, bottom: int) -> None:
        self.left = left
        self.top = top
        self.right = right
        self.bottom = bottom

    def width(self) -> int:
        return self.right - self.left

    def height(self) -> int:
        return self.bottom - self.top


class _DummyWindow:
    def __init__(
        self,
        title: str = "",
        *,
        process_id: int = 100,
        class_name: str = "",
        control_type: str = "",
        rect: _DummyRect | None = None,
    ) -> None:
        self._title = title
        self.handle = 100
        self.element_info = type(
            "ElementInfo",
            (),
            {
                "process_id": process_id,
                "class_name": class_name,
                "control_type": control_type,
            },
        )()
        self._children = []
        self._rect = rect or _DummyRect(0, 0, 400, 40)

    def window_text(self) -> str:
        return self._title

    def close(self) -> None:
        return None

    def children(self):
        return list(self._children)

    def descendants(self):
        descendants = []
        for child in self._children:
            descendants.append(child)
            descendants.extend(child.descendants())
        return descendants

    def rectangle(self) -> _DummyRect:
        return self._rect

    def is_visible(self) -> bool:
        return True


class _DummyButton:
    def __init__(
        self,
        title: str = "",
        *,
        control_type: str = "Button",
        process_id: int = 100,
        rect: _DummyRect | None = None,
        enabled: bool = True,
    ) -> None:
        self._title = title
        self.clicked = False
        self.handle = 101
        self._enabled = enabled
        self.element_info = type(
            "ElementInfo",
            (),
            {
                "process_id": process_id,
                "class_name": "",
                "control_type": control_type,
            },
        )()
        self._rect = rect or _DummyRect(0, 0, 120, 32)

    def click_input(self) -> None:
        self.clicked = True

    def window_text(self) -> str:
        return self._title

    def children(self):
        return []

    def descendants(self):
        return []

    def rectangle(self) -> _DummyRect:
        return self._rect

    def is_enabled(self) -> bool:
        return self._enabled


class _DummyComboBox(_DummyWindow):
    def __init__(
        self,
        selected_text: str = "",
        *,
        process_id: int = 100,
        rect: _DummyRect | None = None,
        direct_select: bool = False,
    ) -> None:
        super().__init__(
            selected_text,
            process_id=process_id,
            control_type="ComboBox",
            rect=rect,
        )
        self._selected_text = selected_text
        self._direct_select = direct_select
        self.clicked = False

    def click_input(self) -> None:
        self.clicked = True

    def selected_text(self) -> str:
        return self._selected_text

    def select(self, value: str) -> None:
        if not self._direct_select:
            raise RuntimeError("direct select unavailable")
        self.set_selected_text(value)

    def set_selected_text(self, value: str) -> None:
        self._selected_text = value
        self._title = value

    def texts(self):
        return [self._selected_text] if self._selected_text else []


class _DummySelectableItem(_DummyButton):
    def __init__(
        self,
        title: str,
        *,
        control_type: str = "ListItem",
        process_id: int = 100,
        rect: _DummyRect | None = None,
        on_click=None,
    ) -> None:
        super().__init__(
            title,
            control_type=control_type,
            process_id=process_id,
            rect=rect,
        )
        self._on_click = on_click

    def click_input(self) -> None:
        super().click_input()
        if self._on_click is not None:
            self._on_click()


class _DummyDialog(_DummyWindow):
    def __init__(self, title: str = "Info", *, process_id: int = 100, texts: list[str] | None = None) -> None:
        super().__init__(title, process_id=process_id)
        self._children = [_DummyWindow(text, process_id=process_id) for text in (texts or [])]


class _VisibleControl:
    def is_visible(self) -> bool:
        return True


def test_shotcut_find_child_by_class_name_falls_back_to_descendants() -> None:
    operator = software_operations.ShotcutOperator(software_operations.OPERATION_PROFILES["shotcut"])
    toolbar = _DummyWindow("toolbar", class_name="QToolBar")
    container = _DummyWindow("container", class_name="QWidget")
    encode_dock = _DummyWindow("encode", class_name="EncodeDock")
    container._children = [encode_dock]
    window = _DummyWindow("Shotcut", class_name="MainWindow")
    window._children = [toolbar, container]

    resolved = operator._find_child_by_class_name(window, "EncodeDock")

    assert resolved is encode_dock


def test_shotcut_find_child_by_class_name_prefers_direct_child() -> None:
    operator = software_operations.ShotcutOperator(software_operations.OPERATION_PROFILES["shotcut"])
    direct_jobs_dock = _DummyWindow("jobs-direct", class_name="JobsDock")
    nested_jobs_dock = _DummyWindow("jobs-nested", class_name="JobsDock")
    container = _DummyWindow("container", class_name="QWidget")
    container._children = [nested_jobs_dock]
    window = _DummyWindow("Shotcut", class_name="MainWindow")
    window._children = [direct_jobs_dock, container]

    resolved = operator._find_child_by_class_name(window, "JobsDock")

    assert resolved is direct_jobs_dock


def test_shotcut_resolve_jobs_root_falls_back_to_main_window(monkeypatch) -> None:
    operator = software_operations.ShotcutOperator(software_operations.OPERATION_PROFILES["shotcut"])
    window = _DummyWindow("Shotcut", class_name="MainWindow")

    monkeypatch.setattr(operator, "_find_child_by_class_name_if_present", lambda current_window, class_name: None)

    resolved = operator._resolve_jobs_root(window)

    assert resolved is window


def test_shotcut_resolve_jobs_root_prefers_jobs_dock(monkeypatch) -> None:
    operator = software_operations.ShotcutOperator(software_operations.OPERATION_PROFILES["shotcut"])
    window = _DummyWindow("Shotcut", class_name="MainWindow")
    jobs_dock = _DummyWindow("jobs", class_name="JobsDock")

    monkeypatch.setattr(operator, "_find_child_by_class_name_if_present", lambda current_window, class_name: jobs_dock)

    resolved = operator._resolve_jobs_root(window)

    assert resolved is jobs_dock


def test_shotcut_assert_task_queued_with_recovery_reuses_current_window(monkeypatch, tmp_path: Path) -> None:
    operator = software_operations.ShotcutOperator(software_operations.OPERATION_PROFILES["shotcut"])
    window = _DummyWindow("Shotcut", class_name="MainWindow")
    toolbar = _DummyWindow("toolbar", class_name="QToolBar")
    output_video_path = tmp_path / "out.mkv"
    jobs_root = _DummyWindow("jobs", class_name="JobsDock")
    calls: list[tuple[object, object, Path]] = []

    def _assert_task_queued(current_window, current_toolbar, current_output):
        calls.append((current_window, current_toolbar, current_output))
        return jobs_root

    monkeypatch.setattr(operator, "_assert_task_queued", _assert_task_queued)
    monkeypatch.setattr(operator, "_connect_main_window", lambda timeout=0.0: (_ for _ in ()).throw(AssertionError("unexpected reconnect")))
    monkeypatch.setattr(software_operations, "bring_window_to_front", lambda *args, **kwargs: None)

    resolved_window, resolved_jobs_root = operator._assert_task_queued_with_recovery(window, toolbar, output_video_path)

    assert resolved_window is window
    assert resolved_jobs_root is jobs_root
    assert calls == [(window, toolbar, output_video_path)]


def test_shotcut_assert_task_queued_with_recovery_reconnects_after_failure(monkeypatch, tmp_path: Path) -> None:
    operator = software_operations.ShotcutOperator(software_operations.OPERATION_PROFILES["shotcut"])
    stale_window = _DummyWindow("stale-shotcut", class_name="MainWindow")
    stale_toolbar = _DummyWindow("stale-toolbar", class_name="QToolBar")
    fresh_window = _DummyWindow("fresh-shotcut", class_name="MainWindow")
    fresh_toolbar = _DummyWindow("fresh-toolbar", class_name="QToolBar")
    output_video_path = tmp_path / "out.mkv"
    jobs_root = _DummyWindow("jobs", class_name="JobsDock")
    calls: list[tuple[object, object, Path]] = []
    brought_to_front: list[object] = []

    def _assert_task_queued(current_window, current_toolbar, current_output):
        calls.append((current_window, current_toolbar, current_output))
        if current_window is stale_window:
            raise software_operations.UiAutomationError("stale window handle")
        return jobs_root

    monkeypatch.setattr(operator, "_assert_task_queued", _assert_task_queued)
    monkeypatch.setattr(operator, "_connect_main_window", lambda timeout=0.0: fresh_window)
    monkeypatch.setattr(operator, "_find_child_by_class_name", lambda current_window, class_name: fresh_toolbar)
    monkeypatch.setattr(software_operations, "bring_window_to_front", lambda current_window, **kwargs: brought_to_front.append(current_window))

    resolved_window, resolved_jobs_root = operator._assert_task_queued_with_recovery(
        stale_window,
        stale_toolbar,
        output_video_path,
    )

    assert resolved_window is fresh_window
    assert resolved_jobs_root is jobs_root
    assert calls == [
        (stale_window, stale_toolbar, output_video_path),
        (fresh_window, fresh_toolbar, output_video_path),
    ]
    assert brought_to_front == [fresh_window]


def test_shotcut_perform_prepares_timeline_before_export(monkeypatch, tmp_path: Path) -> None:
    operator = software_operations.ShotcutOperator(software_operations.OPERATION_PROFILES["shotcut"])
    input_video_path = tmp_path / "input.mp4"
    input_video_path.write_bytes(b"video")
    output_video_path = tmp_path / "out.mp4"
    window = _DummyWindow("Shotcut", class_name="MainWindow")
    toolbar = _DummyWindow("toolbar", class_name="QToolBar")
    encode_dock = _DummyWindow("encode", class_name="EncodeDock")
    steps: list[tuple[str, object]] = []

    monkeypatch.setattr(operator, "_connect_main_window", lambda timeout=30.0: window)
    monkeypatch.setattr(operator, "_dismiss_recovery_dialog_if_present", lambda current_window, timeout=2.0: False)
    monkeypatch.setattr(operator, "_dismiss_save_changes_dialog_if_present", lambda current_window, timeout=0.8: False)
    monkeypatch.setattr(software_operations, "bring_window_to_front", lambda *args, **kwargs: None)
    monkeypatch.setattr(
        operator,
        "_find_child_by_class_name",
        lambda current_window, class_name: toolbar if class_name == "QToolBar" else encode_dock,
    )
    monkeypatch.setattr(operator, "_open_input_dialog", lambda current_window, current_toolbar: steps.append(("open_input_dialog", current_toolbar)))
    monkeypatch.setattr(
        software_operations,
        "fill_file_dialog",
        lambda path, **kwargs: steps.append(("fill_file_dialog", path)),
    )
    monkeypatch.setattr(operator, "_append_selected_clip_to_timeline", lambda current_window: steps.append(("append_timeline", current_window)))
    monkeypatch.setattr(operator, "_wait_for_export_ready", lambda current_encode_dock: steps.append(("wait_export_ready", current_encode_dock)))
    monkeypatch.setattr(operator, "_show_output_pane", lambda current_toolbar, current_encode_dock: steps.append(("show_output_pane", (current_toolbar, current_encode_dock))))
    monkeypatch.setattr(operator, "_export_clip", lambda current_encode_dock, path: steps.append(("export_clip", path)))
    monkeypatch.setattr(
        operator,
        "_begin_background_wait",
        lambda current_window, phase="", start_waiter=None, start_description=None: (
            steps.append(("minimize", phase)),
            start_waiter() if start_waiter is not None else None,
        )[-1],
    )
    monkeypatch.setattr(operator, "_wait_for_export_completion", lambda path, jobs_dock=None: steps.append(("wait_complete", (path, jobs_dock))))
    monkeypatch.setattr(operator, "close", lambda: steps.append(("close", None)))

    operator.perform(input_video_path, output_video_path)

    assert [name for name, _ in steps] == [
        "open_input_dialog",
        "fill_file_dialog",
        "append_timeline",
        "wait_export_ready",
        "show_output_pane",
        "export_clip",
        "minimize",
        "wait_complete",
        "close",
    ]


def test_shotcut_open_input_clip_reconnects_after_dialog_selection(monkeypatch, tmp_path: Path) -> None:
    operator = software_operations.ShotcutOperator(software_operations.OPERATION_PROFILES["shotcut"])
    input_video_path = tmp_path / "input.mp4"
    input_video_path.write_bytes(b"video")
    original_window = _DummyWindow("Shotcut", class_name="MainWindow")
    reopened_window = _DummyWindow("Shotcut", class_name="MainWindow")
    toolbar = _DummyWindow("toolbar", class_name="QToolBar")
    steps: list[tuple[str, object]] = []

    monkeypatch.setattr(software_operations, "bring_window_to_front", lambda *args, **kwargs: None)
    monkeypatch.setattr(operator, "_dismiss_save_changes_dialog_if_present", lambda current_window, timeout=0.8: False)
    monkeypatch.setattr(
        operator,
        "_find_child_by_class_name",
        lambda current_window, class_name: steps.append(("find_child", class_name)) or toolbar,
    )
    monkeypatch.setattr(operator, "_open_input_dialog", lambda current_window, current_toolbar: steps.append(("open_dialog", current_toolbar)))
    monkeypatch.setattr(
        software_operations,
        "fill_file_dialog",
        lambda path, **kwargs: steps.append(("fill_dialog", path)),
    )
    monkeypatch.setattr(
        operator,
        "_connect_main_window",
        lambda timeout=8.0: steps.append(("reconnect", timeout)) or reopened_window,
    )

    resolved_window = operator._open_input_clip(original_window, input_video_path)

    assert resolved_window is reopened_window
    assert [name for name, _ in steps] == [
        "find_child",
        "open_dialog",
        "fill_dialog",
        "reconnect",
    ]


def test_shotcut_open_input_clip_dismisses_save_changes_before_reusing_session(monkeypatch, tmp_path: Path) -> None:
    operator = software_operations.ShotcutOperator(software_operations.OPERATION_PROFILES["shotcut"])
    input_video_path = tmp_path / "input.mp4"
    input_video_path.write_bytes(b"video")
    original_window = _DummyWindow("Shotcut", class_name="MainWindow")
    reopened_window = _DummyWindow("Shotcut", class_name="MainWindow")
    toolbar = _DummyWindow("toolbar", class_name="QToolBar")
    steps: list[tuple[str, object]] = []

    monkeypatch.setattr(software_operations, "bring_window_to_front", lambda *args, **kwargs: None)
    monkeypatch.setattr(
        operator,
        "_dismiss_save_changes_dialog_if_present",
        lambda current_window, timeout=0.8: steps.append(("dismiss_save_changes", timeout)) or True,
    )
    monkeypatch.setattr(
        operator,
        "_find_child_by_class_name",
        lambda current_window, class_name: steps.append(("find_child", class_name)) or toolbar,
    )
    monkeypatch.setattr(operator, "_open_input_dialog", lambda current_window, current_toolbar: steps.append(("open_dialog", current_toolbar)))
    monkeypatch.setattr(
        software_operations,
        "fill_file_dialog",
        lambda path, **kwargs: steps.append(("fill_dialog", path)),
    )
    monkeypatch.setattr(
        operator,
        "_connect_main_window",
        lambda timeout=8.0: steps.append(("reconnect", timeout)) or reopened_window,
    )

    resolved_window = operator._open_input_clip(original_window, input_video_path)

    assert resolved_window is reopened_window
    assert [name for name, _ in steps] == [
        "dismiss_save_changes",
        "reconnect",
        "find_child",
        "open_dialog",
        "fill_dialog",
        "reconnect",
    ]


def test_shotcut_export_current_timeline_stays_minimized_while_waiting(monkeypatch, tmp_path: Path) -> None:
    operator = software_operations.ShotcutOperator(software_operations.OPERATION_PROFILES["shotcut"])
    output_video_path = tmp_path / "out.mp4"
    output_video_path.write_bytes(b"encoded")
    initial_window = _DummyWindow("Shotcut", class_name="MainWindow")
    toolbar = _DummyWindow("toolbar", class_name="QToolBar")
    encode_dock = _DummyWindow("encode", class_name="EncodeDock")
    steps: list[tuple[str, object]] = []

    def _find_child(current_window, class_name):
        steps.append(("find_child", (current_window, class_name)))
        if class_name == "QToolBar":
            return toolbar
        if class_name == "EncodeDock":
            return encode_dock
        raise AssertionError(class_name)

    monkeypatch.setattr(software_operations, "bring_window_to_front", lambda *args, **kwargs: None)
    monkeypatch.setattr(operator, "_find_child_by_class_name", _find_child)
    monkeypatch.setattr(operator, "_wait_for_export_ready", lambda current_encode_dock: steps.append(("wait_ready", current_encode_dock)))
    monkeypatch.setattr(
        operator,
        "_show_output_pane",
        lambda current_toolbar, current_encode_dock: steps.append(("show_output", (current_toolbar, current_encode_dock))),
    )
    monkeypatch.setattr(operator, "_export_clip", lambda current_encode_dock, path: steps.append(("export_clip", path)))
    monkeypatch.setattr(
        operator,
        "_begin_background_wait",
        lambda current_window, phase="", start_waiter=None, start_description=None: (
            steps.append(("background_wait", phase)),
            start_waiter() if start_waiter is not None else None,
        )[-1],
    )
    monkeypatch.setattr(
        operator,
        "_wait_for_export_completion",
        lambda path, jobs_dock=None: steps.append(("wait_complete", (path, jobs_dock))),
    )

    resolved_window = operator._export_current_timeline(initial_window, output_video_path)

    assert resolved_window is initial_window
    assert [name for name, _ in steps] == [
        "find_child",
        "find_child",
        "wait_ready",
        "show_output",
        "export_clip",
        "background_wait",
        "wait_complete",
    ]


def test_shotcut_dismiss_recovery_dialog_clicks_no(monkeypatch) -> None:
    operator = software_operations.ShotcutOperator(software_operations.OPERATION_PROFILES["shotcut"])
    owner_window = _DummyWindow("Shotcut", process_id=321)
    dialog = _DummyDialog("Shotcut", process_id=321, texts=["存在自动保存的文件。您想恢复它们吗？"])
    no_button = _DummyButton("否(N)", process_id=321)
    dialog._children.append(no_button)

    dialog_snapshots = [[dialog], []]
    monkeypatch.setattr(
        operator,
        "_iter_process_top_level_windows",
        lambda process_id: dialog_snapshots.pop(0) if dialog_snapshots else [],
    )
    monkeypatch.setattr(software_operations, "bring_window_to_front", lambda *args, **kwargs: None)
    monkeypatch.setattr(software_operations.time, "sleep", lambda _seconds: None)

    assert operator._dismiss_recovery_dialog_if_present(owner_window, timeout=0.2) is True
    assert no_button.clicked is True


def test_shotcut_open_input_dialog_retries_after_recovery_prompt(monkeypatch) -> None:
    operator = software_operations.ShotcutOperator(software_operations.OPERATION_PROFILES["shotcut"])
    main_window = _DummyWindow("Shotcut", class_name="MainWindow")
    toolbar = _DummyWindow("toolbar", class_name="QToolBar")
    window = _DummyWindow("Shotcut", class_name="MainWindow")
    open_button = _DummyButton("Open File")
    open_button.is_visible = lambda: True
    steps: list[tuple[str, object]] = []
    attempts = {"count": 0}

    monkeypatch.setattr(
        software_operations,
        "wait_for_text_control",
        lambda *args, **kwargs: open_button,
    )
    monkeypatch.setattr(software_operations.time, "sleep", lambda _seconds: None)
    monkeypatch.setattr(software_operations, "bring_window_to_front", lambda *args, **kwargs: None)
    hotkeys: list[str] = []
    monkeypatch.setattr(software_operations, "send_hotkey", lambda keys: hotkeys.append(keys))

    def _wait_for_file_dialog(**kwargs):
        attempts["count"] += 1
        if attempts["count"] <= 2:
            raise software_operations.UiAutomationError("dialog blocked by recovery prompt")
        return object()

    monkeypatch.setattr(operator, "_wait_for_open_file_dialog", lambda timeout=5.0: _wait_for_file_dialog())
    monkeypatch.setattr(operator, "_dismiss_save_changes_dialog_if_present", lambda owner_window, timeout=0.4: False)
    monkeypatch.setattr(
        operator,
        "_dismiss_recovery_dialog_if_present",
        lambda owner_window, timeout=0.4: steps.append(("dismiss_recovery", timeout)) or (attempts["count"] == 1),
    )
    monkeypatch.setattr(operator, "_connect_main_window", lambda timeout=8.0: steps.append(("reconnect", timeout)) or window)
    monkeypatch.setattr(
        operator,
        "_find_child_by_class_name",
        lambda current_window, class_name: steps.append(("find_child", class_name)) or toolbar,
    )

    operator._open_input_dialog(main_window, toolbar)

    assert attempts["count"] == 3
    assert open_button.clicked is True
    assert hotkeys == ["^o"]
    assert [name for name, _ in steps] == [
        "dismiss_recovery",
        "dismiss_recovery",
        "dismiss_recovery",
        "reconnect",
        "find_child",
        "dismiss_recovery",
    ]


def test_shotcut_close_dismisses_save_prompt_with_no(monkeypatch) -> None:
    operator = software_operations.ShotcutOperator(software_operations.OPERATION_PROFILES["shotcut"])
    main_window = _DummyWindow("Shotcut", process_id=321)
    main_window.handle = 100
    dialog = _DummyDialog("Shotcut", process_id=321, texts=["项目已被修改", "你想保存你的修改吗？"])
    dialog.handle = 200
    no_button = _DummyButton("否(N)", process_id=321)
    dialog._children.append(no_button)
    generic_dismiss_calls: list[tuple[float, object]] = []
    hotkeys: list[str] = []
    close_requests: list[object] = []

    monkeypatch.setattr(operator, "_connect_main_window", lambda timeout=2.0: main_window)
    dialog_snapshots = [[dialog], []]
    monkeypatch.setattr(
        operator,
        "_iter_process_top_level_windows",
        lambda process_id: dialog_snapshots.pop(0) if dialog_snapshots else [],
    )
    monkeypatch.setattr(software_operations, "bring_window_to_front", lambda *args, **kwargs: None)
    monkeypatch.setattr(software_operations.time, "sleep", lambda _seconds: None)
    monkeypatch.setattr(software_operations, "send_hotkey", lambda keys: hotkeys.append(keys))
    monkeypatch.setattr(software_operations, "request_window_close", lambda window: close_requests.append(window) or True)
    monkeypatch.setattr(
        software_operations,
        "dismiss_close_prompts",
        lambda timeout=2.0, owner_window=None: generic_dismiss_calls.append((timeout, owner_window)),
    )

    operator.close()

    assert no_button.clicked is True
    assert hotkeys == [software_operations.SHOTCUT_CLOSE_HOTKEY]
    assert close_requests == [main_window]
    assert generic_dismiss_calls == [(1.5, main_window)]


def test_shotcut_dismiss_save_prompt_falls_back_to_keyboard_no(monkeypatch) -> None:
    operator = software_operations.ShotcutOperator(software_operations.OPERATION_PROFILES["shotcut"])
    owner_window = _DummyWindow("Shotcut", process_id=321)
    owner_window.handle = 100
    dialog = _DummyDialog("Shotcut", process_id=321, texts=["椤圭洰宸茶淇敼", "浣犳兂淇濆瓨浣犵殑淇敼鍚楋紵"])
    dialog.handle = 200
    hotkeys: list[str] = []
    generic_dismiss_calls: list[tuple[float, object]] = []
    dialog_snapshots = [[dialog], []]
    dialog_state = {"present": True}
    dialog_state = {"present": True}

    monkeypatch.setattr(
        operator,
        "_iter_process_top_level_windows",
        lambda process_id: dialog_snapshots.pop(0) if dialog_snapshots else [],
    )
    monkeypatch.setattr(operator, "_dialog_matches", lambda current_dialog, text_patterns=(): current_dialog is dialog)
    monkeypatch.setattr(software_operations, "bring_window_to_front", lambda *args, **kwargs: None)
    monkeypatch.setattr(software_operations, "find_text_control", lambda *args, **kwargs: (_ for _ in ()).throw(RuntimeError("not exposed")))
    monkeypatch.setattr(software_operations, "send_hotkey", lambda keys: hotkeys.append(keys))
    monkeypatch.setattr(software_operations.time, "sleep", lambda _seconds: None)
    monkeypatch.setattr(
        software_operations,
        "dismiss_close_prompts",
        lambda timeout=0.8, owner_window=None: generic_dismiss_calls.append((timeout, owner_window)) or dialog_state.update({"present": False}),
    )

    assert operator._dismiss_save_changes_dialog_if_present(owner_window, timeout=0.2) is True
    assert hotkeys == list(software_operations.SHOTCUT_DONT_SAVE_DIRECT_KEYS)
    assert generic_dismiss_calls == [(0.8, dialog)]


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
    output_video_path = tmp_path / "out.mkv"
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
    output_video_path = tmp_path / "out.mkv"
    output_video_path.write_bytes(b"video")

    monkeypatch.setattr(operator, "_connect_main_window", lambda timeout=0.5: _DummyWindow("Avidemux"))
    monkeypatch.setattr(operator, "_dismiss_export_success_dialog_if_present", lambda *args, **kwargs: True)
    monkeypatch.setattr(software_operations.time, "sleep", lambda _seconds: None)

    operator._wait_for_export_completion(output_video_path)


def test_avidemux_handle_export_overwrite_prompt_accepts_top_level_dialog(monkeypatch, tmp_path: Path) -> None:
    operator = software_operations.AvidemuxOperator(software_operations.OPERATION_PROFILES["avidemux"])
    input_video_path = tmp_path / "input.mp4"
    input_video_path.write_bytes(b"video")
    output_video_path = tmp_path / "out.mkv"
    calls: list[str] = []

    monkeypatch.setattr(
        software_operations,
        "accept_overwrite_confirmation",
        lambda timeout=0.8, owner_window=None, poll_interval=0.05: calls.append("generic") or True,
    )
    monkeypatch.setattr(
        operator,
        "_handle_embedded_overwrite_prompt",
        lambda *args, **kwargs: calls.append("embedded"),
    )

    operator._handle_export_overwrite_prompt(_DummyWindow("Avidemux"), input_video_path, output_video_path)

    assert calls == ["generic"]


def test_avidemux_handle_embedded_overwrite_prompt_clicks_overwrite_button(monkeypatch, tmp_path: Path) -> None:
    operator = software_operations.AvidemuxOperator(software_operations.OPERATION_PROFILES["avidemux"])
    input_video_path = tmp_path / "input.mp4"
    input_video_path.write_bytes(b"video")
    output_video_path = tmp_path / "out.mkv"
    overwrite_button = _DummyButton("Overwrite(O)")

    monkeypatch.setattr(operator, "_embedded_overwrite_prompt_text", lambda window: f"Overwrite file {output_video_path.name}?")
    monkeypatch.setattr(operator, "_find_bottom_button", lambda window, patterns: overwrite_button)
    monkeypatch.setattr(software_operations.time, "sleep", lambda _seconds: None)

    operator._handle_embedded_overwrite_prompt(_DummyWindow("Avidemux"), input_video_path, output_video_path, timeout=0.2)

    assert overwrite_button.clicked is True


def test_avidemux_minimize_encode_dialog_to_tray_clicks_tray_button(monkeypatch) -> None:
    operator = software_operations.AvidemuxOperator(software_operations.OPERATION_PROFILES["avidemux"])
    window = _DummyWindow("Avidemux", process_id=4321)
    dialog = _DummyWindow("正在编码...", process_id=4321)
    tray_button = _DummyButton("缩到工具栏上", process_id=4321)
    dialog._children.append(tray_button)

    monkeypatch.setattr(operator, "_find_export_progress_dialog", lambda process_id, timeout=2.0: dialog)
    monkeypatch.setattr(software_operations, "bring_window_to_front", lambda *args, **kwargs: None)
    monkeypatch.setattr(software_operations.time, "sleep", lambda _seconds: None)

    assert operator._minimize_encode_dialog_to_tray(window) is True
    assert tray_button.clicked is True
    assert operator._active_process_id == 4321
    assert operator._encode_dialog_minimized_to_tray is True


def test_avidemux_minimize_encode_dialog_to_tray_uses_geometry_fallback(monkeypatch) -> None:
    operator = software_operations.AvidemuxOperator(software_operations.OPERATION_PROFILES["avidemux"])
    window = _DummyWindow("Avidemux", process_id=4321)
    dialog = _DummyWindow("encoding", process_id=4321, rect=_DummyRect(0, 0, 520, 420))
    tray_candidate = _DummyButton("", process_id=4321, rect=_DummyRect(18, 352, 132, 388))
    dialog._children.append(tray_candidate)

    monkeypatch.setattr(operator, "_find_export_progress_dialog", lambda process_id, timeout=2.0: dialog)
    monkeypatch.setattr(software_operations, "bring_window_to_front", lambda *args, **kwargs: None)
    monkeypatch.setattr(software_operations.time, "sleep", lambda _seconds: None)

    assert operator._minimize_encode_dialog_to_tray(window) is True
    assert tray_candidate.clicked is True


def test_avidemux_find_export_progress_dialog_falls_back_to_desktop_popup(monkeypatch) -> None:
    operator = software_operations.AvidemuxOperator(software_operations.OPERATION_PROFILES["avidemux"])
    main_window = _DummyWindow("Avidemux", process_id=4321, rect=_DummyRect(60, 60, 920, 700))
    progress_dialog = _DummyWindow("正在编码...", process_id=9999, rect=_DummyRect(160, 120, 680, 540))
    progress_dialog._children.append(_DummyButton("缩到工具栏上", process_id=9999, rect=_DummyRect(18, 462, 132, 498)))

    monkeypatch.setattr(operator, "_iter_process_top_level_windows", lambda process_id: [main_window])
    monkeypatch.setattr(operator, "_iter_desktop_top_level_windows", lambda: [main_window, progress_dialog])
    monkeypatch.setattr(software_operations.time, "sleep", lambda _seconds: None)

    resolved = operator._find_export_progress_dialog(4321, timeout=0.2)

    assert resolved is progress_dialog


def test_avidemux_save_output_minimizes_to_tray_after_output_starts(monkeypatch, tmp_path: Path) -> None:
    operator = software_operations.AvidemuxOperator(software_operations.OPERATION_PROFILES["avidemux"])
    input_video_path = tmp_path / "input.mp4"
    input_video_path.write_bytes(b"video")
    output_video_path = tmp_path / "out.mkv"
    output_video_path.write_bytes(b"encoded")
    save_button = _DummyButton("Save")

    monkeypatch.setattr(software_operations, "bring_window_to_front", lambda *args, **kwargs: None)
    monkeypatch.setattr(software_operations, "send_hotkey", lambda keys: None)
    monkeypatch.setattr(operator, "_focus_dialog_filename_entry", lambda window: None)
    monkeypatch.setattr(software_operations, "paste_text_via_clipboard", lambda text, replace_existing=True: None)
    monkeypatch.setattr(operator, "_find_bottom_button", lambda window, patterns: save_button)
    monkeypatch.setattr(operator, "_handle_export_overwrite_prompt", lambda *args, **kwargs: None)
    monkeypatch.setattr(operator, "_minimize_encode_dialog_to_tray", lambda window, timeout=5.0: True)
    monkeypatch.setattr(software_operations.time, "sleep", lambda _seconds: None)

    minimized = operator._save_output(_DummyWindow("Avidemux"), input_video_path, output_video_path)

    assert save_button.clicked is True
    assert minimized is True


def test_avidemux_perform_uses_x265_and_skips_generic_minimize_when_tray_minimized(monkeypatch, tmp_path: Path) -> None:
    operator = software_operations.AvidemuxOperator(software_operations.OPERATION_PROFILES["avidemux"])
    input_video_path = tmp_path / "input.mp4"
    input_video_path.write_bytes(b"video")
    output_dir = tmp_path / "exports"
    output_dir.mkdir()
    output_video_path = output_dir / "out.mkv"
    window = _DummyWindow("Avidemux", process_id=4321)
    steps: list[tuple[str, object]] = []

    monkeypatch.setattr(operator, "_connect_main_window", lambda timeout=30.0: window)
    monkeypatch.setattr(software_operations, "bring_window_to_front", lambda *args, **kwargs: None)
    monkeypatch.setattr(operator, "_dismiss_thanks_popup_if_present", lambda *args, **kwargs: False)
    monkeypatch.setattr(operator, "_open_input_clip", lambda current_window, path: steps.append(("open_input", path)))
    monkeypatch.setattr(
        operator,
        "_select_main_combo_value",
        lambda current_window, combo_index, target_text: steps.append(("select_combo", (combo_index, target_text))),
    )
    monkeypatch.setattr(
        operator,
        "_save_output",
        lambda current_window, input_path, output_path: steps.append(("save_output", output_path)) or True,
    )
    monkeypatch.setattr(
        operator,
        "_begin_background_wait",
        lambda current_window, phase="": steps.append(("generic_minimize", phase)),
    )
    monkeypatch.setattr(
        operator,
        "_wait_for_export_completion",
        lambda output_path: steps.append(("wait_complete", output_path)),
    )
    monkeypatch.setattr(operator, "close", lambda: steps.append(("close", None)))
    monkeypatch.setattr(operator, "_delete_generated_output_artifact", lambda output_path, input_path: steps.append(("cleanup", output_path)))

    operator.perform(input_video_path, output_video_path)

    assert ("select_combo", (0, software_operations.AVIDEMUX_VIDEO_CODEC_TEXT)) in steps
    assert ("select_combo", (2, software_operations.AVIDEMUX_MUXER_TEXT)) in steps
    assert ("generic_minimize", "encode") not in steps


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


def test_avidemux_close_restores_tray_minimized_windows_before_closing(monkeypatch) -> None:
    operator = software_operations.AvidemuxOperator(software_operations.OPERATION_PROFILES["avidemux"])
    operator._active_process_id = 4321
    operator._encode_dialog_minimized_to_tray = True
    calls: list[str] = []

    class _CloseTrackingWindow(_DummyWindow):
        def close(self) -> None:
            calls.append("close")

    window = _CloseTrackingWindow("Avidemux", process_id=4321)

    monkeypatch.setattr(operator, "_restore_process_windows", lambda process_id, timeout=1.5: calls.append("restore") or True)
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

    assert calls[:3] == ["restore", "dismiss", "close"]
    assert operator._active_process_id is None
    assert operator._encode_dialog_minimized_to_tray is False


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
    assert recorded["allow_direct_selection"] is False
    assert recorded["loaded"] == input_video_path.resolve(strict=False)


def test_handbrake_select_preset_uses_combo_dropdown_when_needed(monkeypatch) -> None:
    operator = software_operations.HandBrakeOperator(software_operations.OPERATION_PROFILES["handbrake"])
    target_text = software_operations.HAND_BRAKE_TARGET_PRESET_TEXT
    window = _DummyWindow("HandBrake", rect=_DummyRect(0, 0, 1200, 800))
    preset_label = _DummyWindow("Preset:", control_type="Text", rect=_DummyRect(40, 56, 130, 92))
    preset_combo = _DummyComboBox(
        "Fast 1080p30",
        rect=_DummyRect(150, 52, 520, 96),
        direct_select=False,
    )
    dropdown_item = _DummySelectableItem(
        target_text,
        rect=_DummyRect(160, 100, 560, 136),
        on_click=lambda: preset_combo.set_selected_text(target_text),
    )
    window._children = [preset_label, preset_combo]

    monkeypatch.setattr(operator, "_iter_process_wrappers", lambda process_id: [dropdown_item])
    monkeypatch.setattr(software_operations, "bring_window_to_front", lambda *args, **kwargs: None)
    monkeypatch.setattr(software_operations.time, "sleep", lambda _seconds: None)

    operator._select_preset(window, target_text)

    assert preset_combo.clicked is True
    assert dropdown_item.clicked is True
    assert preset_combo.selected_text() == target_text


def test_handbrake_select_preset_falls_back_to_keyboard_entry(monkeypatch) -> None:
    operator = software_operations.HandBrakeOperator(software_operations.OPERATION_PROFILES["handbrake"])
    target_text = software_operations.HAND_BRAKE_TARGET_PRESET_TEXT
    window = _DummyWindow("HandBrake", rect=_DummyRect(0, 0, 1200, 800))
    preset_label = _DummyWindow("Preset:", control_type="Text", rect=_DummyRect(40, 56, 130, 92))
    preset_combo = _DummyComboBox(
        "Fast 1080p30",
        rect=_DummyRect(150, 52, 520, 96),
        direct_select=False,
    )
    window._children = [preset_label, preset_combo]

    monkeypatch.setattr(operator, "_iter_process_wrappers", lambda process_id: [])
    monkeypatch.setattr(operator, "_set_combo_text_with_keyboard", lambda current_window, combo, value: combo.set_selected_text(value))
    monkeypatch.setattr(software_operations, "bring_window_to_front", lambda *args, **kwargs: None)
    monkeypatch.setattr(software_operations.time, "sleep", lambda _seconds: None)

    operator._select_preset(window, target_text)

    assert preset_combo.clicked is True
    assert preset_combo.selected_text() == target_text


def test_handbrake_select_preset_supports_button_style_selector(monkeypatch) -> None:
    operator = software_operations.HandBrakeOperator(software_operations.OPERATION_PROFILES["handbrake"])
    target_text = software_operations.HAND_BRAKE_TARGET_PRESET_TEXT
    window = _DummyWindow("HandBrake", rect=_DummyRect(0, 0, 1200, 800))
    preset_label = _DummyWindow("Preset:", control_type="Text", rect=_DummyRect(10, 90, 60, 126))
    preset_button = _DummyButton(
        "Fast 1080p30",
        control_type="Button",
        rect=_DummyRect(70, 88, 360, 130),
    )
    dropdown_item = _DummySelectableItem(
        target_text,
        rect=_DummyRect(380, 420, 720, 456),
        on_click=lambda: setattr(preset_button, "_title", target_text),
    )
    window._children = [preset_label, preset_button]

    monkeypatch.setattr(operator, "_iter_process_wrappers", lambda process_id: [dropdown_item])
    monkeypatch.setattr(software_operations, "bring_window_to_front", lambda *args, **kwargs: None)
    monkeypatch.setattr(software_operations.time, "sleep", lambda _seconds: None)

    operator._select_preset(window, target_text)

    assert preset_button.clicked is True
    assert dropdown_item.clicked is True
    assert preset_button.window_text() == target_text


def test_handbrake_select_preset_accepts_dropdown_click_when_selector_text_stays_empty(monkeypatch) -> None:
    operator = software_operations.HandBrakeOperator(software_operations.OPERATION_PROFILES["handbrake"])
    target_text = software_operations.HAND_BRAKE_TARGET_PRESET_TEXT
    window = _DummyWindow("HandBrake", rect=_DummyRect(0, 0, 1200, 800))
    preset_label = _DummyWindow("Preset:", control_type="Text", rect=_DummyRect(10, 90, 60, 126))
    preset_button = _DummyButton(
        "",
        control_type="Button",
        rect=_DummyRect(70, 88, 360, 130),
    )
    dropdown_item = _DummySelectableItem(
        target_text,
        rect=_DummyRect(380, 420, 720, 456),
    )
    window._children = [preset_label, preset_button]

    monkeypatch.setattr(operator, "_iter_process_wrappers", lambda process_id: [dropdown_item])
    monkeypatch.setattr(software_operations, "bring_window_to_front", lambda *args, **kwargs: None)
    monkeypatch.setattr(software_operations.time, "sleep", lambda _seconds: None)

    operator._select_preset(window, target_text)

    assert preset_button.clicked is True
    assert dropdown_item.clicked is True


def test_handbrake_find_combo_dropdown_item_handles_multiple_handleless_wrappers(monkeypatch) -> None:
    operator = software_operations.HandBrakeOperator(software_operations.OPERATION_PROFILES["handbrake"])
    target_text = software_operations.HAND_BRAKE_TARGET_PRESET_TEXT
    window = _DummyWindow("HandBrake", process_id=4321, rect=_DummyRect(0, 0, 1200, 800))
    combo = _DummyButton("Fast 1080p30", control_type="Button", process_id=4321, rect=_DummyRect(70, 88, 360, 130))
    first_item = _DummySelectableItem("HQ 2160p60 4K AV1 Surround", control_type="Text", process_id=4321, rect=_DummyRect(380, 420, 720, 456))
    target_item = _DummySelectableItem(target_text, control_type="Text", process_id=4321, rect=_DummyRect(380, 456, 720, 492))
    first_item.handle = None
    target_item.handle = None
    monkeypatch.setattr(operator, "_iter_process_wrappers", lambda process_id: [first_item, target_item])

    resolved = operator._find_combo_dropdown_item(window, combo, target_text)

    assert resolved is target_item


def test_handbrake_preset_combo_fallback_prefers_preset_like_value_over_title_selector() -> None:
    operator = software_operations.HandBrakeOperator(software_operations.OPERATION_PROFILES["handbrake"])
    window = _DummyWindow("HandBrake", rect=_DummyRect(0, 0, 1200, 800))
    title_combo = _DummyComboBox(
        "1  (00:03:21)",
        rect=_DummyRect(120, 40, 360, 82),
        direct_select=False,
    )
    preset_combo = _DummyComboBox(
        "Fast 1080p30",
        rect=_DummyRect(150, 90, 520, 132),
        direct_select=False,
    )
    window._children = [title_combo, preset_combo]

    resolved = operator._preset_combo(window)

    assert resolved is preset_combo


def test_handbrake_dismiss_update_dialog_clicks_no(monkeypatch) -> None:
    operator = software_operations.HandBrakeOperator(software_operations.OPERATION_PROFILES["handbrake"])
    dialog = _DummyWindow("检查更新?")
    no_button = _DummyButton("否(N)")

    monkeypatch.setattr(software_operations, "wait_for_window", lambda patterns, timeout=0.5: dialog)
    monkeypatch.setattr(software_operations, "bring_window_to_front", lambda *args, **kwargs: None)
    monkeypatch.setattr(software_operations, "find_text_control", lambda *args, **kwargs: no_button)
    monkeypatch.setattr(software_operations, "click_control", lambda control, post_click_sleep=0.4: control.click_input())
    monkeypatch.setattr(software_operations.time, "sleep", lambda _seconds: None)

    assert operator._dismiss_update_dialog_if_present(timeout=0.2) is True
    assert no_button.clicked is True


def test_handbrake_source_loaded_accepts_auto_renamed_destination(monkeypatch, tmp_path: Path) -> None:
    operator = software_operations.HandBrakeOperator(software_operations.OPERATION_PROFILES["handbrake"])
    input_video_path = tmp_path / "4K_big.mp4"
    input_video_path.write_bytes(b"video")

    monkeypatch.setattr(
        operator,
        "_read_save_path",
        lambda window: r"C:\Users\tester\Videos\4K_big (1).m4v",
    )

    assert operator._source_loaded(_DummyWindow("HandBrake"), input_video_path) is True


def test_handbrake_source_loaded_accepts_ready_state_with_destination(monkeypatch, tmp_path: Path) -> None:
    operator = software_operations.HandBrakeOperator(software_operations.OPERATION_PROFILES["handbrake"])
    input_video_path = tmp_path / "4K_big.mp4"
    input_video_path.write_bytes(b"video")

    monkeypatch.setattr(
        operator,
        "_read_save_path",
        lambda window: r"C:\Users\tester\Videos\custom_destination.mp4",
    )
    monkeypatch.setattr(
        operator,
        "_status_line",
        lambda window, patterns: "Ready" if patterns == software_operations.HAND_BRAKE_READY_STATUS_PATTERNS else "",
    )
    monkeypatch.setattr(operator, "_start_encode_button_ready", lambda window: True)

    assert operator._source_loaded(_DummyWindow("HandBrake"), input_video_path) is True


def test_handbrake_source_loaded_accepts_ready_state_with_visible_source_name(monkeypatch, tmp_path: Path) -> None:
    operator = software_operations.HandBrakeOperator(software_operations.OPERATION_PROFILES["handbrake"])
    input_video_path = tmp_path / "4K_big.mp4"
    input_video_path.write_bytes(b"video")
    source_item = _DummyWindow("4K_big.mp4", control_type="Text")
    window = _DummyWindow("HandBrake")
    window._children.append(source_item)

    monkeypatch.setattr(operator, "_read_save_path", lambda current_window: "")
    monkeypatch.setattr(
        operator,
        "_status_line",
        lambda current_window, patterns: "Ready" if patterns == software_operations.HAND_BRAKE_READY_STATUS_PATTERNS else "",
    )
    monkeypatch.setattr(operator, "_start_encode_button_ready", lambda current_window: False)

    assert operator._source_loaded(window, input_video_path) is True


def test_handbrake_save_path_edit_prefers_field_left_of_browse_button() -> None:
    operator = software_operations.HandBrakeOperator(software_operations.OPERATION_PROFILES["handbrake"])
    window = _DummyWindow("HandBrake", rect=_DummyRect(0, 0, 1200, 800))
    destination_edit = _DummyWindow(
        r"C:\Users\tester\Videos\out.mp4",
        control_type="Edit",
        rect=_DummyRect(120, 680, 930, 720),
    )
    unrelated_edit = _DummyWindow(
        "Preset",
        control_type="Edit",
        rect=_DummyRect(100, 120, 360, 155),
    )
    browse_button = _DummyButton(
        "Browse",
        rect=_DummyRect(950, 678, 1090, 722),
    )
    window._children = [unrelated_edit, destination_edit, browse_button]

    resolved = operator._save_path_edit(window)

    assert resolved is destination_edit


def test_handbrake_start_encode_button_ready_finds_enabled_button_without_old_coords() -> None:
    operator = software_operations.HandBrakeOperator(software_operations.OPERATION_PROFILES["handbrake"])
    window = _DummyWindow("HandBrake", rect=_DummyRect(0, 0, 1200, 800))
    start_button = _DummyButton(
        "Start Encode",
        rect=_DummyRect(360, 88, 520, 132),
        enabled=True,
    )
    window._children = [start_button]

    assert operator._start_encode_button_ready(window) is True


def test_handbrake_wait_for_export_completion_accepts_existing_window(monkeypatch, tmp_path: Path) -> None:
    operator = software_operations.HandBrakeOperator(software_operations.OPERATION_PROFILES["handbrake"])
    output_video_path = tmp_path / "out.mp4"
    output_video_path.write_bytes(b"x" * 2048)
    window = _DummyWindow("HandBrake")

    monkeypatch.setattr(operator, "_connect_main_window", lambda timeout=2.0: window)
    monkeypatch.setattr(operator, "_status_line", lambda current_window, patterns: "")
    monkeypatch.setattr(software_operations.time, "sleep", lambda _seconds: None)

    operator._wait_for_export_completion(output_video_path, window)


def test_handbrake_perform_selects_requested_preset_before_encoding(monkeypatch, tmp_path: Path) -> None:
    operator = software_operations.HandBrakeOperator(software_operations.OPERATION_PROFILES["handbrake"])
    input_video_path = tmp_path / "4K_big.mp4"
    input_video_path.write_bytes(b"video")
    output_dir = tmp_path / "exports"
    output_dir.mkdir()
    output_video_path = output_dir / "handbrake_out.mp4"
    window = _DummyWindow("HandBrake")
    steps: list[tuple[str, object]] = []

    def _unexpected_x264(_window) -> None:
        raise AssertionError("_ensure_x264_encoder should not run once preset selection is enabled.")

    monkeypatch.setattr(operator, "_connect_main_window", lambda timeout=15.0: window)
    monkeypatch.setattr(software_operations, "bring_window_to_front", lambda *args, **kwargs: None)
    monkeypatch.setattr(
        operator,
        "_open_input_clip",
        lambda current_window, path: steps.append(("open_input", path)),
    )
    monkeypatch.setattr(
        operator,
        "_select_preset",
        lambda current_window, target_text: steps.append(("select_preset", target_text)),
    )
    monkeypatch.setattr(operator, "_ensure_x264_encoder", _unexpected_x264)
    monkeypatch.setattr(
        operator,
        "_set_output_path",
        lambda current_window, path: steps.append(("set_output", path)),
    )
    monkeypatch.setattr(
        operator,
        "_start_encode",
        lambda current_window, path: steps.append(("start_encode", path)),
    )
    monkeypatch.setattr(
        operator,
        "_begin_background_wait",
        lambda current_window, phase="", start_waiter=None, start_description=None: steps.append(("minimize", phase)),
    )
    monkeypatch.setattr(
        operator,
        "_wait_for_export_completion",
        lambda path, current_window=None: steps.append(("wait_complete", path)),
    )
    monkeypatch.setattr(operator, "close", lambda: steps.append(("close", None)))
    monkeypatch.setattr(
        operator,
        "_delete_generated_output_artifact",
        lambda output_path, input_path: steps.append(("cleanup_output", output_path)),
    )

    operator.perform(input_video_path, output_video_path)

    assert [name for name, _ in steps] == [
        "open_input",
        "select_preset",
        "set_output",
        "start_encode",
        "minimize",
        "wait_complete",
        "close",
        "cleanup_output",
    ]
    assert steps[1] == ("select_preset", software_operations.HAND_BRAKE_TARGET_PRESET_TEXT)


def test_shutter_encoder_perform_minimizes_before_waiting_for_job_start(monkeypatch, tmp_path: Path) -> None:
    operator = software_operations.ShutterEncoderOperator(software_operations.OPERATION_PROFILES["shutter_encoder"])
    input_video_path = tmp_path / "source.mp4"
    input_video_path.write_bytes(b"video")
    output_dir = tmp_path / "exports"
    output_dir.mkdir()
    output_video_path = output_dir / "out.mp4"
    window = _DummyWindow("Shutter Encoder")
    known_outputs = {tmp_path / "generated.mp4": (0.0, 0)}
    generated_output = tmp_path / "generated.mp4"
    generated_output.write_bytes(b"encoded")
    steps: list[tuple[str, object]] = []

    monkeypatch.setattr(operator, "_snapshot_output_candidates", lambda path: known_outputs)
    monkeypatch.setattr(operator, "_connect_main_window", lambda timeout=2.0: window)
    monkeypatch.setattr(software_operations, "bring_window_to_front", lambda *args, **kwargs: None)
    monkeypatch.setattr(operator, "_dismiss_update_dialog_if_present", lambda timeout=2.0: False)
    monkeypatch.setattr(operator, "_open_input_dialog", lambda current_window: steps.append(("open_input_dialog", current_window)))
    monkeypatch.setattr(
        software_operations,
        "fill_file_dialog",
        lambda path, **kwargs: steps.append(("fill_file_dialog", path)),
    )
    monkeypatch.setattr(software_operations.time, "sleep", lambda _seconds: None)
    monkeypatch.setattr(operator, "_wait_for_imported_source", lambda current_window, path: steps.append(("wait_import", path)))
    monkeypatch.setattr(operator, "_select_function", lambda current_window, text: steps.append(("select_function", text)))
    monkeypatch.setattr(operator, "_start_function", lambda current_window: steps.append(("start_function", current_window)))

    def _begin_background_wait(current_window, phase="", start_waiter=None, start_description=None):
        steps.append(("minimize", phase))
        if start_waiter is not None:
            start_waiter()

    monkeypatch.setattr(operator, "_begin_background_wait", _begin_background_wait)
    monkeypatch.setattr(operator, "_wait_for_job_to_start", lambda path, outputs: steps.append(("wait_job_start", path)))
    monkeypatch.setattr(operator, "_wait_for_completion", lambda path, outputs: steps.append(("wait_complete", path)) or generated_output)
    monkeypatch.setattr(operator, "close", lambda: steps.append(("close", None)))
    monkeypatch.setattr(software_operations.shutil, "copy2", lambda src, dst: dst.write_bytes(src.read_bytes()))
    monkeypatch.setattr(
        operator,
        "_delete_generated_output_artifact",
        lambda source_path, input_path, requested_path: steps.append(("cleanup_output", source_path)),
    )

    operator.perform(input_video_path, output_video_path)

    assert [name for name, _ in steps] == [
        "open_input_dialog",
        "fill_file_dialog",
        "wait_import",
        "select_function",
        "start_function",
        "minimize",
        "wait_job_start",
        "wait_complete",
        "close",
        "cleanup_output",
    ]
    assert steps[3] == ("select_function", software_operations.SHUTTER_ENCODER_TARGET_FUNCTION_TEXT)


def test_shutter_encoder_close_closes_remaining_process_windows_after_main_window_disappears(monkeypatch) -> None:
    operator = software_operations.ShutterEncoderOperator(software_operations.OPERATION_PROFILES["shutter_encoder"])
    close_calls: list[str] = []
    terminated: list[list[str]] = []

    class _CloseTrackingWindow(_DummyWindow):
        def __init__(self, title: str, *, process_id: int) -> None:
            super().__init__(title, process_id=process_id)

        def close(self) -> None:
            close_calls.append(self.window_text())
            if self in visible_windows:
                visible_windows.remove(self)

    main_window = _CloseTrackingWindow("Shutter Encoder", process_id=4321)
    qr_window = _CloseTrackingWindow("Support / QR", process_id=4321)
    visible_windows = [main_window, qr_window]

    monkeypatch.setattr(operator, "_connect_main_window", lambda timeout=2.0: main_window)
    monkeypatch.setattr(operator, "_iter_process_top_level_windows", lambda process_id: list(visible_windows))
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

    monkeypatch.setattr(software_operations.time, "sleep", _sleep)

    operator.close()

    assert close_calls[:2] == ["Shutter Encoder", "Support / QR"]
    assert terminated == []


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

    monkeypatch.setattr(operator, "_render_still_active", lambda render_dialog: (False, False, False, False, None))
    monkeypatch.setattr(operator, "_render_marked_finished", lambda render_dialog: True)

    operator._wait_for_render_completion(object(), output_video_path)


def test_kdenlive_wait_for_render_completion_accepts_stable_output_fallback(monkeypatch, tmp_path: Path) -> None:
    operator = software_operations.KdenliveOperator(software_operations.OPERATION_PROFILES["kdenlive"])
    output_video_path = tmp_path / "finished.mp4"
    output_video_path.write_bytes(b"x" * (software_operations.KDENLIVE_MIN_OUTPUT_BYTES + 1))

    monkeypatch.setattr(operator, "_render_still_active", lambda render_dialog: (False, False, False, False, None))
    monkeypatch.setattr(operator, "_render_marked_finished", lambda render_dialog: False)

    operator._wait_for_render_completion(object(), output_video_path)


def test_kdenlive_perform_minimizes_immediately_after_render_start(monkeypatch, tmp_path: Path) -> None:
    operator = software_operations.KdenliveOperator(software_operations.OPERATION_PROFILES["kdenlive"])
    input_video_path = tmp_path / "source.mp4"
    input_video_path.write_bytes(b"video")
    output_dir = tmp_path / "exports"
    output_dir.mkdir()
    output_video_path = output_dir / "out.mp4"
    window = _DummyWindow("Kdenlive")
    render_dialog = _DummyWindow("Rendering - Kdenlive")
    clip_item = _DummyButton("clip", control_type="Text")
    steps: list[tuple[str, object]] = []

    monkeypatch.setattr(operator, "_connect_main_window", lambda timeout=30.0: window)
    monkeypatch.setattr(operator, "_dismiss_recovery_dialog_if_present", lambda current_window, timeout=3.0: False)
    monkeypatch.setattr(software_operations, "bring_window_to_front", lambda *args, **kwargs: None)
    monkeypatch.setattr(operator, "_open_input_clip", lambda current_window, path: steps.append(("open_input_clip", path)) or window)
    monkeypatch.setattr(operator, "_insert_clip_to_timeline", lambda current_window: steps.append(("insert_timeline", current_window)))
    monkeypatch.setattr(operator, "_render_current_timeline", lambda current_window, path: steps.append(("render_current_timeline", path)) or window)
    monkeypatch.setattr(operator, "close", lambda: steps.append(("close", None)))

    operator.perform(input_video_path, output_video_path)

    assert [name for name, _ in steps] == [
        "open_input_clip",
        "insert_timeline",
        "render_current_timeline",
        "close",
    ]


def test_kdenlive_open_input_clip_imports_and_selects_requested_video(monkeypatch, tmp_path: Path) -> None:
    operator = software_operations.KdenliveOperator(software_operations.OPERATION_PROFILES["kdenlive"])
    input_video_path = tmp_path / "source.mp4"
    input_video_path.write_bytes(b"video")
    current_window = _DummyWindow("Kdenlive")
    reconnected_window = _DummyWindow("Kdenlive")
    clip_item = _DummyButton("clip", control_type="Text")
    steps: list[tuple[str, object]] = []

    monkeypatch.setattr(operator, "_dismiss_kdenlive_save_dialog_if_present", lambda owner_window, timeout=1.0: False)
    monkeypatch.setattr(operator, "_open_import_dialog", lambda window: steps.append(("open_import_dialog", window)))
    monkeypatch.setattr(
        software_operations,
        "fill_file_dialog",
        lambda path, **kwargs: steps.append(("fill_file_dialog", path)),
    )
    monkeypatch.setattr(software_operations.time, "sleep", lambda _seconds: None)
    monkeypatch.setattr(operator, "_connect_main_window", lambda timeout=6.0: steps.append(("reconnect", timeout)) or reconnected_window)
    monkeypatch.setattr(software_operations, "bring_window_to_front", lambda *args, **kwargs: None)
    monkeypatch.setattr(operator, "_wait_for_imported_clip", lambda window, path: steps.append(("wait_clip", path)) or clip_item)
    monkeypatch.setattr(software_operations, "click_control", lambda control, post_click_sleep=0.0: steps.append(("click_clip", control)))

    resolved_window = operator._open_input_clip(current_window, input_video_path)

    assert resolved_window is reconnected_window
    assert operator._main_window is reconnected_window
    assert [name for name, _ in steps] == [
        "open_import_dialog",
        "fill_file_dialog",
        "reconnect",
        "wait_clip",
        "click_clip",
    ]


def test_kdenlive_open_input_clip_dismisses_save_dialog_before_next_import(monkeypatch, tmp_path: Path) -> None:
    operator = software_operations.KdenliveOperator(software_operations.OPERATION_PROFILES["kdenlive"])
    input_video_path = tmp_path / "source.mp4"
    input_video_path.write_bytes(b"video")
    current_window = _DummyWindow("Kdenlive")
    reconnected_window = _DummyWindow("Kdenlive")
    clip_item = _DummyButton("clip", control_type="Text")
    steps: list[tuple[str, object]] = []

    monkeypatch.setattr(
        operator,
        "_dismiss_kdenlive_save_dialog_if_present",
        lambda owner_window, timeout=1.0: steps.append(("dismiss_save", timeout)) or True,
    )
    monkeypatch.setattr(operator, "_open_import_dialog", lambda window: steps.append(("open_import_dialog", window)))
    monkeypatch.setattr(
        software_operations,
        "fill_file_dialog",
        lambda path, **kwargs: steps.append(("fill_file_dialog", path)),
    )
    monkeypatch.setattr(software_operations.time, "sleep", lambda _seconds: None)
    monkeypatch.setattr(operator, "_connect_main_window", lambda timeout=6.0: steps.append(("reconnect", timeout)) or reconnected_window)
    monkeypatch.setattr(software_operations, "bring_window_to_front", lambda *args, **kwargs: None)
    monkeypatch.setattr(operator, "_wait_for_imported_clip", lambda window, path: steps.append(("wait_clip", path)) or clip_item)
    monkeypatch.setattr(software_operations, "click_control", lambda control, post_click_sleep=0.0: steps.append(("click_clip", control)))

    resolved_window = operator._open_input_clip(current_window, input_video_path)

    assert resolved_window is reconnected_window
    assert operator._main_window is reconnected_window
    assert [name for name, _ in steps] == [
        "dismiss_save",
        "reconnect",
        "open_import_dialog",
        "fill_file_dialog",
        "reconnect",
        "wait_clip",
        "click_clip",
    ]


def test_kdenlive_render_current_timeline_minimizes_and_waits(monkeypatch, tmp_path: Path) -> None:
    operator = software_operations.KdenliveOperator(software_operations.OPERATION_PROFILES["kdenlive"])
    output_video_path = tmp_path / "out.mp4"
    render_dialog = _DummyWindow("Rendering - Kdenlive")
    current_window = _DummyWindow("Kdenlive")
    steps: list[tuple[str, object]] = []

    monkeypatch.setattr(operator, "_open_render_dialog", lambda window: steps.append(("open_render_dialog", window)) or render_dialog)
    monkeypatch.setattr(operator, "_set_render_output_path", lambda dialog, path: steps.append(("set_render_output", path)))
    monkeypatch.setattr(operator, "_show_job_queue_tab", lambda dialog: steps.append(("show_job_queue", dialog)))
    monkeypatch.setattr(operator, "_start_render", lambda dialog: steps.append(("start_render", dialog)))
    monkeypatch.setattr(
        operator,
        "_confirm_render_transition_after_minimizing",
        lambda dialog: steps.append(("confirm_transition", dialog)),
    )
    monkeypatch.setattr(
        operator,
        "_begin_background_wait",
        lambda dialog, phase="", start_waiter=None, start_description=None: (
            steps.append(("minimize", phase)),
            start_waiter() if start_waiter is not None else None,
        )[-1],
    )
    monkeypatch.setattr(operator, "_wait_for_render_completion", lambda dialog, path: steps.append(("wait_complete", path)))

    resolved_window = operator._render_current_timeline(current_window, output_video_path)

    assert resolved_window is current_window
    assert operator._main_window is current_window
    assert [name for name, _ in steps] == [
        "open_render_dialog",
        "set_render_output",
        "show_job_queue",
        "start_render",
        "minimize",
        "confirm_transition",
        "wait_complete",
    ]


def test_kdenlive_close_render_dialog_only_preserves_main_window(monkeypatch) -> None:
    operator = software_operations.KdenliveOperator(software_operations.OPERATION_PROFILES["kdenlive"])
    main_window = _DummyWindow("Kdenlive")
    render_dialog = _DummyWindow("Rendering - Kdenlive")
    operator._main_window = main_window
    operator._render_dialog = render_dialog
    steps: list[tuple[str, object]] = []

    monkeypatch.setattr(software_operations, "bring_window_to_front", lambda *args, **kwargs: None)
    monkeypatch.setattr(software_operations, "find_text_control", lambda *args, **kwargs: _DummyButton())
    monkeypatch.setattr(software_operations, "click_control", lambda control, post_click_sleep=0.5: steps.append(("click_close", control)))
    monkeypatch.setattr(
        operator,
        "_dismiss_kdenlive_save_dialog_if_present",
        lambda owner_window, timeout=2.0: steps.append(("dismiss_save", timeout, owner_window)) or True,
    )
    monkeypatch.setattr(
        software_operations,
        "dismiss_close_prompts",
        lambda timeout=2.0, owner_window=None: steps.append(("dismiss_prompts", owner_window)),
    )
    monkeypatch.setattr(operator, "_connect_main_window", lambda timeout=6.0: steps.append(("reconnect", timeout)) or main_window)

    operator._close_render_dialog_only()

    assert operator._render_dialog is None
    assert operator._main_window is main_window
    assert [item[0] for item in steps] == [
        "click_close",
        "dismiss_save",
        "dismiss_prompts",
        "dismiss_save",
        "reconnect",
    ]


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
    monkeypatch.setattr(operator, "_connect_main_window", lambda timeout=6.0: main_window)
    monkeypatch.setattr(operator, "_wait_for_main_window_to_close", lambda timeout=3.0: True)

    operator.close()

    assert dismiss_calls
    assert dismiss_calls[0][0] == "targeted"
    assert dismiss_calls[0][1] == 2.0
    assert any(call[0] == "generic" and call[1] == 1.5 for call in dismiss_calls)
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


def test_kdenlive_dismiss_save_dialog_clicks_no_alias(monkeypatch) -> None:
    operator = software_operations.KdenliveOperator(software_operations.OPERATION_PROFILES["kdenlive"])
    owner_window = _DummyWindow("Kdenlive", process_id=321)
    owner_window.handle = 100
    dialog = _DummyDialog("Warning - Kdenlive", process_id=321, texts=["Save changes to document?"])
    dialog.handle = 200
    no_button = _DummyButton("No(N)", process_id=321)
    dialog._children.append(no_button)

    monkeypatch.setattr(operator, "_iter_process_top_level_windows", lambda process_id: [dialog])
    monkeypatch.setattr(operator, "_dialog_matches", lambda *args, **kwargs: True)
    monkeypatch.setattr(software_operations, "bring_window_to_front", lambda *args, **kwargs: None)
    monkeypatch.setattr(software_operations.time, "sleep", lambda _seconds: None)

    assert operator._dismiss_kdenlive_save_dialog_if_present(owner_window, timeout=0.2) is True
    assert no_button.clicked is True


def test_kdenlive_dismiss_save_dialog_finds_nested_warning_dialog(monkeypatch) -> None:
    operator = software_operations.KdenliveOperator(software_operations.OPERATION_PROFILES["kdenlive"])
    owner_window = _DummyWindow("Kdenlive", process_id=321)
    owner_window.handle = 100
    render_dialog = _DummyDialog("Rendering - Kdenlive", process_id=321)
    render_dialog.handle = 150
    nested_warning = _DummyDialog("Warning - Kdenlive", process_id=321, texts=["Save changes to document?"])
    nested_warning.handle = 200
    discard_button = _DummyButton("Don't Save", process_id=321)
    nested_warning._children.append(discard_button)
    render_dialog._children.append(nested_warning)

    monkeypatch.setattr(operator, "_iter_process_top_level_windows", lambda process_id: [render_dialog])
    monkeypatch.setattr(software_operations, "bring_window_to_front", lambda *args, **kwargs: None)
    monkeypatch.setattr(software_operations.time, "sleep", lambda _seconds: None)

    assert operator._dismiss_kdenlive_save_dialog_if_present(owner_window, timeout=0.2) is True
    assert discard_button.clicked is True


def test_kdenlive_dismiss_save_dialog_falls_back_to_keyboard_shortcuts(monkeypatch) -> None:
    operator = software_operations.KdenliveOperator(software_operations.OPERATION_PROFILES["kdenlive"])
    owner_window = _DummyWindow("Kdenlive", process_id=321)
    owner_window.handle = 100
    dialog = _DummyDialog("Warning - Kdenlive", process_id=321, texts=["Save changes to document?"])
    dialog.handle = 200
    hotkeys: list[str] = []
    generic_dismiss_calls: list[tuple[float, object]] = []
    dialog_snapshots = [[dialog], []]
    dialog_state = {"present": True}

    monkeypatch.setattr(
        operator,
        "_iter_process_top_level_windows",
        lambda process_id: dialog_snapshots.pop(0) if dialog_snapshots else [],
    )
    monkeypatch.setattr(software_operations, "bring_window_to_front", lambda *args, **kwargs: None)
    monkeypatch.setattr(software_operations, "send_hotkey", lambda keys: hotkeys.append(keys))
    monkeypatch.setattr(software_operations.time, "sleep", lambda _seconds: None)
    monkeypatch.setattr(
        software_operations,
        "dismiss_close_prompts",
        lambda timeout=0.8, owner_window=None: generic_dismiss_calls.append((timeout, owner_window)) or dialog_state.update({"present": False}),
    )
    monkeypatch.setattr(
        operator,
        "_first_matching_control",
        lambda root, patterns, control_types=(): None,
    )
    monkeypatch.setattr(
        operator,
        "_control_present",
        lambda root, patterns, control_types=(): root is dialog and dialog_state["present"],
    )

    assert operator._dismiss_kdenlive_save_dialog_if_present(owner_window, timeout=0.2) is True
    assert hotkeys == list(software_operations.KDENLIVE_DONT_SAVE_DIRECT_KEYS)
    assert generic_dismiss_calls == [(0.8, dialog)]


def test_kdenlive_dismiss_recovery_dialog_clicks_do_not_recover(monkeypatch) -> None:
    operator = software_operations.KdenliveOperator(software_operations.OPERATION_PROFILES["kdenlive"])
    owner_window = _DummyWindow("Kdenlive", process_id=321)
    owner_window.handle = 100
    dialog = _DummyDialog("File Recovery - Kdenlive", process_id=321, texts=["Auto-saved file exists. Do you want to recover now?"])
    dialog.handle = 200
    do_not_recover_button = _DummyButton("Do not recover", process_id=321)
    dialog._children.append(do_not_recover_button)

    monkeypatch.setattr(operator, "_iter_process_top_level_windows", lambda process_id: [dialog])
    monkeypatch.setattr(operator, "_dialog_matches", lambda *args, **kwargs: True)
    monkeypatch.setattr(software_operations, "bring_window_to_front", lambda *args, **kwargs: None)
    monkeypatch.setattr(software_operations.time, "sleep", lambda _seconds: None)

    assert operator._dismiss_recovery_dialog_if_present(owner_window, timeout=0.2) is True
    assert do_not_recover_button.clicked is True


def test_kdenlive_accept_overwrite_dialog_clicks_overwrite(monkeypatch) -> None:
    operator = software_operations.KdenliveOperator(software_operations.OPERATION_PROFILES["kdenlive"])
    owner_window = _DummyWindow("Rendering - Kdenlive", process_id=321)
    owner_window.handle = 100
    dialog = _DummyDialog("Warning - Kdenlive", process_id=321, texts=["Output file already exists. Do you want to overwrite it?"])
    dialog.handle = 200
    overwrite_button = _DummyButton("Overwrite(O)", process_id=321)
    dialog._children.append(overwrite_button)
    owner_window._children.append(dialog)

    monkeypatch.setattr(operator, "_iter_process_top_level_windows", lambda process_id: [owner_window])
    monkeypatch.setattr(operator, "_control_present", lambda root, patterns, control_types=(): root is dialog)
    monkeypatch.setattr(software_operations, "bring_window_to_front", lambda *args, **kwargs: None)
    monkeypatch.setattr(software_operations.time, "sleep", lambda _seconds: None)

    assert operator._accept_kdenlive_overwrite_dialog_if_present(owner_window, timeout=0.2) is True
    assert overwrite_button.clicked is True


def test_kdenlive_open_render_dialog_falls_back_to_window_search(monkeypatch) -> None:
    operator = software_operations.KdenliveOperator(software_operations.OPERATION_PROFILES["kdenlive"])
    window = _DummyWindow("Kdenlive", process_id=321)
    container = _DummyWindow("render-container", process_id=321)
    operator._main_window = window

    monkeypatch.setattr(software_operations, "bring_window_to_front", lambda *args, **kwargs: None)
    monkeypatch.setattr(software_operations, "send_hotkey", lambda keys: None)
    monkeypatch.setattr(software_operations.time, "sleep", lambda _seconds: None)
    monkeypatch.setattr(
        software_operations,
        "wait_for_desktop_text_control",
        lambda *args, **kwargs: (_ for _ in ()).throw(software_operations.UiAutomationError("not found")),
    )
    monkeypatch.setattr(operator, "_iter_process_top_level_windows", lambda process_id: [window])
    monkeypatch.setattr(
        operator,
        "_locate_render_container",
        lambda root, timeout: container if root is window else (_ for _ in ()).throw(AssertionError("unexpected root")),
    )
    monkeypatch.setattr(operator, "_resolve_render_dialog_top_level", lambda render_container: None)

    resolved = operator._open_render_dialog(window)

    assert resolved is container


def test_kdenlive_open_import_dialog_uses_extended_confirm_patterns(monkeypatch) -> None:
    operator = software_operations.KdenliveOperator(software_operations.OPERATION_PROFILES["kdenlive"])
    button = _DummyButton("Add Clip or Folder")
    captured: dict[str, object] = {}

    monkeypatch.setattr(operator, "_dismiss_recovery_dialog_if_present", lambda owner_window, timeout=0.8: False)
    monkeypatch.setattr(software_operations, "wait_for_text_control", lambda *args, **kwargs: button)
    monkeypatch.setattr(software_operations.time, "sleep", lambda _seconds: None)
    monkeypatch.setattr(
        software_operations,
        "wait_for_file_dialog",
        lambda **kwargs: captured.update(kwargs) or object(),
    )

    operator._open_import_dialog(_DummyWindow("Kdenlive"))

    assert button.clicked is True
    assert captured["confirm_patterns"] == software_operations.KDENLIVE_OPEN_CONFIRM_PATTERNS


def test_kdenlive_start_render_checks_kdenlive_specific_overwrite_dialog_when_generic_helper_misses(monkeypatch) -> None:
    operator = software_operations.KdenliveOperator(software_operations.OPERATION_PROFILES["kdenlive"])
    render_button = _DummyButton("Render to File")
    render_dialog = _DummyWindow("Rendering - Kdenlive", process_id=321)
    helper_calls: list[tuple[object, float]] = []

    monkeypatch.setattr(software_operations, "bring_window_to_front", lambda *args, **kwargs: None)
    monkeypatch.setattr(operator, "_find_render_to_file_button", lambda current_dialog: render_button)
    monkeypatch.setattr(software_operations, "accept_overwrite_confirmation", lambda *args, **kwargs: False)
    monkeypatch.setattr(
        operator,
        "_accept_kdenlive_overwrite_dialog_if_present",
        lambda owner_window, timeout=2.5: helper_calls.append((owner_window, timeout)) or True,
    )

    operator._start_render(render_dialog)

    assert render_button.clicked is True
    assert helper_calls == [(render_dialog, 2.5)]
