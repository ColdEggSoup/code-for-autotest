from __future__ import annotations

import ui_automation


class _DummyRect:
    def __init__(self, *, top: int, left: int, right: int | None = None, bottom: int | None = None) -> None:
        self.top = top
        self.left = left
        self.right = right if right is not None else left + 120
        self.bottom = bottom if bottom is not None else top + 30


class _DummyElementInfo:
    def __init__(
        self,
        process_id: int,
        *,
        handle: int | None = None,
        control_type: str = "",
        class_name: str = "",
    ) -> None:
        self.process_id = process_id
        self.handle = handle
        self.control_type = control_type
        self.class_name = class_name


class _DummyWindow:
    def __init__(
        self,
        title: str,
        process_id: int,
        *,
        top: int,
        left: int,
        right: int | None = None,
        bottom: int | None = None,
        handle: int | None = None,
        control_type: str = "",
        class_name: str = "",
    ) -> None:
        self._title = title
        self._rect = _DummyRect(top=top, left=left, right=right, bottom=bottom)
        self.element_info = _DummyElementInfo(
            process_id,
            handle=handle,
            control_type=control_type,
            class_name=class_name,
        )
        self.handle = handle
        self._children: list[_DummyWindow] = []

    def window_text(self) -> str:
        return self._title

    def rectangle(self) -> _DummyRect:
        return self._rect

    def children(self) -> list["_DummyWindow"]:
        return list(self._children)

    def descendants(self) -> list["_DummyWindow"]:
        descendants: list[_DummyWindow] = []
        for child in self._children:
            descendants.append(child)
            descendants.extend(child.descendants())
        return descendants


class _DummyDesktop:
    def __init__(self, windows: list[_DummyWindow]) -> None:
        self._windows = windows

    def __call__(self, backend: str = "uia") -> "_DummyDesktop":
        assert backend == "uia"
        return self

    def windows(self) -> list[_DummyWindow]:
        return list(self._windows)


class _DummyMethodWindow:
    def __init__(self, *, process_id: int, handle: int) -> None:
        self._process_id = process_id
        self._handle = handle

    def process_id(self) -> int:
        return self._process_id

    def handle(self) -> int:
        return self._handle


class _DummyClickableControl:
    def __init__(self, *, title: str = "Control", control_type: str = "Button") -> None:
        self._title = title
        self._control_type = control_type
        self.invoke_calls = 0
        self.click_input_calls = 0

        self.element_info = type(
            "ElementInfo",
            (),
            {
                "control_type": control_type,
                "class_name": "",
            },
        )()

    def window_text(self) -> str:
        return self._title

    def invoke(self) -> None:
        self.invoke_calls += 1

    def click_input(self) -> None:
        self.click_input_calls += 1


class _DummyInputOnlyControl(_DummyClickableControl):
    def invoke(self) -> None:
        raise NotImplementedError("invoke unavailable")


class _DummyDesktopErrorControl(_DummyClickableControl):
    def invoke(self) -> None:
        raise NotImplementedError("invoke unavailable")

    def click_input(self) -> None:
        self.click_input_calls += 1
        raise RuntimeError("There is no active desktop required for moving mouse cursor!\n")


def test_dismiss_close_prompts_only_handles_owner_process(monkeypatch) -> None:
    unrelated_dialog = _DummyWindow("Warning", 1001, top=10, left=10)
    owner_window = _DummyWindow("Shotcut", 2002, top=20, left=20)
    owner_dialog = _DummyWindow("Question", 2002, top=30, left=30)
    desktop = _DummyDesktop([unrelated_dialog, owner_dialog])
    clicked: list[_DummyWindow] = []

    monkeypatch.setattr(ui_automation, "_import_pywinauto", lambda: (None, desktop, None))

    def _click_text_control(dialog, *_args, **_kwargs) -> None:
        clicked.append(dialog)
        desktop._windows = [window for window in desktop._windows if window is not dialog]

    monkeypatch.setattr(ui_automation, "click_text_control", _click_text_control)

    ui_automation.dismiss_close_prompts(timeout=0.2, owner_window=owner_window)

    assert clicked == [owner_dialog]


def test_dismiss_close_prompts_handles_chinese_warning_dialog(monkeypatch) -> None:
    owner_window = _DummyWindow("Kdenlive", 2002, top=20, left=20)
    owner_dialog = _DummyWindow("警告 - Kdenlive", 2002, top=30, left=30)
    desktop = _DummyDesktop([owner_dialog])
    clicked: list[_DummyWindow] = []

    monkeypatch.setattr(ui_automation, "_import_pywinauto", lambda: (None, desktop, None))
    monkeypatch.setattr(ui_automation, "_wait_for_window_to_close", lambda *_args, **_kwargs: True)

    def _click_text_control(dialog, patterns, *_args, **_kwargs) -> None:
        assert any("不保存" in pattern for pattern in patterns)
        clicked.append(dialog)
        desktop._windows = [window for window in desktop._windows if window is not dialog]

    monkeypatch.setattr(ui_automation, "click_text_control", _click_text_control)

    ui_automation.dismiss_close_prompts(timeout=0.2, owner_window=owner_window)

    assert clicked == [owner_dialog]


def test_numeric_window_metadata_accepts_bound_methods() -> None:
    window = _DummyMethodWindow(process_id=2002, handle=5566)

    assert ui_automation._get_process_id(window) == 2002
    assert ui_automation._get_window_handle(window) == 5566


def test_is_file_dialog_candidate_rejects_editor_window_with_matching_buttons() -> None:
    editor_window = _DummyWindow(
        "test_software_operations.py - code-for-autotest - Visual Studio Code",
        2002,
        top=20,
        left=20,
    )
    editor_window._children.extend(
        [
            _DummyWindow("", 2002, top=80, left=120, control_type="Edit"),
            _DummyWindow("打开更改", 2002, top=80, left=320, control_type="Button"),
        ]
    )

    assert ui_automation._is_file_dialog_candidate(
        editor_window,
        ui_automation.DEFAULT_CONFIRM_PATTERNS,
        ui_automation.DEFAULT_DIALOG_PATTERNS,
    ) is False


def test_is_file_dialog_candidate_accepts_standard_open_dialog() -> None:
    dialog = _DummyWindow("打开", 2002, top=20, left=20, class_name="#32770")
    dialog._children.extend(
        [
            _DummyWindow("文件名(N):", 2002, top=80, left=40, control_type="Text"),
            _DummyWindow("", 2002, top=80, left=180, control_type="Edit"),
            _DummyWindow("打开(O)", 2002, top=80, left=420, control_type="Button"),
        ]
    )

    assert ui_automation._is_file_dialog_candidate(
        dialog,
        ui_automation.DEFAULT_CONFIRM_PATTERNS,
        ui_automation.DEFAULT_DIALOG_PATTERNS,
    ) is True


def test_wait_for_file_dialog_ignores_foreground_editor_window(monkeypatch) -> None:
    editor_window = _DummyWindow(
        "test_software_operations.py - code-for-autotest - Visual Studio Code",
        2002,
        top=20,
        left=20,
    )
    editor_window._children.extend(
        [
            _DummyWindow("", 2002, top=80, left=120, control_type="Edit"),
            _DummyWindow("打开更改", 2002, top=80, left=320, control_type="Button"),
        ]
    )
    nested_dialog = _DummyWindow("打开", 2002, top=40, left=40, class_name="#32770")

    monkeypatch.setattr(
        ui_automation,
        "wait_for_window",
        lambda patterns, timeout=0.75: (_ for _ in ()).throw(ui_automation.UiAutomationError("not found")),
    )
    monkeypatch.setattr(ui_automation, "get_foreground_window", lambda: editor_window)
    monkeypatch.setattr(ui_automation, "_find_nested_file_dialog", lambda dialog_patterns, confirm_patterns: nested_dialog)

    assert ui_automation.wait_for_file_dialog(timeout=0.2) is nested_dialog


def test_request_window_close_posts_wm_close_to_target_handle(monkeypatch) -> None:
    window = _DummyWindow("Shotcut", 2002, top=20, left=20, handle=5566)
    posted_messages: list[tuple[int, int, int, int]] = []

    class _DummyUser32:
        def PostMessageW(self, handle: int, message: int, wparam: int, lparam: int) -> int:
            posted_messages.append((handle, message, wparam, lparam))
            return 1

    class _DummyWindll:
        user32 = _DummyUser32()

    monkeypatch.setattr(ui_automation.ctypes, "windll", _DummyWindll())
    monkeypatch.setattr(ui_automation, "_is_window_handle_visible", lambda handle: handle == 5566)

    closed = ui_automation.request_window_close(window)

    assert closed is True
    assert posted_messages == [(5566, ui_automation.WM_CLOSE, 0, 0)]


def test_request_window_close_ignores_disappeared_window(monkeypatch) -> None:
    window = _DummyWindow("Shotcut", 2002, top=20, left=20, handle=5566)

    class _DummyUser32:
        def PostMessageW(self, handle: int, message: int, wparam: int, lparam: int) -> int:
            return 0

    class _DummyWindll:
        user32 = _DummyUser32()

    visibility_checks: list[int] = []

    def _is_window_handle_visible(handle: int) -> bool:
        visibility_checks.append(handle)
        return False

    monkeypatch.setattr(ui_automation.ctypes, "windll", _DummyWindll())
    monkeypatch.setattr(ui_automation, "_is_window_handle_visible", _is_window_handle_visible)
    monkeypatch.setattr(ui_automation, "_is_process_running", lambda process_id: False)

    closed = ui_automation.request_window_close(window)

    assert closed is False
    assert visibility_checks == [5566]


def test_click_text_control_uses_click_input(monkeypatch) -> None:
    control = _DummyClickableControl()

    monkeypatch.setattr(ui_automation, "find_text_control", lambda *args, **kwargs: control)
    monkeypatch.setattr(ui_automation.time, "sleep", lambda _seconds: None)

    ui_automation.click_text_control(object(), r"^Control$")

    assert control.invoke_calls == 0
    assert control.click_input_calls == 1


def test_click_control_falls_back_to_click_input(monkeypatch) -> None:
    control = _DummyInputOnlyControl()

    monkeypatch.setattr(ui_automation.time, "sleep", lambda _seconds: None)

    ui_automation.click_control(control)

    assert control.click_input_calls == 1


def test_click_control_reports_inactive_desktop_as_ui_automation_error(monkeypatch) -> None:
    control = _DummyDesktopErrorControl()

    monkeypatch.setattr(ui_automation.time, "sleep", lambda _seconds: None)

    try:
        ui_automation.click_control(control)
    except ui_automation.UiAutomationError as exc:
        assert "interactive desktop" in str(exc)
    else:
        raise AssertionError("Expected click_control to raise UiAutomationError when click_input needs an active desktop.")


def test_find_file_dialog_edit_prefers_lower_input_near_confirm_button() -> None:
    dialog = _DummyWindow("Open", 2002, top=0, left=0, right=1000, bottom=700, control_type="Window")
    location_edit = _DummyWindow("", 2002, top=90, left=180, right=820, bottom=120, control_type="Edit")
    filename_combo = _DummyWindow("", 2002, top=560, left=180, right=820, bottom=595, control_type="ComboBox")
    confirm_button = _DummyWindow("Open", 2002, top=560, left=850, right=940, bottom=595, control_type="Button")
    dialog._children = [location_edit, filename_combo, confirm_button]

    resolved = ui_automation._find_file_dialog_edit(dialog, (r"^Open$",))

    assert resolved is filename_combo


def test_fill_file_dialog_uses_bottom_filename_input_when_label_missing(monkeypatch, tmp_path) -> None:
    file_path = tmp_path / "input.mp4"
    file_path.write_bytes(b"video")
    dialog = _DummyWindow("Open", 2002, top=0, left=0, right=1000, bottom=700, handle=5566, control_type="Window")
    location_edit = _DummyWindow("", 2002, top=90, left=180, right=820, bottom=120, control_type="Edit")
    filename_combo = _DummyWindow("", 2002, top=560, left=180, right=820, bottom=595, control_type="ComboBox")
    confirm_button = _DummyWindow("Open", 2002, top=560, left=850, right=940, bottom=595, control_type="Button")
    dialog._children = [location_edit, filename_combo, confirm_button]
    captured: dict[str, object] = {}

    monkeypatch.setattr(ui_automation, "wait_for_file_dialog", lambda **kwargs: dialog)
    monkeypatch.setattr(ui_automation, "bring_window_to_front", lambda *args, **kwargs: None)
    monkeypatch.setattr(ui_automation, "_dialog_contains_file", lambda *args, **kwargs: False)
    monkeypatch.setattr(
        ui_automation,
        "_set_edit_value",
        lambda edit_wrapper, value: captured.update({"edit": edit_wrapper, "value": value}),
    )
    monkeypatch.setattr(ui_automation, "click_control", lambda control, post_click_sleep=0.4: captured.setdefault("confirm", control))
    monkeypatch.setattr(ui_automation, "_wait_for_window_to_close", lambda *args, **kwargs: True)

    ui_automation.fill_file_dialog(file_path, confirm_patterns=(r"^Open$",), must_exist=True)

    assert captured["edit"] is filename_combo
    assert str(captured["value"]).endswith("input.mp4")
    assert captured["confirm"] is confirm_button


def test_fill_file_dialog_can_skip_waiting_for_dialog_close_after_confirm(monkeypatch, tmp_path) -> None:
    file_path = tmp_path / "output.mp4"
    dialog = _DummyWindow("Save", 2002, top=0, left=0, right=1000, bottom=700, handle=5566, control_type="Window")
    filename_combo = _DummyWindow("", 2002, top=560, left=180, right=820, bottom=595, control_type="ComboBox")
    confirm_button = _DummyWindow("Save", 2002, top=560, left=850, right=940, bottom=595, control_type="Button")
    dialog._children = [filename_combo, confirm_button]
    captured: dict[str, object] = {}

    monkeypatch.setattr(ui_automation, "wait_for_file_dialog", lambda **kwargs: dialog)
    monkeypatch.setattr(ui_automation, "bring_window_to_front", lambda *args, **kwargs: None)
    monkeypatch.setattr(ui_automation, "_dialog_contains_file", lambda *args, **kwargs: False)
    monkeypatch.setattr(
        ui_automation,
        "_set_edit_value",
        lambda edit_wrapper, value: captured.update({"edit": edit_wrapper, "value": value}),
    )
    monkeypatch.setattr(ui_automation, "click_control", lambda control, post_click_sleep=0.4: captured.setdefault("confirm", control))
    monkeypatch.setattr(
        ui_automation,
        "_wait_for_window_to_close",
        lambda *args, **kwargs: (_ for _ in ()).throw(AssertionError("should not wait for dialog close")),
    )
    monkeypatch.setattr(
        ui_automation,
        "accept_overwrite_confirmation",
        lambda timeout=0.0, owner_window=None, poll_interval=0.05: captured.setdefault("overwrite_timeout", timeout) or False,
    )

    ui_automation.fill_file_dialog(
        file_path,
        confirm_patterns=(r"^Save$",),
        must_exist=False,
        wait_for_dialog_to_close=False,
        overwrite_confirmation_timeout=0.5,
    )

    assert captured["edit"] is filename_combo
    assert str(captured["value"]).endswith("output.mp4")
    assert captured["confirm"] is confirm_button
    assert captured["overwrite_timeout"] == 0.5
