from __future__ import annotations

import ui_automation


class _DummyRect:
    def __init__(self, *, top: int, left: int) -> None:
        self.top = top
        self.left = left


class _DummyElementInfo:
    def __init__(self, process_id: int, *, handle: int | None = None) -> None:
        self.process_id = process_id
        self.handle = handle


class _DummyWindow:
    def __init__(self, title: str, process_id: int, *, top: int, left: int, handle: int | None = None) -> None:
        self._title = title
        self._rect = _DummyRect(top=top, left=left)
        self.element_info = _DummyElementInfo(process_id, handle=handle)
        self.handle = handle

    def window_text(self) -> str:
        return self._title

    def rectangle(self) -> _DummyRect:
        return self._rect


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
