from __future__ import annotations

import ctypes
from ctypes import wintypes
from dataclasses import dataclass
import logging
from pathlib import Path
import re
import threading
import time
from typing import Iterable, Sequence


class UiAutomationError(RuntimeError):
    """Raised when the target UI cannot be located or manipulated safely."""


@dataclass(frozen=True)
class UiActionResult:
    target: str
    changed: bool
    state: str


logger = logging.getLogger(__name__)


HWND_TOPMOST = -1
HWND_NOTOPMOST = -2
SWP_NOMOVE = 0x0002
SWP_NOSIZE = 0x0001
SWP_SHOWWINDOW = 0x0040
SW_RESTORE = 9
SW_MINIMIZE = 6
WM_CLOSE = 0x0010
DEFAULT_DIALOG_PATTERNS = (
    r"^Open$",
    r"^Save$",
    r"^Save As$",
    r"^Select",
    r"^Browse",
    r"^\u6253\u5f00$",
    r"^\u6253\u5f00\u6587\u4ef6$",
    r"^\u4fdd\u5b58$",
    r"^\u53e6\u5b58\u4e3a$",
)
DEFAULT_CONFIRM_PATTERNS = (
    r"^Open$",
    r"^Save$",
    r"^OK$",
    r"^Select Folder$",
    r"^\u6253\u5f00",
    r"^\u4fdd\u5b58",
)
OVERWRITE_DIALOG_TITLE_PATTERNS = (
    r"^Confirm Save As$",
    r"^Confirm$",
    r"^Replace$",
    r"^\u786e\u8ba4\u53e6\u5b58\u4e3a$",
    r"^\u786e\u8ba4$",
)
OVERWRITE_DIALOG_TEXT_PATTERNS = (
    r"exists",
    r"replace",
    r"overwrite",
    r"\u5df2\u5b58\u5728",
    r"\u66ff\u6362",
    r"\u8986\u76d6",
)
OVERWRITE_CONFIRM_PATTERNS = (
    r"^Yes(?:\b|[(])",
    r"^Overwrite(?:\b|[(])",
    r"^Replace(?:\b|[(])",
    r"^\u662f(?:$|[(])",
    r"^\u8986\u76d6(?:$|[(])",
)
FILE_NAME_LABEL_PATTERNS = (
    r"^File name",
    r"^Name:?$",
    r"^\u6587\u4ef6\u540d",
    r"^\u540d\u79f0(?:\uff1a|:)?$",
)
FILE_LIST_CONTROL_TYPES = ("DataItem", "ListItem", "TreeItem", "Text")
COMMON_CLOSE_DIALOG_PATTERNS = (
    r"save",
    r"confirm",
    r"warning",
    r"question",
    "\u4fdd\u5b58",
    "\u786e\u8ba4",
    "\u8b66\u544a",
    "\u95ee\u9898",
)
COMMON_DISCARD_PATTERNS = (
    r"^don't save(?:\b|[(])",
    r"^discard(?:\b|[(])",
    r"^no(?:\b|[(])",
    r"^close without saving(?:\b|[(])",
    "^\u4e0d\u4fdd\u5b58(?:$|[(])",
    "^\u5426(?:$|[(])",
)
CF_UNICODETEXT = 13
GMEM_MOVEABLE = 0x0002


def _import_pywinauto():
    try:
        from pywinauto import Application, Desktop, keyboard
    except ImportError as exc:
        raise UiAutomationError(
            "pywinauto is not installed. Run: pip install -r requirements.txt"
        ) from exc
    return Application, Desktop, keyboard


def _compiled_patterns(patterns: str | Sequence[str]) -> tuple[re.Pattern[str], ...]:
    normalized = (patterns,) if isinstance(patterns, str) else tuple(patterns)
    assert normalized, "At least one text pattern is required."
    return tuple(re.compile(pattern, re.IGNORECASE) for pattern in normalized)


def _safe_window_text(wrapper) -> str:
    try:
        return (wrapper.window_text() or "").strip()
    except Exception:
        return ""


def _control_type(wrapper) -> str:
    try:
        return (getattr(wrapper.element_info, "control_type", "") or "").strip()
    except Exception:
        return ""


def _class_name(wrapper) -> str:
    try:
        return (getattr(wrapper.element_info, "class_name", "") or "").strip()
    except Exception:
        return ""


def _wrapper_rect(wrapper):
    try:
        return wrapper.rectangle()
    except Exception as exc:  # pragma: no cover - backend specific
        raise UiAutomationError(f"Failed to read rectangle for UI element: {wrapper}") from exc


def _iter_controls(root) -> Iterable:
    yield root
    for child in root.descendants():
        yield child


def _matches_text(text: str, patterns: Sequence[re.Pattern[str]]) -> bool:
    return any(pattern.search(text) for pattern in patterns)


def _set_window_topmost(handle: int, enabled: bool) -> None:
    user32 = ctypes.windll.user32
    z_order = HWND_TOPMOST if enabled else HWND_NOTOPMOST
    user32.SetWindowPos(handle, z_order, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW)


def bring_window_to_front(window, *, keep_topmost: bool = True, force_foreground: bool = True) -> None:
    handle = window.handle
    user32 = ctypes.windll.user32
    user32.ShowWindow(handle, SW_RESTORE)
    if keep_topmost:
        _set_window_topmost(handle, True)
    if force_foreground:
        user32.SetForegroundWindow(handle)
    window.set_focus()
    time.sleep(0.3)


def minimize_window(window) -> None:
    handle = window.handle
    user32 = ctypes.windll.user32
    _set_window_topmost(handle, False)
    user32.ShowWindow(handle, SW_MINIMIZE)
    time.sleep(0.2)


def get_foreground_window():
    _, Desktop, _ = _import_pywinauto()
    deadline = time.monotonic() + 1.0
    last_error: Exception | None = None
    while time.monotonic() < deadline:
        handle = ctypes.windll.user32.GetForegroundWindow()
        if not handle:
            last_error = UiAutomationError("No foreground window is available.")
            time.sleep(0.05)
            continue
        try:
            return Desktop(backend="uia").window(handle=handle)
        except Exception as exc:
            last_error = exc
            time.sleep(0.05)
    raise UiAutomationError("Could not resolve the current foreground window.") from last_error


def clear_window_topmost(window) -> None:
    _set_window_topmost(window.handle, False)


def connect_window(window_title_re: str, timeout: float = 15.0):
    _, Desktop, _ = _import_pywinauto()
    desktop = Desktop(backend="uia")
    window = desktop.window(title_re=window_title_re)
    try:
        window.wait("visible", timeout=timeout)
        return window
    except Exception as exc:
        raise UiAutomationError(f"Could not connect to a visible window matching '{window_title_re}'.") from exc


def wait_for_window(patterns: str | Sequence[str], timeout: float = 15.0):
    _, Desktop, _ = _import_pywinauto()
    compiled = _compiled_patterns(patterns)
    deadline = time.monotonic() + timeout
    last_seen_titles: list[str] = []
    while time.monotonic() < deadline:
        desktop = Desktop(backend="uia")
        windows = list(desktop.windows())
        last_seen_titles = [_safe_window_text(window) for window in windows if _safe_window_text(window)]
        candidates = []
        for window in windows:
            title = _safe_window_text(window)
            if not title or not _matches_text(title, compiled):
                continue
            rect = _wrapper_rect(window)
            candidates.append((rect.top, rect.left, window))
        if candidates:
            candidates.sort(key=lambda item: (item[0], item[1]))
            return candidates[0][2]
        time.sleep(0.25)
    raise UiAutomationError(
        f"Could not find a top-level window matching {tuple(pattern.pattern for pattern in compiled)}. "
        f"Seen titles: {last_seen_titles[:10]}"
    )


def find_text_control(root, patterns: str | Sequence[str], *, control_types: Sequence[str] = ()):
    compiled = _compiled_patterns(patterns)
    allowed_types = {value.lower() for value in control_types}
    candidates = []
    for wrapper in _iter_controls(root):
        title = _safe_window_text(wrapper)
        if not title or not _matches_text(title, compiled):
            continue
        control_type = _control_type(wrapper).lower()
        if allowed_types and control_type not in allowed_types:
            continue
        rect = _wrapper_rect(wrapper)
        candidates.append((rect.top, rect.left, wrapper))
    if not candidates:
        raise UiAutomationError(
            f"Could not find a control matching {tuple(pattern.pattern for pattern in compiled)} "
            f"under '{_safe_window_text(root)}'."
        )
    candidates.sort(key=lambda item: (item[0], item[1]))
    return candidates[0][2]


def _inactive_desktop_error(exc: Exception) -> bool:
    return "There is no active desktop required for moving mouse cursor!" in str(exc)


def click_control(control, *, post_click_sleep: float = 0.4) -> None:
    try:
        control.click_input()
        time.sleep(post_click_sleep)
        return
    except Exception as exc:
        if _inactive_desktop_error(exc):
            raise UiAutomationError(
                f"Could not activate control '{_safe_window_text(control) or '<untitled>'}' with click_input() "
                "because there is no interactive desktop."
            ) from exc
        raise UiAutomationError(
            f"Could not activate control '{_safe_window_text(control) or '<untitled>'}' with click_input()."
        ) from exc


def click_text_control(root, patterns: str | Sequence[str], *, control_types: Sequence[str] = ()) -> None:
    control = find_text_control(root, patterns, control_types=control_types)
    click_control(control)


def control_exists(root, patterns: str | Sequence[str], *, control_types: Sequence[str] = ()) -> bool:
    try:
        find_text_control(root, patterns, control_types=control_types)
        return True
    except UiAutomationError:
        return False


def wait_for_text_control(
    root,
    patterns: str | Sequence[str],
    *,
    control_types: Sequence[str] = (),
    timeout: float = 15.0,
    poll_interval: float = 0.25,
):
    deadline = time.monotonic() + timeout
    last_error: UiAutomationError | None = None
    while time.monotonic() < deadline:
        try:
            return find_text_control(root, patterns, control_types=control_types)
        except UiAutomationError as exc:
            last_error = exc
            time.sleep(poll_interval)
    if last_error is not None:
        raise last_error
    raise UiAutomationError(f"Could not find control matching {patterns!r}.")


def find_desktop_text_control(patterns: str | Sequence[str], *, control_types: Sequence[str] = ()):
    _, Desktop, _ = _import_pywinauto()
    compiled = _compiled_patterns(patterns)
    allowed_types = {value.lower() for value in control_types}
    candidates = []
    for window in Desktop(backend="uia").windows():
        for wrapper in _iter_controls(window):
            title = _safe_window_text(wrapper)
            if not title or not _matches_text(title, compiled):
                continue
            control_type = _control_type(wrapper).lower()
            if allowed_types and control_type not in allowed_types:
                continue
            try:
                rect = _wrapper_rect(wrapper)
            except UiAutomationError:
                continue
            candidates.append((rect.top, rect.left, wrapper))
    if not candidates:
        raise UiAutomationError(
            f"Could not find a desktop control matching {tuple(pattern.pattern for pattern in compiled)}."
        )
    candidates.sort(key=lambda item: (item[0], item[1]))
    return candidates[0][2]


def wait_for_desktop_text_control(
    patterns: str | Sequence[str],
    *,
    control_types: Sequence[str] = (),
    timeout: float = 15.0,
    poll_interval: float = 0.25,
):
    deadline = time.monotonic() + timeout
    last_error: UiAutomationError | None = None
    while time.monotonic() < deadline:
        try:
            return find_desktop_text_control(patterns, control_types=control_types)
        except UiAutomationError as exc:
            last_error = exc
            time.sleep(poll_interval)
    if last_error is not None:
        raise last_error
    raise UiAutomationError(f"Could not find desktop control matching {patterns!r}.")


def send_hotkey(keys: str) -> None:
    _, _, keyboard = _import_pywinauto()
    keyboard.send_keys(keys, pause=0.05, with_spaces=True)
    time.sleep(0.25)


def _is_process_running(process_id: int | None) -> bool:
    if not process_id:
        return False
    try:
        _, Desktop, _ = _import_pywinauto()
        for window in Desktop(backend="uia").windows():
            if _get_process_id(window) == process_id:
                return True
    except Exception:
        return False
    return False


def request_window_close(window) -> bool:
    handle = _get_window_handle(window)
    title = _safe_window_text(window) or "<untitled>"
    process_id = _get_process_id(window)
    if handle is None:
        if not _is_process_running(process_id):
            logger.info(
                "Skipping WM_CLOSE because window '%s' no longer exposes a handle and process=%s is not running.",
                title,
                process_id or "<unknown>",
            )
            return False
        raise UiAutomationError(f"Could not resolve a native window handle for '{_safe_window_text(window)}'.")
    if not _is_window_handle_visible(handle):
        logger.info(
            "Skipping WM_CLOSE because window '%s' handle=%s is no longer visible.",
            title,
            handle,
        )
        return False
    result = ctypes.windll.user32.PostMessageW(handle, WM_CLOSE, 0, 0)
    if not result:
        if not _is_window_handle_visible(handle) or not _is_process_running(process_id):
            logger.info(
                "WM_CLOSE target '%s' disappeared while closing. handle=%s process=%s",
                title,
                handle,
                process_id or "<unknown>",
            )
            return False
        raise UiAutomationError(
            f"Windows rejected the close request for '{title}' (handle={handle})."
        )
    logger.info("Posted WM_CLOSE to window '%s' handle=%s.", title, handle)
    return True


def _find_first_control(root, control_types: Sequence[str]):
    allowed_types = {value.lower() for value in control_types}
    candidates = []
    for wrapper in _iter_controls(root):
        control_type = _control_type(wrapper).lower()
        if control_type not in allowed_types:
            continue
        rect = _wrapper_rect(wrapper)
        candidates.append((rect.top, rect.left, wrapper))
    if not candidates:
        raise UiAutomationError(f"Could not find a control of types {tuple(control_types)}.")
    candidates.sort(key=lambda item: (item[0], item[1]))
    return candidates[0][2]


def _find_last_matching_control(root, patterns: str | Sequence[str], *, control_types: Sequence[str]):
    compiled = _compiled_patterns(patterns)
    allowed_types = {value.lower() for value in control_types}
    candidates = []
    for wrapper in _iter_controls(root):
        title = _safe_window_text(wrapper)
        if not title or not _matches_text(title, compiled):
            continue
        control_type = _control_type(wrapper).lower()
        if allowed_types and control_type not in allowed_types:
            continue
        rect = _wrapper_rect(wrapper)
        candidates.append((rect.bottom, rect.right, wrapper))
    if not candidates:
        raise UiAutomationError(
            f"Could not find a control matching {tuple(pattern.pattern for pattern in compiled)} "
            f"under '{_safe_window_text(root)}'."
        )
    candidates.sort(key=lambda item: (item[0], item[1]), reverse=True)
    return candidates[0][2]


def _find_labeled_edit(root, label_patterns: Sequence[str]):
    label = find_text_control(root, label_patterns, control_types=("Text", "Group", "Pane"))
    label_rect = _wrapper_rect(label)
    candidates = []
    for wrapper in _iter_controls(root):
        control_type = _control_type(wrapper).lower()
        if control_type not in {"edit", "combobox"}:
            continue
        rect = _wrapper_rect(wrapper)
        if rect.top < label_rect.top - 40 or rect.bottom > label_rect.bottom + 80:
            continue
        if rect.left <= label_rect.right:
            continue
        score = (abs(rect.top - label_rect.top), rect.left - label_rect.right, rect.top, rect.left)
        candidates.append((score, wrapper))
    if not candidates:
        raise UiAutomationError(
            f"Could not find an edit field next to labels {tuple(label_patterns)} under '{_safe_window_text(root)}'."
        )
    candidates.sort(key=lambda item: item[0])
    return candidates[0][1]


def find_labeled_edit(root, label_patterns: Sequence[str]):
    return _find_labeled_edit(root, label_patterns)


def _find_file_dialog_edit(dialog, confirm_patterns: Sequence[str]):
    try:
        return _find_labeled_edit(dialog, FILE_NAME_LABEL_PATTERNS)
    except UiAutomationError:
        pass

    confirm_rect = None
    try:
        confirm_button = _find_last_matching_control(dialog, confirm_patterns, control_types=("Button",))
        confirm_rect = _wrapper_rect(confirm_button)
    except Exception:
        confirm_rect = None

    candidates = []
    for wrapper in _iter_controls(dialog):
        control_type = _control_type(wrapper).lower()
        if control_type not in {"edit", "combobox"}:
            continue
        rect = _wrapper_rect(wrapper)
        width = rect.right - rect.left
        height = rect.bottom - rect.top
        if width < 60 or height < 14:
            continue
        if confirm_rect is not None:
            center_y_distance = abs(((rect.top + rect.bottom) // 2) - ((confirm_rect.top + confirm_rect.bottom) // 2))
            score = (
                0 if rect.left < confirm_rect.left else 1,
                0 if rect.bottom >= confirm_rect.top - 120 else 1,
                center_y_distance,
                -width,
                -rect.top,
                rect.left,
            )
        else:
            score = (0, 0, 0, -width, -rect.top, rect.left)
        candidates.append((score, wrapper))

    if not candidates:
        return _find_first_control(dialog, ("Edit", "ComboBox"))
    candidates.sort(key=lambda item: item[0])
    return candidates[0][1]


def _dialog_contains_file(dialog, file_path: Path) -> bool:
    candidates = {file_path.name.casefold(), file_path.stem.casefold()}
    for wrapper in _iter_controls(dialog):
        text = _safe_window_text(wrapper).casefold()
        if text and text in candidates:
            return True
    return False


def _find_dialog_file_item(dialog, file_path: Path):
    candidates = []
    target_names = {file_path.name.casefold(), file_path.stem.casefold()}
    type_priority = {"dataitem": 0, "listitem": 1, "treeitem": 2, "text": 3}
    allowed_types = {value.lower() for value in FILE_LIST_CONTROL_TYPES}
    for wrapper in _iter_controls(dialog):
        text = _safe_window_text(wrapper).casefold()
        if not text or text not in target_names:
            continue
        control_type = _control_type(wrapper).lower()
        if control_type not in allowed_types:
            continue
        rect = _wrapper_rect(wrapper)
        candidates.append((type_priority.get(control_type, 99), rect.top, rect.left, wrapper))
    if not candidates:
        raise UiAutomationError(f"Could not find file item '{file_path.name}' in the current dialog list.")
    candidates.sort(key=lambda item: (item[0], item[1], item[2]))
    return candidates[0][3]


def _try_select_file_from_dialog(dialog, file_path: Path, confirm_patterns: Sequence[str]) -> bool:
    try:
        file_item = _find_dialog_file_item(dialog, file_path)
    except UiAutomationError:
        return False

    logger.info("Selecting visible file item from dialog list: %s", file_path.name)
    bring_window_to_front(dialog, keep_topmost=False)
    click_control(file_item, post_click_sleep=0.2)
    dialog_handle = _get_window_handle(dialog)

    confirm_button = _find_last_matching_control(dialog, confirm_patterns, control_types=("Button",))
    deadline = time.monotonic() + 1.0
    while time.monotonic() < deadline:
        try:
            if confirm_button.is_enabled():
                logger.info(
                    "Clicking file dialog confirmation button after file selection: %s",
                    _safe_window_text(confirm_button) or "<untitled>",
                )
                click_control(confirm_button)
                if _wait_for_window_to_close(dialog_handle, timeout=0.8):
                    return True
                logger.info("File dialog stayed open after clicking the selected-file confirmation button.")
                break
        except Exception:
            logger.info("Confirmation button state could not be read reliably. Clicking it directly.")
            click_control(confirm_button)
            if _wait_for_window_to_close(dialog_handle, timeout=0.8):
                return True
            logger.info("File dialog stayed open after directly clicking the confirmation button.")
            break
        time.sleep(0.05)

    logger.info("Trying double-click on the selected file item.")
    try:
        file_item.double_click_input()
        if _wait_for_window_to_close(dialog_handle, timeout=0.8):
            return True
        logger.info("File dialog stayed open after double-clicking the selected file item.")
    except Exception:
        pass

    logger.info("Trying Enter on the selected file item.")
    try:
        bring_window_to_front(dialog, keep_topmost=False)
        click_control(file_item, post_click_sleep=0.0)
    except Exception:
        bring_window_to_front(dialog, keep_topmost=False)
    send_hotkey("{ENTER}")
    if _wait_for_window_to_close(dialog_handle, timeout=0.8):
        return True

    logger.info("Trying Alt+O on the selected file item.")
    bring_window_to_front(dialog, keep_topmost=False)
    send_hotkey("%o")
    if _wait_for_window_to_close(dialog_handle, timeout=0.8):
        return True
    return False


def _is_file_dialog_candidate(
    window,
    confirm_patterns: Sequence[str],
    dialog_patterns: Sequence[str] = DEFAULT_DIALOG_PATTERNS,
) -> bool:
    title = _safe_window_text(window)
    class_name = _class_name(window)
    control_type = _control_type(window).lower()
    title_matches = bool(title and _matches_text(title, _compiled_patterns(dialog_patterns)))

    try:
        find_text_control(window, confirm_patterns, control_types=("Button",))
    except UiAutomationError:
        return False

    try:
        _find_first_control(window, ("Edit", "ComboBox"))
    except UiAutomationError:
        return False

    file_name_label_present = _dialog_has_text(window, FILE_NAME_LABEL_PATTERNS)
    if title_matches:
        return True
    if class_name == "#32770":
        return True
    if control_type == "window" and file_name_label_present:
        return True
    return False


def _dialog_has_text(window, patterns: Sequence[str]) -> bool:
    compiled = _compiled_patterns(patterns)
    for wrapper in _iter_controls(window):
        text = _safe_window_text(wrapper)
        if text and _matches_text(text, compiled):
            return True
    return False


def _is_overwrite_dialog_candidate(window) -> bool:
    try:
        class_name = _class_name(window)
        title = _safe_window_text(window)
        title_matches = bool(title and _matches_text(title, _compiled_patterns(OVERWRITE_DIALOG_TITLE_PATTERNS)))
        if not title_matches and class_name != "#32770":
            return False
        find_text_control(window, OVERWRITE_CONFIRM_PATTERNS, control_types=("Button",))
    except Exception:
        return False
    if title_matches:
        return True
    return _dialog_has_text(window, OVERWRITE_DIALOG_TEXT_PATTERNS)


def _find_nested_file_dialog(dialog_patterns: Sequence[str], confirm_patterns: Sequence[str]):
    _, Desktop, _ = _import_pywinauto()
    compiled = _compiled_patterns(dialog_patterns)
    candidates = []
    for top_window in Desktop(backend="uia").windows():
        try:
            children = list(top_window.children())
        except Exception:
            continue
        for child in children:
            class_name = (getattr(child.element_info, "class_name", "") or "").strip()
            control_type = _control_type(child).lower()
            title = _safe_window_text(child)
            looks_like_dialog_window = class_name == "#32770" or control_type == "window"
            title_matches = bool(title and _matches_text(title, compiled))
            if not looks_like_dialog_window and not title_matches:
                continue
            try:
                is_candidate = _is_file_dialog_candidate(child, confirm_patterns, dialog_patterns)
            except Exception:
                continue
            if not is_candidate:
                continue
            try:
                rect = _wrapper_rect(child)
            except Exception:
                continue
            candidates.append((rect.top, rect.left, child))
    if not candidates:
        raise UiAutomationError("Could not locate a nested file dialog window.")
    candidates.sort(key=lambda item: (item[0], item[1]))
    return candidates[0][2]


def _find_nested_overwrite_dialog(root_window=None):
    _, Desktop, _ = _import_pywinauto()
    owner_windows = []
    if root_window is not None:
        owner_windows.append(root_window)
        try:
            owner_windows.append(root_window.top_level_parent())
        except Exception:
            pass
    candidates = []
    top_windows = owner_windows or list(Desktop(backend="uia").windows())
    for top_window in top_windows:
        try:
            children = list(top_window.children())
        except Exception:
            continue
        for child in children:
            class_name = _class_name(child)
            control_type = _control_type(child).lower()
            if class_name != "#32770" and control_type != "window":
                continue
            try:
                is_candidate = _is_overwrite_dialog_candidate(child)
            except Exception:
                continue
            if not is_candidate:
                continue
            try:
                rect = _wrapper_rect(child)
            except Exception:
                continue
            candidates.append((rect.top, rect.left, child))
    if not candidates:
        raise UiAutomationError("Could not locate an overwrite confirmation dialog.")
    candidates.sort(key=lambda item: (item[0], item[1]))
    return candidates[0][2]


def _normalize_edit_value(value: str) -> str:
    return value.strip().strip('"').replace("/", "\\").casefold()


def _edit_value_matches(actual_value: str, expected_value: str) -> bool:
    normalized_actual = _normalize_edit_value(actual_value)
    normalized_expected = _normalize_edit_value(expected_value)
    if normalized_actual == normalized_expected:
        return True
    expected_name = _normalize_edit_value(Path(expected_value).name)
    return bool(expected_name) and normalized_actual == expected_name


def _resolve_text_entry_target(edit_wrapper):
    control_type = _control_type(edit_wrapper).lower()
    if control_type != "combobox":
        return edit_wrapper
    candidates = []
    for child in edit_wrapper.descendants():
        child_type = _control_type(child).lower()
        if child_type not in {"edit", "document"}:
            continue
        rect = _wrapper_rect(child)
        width = rect.right - rect.left
        candidates.append((-width, rect.top, rect.left, child))
    if not candidates:
        return edit_wrapper
    candidates.sort(key=lambda item: (item[0], item[1], item[2]))
    return candidates[0][3]


def _read_edit_value(edit_wrapper) -> str:
    target = _resolve_text_entry_target(edit_wrapper)
    readers = (
        lambda: target.get_value(),
        lambda: target.iface_value.CurrentValue,
        lambda: target.window_text(),
        lambda: target.texts()[0] if target.texts() else "",
    )
    for reader in readers:
        try:
            value = reader()
        except Exception:
            continue
        if isinstance(value, str):
            return value
    return ""


def _open_clipboard_with_retry(retries: int = 10, delay_seconds: float = 0.1) -> None:
    user32 = ctypes.windll.user32
    user32.OpenClipboard.argtypes = [wintypes.HWND]
    user32.OpenClipboard.restype = wintypes.BOOL
    for _ in range(retries):
        if user32.OpenClipboard(None):
            return
        time.sleep(delay_seconds)
    raise UiAutomationError("Could not open the Windows clipboard.")


def _set_clipboard_text(value: str) -> None:
    user32 = ctypes.windll.user32
    kernel32 = ctypes.windll.kernel32
    user32.EmptyClipboard.argtypes = []
    user32.EmptyClipboard.restype = wintypes.BOOL
    user32.SetClipboardData.argtypes = [wintypes.UINT, wintypes.HANDLE]
    user32.SetClipboardData.restype = wintypes.HANDLE
    user32.CloseClipboard.argtypes = []
    user32.CloseClipboard.restype = wintypes.BOOL
    kernel32.GlobalAlloc.argtypes = [wintypes.UINT, ctypes.c_size_t]
    kernel32.GlobalAlloc.restype = wintypes.HGLOBAL
    kernel32.GlobalLock.argtypes = [wintypes.HGLOBAL]
    kernel32.GlobalLock.restype = wintypes.LPVOID
    kernel32.GlobalUnlock.argtypes = [wintypes.HGLOBAL]
    kernel32.GlobalUnlock.restype = wintypes.BOOL
    kernel32.GlobalFree.argtypes = [wintypes.HGLOBAL]
    kernel32.GlobalFree.restype = wintypes.HGLOBAL

    buffer = ctypes.create_unicode_buffer(value + "\0")
    buffer_size = ctypes.sizeof(buffer)
    handle = None
    _open_clipboard_with_retry()
    try:
        if not user32.EmptyClipboard():
            raise UiAutomationError("Could not empty the Windows clipboard.")
        handle = kernel32.GlobalAlloc(GMEM_MOVEABLE, buffer_size)
        if not handle:
            raise UiAutomationError("Could not allocate clipboard memory.")
        locked = kernel32.GlobalLock(handle)
        if not locked:
            raise UiAutomationError("Could not lock clipboard memory.")
        try:
            ctypes.memmove(locked, ctypes.addressof(buffer), buffer_size)
        finally:
            kernel32.GlobalUnlock(handle)
        if not user32.SetClipboardData(CF_UNICODETEXT, handle):
            raise UiAutomationError("Could not update the Windows clipboard.")
        handle = None
    finally:
        user32.CloseClipboard()
        if handle:
            kernel32.GlobalFree(handle)


def get_clipboard_text() -> str:
    user32 = ctypes.windll.user32
    kernel32 = ctypes.windll.kernel32
    user32.GetClipboardData.argtypes = [wintypes.UINT]
    user32.GetClipboardData.restype = wintypes.HANDLE
    user32.CloseClipboard.argtypes = []
    user32.CloseClipboard.restype = wintypes.BOOL
    kernel32.GlobalLock.argtypes = [wintypes.HGLOBAL]
    kernel32.GlobalLock.restype = wintypes.LPVOID
    kernel32.GlobalUnlock.argtypes = [wintypes.HGLOBAL]
    kernel32.GlobalUnlock.restype = wintypes.BOOL

    _open_clipboard_with_retry()
    try:
        handle = user32.GetClipboardData(CF_UNICODETEXT)
        if not handle:
            return ""
        locked = kernel32.GlobalLock(handle)
        if not locked:
            raise UiAutomationError("Could not lock clipboard memory for reading.")
        try:
            return (ctypes.wstring_at(locked) or "").strip()
        finally:
            kernel32.GlobalUnlock(handle)
    finally:
        user32.CloseClipboard()


def _set_edit_value_with_clipboard(edit_wrapper, value: str) -> None:
    target = _resolve_text_entry_target(edit_wrapper)
    bring_window_to_front(target.top_level_parent(), keep_topmost=False)
    try:
        rect = _wrapper_rect(target)
        width = max(1, rect.right - rect.left)
        height = max(1, rect.bottom - rect.top)
        click_x = max(5, min(18, width // 10))
        click_y = max(1, height // 2)
        target.click_input(coords=(click_x, click_y))
    except Exception:
        try:
            target.click_input()
        except Exception:
            pass
    send_hotkey("^a")
    send_hotkey("{BACKSPACE}")
    _set_clipboard_text(value)
    send_hotkey("^v")
    time.sleep(0.25)


def _is_window_handle_visible(handle: int) -> bool:
    _, Desktop, _ = _import_pywinauto()
    try:
        for window in Desktop(backend="uia").windows():
            if _get_window_handle(window) == handle:
                return True
    except Exception:
        return False
    return False


def _read_int_like(value) -> int | None:
    current = value
    for _ in range(2):
        if callable(current):
            try:
                current = current()
            except Exception:
                return None
            continue
        break
    if current in (None, ""):
        return None
    try:
        return int(current)
    except (TypeError, ValueError):
        return None


def _get_window_handle(window) -> int | None:
    readers = (
        lambda: getattr(getattr(window, "element_info", None), "handle", None),
        lambda: getattr(window, "handle", None),
    )
    for reader in readers:
        try:
            handle = _read_int_like(reader())
        except Exception:
            continue
        if handle:
            return handle
    return None


def _get_process_id(window) -> int | None:
    readers = (
        lambda: getattr(getattr(window, "element_info", None), "process_id", None),
        lambda: getattr(window, "process_id", None),
    )
    for reader in readers:
        try:
            process_id = _read_int_like(reader())
        except Exception:
            continue
        if process_id:
            return process_id
    return None


def _wait_for_window_to_close(window_or_handle, timeout: float = 1.0, poll_interval: float = 0.05) -> bool:
    deadline = time.monotonic() + timeout
    handle = window_or_handle if isinstance(window_or_handle, int) else _get_window_handle(window_or_handle)
    if not handle:
        return True
    while time.monotonic() < deadline:
        if not _is_window_handle_visible(handle):
            return True
        time.sleep(poll_interval)
    return not _is_window_handle_visible(handle)


def paste_text_via_clipboard(value: str, *, replace_existing: bool = True) -> None:
    _set_clipboard_text(value)
    if replace_existing:
        send_hotkey("^a")
        send_hotkey("{BACKSPACE}")
    send_hotkey("^v")
    time.sleep(0.25)


def wait_for_file_dialog(
    *,
    dialog_patterns: Sequence[str] = DEFAULT_DIALOG_PATTERNS,
    confirm_patterns: Sequence[str] = DEFAULT_CONFIRM_PATTERNS,
    timeout: float = 15.0,
):
    logger.info(
        "Waiting for file dialog. dialog_patterns=%s confirm_patterns=%s timeout=%.1fs",
        tuple(dialog_patterns),
        tuple(confirm_patterns),
        timeout,
    )
    deadline = time.monotonic() + timeout
    last_error: UiAutomationError | None = None
    while time.monotonic() < deadline:
        try:
            window = wait_for_window(dialog_patterns, timeout=0.75)
            if _is_file_dialog_candidate(window, confirm_patterns, dialog_patterns):
                logger.info("Matched top-level file dialog: %s", _safe_window_text(window) or "<untitled>")
                return window
        except UiAutomationError as exc:
            last_error = exc
        try:
            foreground_window = get_foreground_window()
            if _is_file_dialog_candidate(foreground_window, confirm_patterns, dialog_patterns):
                logger.info("Matched foreground file dialog: %s", _safe_window_text(foreground_window) or "<untitled>")
                return foreground_window
        except UiAutomationError:
            pass
        try:
            nested_dialog = _find_nested_file_dialog(dialog_patterns, confirm_patterns)
            logger.info("Matched nested file dialog: %s", _safe_window_text(nested_dialog) or "<untitled>")
            return nested_dialog
        except UiAutomationError:
            pass
        time.sleep(0.25)
    if last_error is not None:
        raise last_error
    raise UiAutomationError("Could not locate a file dialog window.")


def _find_overwrite_dialog(owner_window=None):
    if owner_window is not None:
        try:
            if _is_overwrite_dialog_candidate(owner_window):
                return owner_window
        except Exception:
            pass
        try:
            nested_dialog = _find_nested_overwrite_dialog(owner_window)
            return nested_dialog
        except Exception:
            pass
    try:
        foreground_window = get_foreground_window()
        try:
            if _is_overwrite_dialog_candidate(foreground_window):
                return foreground_window
        except Exception:
            pass
    except Exception:
        pass
    _, Desktop, _ = _import_pywinauto()
    compiled_titles = _compiled_patterns(OVERWRITE_DIALOG_TITLE_PATTERNS)
    candidates = []
    for window in Desktop(backend="uia").windows():
        title = _safe_window_text(window)
        class_name = _class_name(window)
        if not (title and _matches_text(title, compiled_titles)) and class_name != "#32770":
            continue
        try:
            is_candidate = _is_overwrite_dialog_candidate(window)
        except Exception:
            continue
        if not is_candidate:
            continue
        try:
            rect = _wrapper_rect(window)
        except Exception:
            continue
        candidates.append((rect.top, rect.left, window))
    if candidates:
        candidates.sort(key=lambda item: (item[0], item[1]))
        return candidates[0][2]
    try:
        return _find_nested_overwrite_dialog()
    except Exception as exc:
        raise UiAutomationError("Could not locate an overwrite confirmation dialog.") from exc


def accept_overwrite_confirmation(timeout: float = 2.0, *, owner_window=None, poll_interval: float = 0.05) -> bool:
    logger.info("Checking for overwrite confirmation dialog. timeout=%.1fs", timeout)
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        try:
            dialog = _find_overwrite_dialog(owner_window)
            logger.info("Overwrite confirmation detected: %s", _safe_window_text(dialog) or "<untitled>")
            bring_window_to_front(dialog, keep_topmost=False)
            confirm_button = find_text_control(dialog, OVERWRITE_CONFIRM_PATTERNS, control_types=("Button",))
            logger.info("Clicking overwrite confirmation button: %s", _safe_window_text(confirm_button) or "<untitled>")
            click_control(confirm_button)
            return True
        except Exception:
            pass
        time.sleep(poll_interval)
    logger.info("No overwrite confirmation dialog appeared.")
    return False


def _set_edit_value(edit_wrapper, value: str) -> None:
    target = _resolve_text_entry_target(edit_wrapper)
    attempts = [
        lambda: target.iface_value.SetValue(value),
    ]
    if hasattr(target, "set_edit_text"):
        attempts.append(lambda: target.set_edit_text(value))
    if hasattr(target, "set_window_text"):
        attempts.append(lambda: target.set_window_text(value))
    attempts.append(lambda: _set_edit_value_with_clipboard(target, value))

    last_error: Exception | None = None
    for attempt in attempts:
        try:
            attempt()
        except Exception as exc:
            last_error = exc
        actual_value = _read_edit_value(target)
        if _edit_value_matches(actual_value, value):
            time.sleep(0.25)
            return
    raise UiAutomationError(
        f"Could not set file dialog input to '{value}'. "
        f"Last observed value: '{_read_edit_value(target)}'."
    ) from last_error


def set_edit_text_value(edit_wrapper, value: str) -> None:
    _set_edit_value(edit_wrapper, value)


def fill_file_dialog(
    file_path: str | Path,
    *,
    dialog_patterns: Sequence[str] = DEFAULT_DIALOG_PATTERNS,
    confirm_patterns: Sequence[str] = DEFAULT_CONFIRM_PATTERNS,
    timeout: float = 15.0,
    must_exist: bool = True,
    allow_direct_selection: bool = True,
    wait_for_dialog_to_close: bool = True,
    overwrite_confirmation_timeout: float | None = None,
) -> None:
    normalized_path = Path(file_path).resolve(strict=False)
    logger.info(
        "Preparing file dialog input. path=%s must_exist=%s timeout=%.1fs",
        normalized_path,
        must_exist,
        timeout,
    )
    assert normalized_path.parent.exists(), f"Parent directory does not exist: {normalized_path.parent}"
    if must_exist:
        assert normalized_path.exists(), f"Path does not exist: {normalized_path}"

    dialog = wait_for_file_dialog(
        dialog_patterns=dialog_patterns,
        confirm_patterns=confirm_patterns,
        timeout=timeout,
    )
    bring_window_to_front(dialog, keep_topmost=False)
    logger.info("Filling dialog window: %s", _safe_window_text(dialog) or "<untitled>")

    if allow_direct_selection and must_exist and _dialog_contains_file(dialog, normalized_path):
        logger.info("Target file is already visible in the current dialog list: %s", normalized_path.name)
        if _try_select_file_from_dialog(dialog, normalized_path, confirm_patterns):
            return
        logger.info("Direct file selection did not succeed. Falling back to edit input.")

    edit = _find_file_dialog_edit(dialog, confirm_patterns)
    logger.info(
        "Resolved file dialog input control. control_type=%s title=%s",
        _control_type(edit) or "<unknown>",
        _safe_window_text(edit) or "<untitled>",
    )
    input_value = str(normalized_path)
    if must_exist and _dialog_contains_file(dialog, normalized_path):
        input_value = normalized_path.name
    logger.info("Setting file dialog text to: %s", input_value)
    try:
        _set_edit_value(edit, input_value)
    except UiAutomationError:
        if must_exist and input_value != normalized_path.name:
            logger.info("Absolute path input was unstable. Falling back to filename only: %s", normalized_path.name)
            _set_edit_value(edit, normalized_path.name)
        else:
            raise
    dialog_handle = _get_window_handle(dialog)
    confirm_button = _find_last_matching_control(dialog, confirm_patterns, control_types=("Button",))
    logger.info("Clicking file dialog confirmation button: %s", _safe_window_text(confirm_button) or "<untitled>")
    click_control(confirm_button)
    if wait_for_dialog_to_close:
        if not _wait_for_window_to_close(dialog_handle, timeout=0.8):
            logger.info("File dialog stayed open after clicking the confirmation button. Retrying with Enter.")
            try:
                target = _resolve_text_entry_target(edit)
                bring_window_to_front(dialog, keep_topmost=False)
                try:
                    rect = _wrapper_rect(target)
                    click_x = max(5, min(18, max(1, (rect.right - rect.left)) // 10))
                    click_y = max(1, max(1, (rect.bottom - rect.top)) // 2)
                    target.click_input(coords=(click_x, click_y))
                except Exception:
                    try:
                        target.click_input()
                    except Exception:
                        pass
            except Exception:
                bring_window_to_front(dialog, keep_topmost=False)
            send_hotkey("{ENTER}")
            if not _wait_for_window_to_close(dialog_handle, timeout=0.8):
                logger.info("File dialog stayed open after Enter. Retrying with Alt+O.")
                bring_window_to_front(dialog, keep_topmost=False)
                send_hotkey("%o")
                if not _wait_for_window_to_close(dialog_handle, timeout=0.8):
                    raise UiAutomationError(
                        f"File dialog confirmation did not close the dialog for '{normalized_path}'."
                    )
    if not must_exist:
        if overwrite_confirmation_timeout is None:
            overwrite_confirmation_timeout = 2.0
        if overwrite_confirmation_timeout > 0:
            accept_overwrite_confirmation(
                timeout=overwrite_confirmation_timeout,
                owner_window=dialog,
                poll_interval=0.05,
            )


def dismiss_close_prompts(timeout: float = 2.0, owner_window=None) -> None:
    deadline = time.monotonic() + timeout
    owner_process_id = _get_process_id(owner_window) if owner_window is not None else None
    _, Desktop, _ = _import_pywinauto()
    compiled_titles = _compiled_patterns(COMMON_CLOSE_DIALOG_PATTERNS)
    while time.monotonic() < deadline:
        candidates = []
        for window in Desktop(backend="uia").windows():
            title = _safe_window_text(window)
            if not title or not _matches_text(title, compiled_titles):
                continue
            if owner_process_id is not None and _get_process_id(window) != owner_process_id:
                continue
            try:
                rect = _wrapper_rect(window)
            except UiAutomationError:
                continue
            candidates.append((rect.top, rect.left, window))
        if not candidates:
            return
        candidates.sort(key=lambda item: (item[0], item[1]))
        dialog = candidates[0][2]
        dialog_handle = _get_window_handle(dialog)
        try:
            click_text_control(dialog, COMMON_DISCARD_PATTERNS, control_types=("Button",))
            if _wait_for_window_to_close(dialog_handle, timeout=0.8):
                continue
        except UiAutomationError:
            pass
        try:
            bring_window_to_front(dialog, keep_topmost=False)
            send_hotkey("%d")
            if _wait_for_window_to_close(dialog_handle, timeout=0.6):
                continue
            send_hotkey("%n")
            if _wait_for_window_to_close(dialog_handle, timeout=0.6):
                continue
        except Exception:
            return
        time.sleep(0.1)


def _vertical_overlap(a, b) -> int:
    return max(0, min(a.bottom, b.bottom) - max(a.top, b.top))


def _find_title_element(root, title_pattern: str):
    compiled = re.compile(title_pattern, re.IGNORECASE)
    candidates = []
    for wrapper in root.descendants():
        title = _safe_window_text(wrapper)
        if not title or not compiled.fullmatch(title):
            continue
        rect = _wrapper_rect(wrapper)
        candidates.append((rect.top, rect.left, wrapper))
    if not candidates:
        raise UiAutomationError(f"Could not find UI text matching '{title_pattern}'.")
    candidates.sort(key=lambda item: (item[0], item[1]))
    return candidates[0][2]


def _get_toggle_state(wrapper) -> bool | None:
    try:
        return bool(wrapper.get_toggle_state())
    except Exception:
        pass
    try:
        return bool(wrapper.iface_toggle.CurrentToggleState)
    except Exception:
        return None


def _find_row_toggle(root, label_wrapper):
    label_rect = _wrapper_rect(label_wrapper)
    best_match = None
    best_score = None
    for candidate in root.descendants():
        control_type = _control_type(candidate).lower()
        if control_type not in {"button", "checkbox"}:
            continue
        rect = _wrapper_rect(candidate)
        if rect.left < label_rect.right:
            continue
        overlap = _vertical_overlap(label_rect, rect)
        if overlap <= 0:
            continue
        score = (rect.left - label_rect.right, -overlap, abs(rect.top - label_rect.top))
        if best_score is None or score < best_score:
            best_score = score
            best_match = candidate
    if best_match is None:
        label = _safe_window_text(label_wrapper) or "<unknown>"
        raise UiAutomationError(f"Could not find a checkbox or toggle next to '{label}'.")
    return best_match


def _set_toggle_state(toggle_wrapper, expected_checked: bool) -> bool:
    current_state = _get_toggle_state(toggle_wrapper)
    if current_state is not None and current_state == expected_checked:
        return False
    toggle_wrapper.click_input()
    time.sleep(0.4)
    current_state = _get_toggle_state(toggle_wrapper)
    if current_state is not None and current_state != expected_checked:
        raise UiAutomationError("Clicked the control, but its checked state did not change as expected.")
    return True


class WindowTopmostKeeper:
    def __init__(self, window_title_re: str, *, interval_seconds: float = 1.0) -> None:
        assert interval_seconds > 0, "interval_seconds must be positive."
        self.window_title_re = window_title_re
        self.interval_seconds = interval_seconds
        self._stop_event = threading.Event()
        self._thread: threading.Thread | None = None

    def start(self) -> None:
        if self._thread is not None:
            return
        self._thread = threading.Thread(target=self._run, name="ai-turbo-topmost-keeper", daemon=True)
        self._thread.start()

    def stop(self) -> None:
        self._stop_event.set()
        if self._thread is not None:
            self._thread.join(timeout=2.0)
            self._thread = None
        try:
            clear_window_topmost(connect_window(self.window_title_re, timeout=1.0))
        except UiAutomationError:
            pass

    def _run(self) -> None:
        while not self._stop_event.is_set():
            try:
                window = connect_window(self.window_title_re, timeout=1.0)
                _set_window_topmost(window.handle, True)
            except UiAutomationError:
                pass
            self._stop_event.wait(self.interval_seconds)


def set_performance_boost_state(
    window_title_re: str,
    app_name: str,
    *,
    enabled: bool,
    timeout: float = 15.0,
) -> list[UiActionResult]:
    Application, _, _ = _import_pywinauto()
    app = Application(backend="uia").connect(title_re=window_title_re, timeout=timeout)
    window = app.window(title_re=window_title_re)
    window.wait("visible", timeout=timeout)
    bring_window_to_front(window)

    boost_label = _find_title_element(window, r"Performance Boost")
    boost_toggle = _find_row_toggle(window, boost_label)
    app_label = _find_title_element(window, re.escape(app_name))
    app_checkbox = _find_row_toggle(window, app_label)

    if enabled:
        boost_changed = _set_toggle_state(boost_toggle, expected_checked=True)
        app_changed = _set_toggle_state(app_checkbox, expected_checked=True)
        boost_state = "checked"
        app_state = "checked"
    else:
        app_changed = _set_toggle_state(app_checkbox, expected_checked=False)
        boost_changed = _set_toggle_state(boost_toggle, expected_checked=False)
        boost_state = "unchecked"
        app_state = "unchecked"

    return [
        UiActionResult(target="performance_boost", changed=boost_changed, state=boost_state),
        UiActionResult(target=app_name, changed=app_changed, state=app_state),
    ]


def set_performance_boost_selection(window_title_re: str, app_name: str, timeout: float = 15.0) -> list[UiActionResult]:
    return set_performance_boost_state(
        window_title_re,
        app_name,
        enabled=True,
        timeout=timeout,
    )





