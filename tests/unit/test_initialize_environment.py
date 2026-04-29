from __future__ import annotations

from dataclasses import replace
from pathlib import Path

import initialize_environment
import pytest


class _DummyRect:
    def __init__(self, *, top: int, left: int, width: int = 100, height: int = 80) -> None:
        self.top = top
        self.left = left
        self._width = width
        self._height = height
        self.right = left + width
        self.bottom = top + height

    def width(self) -> int:
        return self._width

    def height(self) -> int:
        return self._height


class _DummyElementInfo:
    def __init__(self, process_id: int, *, control_type: str = "Window") -> None:
        self.process_id = process_id
        self.control_type = control_type


class _DummyWindow:
    def __init__(
        self,
        title: str,
        process_id: int,
        *,
        top: int,
        left: int,
        width: int = 100,
        height: int = 80,
        control_type: str = "Window",
        enabled: bool = True,
        children: list["_DummyWindow"] | None = None,
    ) -> None:
        self._title = title
        self._rect = _DummyRect(top=top, left=left, width=width, height=height)
        self.element_info = _DummyElementInfo(process_id, control_type=control_type)
        self._enabled = enabled
        self._children = list(children or [])
        self.clicks: list[tuple[int, int] | None] = []
        self._top_level_parent: _DummyWindow = self
        for child in self._children:
            child._top_level_parent = self

    def window_text(self) -> str:
        return self._title

    def rectangle(self) -> _DummyRect:
        return self._rect

    def is_visible(self) -> bool:
        return True

    def is_enabled(self) -> bool:
        return self._enabled

    def descendants(self) -> list["_DummyWindow"]:
        return list(self._children)

    def click_input(self, coords: tuple[int, int] | None = None) -> None:
        self.clicks.append(coords)

    def top_level_parent(self) -> "_DummyWindow":
        return self._top_level_parent


class _DummyDesktop:
    def __init__(self, windows: list[_DummyWindow]) -> None:
        self._windows = windows

    def __call__(self, backend: str = "uia") -> "_DummyDesktop":
        assert backend == "uia"
        return self

    def windows(self) -> list[_DummyWindow]:
        return list(self._windows)


def _build_profile(tmp_path: Path) -> initialize_environment.SoftwareInstallProfile:
    return initialize_environment.SoftwareInstallProfile(
        software="avidemux",
        installer_path=tmp_path / "installer.exe",
        shortcut_path=tmp_path / "avidemux.lnk",
        install_root=tmp_path,
        display_name_patterns=("avidemux",),
        executable_names=("avidemux.exe",),
        target_candidates=(tmp_path / "avidemux.exe",),
        silent_argument_sets=(),
    )


def test_detect_installed_executable_ignores_incomplete_avidemux_install(monkeypatch, tmp_path: Path) -> None:
    profile = _build_profile(tmp_path)
    executable = tmp_path / "avidemux.exe"
    executable.write_bytes(b"exe")

    monkeypatch.setattr(initialize_environment, "_search_custom_install_root", lambda profile: None)
    monkeypatch.setattr(initialize_environment, "_search_registry_for_executable", lambda profile: None)
    monkeypatch.setattr(initialize_environment, "_search_common_roots", lambda profile: None)

    assert initialize_environment.detect_installed_executable(profile) is None


def test_detect_installed_executable_accepts_complete_avidemux_install(monkeypatch, tmp_path: Path) -> None:
    profile = _build_profile(tmp_path)
    executable = tmp_path / "avidemux.exe"
    executable.write_bytes(b"exe")
    (tmp_path / "plugins" / "demuxers").mkdir(parents=True)
    (tmp_path / "plugins" / "muxers").mkdir(parents=True)
    (tmp_path / "plugins" / "videoDecoders").mkdir(parents=True)

    monkeypatch.setattr(initialize_environment, "_search_custom_install_root", lambda profile: None)
    monkeypatch.setattr(initialize_environment, "_search_registry_for_executable", lambda profile: None)
    monkeypatch.setattr(initialize_environment, "_search_common_roots", lambda profile: None)

    assert initialize_environment.detect_installed_executable(profile) == executable


def test_avidemux_profile_uses_clean_runtime_install_root() -> None:
    profile = next(p for p in initialize_environment.SOFTWARE_INSTALL_PROFILES if p.software == "avidemux")

    assert profile.install_root.name == "avidemux_runtime"
    assert any(candidate.parent.name == "avidemux_runtime" for candidate in profile.target_candidates)


def test_blender_installer_commands_include_repo_local_install_root() -> None:
    profile = next(p for p in initialize_environment.SOFTWARE_INSTALL_PROFILES if p.software == "blender")

    commands = initialize_environment.installer_commands(profile)

    expected_arguments = {
        f"INSTALLDIR={profile.install_root}",
        f"INSTALLFOLDER={profile.install_root}",
        f"TARGETDIR={profile.install_root}",
    }
    assert len(commands) == 2
    assert expected_arguments.issubset(set(commands[0]))
    assert expected_arguments.issubset(set(commands[1]))
    assert "/qn" in commands[1]


def test_ensure_software_ready_reuses_existing_install(monkeypatch, tmp_path: Path) -> None:
    profile = _build_profile(tmp_path)
    installed_executable = tmp_path / "existing" / "avidemux.exe"
    installed_executable.parent.mkdir(parents=True)
    installed_executable.write_bytes(b"exe")

    run_installer_calls: list[Path] = []
    shortcut_targets: list[Path] = []

    def _run_installer(
        _profile,
        *,
        timeout_seconds=initialize_environment.COMMON_INSTALL_TIMEOUT_SECONDS,
        recorder=None,
    ) -> Path:
        _ = timeout_seconds
        _ = recorder
        run_installer_calls.append(_profile.install_root)
        return _profile.install_root / _profile.executable_names[0]

    monkeypatch.setattr(initialize_environment, "detect_installed_executable", lambda _profile: installed_executable)
    monkeypatch.setattr(initialize_environment, "run_installer", _run_installer)
    monkeypatch.setattr(
        initialize_environment,
        "create_or_update_shortcut",
        lambda _shortcut_path, executable_path: shortcut_targets.append(executable_path),
    )

    executable = initialize_environment.ensure_software_ready(profile)

    assert executable == installed_executable
    assert run_installer_calls == []
    assert shortcut_targets == [installed_executable]


def test_list_candidate_installer_windows_prioritizes_modal_warning(monkeypatch, tmp_path: Path) -> None:
    profile = _build_profile(tmp_path)
    main_window = _DummyWindow("Avidemux VC++ 64bits Setup", 4242, top=10, left=10, width=1200, height=800)
    warning_window = _DummyWindow("Warning", 4242, top=100, left=100, width=400, height=200)
    desktop = _DummyDesktop([main_window, warning_window])

    monkeypatch.setattr(initialize_environment, "_import_pywinauto_desktop", lambda: desktop)

    windows = initialize_environment._list_candidate_installer_windows(
        profile,
        initialize_environment.ShellStartedProcess(4242),
    )

    assert windows[0] is warning_window


def test_advance_installer_window_accepts_license_then_clicks_next(monkeypatch, tmp_path: Path) -> None:
    profile = _build_profile(tmp_path)
    window = _DummyWindow("Avidemux VC++ 64bits Setup", 4242, top=10, left=10, width=1200, height=800)
    actions: list[str] = []

    monkeypatch.setattr(initialize_environment, "bring_window_to_front", lambda *args, **kwargs: None)
    monkeypatch.setattr(initialize_environment, "_ensure_install_destination", lambda window, profile: False)
    monkeypatch.setattr(initialize_environment, "_window_has_completion_text", lambda window: False)
    monkeypatch.setattr(initialize_environment.time, "sleep", lambda _seconds: None)

    def _accept_license_terms(_window) -> bool:
        actions.append("accept")
        return True

    def _advance_installer_buttons(_window) -> bool:
        actions.append("advance")
        return len(actions) >= 2

    monkeypatch.setattr(initialize_environment, "_accept_license_terms", _accept_license_terms)
    monkeypatch.setattr(initialize_environment, "_advance_installer_buttons", _advance_installer_buttons)

    progressed = initialize_environment._advance_installer_window(window, profile)

    assert progressed is True
    assert actions == ["advance", "accept", "advance"]


def test_wait_for_installer_completion_closes_finish_window_before_returning(
    monkeypatch,
    tmp_path: Path,
) -> None:
    profile = _build_profile(tmp_path)
    installed_executable = tmp_path / "avidemux.exe"
    installed_executable.write_bytes(b"exe")
    finish_button = _DummyWindow("Finish", 4242, top=100, left=120, width=90, height=28, control_type="Button")
    window = _DummyWindow(
        "Avidemux VC++ 64bits Setup",
        4242,
        top=10,
        left=10,
        width=640,
        height=420,
        children=[finish_button],
    )
    window_batches = iter(([window], []))
    detected_paths = iter((installed_executable,))

    monkeypatch.setattr(initialize_environment, "bring_window_to_front", lambda *args, **kwargs: None)
    monkeypatch.setattr(initialize_environment, "detect_installed_executable", lambda _profile: next(detected_paths))
    monkeypatch.setattr(
        initialize_environment,
        "_list_candidate_installer_windows",
        lambda _profile, _process: next(window_batches),
    )
    monkeypatch.setattr(initialize_environment.time, "sleep", lambda _seconds: None)

    executable = initialize_environment._wait_for_installer_completion(
        profile,
        initialize_environment.ShellStartedProcess(4242),
        timeout_seconds=30.0,
    )

    assert executable == installed_executable
    assert finish_button.clicks == [None]


def test_find_row_toggle_prefers_checkbox_left_of_label(monkeypatch, tmp_path: Path) -> None:
    _ = monkeypatch
    _ = tmp_path
    label = _DummyWindow(
        "I accept this license.",
        4242,
        top=60,
        left=100,
        width=180,
        height=22,
        control_type="Text",
    )
    large_button = _DummyWindow(
        "",
        4242,
        top=44,
        left=80,
        width=260,
        height=52,
        control_type="Button",
    )
    checkbox = _DummyWindow(
        "",
        4242,
        top=62,
        left=74,
        width=16,
        height=16,
        control_type="CheckBox",
    )
    root = _DummyWindow(
        "Avidemux Setup",
        4242,
        top=10,
        left=10,
        width=500,
        height=320,
        children=[label, large_button, checkbox],
    )

    match = initialize_environment._find_row_toggle(root, label)

    assert match is checkbox


def test_accept_license_terms_uses_label_hitbox_for_generic_toggle(monkeypatch, tmp_path: Path) -> None:
    _ = tmp_path
    window = _DummyWindow("Avidemux Setup", 4242, top=10, left=10, width=500, height=320)
    label = _DummyWindow(
        "I accept this license.",
        4242,
        top=60,
        left=100,
        width=180,
        height=22,
        control_type="Text",
    )
    generic_button = _DummyWindow(
        "",
        4242,
        top=58,
        left=72,
        width=20,
        height=20,
        control_type="Button",
    )
    clicks: list[str] = []

    def _find_matching_controls(_root, _patterns, *, control_types):
        if control_types == ("Text", "Pane", "Group"):
            return [label]
        return []

    monkeypatch.setattr(initialize_environment, "_find_matching_controls", _find_matching_controls)
    monkeypatch.setattr(initialize_environment, "_find_row_toggle", lambda _root, _label: generic_button)
    monkeypatch.setattr(
        initialize_environment,
        "_click_label_toggle_hitbox",
        lambda _window, _label: clicks.append("hitbox"),
    )
    monkeypatch.setattr(
        initialize_environment,
        "_set_toggle_checked",
        lambda _toggle: (_ for _ in ()).throw(AssertionError("should not click generic toggle directly")),
    )

    assert initialize_environment._accept_license_terms(window) is True
    assert clicks == ["hitbox"]


def test_wait_for_installer_completion_fails_fast_for_non_admin_elevated_installer(
    monkeypatch,
    tmp_path: Path,
) -> None:
    profile = _build_profile(tmp_path)
    process = initialize_environment.ShellStartedProcess(None, requires_elevation=True)
    monotonic_values = iter((0.0, 0.0, 30.0, 95.0))

    monkeypatch.setattr(initialize_environment, "detect_installed_executable", lambda _profile: None)
    monkeypatch.setattr(initialize_environment, "_list_candidate_installer_windows", lambda _profile, _process: [])
    monkeypatch.setattr(initialize_environment, "_current_process_is_elevated", lambda: False)
    monkeypatch.setattr(initialize_environment.time, "sleep", lambda _seconds: None)
    monkeypatch.setattr(initialize_environment.time, "monotonic", lambda: next(monotonic_values))

    with pytest.raises(AssertionError, match="run initialize_environment\\.cmd as Administrator"):
        initialize_environment._wait_for_installer_completion(profile, process, timeout_seconds=900.0)


def test_run_installer_skips_interactive_fallback_after_non_admin_elevated_failure(
    monkeypatch,
    tmp_path: Path,
) -> None:
    installer_path = tmp_path / "installer.exe"
    installer_path.write_bytes(b"installer")
    profile = replace(
        _build_profile(tmp_path),
        installer_path=installer_path,
        silent_argument_sets=(("/S",),),
        prefer_silent=True,
    )
    launch_calls: list[list[str]] = []
    terminated_pids: list[int | None] = []

    monkeypatch.setattr(
        initialize_environment,
        "installer_commands",
        lambda _profile: [[str(installer_path), "/S"], [str(installer_path)]],
    )

    def _launch(command, *, cwd):
        _ = cwd
        launch_calls.append(list(command))
        return initialize_environment.ShellStartedProcess(4242, requires_elevation=True)

    monkeypatch.setattr(initialize_environment, "launch_installer_process", _launch)
    monkeypatch.setattr(
        initialize_environment,
        "_wait_for_installer_completion",
        lambda _profile, _process, *, timeout_seconds: (_ for _ in ()).throw(
            AssertionError("silent install stalled behind elevation")
        ),
    )
    monkeypatch.setattr(initialize_environment, "_current_process_is_elevated", lambda: False)
    monkeypatch.setattr(
        initialize_environment,
        "terminate_process_tree",
        lambda process: terminated_pids.append(getattr(process, "pid", None)),
    )

    with pytest.raises(AssertionError, match="Skipping interactive installer fallback"):
        initialize_environment.run_installer(profile)

    assert launch_calls == [[str(installer_path), "/S"]]
    assert terminated_pids == [4242]
