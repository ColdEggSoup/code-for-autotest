from __future__ import annotations

import argparse
import ctypes
from datetime import datetime
import importlib.util
import json
import logging
import os
from dataclasses import dataclass
from pathlib import Path
import re
import subprocess
import sys
import time
from typing import Sequence
import winreg
from ctypes import wintypes

from workspace_runtime import configure_workspace_runtime

configure_workspace_runtime()

from automation_components import REPO_ROOT, WHITELIST_APP_DIR
from ui_automation import UiAutomationError, bring_window_to_front, find_labeled_edit, set_edit_text_value


logger = logging.getLogger(__name__)

INSTALLER_DIR = WHITELIST_APP_DIR / "app"
PYTHON_PACKAGE_NAMES = ("psutil", "openpyxl", "pywinauto", "pyautogui")
ACCEPTABLE_INSTALL_EXIT_CODES = {0, 1641, 3010}
COMMON_INSTALL_TIMEOUT_SECONDS = 900.0
INSTALLER_WINDOW_POST_EXIT_GRACE_SECONDS = 20.0
INSTALLER_POLL_SECONDS = 1.0
INSTALLER_ACTION_SETTLE_SECONDS = 1.0
PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
SYNCHRONIZE = 0x00100000
STILL_ACTIVE = 259
UNINSTALL_REGISTRY_ROOTS = (
    (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"),
    (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"),
    (winreg.HKEY_CURRENT_USER, r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"),
)
INSTALLER_ACCEPT_LABEL_PATTERNS = (
    r"^i accept.*$",
    r"^i agree.*$",
    r"^accept.*agreement.*$",
    r"^accept.*license.*$",
    r"^accept.*terms.*$",
    r"^agree.*terms.*$",
    r"^\u540c\u610f.*$",
    r"^\u63a5\u53d7.*$",
    r"^\u6211\u540c\u610f.*$",
    r"^\u6211\u63a5\u53d7.*$",
    r"^\u63a5\u53d7\u534f\u8bae.*$",
    r"^\u540c\u610f\u8bb8\u53ef.*$",
)
INSTALLER_NEXT_BUTTON_PATTERNS = (
    r"^next(?:\b|[\s(<]).*$",
    r"^\u4e0b\u4e00\u6b65.*$",
)
INSTALLER_INSTALL_BUTTON_PATTERNS = (
    r"^install(?:\b|[\s(<]).*$",
    r"^\u5f00\u59cb\u5b89\u88c5.*$",
    r"^\u5b89\u88c5.*$",
)
INSTALLER_CONFIRM_BUTTON_PATTERNS = (
    r"^yes(?:\b|[\s(<]).*$",
    r"^ok(?:\b|[\s(<]).*$",
    r"^continue(?:\b|[\s(<]).*$",
    r"^\u7ee7\u7eed.*$",
    r"^\u786e\u5b9a.*$",
    r"^\u662f.*$",
)
INSTALLER_DESTINATION_LABEL_PATTERNS = (
    r"^destination(?: folder| location)?$",
    r"^select destination location$",
    r"^install folder$",
    r"^installation folder$",
    r"^install(?:ation)? (?:folder|path|location)$",
    r"^folder$",
    r"^path$",
    r"^location$",
    r"^target(?: folder| path| location)?$",
    r"^\u5b89\u88c5\u6587\u4ef6\u5939$",
    r"^\u5b89\u88c5(?:\u8def\u5f84|\u4f4d\u7f6e|\u76ee\u5f55)$",
    r"^\u76ee\u6807(?:\u6587\u4ef6\u5939|\u8def\u5f84|\u4f4d\u7f6e)$",
    r"^\u6587\u4ef6\u5939$",
    r"^\u8def\u5f84$",
    r"^\u4f4d\u7f6e$",
)
INSTALLER_BROWSE_BUTTON_PATTERNS = (
    r"^browse(?:\.{3}|\u2026)?$",
    r"^\u6d4f\u89c8(?:\.{3}|\u2026)?$",
    r"^change(?:\.{3}|\u2026)?$",
    r"^\u66f4\u6539(?:\.{3}|\u2026)?$",
)
INSTALLER_FINISH_BUTTON_PATTERNS = (
    r"^finish(?:\b|[\s(<]).*$",
    r"^done(?:\b|[\s(<]).*$",
    r"^\u5b8c\u6210.*$",
)
INSTALLER_CLOSE_BUTTON_PATTERNS = (
    r"^close(?:\b|[\s(<]).*$",
    r"^exit(?:\b|[\s(<]).*$",
    r"^\u5173\u95ed.*$",
    r"^\u7ed3\u675f.*$",
)
INSTALLER_MODAL_TITLE_PATTERNS = (
    r"^warning$",
    r"^confirm(?:ation)?$",
    r"^question$",
    r"^\u8b66\u544a$",
    r"^\u786e\u8ba4$",
    r"^\u63d0\u793a$",
    r"^\u95ee\u9898$",
)
INSTALLER_COMPLETION_TEXT_PATTERNS = (
    r".*\bcompleted\b.*",
    r".*\bsuccessfully\b.*",
    r".*\binstalled\b.*",
    r".*\bfinished\b.*",
    r".*\bsetup was completed\b.*",
    r".*\binstallation complete\b.*",
    r".*\binstallation successful\b.*",
    r".*\u5b89\u88c5\u5b8c\u6210.*",
    r".*\u5df2\u5b8c\u6210.*",
    r".*\u5df2\u6210\u529f.*",
    r".*\u5b8c\u6210\u5b89\u88c5.*",
)


def _program_files_candidates(*parts: str) -> tuple[Path, ...]:
    candidates: list[Path] = []
    roots = [
        os.environ.get("ProgramFiles", r"C:\Program Files"),
        os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)"),
        os.environ.get("LOCALAPPDATA", ""),
    ]
    seen: set[Path] = set()
    for root in roots:
        if not root:
            continue
        candidate = Path(root, *parts)
        if candidate in seen:
            continue
        seen.add(candidate)
        candidates.append(candidate)
    return tuple(candidates)


@dataclass(frozen=True)
class SoftwareInstallProfile:
    software: str
    installer_path: Path
    shortcut_path: Path
    install_root: Path
    display_name_patterns: tuple[str, ...]
    executable_names: tuple[str, ...]
    target_candidates: tuple[Path, ...]
    silent_argument_sets: tuple[tuple[str, ...], ...]
    prefer_silent: bool = False
    silent_directory_mode: str | None = None
    msi_directory_properties: tuple[str, ...] = ()


@dataclass(frozen=True)
class InitializationRunPaths:
    run_root: Path
    log_path: Path
    summary_path: Path
    events_path: Path


class InitializationRecorder:
    def __init__(self, paths: InitializationRunPaths):
        self.paths = paths
        self._events: list[dict[str, object]] = []

    def record(self, stage: str, status: str, **fields: object) -> None:
        event = {
            "timestamp": datetime.now().isoformat(timespec="seconds"),
            "stage": stage,
            "status": status,
        }
        event.update(fields)
        self._events.append(event)
        with self.paths.events_path.open("a", encoding="utf-8") as handle:
            handle.write(json.dumps(event, ensure_ascii=True) + "\n")
        self._write_summary()
        print(self._format_console_event(event), flush=True)

    def _format_console_event(self, event: dict[str, object]) -> str:
        pieces = [f"[init][{event['status']}]", str(event["stage"])]
        if event.get("software"):
            pieces.append(f"software={event['software']}")
        if event.get("executable_path"):
            pieces.append(f"exe={event['executable_path']}")
        if event.get("log_path"):
            pieces.append(f"log={event['log_path']}")
        if event.get("summary_path"):
            pieces.append(f"summary={event['summary_path']}")
        if event.get("error"):
            pieces.append(f"error={event['error']}")
        return " ".join(pieces)

    def _latest_initialization_status(self) -> str:
        for event in reversed(self._events):
            if event["stage"] == "initialization":
                return str(event["status"])
        return "running"

    def _write_summary(self) -> None:
        software_rows: dict[str, dict[str, str]] = {}
        timeline: list[str] = []
        for event in self._events:
            stage = str(event["stage"])
            status = str(event["status"])
            software = str(event.get("software", "") or "")
            if software and stage in {"software", "installer", "shortcut"}:
                row = software_rows.setdefault(
                    software,
                    {
                        "software_status": "",
                        "installer_status": "",
                        "shortcut_status": "",
                        "executable_path": "",
                        "error": "",
                    },
                )
                if stage == "software":
                    row["software_status"] = status
                    if event.get("executable_path"):
                        row["executable_path"] = str(event["executable_path"])
                elif stage == "installer":
                    row["installer_status"] = status
                elif stage == "shortcut":
                    row["shortcut_status"] = status
                if event.get("error"):
                    row["error"] = str(event["error"])

            details = [str(event["timestamp"]), stage, status]
            if software:
                details.append(f"software={software}")
            if event.get("executable_path"):
                details.append(f"exe={event['executable_path']}")
            if event.get("error"):
                details.append(f"error={event['error']}")
            timeline.append("- " + " | ".join(details))

        lines = [
            "# Initialization Summary",
            "",
            f"- status: {self._latest_initialization_status()}",
            f"- run_root: {self.paths.run_root}",
            f"- progress_log: {self.paths.events_path}",
            f"- full_log: {self.paths.log_path}",
            "",
            "## Software Status",
            "",
            "| software | software_status | installer_status | shortcut_status | executable | error |",
            "| --- | --- | --- | --- | --- | --- |",
        ]
        if software_rows:
            for software in sorted(software_rows):
                row = software_rows[software]
                executable_path = row["executable_path"].replace("|", "\\|")
                error = row["error"].replace("|", "\\|")
                lines.append(
                    f"| {software} | {row['software_status']} | {row['installer_status']} | "
                    f"{row['shortcut_status']} | {executable_path} | {error} |"
                )
        else:
            lines.append("|  | not_started |  |  |  |  |")
        lines.extend([
            "",
            "## Timeline",
            "",
            *timeline,
            "",
        ])
        self.paths.summary_path.write_text("\n".join(lines), encoding="utf-8")


class ShellStartedProcess:
    def __init__(self, pid: int | None):
        self.pid = pid
        self._returncode: int | None = None
        self._handle = None
        if pid is not None:
            self._handle = ctypes.windll.kernel32.OpenProcess(
                PROCESS_QUERY_LIMITED_INFORMATION | SYNCHRONIZE,
                False,
                pid,
            )

    def poll(self) -> int | None:
        if self._returncode is not None:
            return self._returncode
        if self.pid is None:
            return None
        if not self._handle:
            completed = subprocess.run(
                ["tasklist", "/FI", f"PID eq {self.pid}", "/FO", "CSV", "/NH"],
                capture_output=True,
                text=True,
                check=False,
            )
            if str(self.pid) not in completed.stdout:
                self._returncode = 0
                return self._returncode
            return None
        exit_code = wintypes.DWORD()
        if not ctypes.windll.kernel32.GetExitCodeProcess(self._handle, ctypes.byref(exit_code)):
            return None
        if exit_code.value == STILL_ACTIVE:
            return None
        self._returncode = int(exit_code.value)
        ctypes.windll.kernel32.CloseHandle(self._handle)
        self._handle = None
        return self._returncode

    def close(self) -> None:
        if self._handle:
            ctypes.windll.kernel32.CloseHandle(self._handle)
            self._handle = None

    def __del__(self) -> None:
        self.close()


def _repo_install_root(software: str) -> Path:
    return (REPO_ROOT / "software" / software).resolve(strict=False)


def _repo_install_candidates(software: str, *relative_parts: str) -> tuple[Path, ...]:
    install_root = _repo_install_root(software)
    candidates = [
        install_root.joinpath(*relative_parts),
        install_root / "bin" / relative_parts[-1],
    ]
    unique: list[Path] = []
    seen: set[Path] = set()
    for candidate in candidates:
        candidate = candidate.resolve(strict=False)
        if candidate in seen:
            continue
        seen.add(candidate)
        unique.append(candidate)
    return tuple(unique)


def build_initialization_run_paths(artifacts_root: Path) -> InitializationRunPaths:
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    run_root = artifacts_root / f"initialize_environment_{timestamp}"
    run_root.mkdir(parents=True, exist_ok=True)
    return InitializationRunPaths(
        run_root=run_root,
        log_path=run_root / "initialize_environment.log",
        summary_path=run_root / "initialize_summary.md",
        events_path=run_root / "initialize_progress.jsonl",
    )


def configure_logging(log_path: Path) -> None:
    handlers = [
        logging.StreamHandler(sys.stdout),
        logging.FileHandler(log_path, encoding="utf-8"),
    ]
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s | %(levelname)s | %(message)s",
        handlers=handlers,
        force=True,
    )


SOFTWARE_INSTALL_PROFILES = (
    SoftwareInstallProfile(
        software="avidemux",
        installer_path=INSTALLER_DIR / "Avidemux_2.8.1VC++64bits.exe",
        shortcut_path=WHITELIST_APP_DIR / "avidemux.lnk",
        install_root=_repo_install_root("avidemux_runtime"),
        display_name_patterns=("avidemux",),
        executable_names=("avidemux.exe",),
        target_candidates=(
            *_repo_install_candidates("avidemux_runtime", "avidemux.exe"),
            *_repo_install_candidates("avidemux", "avidemux.exe"),
            *_program_files_candidates("Avidemux 2.8 VC++ 64bits", "avidemux.exe"),
        ),
        silent_argument_sets=(
            ("/S",),
            ("/quiet", "/norestart"),
            ("/passive", "/norestart"),
        ),
        prefer_silent=True,
        silent_directory_mode="nsis",
    ),
    SoftwareInstallProfile(
        software="handbrake",
        installer_path=INSTALLER_DIR / "HandBrake-1.10.2-x86_64-Win_GUI.exe",
        shortcut_path=WHITELIST_APP_DIR / "HandBrake.lnk",
        install_root=_repo_install_root("handbrake"),
        display_name_patterns=("handbrake",),
        executable_names=("HandBrake.exe",),
        target_candidates=(
            *_repo_install_candidates("handbrake", "HandBrake.exe"),
            *_program_files_candidates("HandBrake", "HandBrake.exe"),
        ),
        silent_argument_sets=(( "/S",),),
        prefer_silent=True,
        silent_directory_mode="nsis",
    ),
    SoftwareInstallProfile(
        software="kdenlive",
        installer_path=INSTALLER_DIR / "kdenlive-25.08.3.exe",
        shortcut_path=WHITELIST_APP_DIR / "Kdenlive.lnk",
        install_root=_repo_install_root("kdenlive"),
        display_name_patterns=("kdenlive",),
        executable_names=("kdenlive.exe",),
        target_candidates=(
            *_repo_install_candidates("kdenlive", "kdenlive.exe"),
            *_program_files_candidates("kdenlive", "bin", "kdenlive.exe"),
        ),
        silent_argument_sets=(
            ("/S",),
            ("/quiet", "/norestart"),
        ),
        prefer_silent=True,
        silent_directory_mode="nsis",
    ),
    SoftwareInstallProfile(
        software="shotcut",
        installer_path=INSTALLER_DIR / "shotcut-win64-25.10.31.exe",
        shortcut_path=WHITELIST_APP_DIR / "shotcut.lnk",
        install_root=_repo_install_root("shotcut"),
        display_name_patterns=("shotcut",),
        executable_names=("shotcut.exe",),
        target_candidates=(
            *_repo_install_candidates("shotcut", "shotcut.exe"),
            *_program_files_candidates("Shotcut", "shotcut.exe"),
        ),
        silent_argument_sets=(
            ("/VERYSILENT", "/SUPPRESSMSGBOXES", "/NORESTART", "/SP-"),
            ("/S",),
        ),
        prefer_silent=True,
        silent_directory_mode="inno",
    ),
    SoftwareInstallProfile(
        software="shutter_encoder",
        installer_path=INSTALLER_DIR / "Shutter Encoder 19.6 Windows 64bits.exe",
        shortcut_path=WHITELIST_APP_DIR / "Shutter Encoder.lnk",
        install_root=_repo_install_root("shutter_encoder"),
        display_name_patterns=("shutter encoder",),
        executable_names=("Shutter Encoder.exe",),
        target_candidates=(
            *_repo_install_candidates("shutter_encoder", "Shutter Encoder.exe"),
            *_program_files_candidates("Shutter Encoder", "Shutter Encoder.exe"),
        ),
        silent_argument_sets=(
            ("/VERYSILENT", "/SUPPRESSMSGBOXES", "/NORESTART", "/SP-"),
            ("/S",),
        ),
        prefer_silent=True,
        silent_directory_mode="inno",
    ),
    SoftwareInstallProfile(
        software="blender",
        installer_path=INSTALLER_DIR / "blender-4.5.5-windows-x64.msi",
        shortcut_path=WHITELIST_APP_DIR / "blender.lnk",
        install_root=_repo_install_root("blender"),
        display_name_patterns=("blender",),
        executable_names=("blender.exe",),
        target_candidates=(
            *_repo_install_candidates("blender", "blender.exe"),
            *_program_files_candidates("Blender Foundation", "Blender 4.5", "blender.exe"),
            *_program_files_candidates("Blender Foundation", "Blender", "blender.exe"),
        ),
        silent_argument_sets=((),),
        msi_directory_properties=("INSTALLDIR", "INSTALLFOLDER", "TARGETDIR"),
    ),
)


def _query_registry_value(key, value_name: str) -> str:
    try:
        value, _ = winreg.QueryValueEx(key, value_name)
    except OSError:
        return ""
    return str(value).strip()


def _display_icon_to_path(raw_value: str) -> Path | None:
    if not raw_value:
        return None
    normalized = raw_value.strip().strip('"')
    match = re.match(r"(?i)^([^,]+?\.exe)", normalized)
    if not match:
        return None
    candidate = Path(match.group(1).strip().strip('"')).resolve(strict=False)
    return candidate if candidate.exists() else None


def _search_install_location(install_location: str, executable_names: tuple[str, ...]) -> Path | None:
    if not install_location:
        return None
    root = Path(install_location).resolve(strict=False)
    if not root.exists():
        return None
    for executable_name in executable_names:
        for candidate in (root / executable_name, root / "bin" / executable_name):
            if candidate.exists():
                return candidate
        try:
            matches = list(root.rglob(executable_name))
        except OSError:
            matches = []
        if matches:
            return matches[0]
    return None


def _search_registry_for_executable(profile: SoftwareInstallProfile) -> Path | None:
    wanted_patterns = tuple(pattern.casefold() for pattern in profile.display_name_patterns)
    for hive, subkey_path in UNINSTALL_REGISTRY_ROOTS:
        try:
            root = winreg.OpenKey(hive, subkey_path)
        except OSError:
            continue
        with root:
            subkey_count = winreg.QueryInfoKey(root)[0]
            for index in range(subkey_count):
                try:
                    child_name = winreg.EnumKey(root, index)
                    child_key = winreg.OpenKey(root, child_name)
                except OSError:
                    continue
                with child_key:
                    display_name = _query_registry_value(child_key, "DisplayName")
                    if not display_name:
                        continue
                    lowered_name = display_name.casefold()
                    if not any(pattern in lowered_name for pattern in wanted_patterns):
                        continue
                    icon_candidate = _display_icon_to_path(_query_registry_value(child_key, "DisplayIcon"))
                    if icon_candidate is not None:
                        return icon_candidate
                    location_candidate = _search_install_location(
                        _query_registry_value(child_key, "InstallLocation"),
                        profile.executable_names,
                    )
                    if location_candidate is not None:
                        return location_candidate
    return None


def _search_common_roots(profile: SoftwareInstallProfile) -> Path | None:
    search_roots = [
        Path(os.environ.get("ProgramFiles", r"C:\Program Files")),
        Path(os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)")),
        Path(os.environ.get("LOCALAPPDATA", "")),
    ]
    for root in search_roots:
        if not root.exists():
            continue
        for executable_name in profile.executable_names:
            try:
                for candidate in root.rglob(executable_name):
                    return candidate
            except OSError:
                continue
    return None


def _search_custom_install_root(profile: SoftwareInstallProfile) -> Path | None:
    root = profile.install_root
    if not root.exists():
        return None
    return _search_install_location(str(root), profile.executable_names)


def _missing_install_components(profile: SoftwareInstallProfile, executable_path: Path) -> tuple[str, ...]:
    if profile.software != "avidemux":
        return ()
    install_root = executable_path.parent
    required_dirs = (
        install_root / "plugins" / "demuxers",
        install_root / "plugins" / "muxers",
        install_root / "plugins" / "videoDecoders",
    )
    missing = [str(path) for path in required_dirs if not path.exists()]
    return tuple(missing)


def _is_valid_installed_executable(profile: SoftwareInstallProfile, executable_path: Path) -> bool:
    if not executable_path.exists():
        return False
    return not _missing_install_components(profile, executable_path)


def detect_installed_executable(profile: SoftwareInstallProfile) -> Path | None:
    for candidate in profile.target_candidates:
        if _is_valid_installed_executable(profile, candidate):
            return candidate
    custom_root_candidate = _search_custom_install_root(profile)
    if custom_root_candidate is not None and _is_valid_installed_executable(profile, custom_root_candidate):
        return custom_root_candidate
    registry_candidate = _search_registry_for_executable(profile)
    if registry_candidate is not None and _is_valid_installed_executable(profile, registry_candidate):
        return registry_candidate
    if profile.shortcut_path.exists():
        target = resolve_shortcut_target(profile.shortcut_path)
        if target is not None and _is_valid_installed_executable(profile, target):
            return target
    common_root_candidate = _search_common_roots(profile)
    if common_root_candidate is not None and _is_valid_installed_executable(profile, common_root_candidate):
        return common_root_candidate
    return None


def resolve_shortcut_target(shortcut_path: Path) -> Path | None:
    if not shortcut_path.exists():
        return None
    script = (
        "$shell = New-Object -ComObject WScript.Shell; "
        f"$shortcut = $shell.CreateShortcut('{powershell_quote(str(shortcut_path))}'); "
        "Write-Output $shortcut.TargetPath"
    )
    completed = subprocess.run(
        ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", script],
        capture_output=True,
        text=True,
        check=False,
    )
    if completed.returncode != 0:
        return None
    target_path = completed.stdout.strip()
    if not target_path:
        return None
    return Path(target_path).resolve(strict=False)


def powershell_quote(value: str) -> str:
    return value.replace("'", "''")


def _silent_directory_arguments(profile: SoftwareInstallProfile) -> tuple[str, ...]:
    if profile.silent_directory_mode is None:
        return ()
    install_root = str(profile.install_root)
    if profile.silent_directory_mode == "nsis":
        return (f"/D={install_root}",)
    if profile.silent_directory_mode == "inno":
        return (f"/DIR={install_root}",)
    raise AssertionError(f"Unsupported silent directory mode: {profile.silent_directory_mode}")


def _msi_directory_arguments(profile: SoftwareInstallProfile) -> tuple[str, ...]:
    if not profile.msi_directory_properties:
        return ()
    install_root = str(profile.install_root)
    return tuple(f"{property_name}={install_root}" for property_name in profile.msi_directory_properties)


def _shell_start_process(command: Sequence[str], *, cwd: Path) -> ShellStartedProcess:
    executable, *arguments = command
    argument_list = ", ".join(f"'{powershell_quote(argument)}'" for argument in arguments)
    script = (
        f"$arguments = @({argument_list}); "
        f"$process = Start-Process -FilePath '{powershell_quote(executable)}' "
        f"-ArgumentList $arguments -WorkingDirectory '{powershell_quote(str(cwd))}' -PassThru; "
        "Write-Output $process.Id"
    )
    completed = subprocess.run(
        ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", script],
        capture_output=True,
        text=True,
        check=False,
    )
    assert completed.returncode == 0, completed.stdout + completed.stderr
    lines = [line.strip() for line in completed.stdout.splitlines() if line.strip()]
    if not lines:
        logger.warning(
            "Windows Shell launched the installer but did not report a process id for command %s. "
            "Falling back to title-based window detection only.",
            command,
        )
        return ShellStartedProcess(None)
    return ShellStartedProcess(int(lines[-1]))


def launch_installer_process(command: Sequence[str], *, cwd: Path):
    try:
        return subprocess.Popen(
            command,
            cwd=str(cwd),
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
    except OSError as exc:
        if getattr(exc, "winerror", None) != 740:
            raise
        logger.info(
            "Installer requires elevation for command %s. Relaunching it through Windows Shell.",
            command,
        )
        return _shell_start_process(command, cwd=cwd)


def create_or_update_shortcut(shortcut_path: Path, target_path: Path) -> None:
    shortcut_path.parent.mkdir(parents=True, exist_ok=True)
    if shortcut_path.exists():
        shortcut_path.unlink()
    working_directory = target_path.parent
    script = (
        "$shell = New-Object -ComObject WScript.Shell; "
        f"$shortcut = $shell.CreateShortcut('{powershell_quote(str(shortcut_path))}'); "
        f"$shortcut.TargetPath = '{powershell_quote(str(target_path))}'; "
        f"$shortcut.WorkingDirectory = '{powershell_quote(str(working_directory))}'; "
        "$shortcut.Arguments = ''; "
        "$shortcut.Save()"
    )
    completed = subprocess.run(
        ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", script],
        capture_output=True,
        text=True,
        check=False,
    )
    assert completed.returncode == 0, completed.stdout + completed.stderr


def _import_pywinauto_desktop():
    try:
        from pywinauto import Desktop
    except ImportError as exc:
        raise AssertionError(
            "pywinauto is required for interactive installer automation. Run initialize_environment without --skip-python first."
        ) from exc
    return Desktop


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


def _wrapper_rect(wrapper):
    return wrapper.rectangle()


def _is_visible(wrapper) -> bool:
    try:
        return bool(wrapper.is_visible())
    except Exception:
        return True


def _is_enabled(wrapper) -> bool:
    try:
        return bool(wrapper.is_enabled())
    except Exception:
        return True


def _process_id(wrapper) -> int | None:
    try:
        process_id = getattr(wrapper.element_info, "process_id", None)
    except Exception:
        return None
    return int(process_id) if process_id else None


def _iter_controls(root):
    yield root
    try:
        for child in root.descendants():
            yield child
    except Exception:
        return


def _compiled_patterns(patterns: Sequence[str]) -> tuple[re.Pattern[str], ...]:
    return tuple(re.compile(pattern, re.IGNORECASE) for pattern in patterns)


def _matches_patterns(text: str, patterns: Sequence[re.Pattern[str]]) -> bool:
    return any(pattern.search(text) for pattern in patterns)


def _vertical_overlap(a, b) -> int:
    return max(0, min(a.bottom, b.bottom) - max(a.top, b.top))


def _installer_title_tokens(profile: SoftwareInstallProfile) -> tuple[str, ...]:
    tokens = {profile.software.replace("_", " ").casefold()}
    for pattern in profile.display_name_patterns:
        normalized = pattern.strip().casefold()
        if normalized:
            tokens.add(normalized)
    for token in re.split(r"[^a-z0-9]+", profile.installer_path.stem.casefold()):
        if len(token) >= 4 and not token.isdigit():
            tokens.add(token)
    return tuple(sorted(tokens))


def _list_candidate_installer_windows(
    profile: SoftwareInstallProfile,
    process: subprocess.Popen[str] | subprocess.Popen[bytes] | ShellStartedProcess,
):
    Desktop = _import_pywinauto_desktop()
    tokens = _installer_title_tokens(profile)
    modal_title_patterns = _compiled_patterns(INSTALLER_MODAL_TITLE_PATTERNS)
    candidates = []
    for window in Desktop(backend="uia").windows():
        if not _is_visible(window):
            continue
        title = _safe_window_text(window)
        lowered_title = title.casefold()
        compact_title = re.sub(r"\s+", "", lowered_title)
        window_pid = _process_id(window)
        token_match = bool(lowered_title) and any(
            token in lowered_title or token.replace(" ", "") in compact_title
            for token in tokens
        )
        pid_match = process.pid is not None and window_pid == process.pid
        if not token_match and not pid_match:
            continue
        modal_title_match = bool(title) and _matches_patterns(title, modal_title_patterns)
        try:
            rect = _wrapper_rect(window)
        except Exception:
            continue
        area = max(1, rect.width() * rect.height())
        candidates.append(
            (
                0 if modal_title_match else 1,
                0 if token_match else 1,
                0 if pid_match else 1,
                -area,
                rect.top,
                rect.left,
                window,
            )
        )
    candidates.sort()
    return [item[-1] for item in candidates]


def _find_matching_controls(root, patterns: Sequence[str], *, control_types: Sequence[str]) -> list:
    compiled = _compiled_patterns(patterns)
    allowed_types = {value.lower() for value in control_types}
    candidates = []
    for wrapper in _iter_controls(root):
        title = _safe_window_text(wrapper)
        if not title or not _matches_patterns(title, compiled):
            continue
        control_type = _control_type(wrapper).lower()
        if allowed_types and control_type not in allowed_types:
            continue
        try:
            rect = _wrapper_rect(wrapper)
        except Exception:
            continue
        candidates.append((rect.top, rect.left, wrapper))
    candidates.sort(key=lambda item: (item[0], item[1]))
    return [item[2] for item in candidates]


def _read_wrapper_value(wrapper) -> str:
    readers = (
        lambda: wrapper.get_value(),
        lambda: wrapper.iface_value.CurrentValue,
        lambda: wrapper.window_text(),
        lambda: wrapper.texts()[0] if wrapper.texts() else "",
    )
    for reader in readers:
        try:
            value = reader()
        except Exception:
            continue
        if isinstance(value, str):
            return value.strip()
    return ""


def _path_like_text(value: str) -> bool:
    lowered = value.strip().casefold()
    return ":\\" in lowered or lowered.startswith("\\\\") or lowered.endswith(("\\", "/")) or "program files" in lowered


def _candidate_destination_edits(window) -> list:
    candidates = []
    browse_present = bool(
        _find_matching_controls(
            window,
            INSTALLER_BROWSE_BUTTON_PATTERNS,
            control_types=("Button", "Custom", "Pane"),
        )
    )
    for wrapper in _iter_controls(window):
        control_type = _control_type(wrapper).lower()
        if control_type not in {"edit", "combobox"}:
            continue
        if not _is_enabled(wrapper):
            continue
        current_value = _read_wrapper_value(wrapper)
        if not current_value and not browse_present:
            continue
        if current_value and not _path_like_text(current_value) and not browse_present:
            continue
        try:
            rect = _wrapper_rect(wrapper)
        except Exception:
            continue
        width = rect.right - rect.left
        candidates.append((-width, rect.top, rect.left, wrapper))
    candidates.sort(key=lambda item: (item[0], item[1], item[2]))
    return [item[3] for item in candidates]


def _is_caption_close_button(root, wrapper) -> bool:
    title = _safe_window_text(wrapper).casefold()
    if title not in {"close", "exit", "\u5173\u95ed", "\u7ed3\u675f"}:
        return False
    try:
        root_rect = _wrapper_rect(root)
        rect = _wrapper_rect(wrapper)
    except Exception:
        return False
    return rect.top <= root_rect.top + 80 and rect.right >= root_rect.right - 80


def _get_toggle_state(wrapper) -> bool | None:
    try:
        return bool(wrapper.get_toggle_state())
    except Exception:
        pass
    try:
        return bool(wrapper.iface_toggle.CurrentToggleState)
    except Exception:
        pass
    try:
        return bool(wrapper.iface_selection_item.CurrentIsSelected)
    except Exception:
        return None


def _click_wrapper(wrapper) -> None:
    wrapper.click_input()
    time.sleep(0.4)


def _set_toggle_checked(toggle_wrapper) -> bool:
    current_state = _get_toggle_state(toggle_wrapper)
    if current_state is True:
        return False
    _click_wrapper(toggle_wrapper)
    current_state = _get_toggle_state(toggle_wrapper)
    if current_state is False:
        raise AssertionError("Installer agreement control did not change to the checked state.")
    return True


def _toggle_control_priority(wrapper) -> tuple[int, int]:
    control_type = _control_type(wrapper).lower()
    if control_type == "checkbox":
        return (0, 0 if _get_toggle_state(wrapper) is not None else 1)
    if control_type == "radiobutton":
        return (1, 0 if _get_toggle_state(wrapper) is not None else 1)
    if control_type in {"custom", "listitem", "pane"}:
        return (2, 0 if _get_toggle_state(wrapper) is not None else 1)
    if control_type == "button":
        return (3, 0 if _get_toggle_state(wrapper) is not None else 1)
    return (4, 1)


def _toggle_hitbox_requires_coordinate_click(wrapper) -> bool:
    control_type = _control_type(wrapper).lower()
    if control_type in {"checkbox", "radiobutton"}:
        return False
    return _get_toggle_state(wrapper) is None


def _click_installer_point(window, screen_x: int, screen_y: int, *, description: str) -> None:
    top_level = window.top_level_parent() if hasattr(window, "top_level_parent") else window
    bring_window_to_front(top_level, keep_topmost=False)
    top_level_rect = _wrapper_rect(top_level)
    rel_x = max(1, min(screen_x - top_level_rect.left, max(1, top_level_rect.right - top_level_rect.left - 2)))
    rel_y = max(1, min(screen_y - top_level_rect.top, max(1, top_level_rect.bottom - top_level_rect.top - 2)))
    logger.info(
        "Clicking installer %s at screen coordinates x=%d y=%d.",
        description,
        screen_x,
        screen_y,
    )
    try:
        top_level.click_input(coords=(rel_x, rel_y))
    except Exception:
        try:
            import pyautogui
        except ImportError as exc:
            raise AssertionError(
                "pyautogui is required for installer coordinate fallback. Run initialize_environment without --skip-python first."
            ) from exc
        pyautogui.FAILSAFE = False
        pyautogui.PAUSE = 0.05
        pyautogui.moveTo(screen_x, screen_y, duration=0.05)
        pyautogui.click(screen_x, screen_y)
    time.sleep(0.4)


def _click_label_toggle_hitbox(window, label_wrapper) -> None:
    top_level = window.top_level_parent() if hasattr(window, "top_level_parent") else window
    top_level_rect = _wrapper_rect(top_level)
    label_rect = _wrapper_rect(label_wrapper)
    label_height = max(1, label_rect.bottom - label_rect.top)
    horizontal_offset = max(18, min(42, int(label_height * 1.2)))
    screen_x = max(top_level_rect.left + 8, label_rect.left - horizontal_offset)
    screen_y = label_rect.top + label_height // 2
    _click_installer_point(window, screen_x, screen_y, description="agreement hitbox")


def _find_row_toggle(root, label_wrapper):
    label_rect = _wrapper_rect(label_wrapper)
    best_match = None
    best_score = None
    for candidate in _iter_controls(root):
        control_type = _control_type(candidate).lower()
        if control_type not in {"checkbox", "radiobutton", "button", "custom", "listitem", "pane"}:
            continue
        try:
            rect = _wrapper_rect(candidate)
        except Exception:
            continue
        overlap = _vertical_overlap(label_rect, rect)
        if overlap <= 0:
            continue
        width = max(1, rect.right - rect.left)
        height = max(1, rect.bottom - rect.top)
        label_height = max(1, label_rect.bottom - label_rect.top)
        label_width = max(1, label_rect.right - label_rect.left)
        large_generic_control = (
            control_type in {"button", "pane", "custom", "listitem"}
            and width > max(48, label_height * 3)
            and height > max(28, label_height + 6)
        )
        if large_generic_control:
            continue
        if rect.right <= label_rect.left:
            horizontal_gap = label_rect.left - rect.right
            side_priority = 0
        elif rect.left >= label_rect.right:
            horizontal_gap = rect.left - label_rect.right
            side_priority = 2
        else:
            horizontal_gap = 0
            side_priority = 1
        center_y = rect.top + height // 2
        label_center_y = label_rect.top + label_height // 2
        score = (
            *_toggle_control_priority(candidate),
            side_priority,
            horizontal_gap,
            abs(center_y - label_center_y),
            -overlap,
            0 if width <= max(40, label_width // 3) else 1,
            abs(rect.left - label_rect.left),
        )
        if best_score is None or score < best_score:
            best_score = score
            best_match = candidate
    return best_match


def _accept_license_terms(window) -> bool:
    direct_controls = _find_matching_controls(
        window,
        INSTALLER_ACCEPT_LABEL_PATTERNS,
        control_types=("CheckBox", "RadioButton", "Button", "Custom", "ListItem", "Pane"),
    )
    for control in sorted(direct_controls, key=_toggle_control_priority):
        if not _is_enabled(control):
            continue
        control_type = _control_type(control).lower()
        logger.info(
            "Selecting installer agreement control: %s type=%s",
            _safe_window_text(control) or "<untitled>",
            control_type or "<unknown>",
        )
        if _toggle_hitbox_requires_coordinate_click(control):
            _click_installer_point(
                window,
                _wrapper_rect(control).left + max(6, (_wrapper_rect(control).right - _wrapper_rect(control).left) // 4),
                _wrapper_rect(control).top + max(4, (_wrapper_rect(control).bottom - _wrapper_rect(control).top) // 2),
                description="agreement control fallback",
            )
        else:
            _set_toggle_checked(control)
        return True

    label_controls = _find_matching_controls(
        window,
        INSTALLER_ACCEPT_LABEL_PATTERNS,
        control_types=("Text", "Pane", "Group"),
    )
    for label in label_controls:
        toggle = _find_row_toggle(window, label)
        if toggle is not None and _is_enabled(toggle):
            logger.info(
                "Selecting installer agreement row next to label: %s type=%s",
                _safe_window_text(label),
                _control_type(toggle).lower() or "<unknown>",
            )
            if _toggle_hitbox_requires_coordinate_click(toggle):
                _click_label_toggle_hitbox(window, label)
            else:
                _set_toggle_checked(toggle)
            return True
        logger.info("Clicking installer agreement row hitbox next to label: %s", _safe_window_text(label))
        _click_label_toggle_hitbox(window, label)
        return True
    return False


def _ensure_install_destination(window, profile: SoftwareInstallProfile) -> bool:
    desired_root = profile.install_root
    desired_root.mkdir(parents=True, exist_ok=True)
    desired_text = str(desired_root)

    try:
        destination_edit = find_labeled_edit(window, INSTALLER_DESTINATION_LABEL_PATTERNS)
    except UiAutomationError:
        destination_edit = None

    if destination_edit is None:
        candidates = _candidate_destination_edits(window)
        destination_edit = candidates[0] if candidates else None

    if destination_edit is None:
        return False

    current_value = _read_wrapper_value(destination_edit)
    if current_value.casefold() == desired_text.casefold():
        return False

    logger.info("Setting installer destination for %s to: %s", profile.software, desired_text)
    set_edit_text_value(destination_edit, desired_text)
    return True


def _click_first_enabled_button(window, patterns: Sequence[str], *, action_name: str) -> bool:
    for control in _find_matching_controls(window, patterns, control_types=("Button", "Custom", "Pane")):
        if not _is_enabled(control):
            continue
        if _is_caption_close_button(window, control):
            continue
        logger.info("Clicking installer %s control: %s", action_name, _safe_window_text(control) or "<untitled>")
        _click_wrapper(control)
        return True
    return False


def _advance_installer_buttons(window) -> bool:
    if _click_first_enabled_button(window, INSTALLER_NEXT_BUTTON_PATTERNS, action_name="next"):
        return True
    if _click_first_enabled_button(window, INSTALLER_INSTALL_BUTTON_PATTERNS, action_name="install"):
        return True
    if _click_first_enabled_button(window, INSTALLER_CONFIRM_BUTTON_PATTERNS, action_name="confirm"):
        return True
    if _click_first_enabled_button(window, INSTALLER_FINISH_BUTTON_PATTERNS, action_name="finish"):
        return True
    return False


def _window_has_completion_text(window) -> bool:
    compiled = _compiled_patterns(INSTALLER_COMPLETION_TEXT_PATTERNS)
    for wrapper in _iter_controls(window):
        text = _safe_window_text(wrapper)
        if text and _matches_patterns(text, compiled):
            return True
    return False


def _advance_installer_window(window, profile: SoftwareInstallProfile) -> bool:
    title = _safe_window_text(window) or "<untitled>"
    logger.info("Processing installer window: %s", title)
    try:
        bring_window_to_front(window, keep_topmost=False)
    except Exception:
        pass

    if _advance_installer_buttons(window):
        return True
    if _ensure_install_destination(window, profile):
        return True
    if _accept_license_terms(window):
        time.sleep(0.3)
        if _advance_installer_buttons(window):
            return True
        return True
    if _window_has_completion_text(window):
        if _click_first_enabled_button(window, INSTALLER_CLOSE_BUTTON_PATTERNS, action_name="close"):
            return True
    return False


def _wait_for_installer_completion(
    profile: SoftwareInstallProfile,
    process: subprocess.Popen[str] | subprocess.Popen[bytes] | ShellStartedProcess,
    *,
    timeout_seconds: float,
) -> Path:
    deadline = time.monotonic() + timeout_seconds
    exit_code: int | None = None
    exit_deadline: float | None = None
    observed_titles: list[str] = []

    while time.monotonic() < deadline:
        installed_executable = detect_installed_executable(profile)
        if installed_executable is not None:
            return installed_executable

        windows = _list_candidate_installer_windows(profile, process)
        if windows:
            observed_titles = [_safe_window_text(window) for window in windows if _safe_window_text(window)]
            progressed = False
            for window in windows:
                progressed = _advance_installer_window(window, profile) or progressed
            time.sleep(INSTALLER_ACTION_SETTLE_SECONDS if progressed else INSTALLER_POLL_SECONDS)
            continue

        if exit_code is None:
            exit_code = process.poll()
            if exit_code is not None:
                exit_deadline = time.monotonic() + INSTALLER_WINDOW_POST_EXIT_GRACE_SECONDS

        if exit_deadline is not None:
            if time.monotonic() < exit_deadline:
                time.sleep(INSTALLER_POLL_SECONDS)
                continue
            installed_executable = detect_installed_executable(profile)
            if installed_executable is not None:
                return installed_executable
            if exit_code not in ACCEPTABLE_INSTALL_EXIT_CODES:
                raise AssertionError(
                    f"Installer exited with code {exit_code} for {profile.software}. "
                    f"Last installer windows: {observed_titles[:5]}"
                )
            break

        time.sleep(INSTALLER_POLL_SECONDS)

    installed_executable = detect_installed_executable(profile)
    if installed_executable is not None:
        return installed_executable
    if process.poll() is None:
        logger.warning("Installer automation timed out for %s. Killing the process tree.", profile.software)
        terminate_process_tree(process)
    raise AssertionError(
        f"Could not install or detect the executable for {profile.software}. "
        f"Last installer windows: {observed_titles[:5]}"
    )


def terminate_process_tree(
    process: subprocess.Popen[bytes] | subprocess.Popen[str] | ShellStartedProcess,
) -> None:
    if process.pid is None:
        if hasattr(process, "close"):
            process.close()
        return
    subprocess.run(
        ["taskkill", "/PID", str(process.pid), "/T", "/F"],
        capture_output=True,
        text=True,
        check=False,
    )
    if hasattr(process, "close"):
        process.close()


def installer_commands(profile: SoftwareInstallProfile) -> list[list[str]]:
    installer = str(profile.installer_path)
    if profile.installer_path.suffix.lower() == ".msi":
        msi_dir_args = list(_msi_directory_arguments(profile))
        return [
            ["msiexec.exe", "/i", installer, *msi_dir_args],
            ["msiexec.exe", "/i", installer, "/qn", "/norestart", *msi_dir_args],
        ]
    silent_dir_args = _silent_directory_arguments(profile)
    silent_commands = []
    for argument_set in profile.silent_argument_sets:
        command = [installer, *argument_set]
        if silent_dir_args:
            command.extend(silent_dir_args)
        silent_commands.append(command)
    interactive_commands = [[installer]]
    commands = [*silent_commands, *interactive_commands] if profile.prefer_silent else [*interactive_commands, *silent_commands]
    return commands


def run_installer(
    profile: SoftwareInstallProfile,
    *,
    timeout_seconds: float = COMMON_INSTALL_TIMEOUT_SECONDS,
    recorder: InitializationRecorder | None = None,
) -> Path:
    assert profile.installer_path.exists(), f"Installer was not found: {profile.installer_path}"
    commands = installer_commands(profile)
    logger.info("Installing %s from %s", profile.software, profile.installer_path)
    if recorder is not None:
        recorder.record("installer", "started", software=profile.software, installer_path=str(profile.installer_path))
    failures: list[str] = []
    for command in commands:
        logger.info("Trying installer command: %s", command)
        process = launch_installer_process(command, cwd=profile.installer_path.parent)
        try:
            installed_executable = _wait_for_installer_completion(
                profile,
                process,
                timeout_seconds=timeout_seconds,
            )
            if hasattr(process, "close"):
                process.close()
            logger.info("Detected installed executable for %s: %s", profile.software, installed_executable)
            if recorder is not None:
                recorder.record(
                    "installer",
                    "completed",
                    software=profile.software,
                    executable_path=str(installed_executable),
                )
            return installed_executable
        except (subprocess.TimeoutExpired, AssertionError, UiAutomationError) as exc:
            logger.warning(
                "Installer command failed for %s with %s. Trying the next command if available.",
                profile.software,
                exc,
            )
            failures.append(f"{command!r}: {exc}")
            terminate_process_tree(process)
            continue
    if recorder is not None:
        recorder.record("installer", "failed", software=profile.software, error=" | ".join(failures))
    raise AssertionError(
        f"Could not install or detect the executable for {profile.software}. "
        f"Attempt summary: {' | '.join(failures)}"
    )


def ensure_software_ready(
    profile: SoftwareInstallProfile,
    *,
    recorder: InitializationRecorder | None = None,
) -> Path:
    logger.info("Checking whether %s is installed.", profile.software)
    custom_root_candidate = _search_custom_install_root(profile)
    if custom_root_candidate is not None:
        missing_components = _missing_install_components(profile, custom_root_candidate)
        if missing_components:
            logger.warning(
                "%s installation looks incomplete at %s. Missing components: %s",
                profile.software,
                custom_root_candidate,
                ", ".join(missing_components),
            )
    installed_executable = detect_installed_executable(profile)
    if installed_executable is None:
        logger.info("%s is not installed. Running the bundled installer.", profile.software)
        if recorder is not None:
            recorder.record("software", "installing", software=profile.software, installer_path=str(profile.installer_path))
        installed_executable = run_installer(profile, recorder=recorder)
    else:
        logger.info("%s is already installed: %s", profile.software, installed_executable)
        if recorder is not None:
            recorder.record(
                "software",
                "already_installed",
                software=profile.software,
                executable_path=str(installed_executable),
            )
    create_or_update_shortcut(profile.shortcut_path, installed_executable)
    logger.info("Shortcut is ready: %s -> %s", profile.shortcut_path, installed_executable)
    if recorder is not None:
        recorder.record(
            "shortcut",
            "completed",
            software=profile.software,
            executable_path=str(installed_executable),
            shortcut_path=str(profile.shortcut_path),
        )
    return installed_executable


def missing_python_packages() -> list[str]:
    missing: list[str] = []
    for package_name in PYTHON_PACKAGE_NAMES:
        if importlib.util.find_spec(package_name) is None:
            missing.append(package_name)
    return missing


def ensure_pip_available() -> None:
    completed = subprocess.run(
        [sys.executable, "-m", "pip", "--version"],
        capture_output=True,
        text=True,
        check=False,
    )
    if completed.returncode == 0:
        return
    logger.info("pip is not available. Bootstrapping it with ensurepip.")
    subprocess.run([sys.executable, "-m", "ensurepip", "--upgrade"], check=True)


def ensure_python_requirements() -> None:
    missing = missing_python_packages()
    if not missing:
        logger.info("All required Python packages are already installed.")
        return
    logger.info("Missing Python packages detected: %s", ", ".join(missing))
    ensure_pip_available()
    subprocess.run([sys.executable, "-m", "pip", "install", "-r", str(REPO_ROOT / "requirements.txt")], check=True)
    missing_after_install = missing_python_packages()
    assert not missing_after_install, f"Some Python packages are still missing after installation: {missing_after_install}"


def ai_turbo_hint() -> tuple[bool, Path]:
    shortcut_path = WHITELIST_APP_DIR / "AI Turbo Engine.lnk"
    if shortcut_path.exists():
        logger.info("Detected AI Turbo Engine shortcut: %s", shortcut_path)
        return True, shortcut_path
    logger.warning(
        "AI Turbo Engine shortcut is missing: %s. "
        "The full pipeline turbo pass will still require AI Turbo Engine to be installed later.",
        shortcut_path,
    )
    return False, shortcut_path


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Initialize the local Windows automation environment.")
    parser.add_argument(
        "--software",
        nargs="*",
        default=[profile.software for profile in SOFTWARE_INSTALL_PROFILES],
        choices=[profile.software for profile in SOFTWARE_INSTALL_PROFILES],
        help="Optional subset of software to verify and install.",
    )
    parser.add_argument(
        "--skip-python",
        action="store_true",
        help="Skip the Python requirements check.",
    )
    parser.add_argument(
        "--artifacts-root",
        default=str(REPO_ROOT / "results" / "initialize_environment_runs"),
        help="Directory where initialization summaries and logs will be written.",
    )
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    artifacts_root = Path(args.artifacts_root).resolve(strict=False)
    run_paths = build_initialization_run_paths(artifacts_root)
    configure_logging(run_paths.log_path)
    recorder = InitializationRecorder(run_paths)
    selected = set(args.software)

    logger.info("Starting environment initialization.")
    recorder.record(
        "initialization",
        "started",
        run_root=str(run_paths.run_root),
        log_path=str(run_paths.log_path),
        summary_path=str(run_paths.summary_path),
        selected_software=",".join(args.software),
    )
    try:
        if not args.skip_python:
            recorder.record("python_requirements", "started")
            ensure_python_requirements()
            recorder.record("python_requirements", "completed")
        else:
            recorder.record("python_requirements", "skipped")

        ai_turbo_present, ai_turbo_shortcut_path = ai_turbo_hint()
        recorder.record(
            "ai_turbo_engine",
            "detected" if ai_turbo_present else "missing",
            shortcut_path=str(ai_turbo_shortcut_path),
        )

        resolved: dict[str, Path] = {}
        for profile in SOFTWARE_INSTALL_PROFILES:
            if profile.software not in selected:
                continue
            recorder.record("software", "started", software=profile.software)
            try:
                resolved[profile.software] = ensure_software_ready(profile, recorder=recorder)
            except Exception as exc:
                recorder.record("software", "failed", software=profile.software, error=str(exc))
                raise
            recorder.record(
                "software",
                "completed",
                software=profile.software,
                executable_path=str(resolved[profile.software]),
            )

        recorder.record("initialization", "completed", software_count=len(resolved))
        print("initialization_status=ok")
        print(f"initialization_run_root={run_paths.run_root}")
        print(f"initialization_summary={run_paths.summary_path}")
        print(f"initialization_log={run_paths.log_path}")
        for software, executable_path in resolved.items():
            print(f"{software}_executable={executable_path}")
        return 0
    except Exception as exc:
        recorder.record("initialization", "failed", error=str(exc))
        print(f"initialization_status=failed")
        print(f"initialization_run_root={run_paths.run_root}")
        print(f"initialization_summary={run_paths.summary_path}")
        print(f"initialization_log={run_paths.log_path}")
        raise


if __name__ == "__main__":
    raise SystemExit(main())
