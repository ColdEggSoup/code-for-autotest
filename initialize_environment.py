from __future__ import annotations

import argparse
import importlib.util
import logging
import os
from dataclasses import dataclass
from pathlib import Path
import re
import subprocess
import sys
import time
import winreg

from workspace_runtime import configure_workspace_runtime

configure_workspace_runtime()

from automation_components import REPO_ROOT, WHITELIST_APP_DIR


logger = logging.getLogger(__name__)

INSTALLER_DIR = WHITELIST_APP_DIR / "app"
PYTHON_PACKAGE_NAMES = ("psutil", "openpyxl", "pywinauto", "pyautogui")
ACCEPTABLE_INSTALL_EXIT_CODES = {0, 1641, 3010}
COMMON_INSTALL_TIMEOUT_SECONDS = 900.0
UNINSTALL_REGISTRY_ROOTS = (
    (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"),
    (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"),
    (winreg.HKEY_CURRENT_USER, r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"),
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
    display_name_patterns: tuple[str, ...]
    executable_names: tuple[str, ...]
    target_candidates: tuple[Path, ...]
    silent_argument_sets: tuple[tuple[str, ...], ...]


SOFTWARE_INSTALL_PROFILES = (
    SoftwareInstallProfile(
        software="avidemux",
        installer_path=INSTALLER_DIR / "Avidemux_2.8.1VC++64bits.exe",
        shortcut_path=WHITELIST_APP_DIR / "avidemux.lnk",
        display_name_patterns=("avidemux",),
        executable_names=("avidemux.exe",),
        target_candidates=_program_files_candidates("Avidemux 2.8 VC++ 64bits", "avidemux.exe"),
        silent_argument_sets=(
            ("/quiet", "/norestart"),
            ("/passive", "/norestart"),
            ("/S",),
        ),
    ),
    SoftwareInstallProfile(
        software="handbrake",
        installer_path=INSTALLER_DIR / "HandBrake-1.10.2-x86_64-Win_GUI.exe",
        shortcut_path=WHITELIST_APP_DIR / "HandBrake.lnk",
        display_name_patterns=("handbrake",),
        executable_names=("HandBrake.exe",),
        target_candidates=_program_files_candidates("HandBrake", "HandBrake.exe"),
        silent_argument_sets=(( "/S",),),
    ),
    SoftwareInstallProfile(
        software="kdenlive",
        installer_path=INSTALLER_DIR / "kdenlive-25.08.3.exe",
        shortcut_path=WHITELIST_APP_DIR / "Kdenlive.lnk",
        display_name_patterns=("kdenlive",),
        executable_names=("kdenlive.exe",),
        target_candidates=_program_files_candidates("kdenlive", "bin", "kdenlive.exe"),
        silent_argument_sets=(
            ("/S",),
            ("/quiet", "/norestart"),
        ),
    ),
    SoftwareInstallProfile(
        software="shotcut",
        installer_path=INSTALLER_DIR / "shotcut-win64-25.10.31.exe",
        shortcut_path=WHITELIST_APP_DIR / "shotcut.lnk",
        display_name_patterns=("shotcut",),
        executable_names=("shotcut.exe",),
        target_candidates=_program_files_candidates("Shotcut", "shotcut.exe"),
        silent_argument_sets=(
            ("/VERYSILENT", "/SUPPRESSMSGBOXES", "/NORESTART", "/SP-"),
            ("/S",),
        ),
    ),
    SoftwareInstallProfile(
        software="shutter_encoder",
        installer_path=INSTALLER_DIR / "Shutter Encoder 19.6 Windows 64bits.exe",
        shortcut_path=WHITELIST_APP_DIR / "Shutter Encoder.lnk",
        display_name_patterns=("shutter encoder",),
        executable_names=("Shutter Encoder.exe",),
        target_candidates=_program_files_candidates("Shutter Encoder", "Shutter Encoder.exe"),
        silent_argument_sets=(
            ("/VERYSILENT", "/SUPPRESSMSGBOXES", "/NORESTART", "/SP-"),
            ("/S",),
        ),
    ),
    SoftwareInstallProfile(
        software="blender",
        installer_path=INSTALLER_DIR / "blender-4.5.5-windows-x64.msi",
        shortcut_path=WHITELIST_APP_DIR / "blender.lnk",
        display_name_patterns=("blender",),
        executable_names=("blender.exe",),
        target_candidates=(
            *_program_files_candidates("Blender Foundation", "Blender 4.5", "blender.exe"),
            *_program_files_candidates("Blender Foundation", "Blender", "blender.exe"),
        ),
        silent_argument_sets=((),),
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


def detect_installed_executable(profile: SoftwareInstallProfile) -> Path | None:
    for candidate in profile.target_candidates:
        if candidate.exists():
            return candidate
    registry_candidate = _search_registry_for_executable(profile)
    if registry_candidate is not None:
        return registry_candidate
    if profile.shortcut_path.exists():
        target = resolve_shortcut_target(profile.shortcut_path)
        if target is not None and target.exists():
            return target
    return _search_common_roots(profile)


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


def terminate_process_tree(process: subprocess.Popen[bytes] | subprocess.Popen[str]) -> None:
    subprocess.run(
        ["taskkill", "/PID", str(process.pid), "/T", "/F"],
        capture_output=True,
        text=True,
        check=False,
    )


def installer_commands(profile: SoftwareInstallProfile) -> list[list[str]]:
    installer = str(profile.installer_path)
    if profile.installer_path.suffix.lower() == ".msi":
        return [["msiexec.exe", "/i", installer, "/qn", "/norestart"]]
    return [[installer, *argument_set] for argument_set in profile.silent_argument_sets]


def run_installer(profile: SoftwareInstallProfile, *, timeout_seconds: float = COMMON_INSTALL_TIMEOUT_SECONDS) -> Path:
    assert profile.installer_path.exists(), f"Installer was not found: {profile.installer_path}"
    commands = installer_commands(profile)
    logger.info("Installing %s from %s", profile.software, profile.installer_path)
    for command in commands:
        logger.info("Trying installer command: %s", command)
        process = subprocess.Popen(
            command,
            cwd=str(profile.installer_path.parent),
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
        try:
            return_code = process.wait(timeout=timeout_seconds)
        except subprocess.TimeoutExpired:
            logger.warning("Installer command timed out for %s. Killing the process tree.", profile.software)
            terminate_process_tree(process)
            continue
        if return_code not in ACCEPTABLE_INSTALL_EXIT_CODES:
            logger.warning(
                "Installer command returned exit code %s for %s. Trying the next command if available.",
                return_code,
                profile.software,
            )
            continue
        time.sleep(5.0)
        installed_executable = detect_installed_executable(profile)
        if installed_executable is not None:
            logger.info("Detected installed executable for %s: %s", profile.software, installed_executable)
            return installed_executable
    raise AssertionError(f"Could not install or detect the executable for {profile.software}.")


def ensure_software_ready(profile: SoftwareInstallProfile) -> Path:
    logger.info("Checking whether %s is installed.", profile.software)
    installed_executable = detect_installed_executable(profile)
    if installed_executable is None:
        logger.info("%s is not installed. Running the bundled installer.", profile.software)
        installed_executable = run_installer(profile)
    else:
        logger.info("%s is already installed: %s", profile.software, installed_executable)
    create_or_update_shortcut(profile.shortcut_path, installed_executable)
    logger.info("Shortcut is ready: %s -> %s", profile.shortcut_path, installed_executable)
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


def ai_turbo_hint() -> None:
    shortcut_path = WHITELIST_APP_DIR / "AI Turbo Engine.lnk"
    if shortcut_path.exists():
        logger.info("Detected AI Turbo Engine shortcut: %s", shortcut_path)
        return
    logger.warning(
        "AI Turbo Engine shortcut is missing: %s. "
        "The full pipeline turbo pass will still require AI Turbo Engine to be installed later.",
        shortcut_path,
    )


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
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s | %(levelname)s | %(message)s",
    )
    args = parse_args(argv)
    selected = set(args.software)

    logger.info("Starting environment initialization.")
    if not args.skip_python:
        ensure_python_requirements()
    ai_turbo_hint()

    resolved: dict[str, Path] = {}
    for profile in SOFTWARE_INSTALL_PROFILES:
        if profile.software not in selected:
            continue
        resolved[profile.software] = ensure_software_ready(profile)

    print("initialization_status=ok")
    for software, executable_path in resolved.items():
        print(f"{software}_executable={executable_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
