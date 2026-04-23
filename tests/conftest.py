from __future__ import annotations

import argparse
from datetime import datetime
import json
import logging
import sys
import time
from pathlib import Path

import pytest

from automation_components import DEFAULT_AI_TURBO_TITLE_RE, REPO_ROOT
from full_test_pipeline import DEFAULT_SOFTWARE_ORDER


def _pytest_artifacts_root(config: pytest.Config) -> Path:
    raw_value = config.getoption("--pytest-artifacts-root")
    candidate = Path(raw_value)
    if not candidate.is_absolute():
        candidate = REPO_ROOT / candidate
    return candidate.resolve(strict=False)


def _resolve_repo_path(raw_value: str) -> str:
    if not raw_value:
        return ""
    candidate = Path(raw_value)
    if not candidate.is_absolute():
        candidate = REPO_ROOT / candidate
    return str(candidate.resolve(strict=False))


def _parse_software_list(raw_value: str) -> list[str]:
    values: list[str] = []
    seen: set[str] = set()
    for item in raw_value.split(","):
        normalized = item.strip()
        if not normalized or normalized in seen:
            continue
        seen.add(normalized)
        values.append(normalized)
    if not values:
        raise pytest.UsageError("--pipeline-software must contain at least one software name.")
    return values


def pytest_addoption(parser: pytest.Parser) -> None:
    group = parser.getgroup("video-pipeline")
    group.addoption(
        "--run-e2e",
        action="store_true",
        default=False,
        help="Run tests that launch real Windows desktop applications.",
    )
    group.addoption(
        "--run-ai-turbo",
        action="store_true",
        default=False,
        help="Run tests that require AI Turbo Engine and Performance Boost automation.",
    )
    group.addoption(
        "--pipeline-workload-name",
        default="1080p_video_processing_speed",
        help="Logical workload label used for pytest-driven pipeline runs.",
    )
    group.addoption(
        "--pipeline-results-root",
        default="results/pytest_runs",
        help="Directory where pytest-driven pipeline runs will store outputs.",
    )
    group.addoption(
        "--pipeline-input-video",
        default="",
        help="Optional explicit input video path. Overrides workload-based test_data lookup.",
    )
    group.addoption(
        "--pipeline-blend-file",
        default="",
        help="Optional explicit Blender .blend path.",
    )
    group.addoption(
        "--pipeline-blender-exe",
        default="",
        help="Optional explicit Blender executable path.",
    )
    group.addoption(
        "--pipeline-ai-turbo-window-title",
        default=DEFAULT_AI_TURBO_TITLE_RE,
        help="Top-level window regex used to connect to AI Turbo Engine.",
    )
    group.addoption(
        "--pipeline-ai-turbo-shortcut",
        default=str(REPO_ROOT / "whitelist app" / "AI Turbo Engine.lnk"),
        help="Optional AI Turbo Engine shortcut path. Leave empty if the app will be opened manually.",
    )
    group.addoption(
        "--pipeline-software",
        default=",".join(DEFAULT_SOFTWARE_ORDER),
        help="Comma-separated software list for the pytest workflow.",
    )
    group.addoption(
        "--pytest-artifacts-root",
        default="results/pytest_reports",
        help="Directory where pytest session summaries and log files will be written.",
    )


def pytest_configure(config: pytest.Config) -> None:
    artifacts_root = _pytest_artifacts_root(config)
    session_dir = artifacts_root / datetime.now().strftime("%Y%m%d_%H%M%S")
    session_dir.mkdir(parents=True, exist_ok=True)
    config._autotest_pytest_session_dir = session_dir
    config._autotest_pytest_results = {}
    config._autotest_pytest_started_at = time.time()

    file_handler = logging.FileHandler(session_dir / "pytest_session.log", encoding="utf-8")
    file_handler.setLevel(logging.INFO)
    file_handler.setFormatter(logging.Formatter("%(asctime)s | %(levelname)s | %(name)s | %(message)s"))
    logging.getLogger().addHandler(file_handler)
    config._autotest_pytest_log_handler = file_handler


def pytest_unconfigure(config: pytest.Config) -> None:
    file_handler = getattr(config, "_autotest_pytest_log_handler", None)
    if file_handler is not None:
        logging.getLogger().removeHandler(file_handler)
        file_handler.close()


def pytest_collection_modifyitems(config: pytest.Config, items: list[pytest.Item]) -> None:
    run_e2e = config.getoption("--run-e2e")
    run_ai_turbo = config.getoption("--run-ai-turbo")
    skip_e2e = pytest.mark.skip(reason="use --run-e2e to run real desktop-automation workflow tests")
    skip_ai_turbo = pytest.mark.skip(reason="use --run-ai-turbo to run AI Turbo workflow tests")
    skip_windows = pytest.mark.skip(reason="Windows desktop automation tests require Windows.")
    for item in items:
        if "windows" in item.keywords and sys.platform != "win32":
            item.add_marker(skip_windows)
        if "e2e" in item.keywords and not run_e2e:
            item.add_marker(skip_e2e)
        if "ai_turbo" in item.keywords and not run_ai_turbo:
            item.add_marker(skip_ai_turbo)


def pytest_runtest_setup(item: pytest.Item) -> None:
    results = item.config._autotest_pytest_results
    results[item.nodeid] = {
        "nodeid": item.nodeid,
        "name": item.name,
        "markers": sorted(marker.name for marker in item.iter_markers()),
        "started_at": datetime.now().isoformat(timespec="seconds"),
        "outcome": "running",
        "duration_seconds": 0.0,
        "properties": {},
    }


@pytest.hookimpl(hookwrapper=True)
def pytest_runtest_makereport(item: pytest.Item, call: pytest.CallInfo[object]):
    outcome = yield
    report = outcome.get_result()
    results = item.config._autotest_pytest_results
    entry = results.setdefault(
        item.nodeid,
        {
            "nodeid": item.nodeid,
            "name": item.name,
            "markers": sorted(marker.name for marker in item.iter_markers()),
            "started_at": datetime.now().isoformat(timespec="seconds"),
            "outcome": "running",
            "duration_seconds": 0.0,
            "properties": {},
        },
    )
    if report.when == "call":
        entry["outcome"] = report.outcome
        entry["duration_seconds"] = round(report.duration, 3)
        entry["finished_at"] = datetime.now().isoformat(timespec="seconds")
        entry["properties"] = {name: value for name, value in report.user_properties}
    elif report.when == "setup" and report.outcome in {"failed", "skipped"}:
        entry["outcome"] = report.outcome
        entry["duration_seconds"] = round(report.duration, 3)
        entry["finished_at"] = datetime.now().isoformat(timespec="seconds")
        entry["properties"] = {name: value for name, value in report.user_properties}


def pytest_terminal_summary(terminalreporter, exitstatus: int, config: pytest.Config) -> None:
    session_dir = config._autotest_pytest_session_dir
    results = list(config._autotest_pytest_results.values())
    summary_json_path = session_dir / "session_summary.json"
    summary_md_path = session_dir / "session_summary.md"
    session_payload = {
        "started_at": datetime.fromtimestamp(config._autotest_pytest_started_at).isoformat(timespec="seconds"),
        "finished_at": datetime.now().isoformat(timespec="seconds"),
        "exitstatus": exitstatus,
        "test_count": len(results),
        "results": results,
    }
    summary_json_path.write_text(json.dumps(session_payload, indent=2, ensure_ascii=True), encoding="utf-8")

    lines = [
        "# Pytest Session Summary",
        "",
        f"- exitstatus: {exitstatus}",
        f"- test_count: {len(results)}",
        f"- session_log: {session_dir / 'pytest_session.log'}",
        "",
        "## Test Results",
        "",
        "| outcome | nodeid | duration_seconds | artifacts |",
        "| --- | --- | --- | --- |",
    ]
    for entry in results:
        properties = entry.get("properties", {})
        artifact_bits = []
        for key in ("pipeline_run_root", "pipeline_report"):
            value = properties.get(key, "")
            if value:
                artifact_bits.append(f"{key}={value}")
        artifact_text = "<br>".join(artifact_bits)
        lines.append(
            f"| {entry.get('outcome', '')} | {entry.get('nodeid', '')} | {entry.get('duration_seconds', 0.0)} | {artifact_text} |"
        )
    summary_md_path.write_text("\n".join(lines) + "\n", encoding="utf-8")

    terminalreporter.write_sep("-", f"pytest session artifacts: {session_dir}")
    terminalreporter.write_line(f"summary: {summary_md_path}")
    terminalreporter.write_line(f"log: {session_dir / 'pytest_session.log'}")


@pytest.fixture
def pipeline_results_root(pytestconfig: pytest.Config) -> Path:
    candidate = Path(pytestconfig.getoption("--pipeline-results-root"))
    if not candidate.is_absolute():
        candidate = REPO_ROOT / candidate
    candidate.mkdir(parents=True, exist_ok=True)
    return candidate.resolve(strict=False)


@pytest.fixture
def pipeline_software(pytestconfig: pytest.Config) -> list[str]:
    return _parse_software_list(pytestconfig.getoption("--pipeline-software"))


@pytest.fixture
def build_pipeline_namespace(
    pytestconfig: pytest.Config,
    pipeline_results_root: Path,
    pipeline_software: list[str],
):
    def factory(**overrides):
        namespace = argparse.Namespace(
            workload_name=pytestconfig.getoption("--pipeline-workload-name"),
            input_video=_resolve_repo_path(pytestconfig.getoption("--pipeline-input-video")),
            software=list(pipeline_software),
            results_root=str(pipeline_results_root),
            ai_turbo_window_title=pytestconfig.getoption("--pipeline-ai-turbo-window-title"),
            ai_turbo_shortcut=_resolve_repo_path(pytestconfig.getoption("--pipeline-ai-turbo-shortcut")),
            blend_file=_resolve_repo_path(pytestconfig.getoption("--pipeline-blend-file")),
            blender_exe=_resolve_repo_path(pytestconfig.getoption("--pipeline-blender-exe")),
            render_mode="frame",
            frame=1,
            skip_baseline=False,
            skip_turbo=False,
        )
        for key, value in overrides.items():
            setattr(namespace, key, value)
        return namespace

    return factory
