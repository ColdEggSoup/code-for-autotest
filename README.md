# Video Duration Automation

[English](README.md) | [简体中文](README.zh-CN.md)

This project monitors processing time for six Windows video tools and writes the results to `csv`.
After one batch of `csv` files is collected, use a separate command to generate one `xlsx` report.

Supported software:

- `shotcut`
- `kdenlive`
- `shutter_encoder`
- `avidemux`
- `handbrake`
- `blender`

## Monitoring Modes

- `shotcut`, `kdenlive`, `shutter_encoder`: monitor child worker process runtime
- `avidemux`, `handbrake`: parse software logs
- `blender`: use Blender-side Python handlers via `blender_listener.py`

## Install

```powershell
pip install -r requirements.txt
```

The initialization script now also writes clearer run artifacts:

```powershell
initialize_environment.cmd
```

Artifact paths:

- `results/initialize_environment_runs/<timestamp>/initialize_summary.md`
- `results/initialize_environment_runs/<timestamp>/initialize_progress.jsonl`
- `results/initialize_environment_runs/<timestamp>/initialize_environment.log`

Use them as follows:

- `initialize_summary.md`: quick status view for the current initialization run
- `initialize_progress.jsonl`: structured machine-readable event stream
- `initialize_environment.log`: full detailed log file

## Pytest Workflow

The repository now includes two `pytest` layers:

- `pytest`: run only the fast unit tests without launching desktop applications
- `pytest -m e2e ... --run-e2e`: run workflow 1 and execute the baseline pass for `shotcut`, `kdenlive`, `shutter_encoder`, `avidemux`, `handbrake`, and `blender`
- `pytest -m "e2e and ai_turbo" ... --run-e2e --run-ai-turbo`: run workflow 1 + workflow 2; the second pass starts AI Turbo Engine, enables `Performance Boost`, checks the matching app, and generates an xlsx report with a `Comparison` sheet

The `Comparison` sheet writes the improvement formula automatically:

- `improvement_percent = (baseline_duration_seconds - turbo_duration_seconds) / baseline_duration_seconds`

Common commands:

```powershell
python -m pytest
python -m pytest -m e2e tests/e2e/test_full_pipeline_workflow.py --run-e2e --pipeline-results-root results/pytest_runs
python -m pytest -m "e2e and ai_turbo" tests/e2e/test_full_pipeline_workflow.py --run-e2e --run-ai-turbo --pipeline-results-root results/pytest_runs
```

Common optional arguments:

- `--pipeline-workload-name`
- `--pipeline-input-video`
- `--pipeline-blend-file`
- `--pipeline-blender-exe`
- `--pipeline-software`
- `--pytest-artifacts-root`

`pytest` now writes two clearer report layers:

- session summary: `results/pytest_reports/<timestamp>/session_summary.md`
- full session log: `results/pytest_reports/<timestamp>/pytest_session.log`
- per-pipeline progress summary: `results/pytest_runs/<run>/pipeline_summary.md`
- per-pipeline structured event stream: `results/pytest_runs/<run>/pipeline_progress.jsonl`

Use them as follows:

- `session_summary.md`: quick pass/fail overview, durations, and artifact paths
- `pipeline_summary.md`: current workflow position plus per-software success/failure
- `pipeline_progress.jsonl`: structured machine-readable event timeline

The terminal also prints concise live progress lines such as:

```text
[pipeline][started] case variant=baseline software=shotcut
[pipeline][completed] case variant=baseline software=shotcut csv=...
[pipeline][started] report report=...
[pipeline][completed] pipeline report=...
```

If AI Turbo Engine is not installed on the current machine yet, run the baseline command first and keep the `ai_turbo` command for later.

## Main Commands

List supported software:

```powershell
python main.py list-software
```

Start a foreground monitor:

```powershell
python main.py start --software shotcut
```

Start a background monitor:

```powershell
python main.py start-bg --software shotcut
```

Check one background session:

```powershell
python main.py status --session-id shotcut_run001
```

Generate one xlsx report from pending csv files:

```powershell
python main.py report
```

## Quick Templates

The examples below use PowerShell. If you use `cmd.exe`, replace:

- `$env:LOCALAPPDATA` with `%LOCALAPPDATA%`
- `$env:APPDATA` with `%APPDATA%`

Foreground templates:

```powershell
python main.py start --software shotcut --name two_4k
python main.py start --software kdenlive --name two_4k
python main.py start --software shutter_encoder --name two_4k
python main.py start --software avidemux --name two_4k
python main.py start --software handbrake --name two_4k
```

Background templates:

```powershell
python main.py start-bg --software shotcut --name two_4k
python main.py start-bg --software kdenlive --name two_4k
python main.py start-bg --software shutter_encoder --name two_4k
python main.py start-bg --software avidemux --name two_4k
python main.py start-bg --software handbrake --name two_4k
```

Report templates:

```powershell
python main.py report --name two_4k
python main.py report --name two_4k --output results\xlsx\two_4k_summary.xlsx
```

## Start Command

`python main.py start` only requires:

```powershell
python main.py start --software <software>
```

Optional arguments:

- `--name`: test label, for example `two_4k`
- `--session-id`: optional explicit session id
- `--log-path`: optional log path override for `avidemux` or `handbrake`
- `--output`: optional csv output path

Behavior:

- foreground mode keeps the current terminal occupied
- it prints progress while monitoring
- it writes `csv` only
- it does not generate `xlsx` automatically

Default naming:

- without `--name`: `results\csv\<software>_runNNN.csv`
- with `--name`: `results\csv\<software>_<name>_runNNN.csv`

Common examples:

```powershell
python main.py start --software avidemux
python main.py start --software shotcut --name two_4k
python main.py start --software handbrake --log-path "$env:APPDATA\HandBrake\logs"
python main.py start --software avidemux --log-path "$env:LOCALAPPDATA\avidemux\admlog.txt"
python main.py start --software shotcut --output results\csv\shotcut_custom.csv
```

Software-specific defaults:

- `shotcut`: no extra path argument required
- `kdenlive`: no extra path argument required
- `shutter_encoder`: no extra path argument required
- `avidemux`: default log path is `%LOCALAPPDATA%\avidemux\admlog.txt`
- `handbrake`: default log directory is `%APPDATA%\HandBrake\logs`

### Shotcut

```powershell
python main.py start --software shotcut --name two_4k
python main.py start-bg --software shotcut --name two_4k
```

### Kdenlive

```powershell
python main.py start --software kdenlive --name two_4k
python main.py start-bg --software kdenlive --name two_4k
```

### Shutter Encoder

```powershell
python main.py start --software shutter_encoder --name two_4k
python main.py start-bg --software shutter_encoder --name two_4k
```

### Avidemux

```powershell
python main.py start --software avidemux --name two_4k
python main.py start-bg --software avidemux --name two_4k
python main.py start --software avidemux --name two_4k --log-path "$env:LOCALAPPDATA\avidemux\admlog.txt"
```

### HandBrake

```powershell
python main.py start --software handbrake --name two_4k
python main.py start-bg --software handbrake --name two_4k
python main.py start --software handbrake --name two_4k --log-path "$env:APPDATA\HandBrake\logs"
```

## Background Mode

Use `start-bg` when you want the monitor to run in the background:

```powershell
python main.py start-bg --software avidemux --name two_4k
python main.py status --session-id avidemux_two_4k_run001
python main.py stop --session-id avidemux_two_4k_run001
```

Useful fields from `status`:

- `progress_state`
- `progress_message`
- `capture_detected`
- `capture_complete`
- `preview_duration_seconds`
- `current_source_path`
- `worker_stdout_path`
- `xlsx_status`

`xlsx_status` is normally `not_generated` during `start` / `start-bg`, because xlsx is created later by `report`.

## Report Command

Use `report` after one batch of csv files has been collected:

```powershell
python main.py report
python main.py report --name two_4k
python main.py report --name two_4k --output results\xlsx\two_4k_summary.xlsx
```

Behavior:

- scans `results\csv` by default
- generates one xlsx file under `results\xlsx`
- moves the participating csv files into a sibling folder with the same stem as the xlsx file
- after those csv files are moved out, the next batch can start again from `run001`
- if `--name` was used during `start`, that value is carried into the xlsx `test_case`

Example output:

- `results\xlsx\two_4k_summary_run001.xlsx`
- `results\xlsx\two_4k_summary_run001\*.csv`

## Xlsx Content

The generated workbook contains:

- `DeviceInfo`: `device_name`, `windows_system_version`, `cpu`, `gpu`, `mem_ram`, `disk`
- `TestResults`: `test_software`, `test_case`, `start_time`, `end_time`, `duration_seconds`

If you want to call the standalone generator directly:

```powershell
python xlsx_report_generator.py --input results --output results\test_report.xlsx --default-test-case log_monitor_case
```

## Software Notes

- Avidemux default log path: `%LOCALAPPDATA%\avidemux\admlog.txt`
- HandBrake default log directory: `%APPDATA%\HandBrake\logs`
- Shotcut, Kdenlive, and Shutter Encoder only capture new matching worker processes by default
- HandBrake ignores generic activity logs and prefers encode logs

## Blender

`blender` can be started directly by `main.py start`.
`main.py` will launch Blender in background render mode, execute `blender_listener.py`, wait for render completion, and then write the csv result.

Foreground auto-launch examples:

```powershell
python main.py start --software blender --blend-file "C:\path\to\project.blend"
python main.py start --software blender --blend-file "C:\path\to\project.blend" --blender-exe "C:\Program Files\Blender Foundation\Blender 4.5\blender.exe"
python main.py start --software blender --blend-file "C:\path\to\project.blend" --render-mode animation
```

Notes:

- `--blend-file` is required for Blender.
- `start-bg` is not supported for Blender.
- `--render-mode frame` uses `--frame 1` by default.
- `--render-mode animation` runs `blender -a`.
- If `--blender-exe` is omitted, `main.py` will try common default Blender install paths first.
