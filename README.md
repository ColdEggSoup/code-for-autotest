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
