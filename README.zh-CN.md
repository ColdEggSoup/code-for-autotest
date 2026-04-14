# 视频处理时长自动化监控

[English](README.md) | [简体中文](README.zh-CN.md)

这个项目用于监控 6 个 Windows 视频软件的处理时长，并将结果写入 `csv`。
当一批 `csv` 收集完成后，再通过单独的命令生成一个 `xlsx` 汇总报告。

支持的软件：

- `shotcut`
- `kdenlive`
- `shutter_encoder`
- `avidemux`
- `handbrake`
- `blender`

## 监控方式

- `shotcut`、`kdenlive`、`shutter_encoder`：监控子进程运行时长
- `avidemux`、`handbrake`：解析软件日志
- `blender`：通过 `blender_listener.py` 在 Blender 内部注册 Python handler

## 安装

```powershell
pip install -r requirements.txt
```

## 主要命令

查看支持的软件：

```powershell
python main.py list-software
```

启动前台监控：

```powershell
python main.py start --software shotcut
```

启动后台监控：

```powershell
python main.py start-bg --software shotcut
```

查看某个后台会话状态：

```powershell
python main.py status --session-id shotcut_run001
```

把待处理的 csv 汇总成一个 xlsx：

```powershell
python main.py report
```

## 快速命令模板

以下示例基于 PowerShell。如果你使用 `cmd.exe`，请将：

- `$env:LOCALAPPDATA` 替换为 `%LOCALAPPDATA%`
- `$env:APPDATA` 替换为 `%APPDATA%`

前台模板：

```powershell
python main.py start --software shotcut --name two_4k
python main.py start --software kdenlive --name two_4k
python main.py start --software shutter_encoder --name two_4k
python main.py start --software avidemux --name two_4k
python main.py start --software handbrake --name two_4k
```

后台模板：

```powershell
python main.py start-bg --software shotcut --name two_4k
python main.py start-bg --software kdenlive --name two_4k
python main.py start-bg --software shutter_encoder --name two_4k
python main.py start-bg --software avidemux --name two_4k
python main.py start-bg --software handbrake --name two_4k
```

汇总模板：

```powershell
python main.py report --name two_4k
python main.py report --name two_4k --output results\xlsx\two_4k_summary.xlsx
```

## Start 命令

`python main.py start` 只要求：

```powershell
python main.py start --software <software>
```

可选参数：

- `--name`：测试标签，例如 `two_4k`
- `--session-id`：可选，自定义会话 id
- `--log-path`：可选，仅用于 `avidemux` 或 `handbrake` 的日志路径覆盖
- `--output`：可选，自定义 csv 输出路径

当前行为：

- 前台模式会占用当前终端
- 运行过程中会输出进度信息
- 只生成 `csv`
- 不会自动生成 `xlsx`

默认命名规则：

- 不带 `--name`：`results\csv\<software>_runNNN.csv`
- 带 `--name`：`results\csv\<software>_<name>_runNNN.csv`

常用示例：

```powershell
python main.py start --software avidemux
python main.py start --software shotcut --name two_4k
python main.py start --software handbrake --log-path "$env:APPDATA\HandBrake\logs"
python main.py start --software avidemux --log-path "$env:LOCALAPPDATA\avidemux\admlog.txt"
python main.py start --software shotcut --output results\csv\shotcut_custom.csv
```

各软件默认说明：

- `shotcut`：不需要额外路径参数
- `kdenlive`：不需要额外路径参数
- `shutter_encoder`：不需要额外路径参数
- `avidemux`：默认日志路径为 `%LOCALAPPDATA%\avidemux\admlog.txt`
- `handbrake`：默认日志目录为 `%APPDATA%\HandBrake\logs`

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

## 后台模式

如果你希望监控在后台运行，可以使用 `start-bg`：

```powershell
python main.py start-bg --software avidemux --name two_4k
python main.py status --session-id avidemux_two_4k_run001
python main.py stop --session-id avidemux_two_4k_run001
```

`status` 常用字段：

- `progress_state`
- `progress_message`
- `capture_detected`
- `capture_complete`
- `preview_duration_seconds`
- `current_source_path`
- `worker_stdout_path`
- `xlsx_status`

`start` / `start-bg` 阶段的 `xlsx_status` 通常为 `not_generated`，因为 xlsx 是后续由 `report` 命令生成的。

## Report 命令

当一批 csv 收集完成后，再使用 `report`：

```powershell
python main.py report
python main.py report --name two_4k
python main.py report --name two_4k --output results\xlsx\two_4k_summary.xlsx
```

当前行为：

- 默认扫描 `results\csv`
- 在 `results\xlsx` 下生成一个 xlsx
- 将参与本次汇总的 csv 移动到与 xlsx 同名的同级文件夹中
- 这些 csv 被移走后，下一批测试会重新从 `run001` 开始计数
- 如果 `start` 时使用了 `--name`，该值会带入 xlsx 的 `test_case`

示例输出：

- `results\xlsx\two_4k_summary_run001.xlsx`
- `results\xlsx\two_4k_summary_run001\*.csv`

## Xlsx 内容

生成的工作簿包含：

- `DeviceInfo`：`device_name`、`windows_system_version`、`cpu`、`gpu`、`mem_ram`、`disk`
- `TestResults`：`test_software`、`test_case`、`start_time`、`end_time`、`duration_seconds`

如果你仍然想直接调用独立生成器：

```powershell
python xlsx_report_generator.py --input results --output results\test_report.xlsx --default-test-case log_monitor_case
```

## 软件说明

- Avidemux 默认日志路径：`%LOCALAPPDATA%\avidemux\admlog.txt`
- HandBrake 默认日志目录：`%APPDATA%\HandBrake\logs`
- Shotcut、Kdenlive、Shutter Encoder 默认只捕获新的匹配子进程
- HandBrake 会忽略通用 activity log，优先使用 encode log

## Blender

`blender` 现在可以直接通过 `main.py start` 启动。
`main.py` 会以后台渲染模式拉起 Blender，执行 `blender_listener.py`，等待渲染完成后再写入 csv 结果。

自动拉起示例：

```powershell
python main.py start --software blender --blend-file "C:\path\to\project.blend"
python main.py start --software blender --blend-file "C:\path\to\project.blend" --blender-exe "C:\Program Files\Blender Foundation\Blender 4.5\blender.exe"
python main.py start --software blender --blend-file "C:\path\to\project.blend" --render-mode animation
```

说明：

- `blender` 必须传 `--blend-file`
- `blender` 不支持 `start-bg`
- `--render-mode frame` 默认使用 `--frame 1`
- `--render-mode animation` 会调用 `blender -a`
- 如果不传 `--blender-exe`，`main.py` 会先尝试常见的默认安装路径
