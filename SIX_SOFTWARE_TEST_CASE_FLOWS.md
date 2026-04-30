# 六个软件测试用例流程说明

这份文档整理了当前默认流水线里六个软件的测试用例执行流程：

1. `shotcut`
2. `kdenlive`
3. `shutter_encoder`
4. `avidemux`
5. `handbrake`
6. `blender`

相关核心代码入口主要在：

- [full_test_pipeline.py](./full_test_pipeline.py)
- [software_operations.py](./software_operations.py)
- [automation_components.py](./automation_components.py)
- [main.py](./main.py)

## 总体结构

默认的软件执行顺序定义在 [full_test_pipeline.py](./full_test_pipeline.py)：

```python
DEFAULT_SOFTWARE_ORDER = ("shotcut", "kdenlive", "shutter_encoder", "avidemux", "handbrake", "blender")
```

这六个软件的测试流程可以分成三类：

1. `shotcut` 和 `kdenlive`
   这两个是“多 case 串行、单次启动复用窗口”的时间线类软件。

2. `shutter_encoder`、`avidemux`、`handbrake`
   这三个走统一的 `run_non_blender_case(...)` 模式，每个 case 独立启动监控、启动软件、执行自动化、等待导出完成。

3. `blender`
   这是单独的一套后台渲染监控流程，不走 `software_operations.py` 的 operator 模式。

## 通用非 Blender 流程

适用软件：

- `shutter_encoder`
- `avidemux`
- `handbrake`

入口函数：

- [run_non_blender_case](./full_test_pipeline.py)

通用流程如下：

1. 生成 `run_name`
2. 创建对应软件的 operator
3. 关闭残留窗口，确保是干净环境
4. 如有需要，配置 AI Turbo Engine Performance Boost
5. 先启动后台 timing monitor
6. 启动软件并等待主窗口出现
7. 执行 `operator.perform(input_video_path, output_video_path)`
8. 根据软件类型，选择：
   - 自动化结束后立即停止 monitor
   - 或等待 monitor 自然完成
9. 做 cleanup，关闭软件

说明：

- `avidemux` 和 `kdenlive` 在自动化结束后会立即停止 monitor
- `shotcut` 在自己的 sequence runner 里也是自动化结束后立刻停止 monitor
- `handbrake` 和 `shutter_encoder` 会等待 monitor 完成

## Shotcut

序列入口：

- [run_shotcut_case_sequence](./full_test_pipeline.py)

操作入口：

- `ShotcutOperator.perform(...)` in [software_operations.py](./software_operations.py)

### Shotcut 的 case 组织方式

Shotcut 不是每个 case 都重启软件，而是：

1. 先启动一次 Shotcut
2. 在同一个窗口会话里依次执行多个 case
3. 每个 case 独立启动自己的 monitor
4. 每个 case 自动化结束后立即停止自己的 monitor

这个 sequence runner 会先断言：

- `cases` 不能为空
- 所有 case 的 `software` 必须都是 `shotcut`
- 辅助小视频 `4K_small.mp4` 必须存在

### Shotcut 的三种 sequence_order

每个 Shotcut case 根据 `sequence_order` 走不同流程：

1. `sequence_order == 1`
   - 打开主输入视频
   - 追加到时间线

2. `sequence_order == 2`
   - 打开 `4K_small.mp4`
   - 追加到时间线

3. `sequence_order == 3`
   - 先把当前选中内容追加一次
   - 再打开主输入视频
   - 再追加一次

### Shotcut UI 自动化主流程

主流程如下：

1. 校验输入文件存在
2. 校验输出目录存在
3. 校验输出后缀必须是 `.mp4`
4. 连接 Shotcut 主窗口
5. 如果有 autosave recovery 弹窗，先 dismiss
6. 打开输入文件选择器
7. 填文件路径并回到主窗口
8. 把导入素材追加到时间线
9. 定位导出相关控件
10. 点击顶部工具栏 `Export`
11. 如有 save-changes 弹窗，先点 `No` / `Don't Save`
12. 等待导出面板里的 `Export Video/Audio` 或等价控件出现
13. 等待导出按钮可用
14. 点击导出按钮并填写输出文件
15. 最小化窗口
16. 最小化后确认输出文件已经开始写入
17. 轮询等待导出完成
18. 关闭 Shotcut

### Shotcut 收尾关闭逻辑

Shotcut 的关闭逻辑比较重，主要是因为它经常会弹保存确认：

1. 先尝试 dismiss save-changes 弹窗
2. 发送 `Alt+F4`
3. 再次 dismiss save-changes 弹窗
4. fallback 到 `WM_CLOSE`
5. 最后再走一次通用 prompt dismissal

当前关闭逻辑已经适配：

- 英文提示文本
- 中文提示文本
- 独立顶层弹窗
- 嵌在 owner window 上的弹窗
- `Button` / `Text` / `Pane` / `Group` / `Custom` 类型的 `No` 或 `Don't Save`

## Kdenlive

序列入口：

- [run_kdenlive_case_sequence](./full_test_pipeline.py)

操作入口：

- `KdenliveOperator.perform(...)` in [software_operations.py](./software_operations.py)

### Kdenlive 的 case 组织方式

Kdenlive 和 Shotcut 一样，也是多 case 复用同一个软件会话。

它会先断言：

- `cases` 不能为空
- 所有 case 的 `software` 必须都是 `kdenlive`
- 辅助小视频 `4K_small.mp4` 必须存在

### Kdenlive 的三种 sequence_order

1. `sequence_order == 1`
   - 打开主输入视频
   - 插入时间线

2. `sequence_order == 2`
   - 先关闭渲染窗口，不关主窗口
   - 打开 `4K_small.mp4`
   - 插入时间线

3. `sequence_order == 3`
   - 先关闭渲染窗口，不关主窗口
   - 直接从项目素材区选中小视频并插入
   - 再选中主视频并插入

### Kdenlive UI 自动化主流程

1. 连接 Kdenlive 主窗口
2. 如有 file recovery 弹窗，先 dismiss
3. 打开导入文件命令
4. 填写文件选择器
5. 导入后重连主窗口
6. 在 project bin 里选中导入素材
7. 插入到时间线
8. 打开 render dialog
9. 切到 `Render Project` 标签
10. 设置输出路径
11. 开始渲染
12. 最小化渲染窗口
13. 最小化后确认 render 已真正开始
14. 等待渲染完成
15. 关闭 render dialog
16. 关闭主窗口
17. dismiss 保存确认和 profile switch 提示

### Kdenlive 特点

- 会在 sequence 里长期持有 `_main_window` 和 `_render_dialog`
- 第二个、第三个 case 之间常常只是关渲染窗口，不关主窗口
- 如果多轮尝试仍然无法关闭，会 fallback 到 `taskkill`

## Shutter Encoder

入口：

- [run_non_blender_case](./full_test_pipeline.py)
- `ShutterEncoderOperator.perform(...)` in [software_operations.py](./software_operations.py)

### Shutter Encoder 主流程

1. 校验输入文件存在
2. 校验输出目录存在
3. 先对输入目录做一次输出文件快照
4. 检查 Shutter Encoder 是否已运行
5. 如果未运行，则从快捷方式启动
6. dismiss update 弹窗
7. 打开 source picker
8. 填写输入视频
9. 导入后重连主窗口
10. 检查导入素材是否在界面中可见
11. 选择功能 `H.265`
12. 启动转码
13. 最小化窗口
14. 最小化后确认任务已经开始
15. 通过目录快照差异等待输出生成完成
16. 关闭 Shutter Encoder
17. 把生成文件复制到验证路径

### Shutter Encoder 特点

- 允许复用已运行窗口
- 生成文件不是固定输出名，所以靠目录快照和新文件识别
- 关闭流程较强，如果失败会升级到 `taskkill`

## Avidemux

入口：

- [run_non_blender_case](./full_test_pipeline.py)
- `AvidemuxOperator.perform(...)` in [software_operations.py](./software_operations.py)

### Avidemux 主流程

1. 校验输入文件存在
2. 校验输出目录存在
3. 校验输出后缀必须是 `.mkv`
4. 连接 Avidemux 主窗口
5. dismiss 初始 `Thanks` 弹窗
6. 打开输入视频
7. 导入后重连主窗口
8. 设置视频编码为 `HEVC (x265)`
9. 设置容器为 `MKV Muxer`
10. 保存输出
11. 尝试把 encode progress dialog 最小化到托盘
12. 如果不能托盘最小化，则最小化主窗口
13. 等待导出完成
14. 关闭 Avidemux
15. 删除不需要的中间产物

### Avidemux 特点

- 有单独的进度对话框
- 支持“最小化到托盘”的专门路径
- 托盘按钮支持文本定位和几何定位双重 fallback

## HandBrake

入口：

- [run_non_blender_case](./full_test_pipeline.py)
- `HandBrakeOperator.perform(...)` in [software_operations.py](./software_operations.py)

### HandBrake 主流程

1. 校验输入文件存在
2. 校验输出目录存在
3. 校验输出后缀必须是 `.mp4`
4. 连接 HandBrake 主窗口
5. dismiss update-check prompt
6. dismiss recoverable-queue prompt
7. 打开源视频
8. 导入后重连主窗口
9. 选择 preset `HQ 2160P60 4K HEVC Surround`
10. 设置输出路径
11. 开始编码
12. 最小化窗口
13. 等待编码完成
14. 关闭 HandBrake
15. 删除不需要的中间产物

### HandBrake 特点

- 主窗口定位不是直接走默认 `connect_window(...)`，而是用自己的主窗口筛选逻辑
- 一开始就会先处理 update 和 recovery 相关弹窗

## Blender

入口：

- [run_blender_case](./full_test_pipeline.py)
- `run_blender_session(...)` in [main.py](./main.py)

### Blender 主流程

Blender 不走 `software_operations.py`，是单独的后台渲染监控模型。

执行流程：

1. 解析 blend 文件路径、blender 可执行文件路径、UI 模式、render 模式、frame
2. 断言 blend 文件存在
3. 如有需要，配置 AI Turbo Engine Performance Boost
4. 通过 `monitor_bridge.run_blender_monitor(...)` 启动 Blender 监控流程
5. 生成 Blender 命令
6. 根据模式分两种：
   - `visible`
     用带 listener 的可视化渲染流程
   - `headless`
     用后台渲染参数直接渲染
7. 如果是 `visible` 模式，会尝试把 Blender 可见窗口最小化
8. 等待渲染结束
9. 收集 CSV 输出
10. 断言 CSV 输出文件存在
11. 返回 case 结果

### Blender 特点

- 它不是“打开 GUI 后点按钮导出”的模式
- 它更像“带监控的渲染任务执行”
- 六个软件里，只有 Blender 不走 `build_operator(...).perform(...)`

## 六个软件对照表

| 软件 | 执行模式 | 是否多 case 复用会话 | 输出类型 | 核心动作 |
| --- | --- | --- | --- | --- |
| `shotcut` | 专用 sequence | 是 | `.mp4` | 时间线导出 |
| `kdenlive` | 专用 sequence | 是 | `.mp4` | 时间线渲染 |
| `shutter_encoder` | 通用 non-blender | 否 | 复制生成文件到目标路径 | H.265 转码 |
| `avidemux` | 通用 non-blender | 否 | `.mkv` | 保存编码输出 |
| `handbrake` | 通用 non-blender | 否 | `.mp4` | 启动编码 |
| `blender` | 专用后台流程 | 否 | CSV 监控结果 | 渲染任务 |

## 建议阅读顺序

如果你想顺代码查某个软件的完整 case 流程，建议按下面顺序看：

1. `full_test_pipeline.py`
2. `automation_components.py`
3. 对应软件的 `software_operations.py` operator
4. Blender 例外，再看 `main.py`

例如：

- Shotcut：`run_shotcut_case_sequence(...)` -> `ShotcutOperator`
- Kdenlive：`run_kdenlive_case_sequence(...)` -> `KdenliveOperator`
- HandBrake：`run_non_blender_case(...)` -> `HandBrakeOperator.perform(...)`
- Blender：`run_blender_case(...)` -> `run_blender_session(...)`
