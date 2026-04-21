# UI 适配新数据分析流程计划

> 计划日期：2026-04-21
>
> 范围：仅规划 `hangbo_gui.py` 及其测试如何适配当前推荐的新数据分析流程，暂不进入实现。

## 1. 背景与目标

当前推荐主流程已经统一为：

```text
data/raw/{year}/{batch}
  -> 预查错
  -> 必要时人工修正并确认继续
  -> data/satisfaction_detail/{year}/{batch}
  -> data/satisfaction_summary/{year}/{batch}
  -> data/sample_summary/{year}/{batch}
  -> data/ppt/{year}/{batch}/{batch}满意度报告.pptx
```

命令入口为：

```bash
uv run python main_pipeline.py --year 2026 --batch 3月
```

GUI 目前仍保留较多旧流程概念：

- 默认目录仍指向 `datas/3月`、`输出结果/3月`、`汇总结果/3月`。
- “一键执行主流程”实际串行调用 `phase_column_preprocess.py`、`survey_stats.py`、`summary_table.py`、`generate_ppt.py` 等分步脚本。
- 页面主线没有显式展示 `预查错`、`样本统计表`、`data/...` 固定目录契约。
- 多月合并与补写年月仍容易被理解为主流程默认步骤，但新流程要求先准备好 `data/raw/{year}/{batch}`，再运行 `main_pipeline.py`。

本次 UI 更新目标是：让普通用户在 GUI 中按新流程理解和执行批次处理；保留必要的高级/分步工具，但默认路径、默认按钮、状态展示和一键执行都围绕 `main_pipeline.py`。

## 2. 设计原则

1. **主流程优先**：GUI 默认入口应执行 `main_pipeline.py --year --batch`，由主流程负责预查错、单月自动补年月、分项、汇总、样本、PPT。
2. **固定目录可解释**：用户输入 `年份` 与 `批次` 后，界面自动展示所有输入输出目录，不再要求用户手工维护分项、汇总、PPT 路径。
3. **预查错可见**：将预查错作为主流程第一步展示，清楚提示阻断问题、日志位置和修正后继续方式。
4. **多月边界清楚**：单月批次可自动补年月；`1-2月`、`Q1`、`1-6月` 等合并批次必须已带正确 `年份` / `月份`，不能整批统一补单一月份。
5. **高级工具后置**：合并、手动补年月、分步生成分项/汇总/PPT 作为“辅助工具/高级操作”，避免用户误以为它们是默认主线。
6. **测试驱动**：先补 GUI 配置、命令生成、状态文案和页面 copy 的单元测试，再改实现。

## 3. 推荐 UI 结构

### 3.1 工作台总览

展示字段调整为：

- 年份：例如 `2026`
- 批次：例如 `3月`、`1-2月`、`Q1`
- 批次类型：单月 / 合并批次，由 `parse_single_month_batch()` 判断
- 原始数据目录：`data/raw/{year}/{batch}`
- 分项输出目录：`data/satisfaction_detail/{year}/{batch}`
- 满意度汇总目录：`data/satisfaction_summary/{year}/{batch}`
- 样本统计目录：`data/sample_summary/{year}/{batch}`
- PPT 输出文件：`data/ppt/{year}/{batch}/{batch}满意度报告.pptx`
- 日志目录：`logs/pipeline/{year}/{batch}`

主要按钮调整为：

- `运行主流程`
- `打开原始数据目录`
- `查看预查错日志`
- `查看输出目录`

### 3.2 批次设置页

将当前“数据源管理”调整为“批次设置 / 数据源准备”：

- 用户只填写 `年份` 和 `批次`。
- 根据 `build_pipeline_paths(year, batch)` 自动生成目录预览。
- 显示标准来源文件清单：`展览.xlsx`、`会议.xlsx`、`酒店.xlsx`、`餐饮.xlsx`、`会展服务商.xlsx`、`旅游.xlsx`。
- 扫描 `data/raw/{year}/{batch}`，显示每个标准来源文件是否存在。
- 明确说明“缺少部分来源文件允许；一个标准来源文件都没有会阻断”。

### 3.3 预查错页

新增或改造为独立页面：

- 显示主流程会检查的阻断项：
  - 原始目录不存在
  - 完全没有标准来源文件
  - 缺少 `问卷数据` sheet
  - Excel 读取失败
  - 合并批次缺少 `年份` / `月份`
  - 存在未映射客户标签记录
- 显示日志路径：
  - `logs/pipeline/{year}/{batch}/precheck.log`
  - `logs/pipeline/{year}/{batch}/unmapped_customer_records.log`
  - `logs/pipeline/{year}/{batch}/pipeline.log`
- 主按钮可先实现为 `运行主流程并执行预查错`，复用主流程输出；如果后续需要独立预查错，再考虑包装 `run_precheck()`。
- 当主流程输出提示“修改完成后输入 y / yes / 继续”时，GUI 需要提供输入/确认机制，不能让后台子进程卡在不可见交互上。

### 3.4 主流程执行页

将“一键执行主流程”改为默认执行：

```bash
uv run python main_pipeline.py --year {year} --batch {batch} --config pipeline.defaults.toml
```

步骤展示改为：

1. 预查错
2. 人工修正确认
3. 满意度分项统计
4. 满意度汇总表
5. 样本统计表
6. PPT 生成

执行进度可先基于 stdout 文本识别：

- `[PRECHECK]` 对应预查错阶段
- `[PIPELINE] 开始生成满意度分项统计...`
- `[PIPELINE] 开始生成满意度汇总表...`
- `[PIPELINE] 开始生成样本统计表...`
- `[PIPELINE] 开始生成 PPT...`
- `[PIPELINE] PPT 已生成：...`

### 3.5 辅助工具页

原“预处理”页调整为“辅助工具”，保留但降低默认优先级：

- `兼容新版调查问卷数据结构`：用于 `phase_column_preprocess.py`。
- `给单月目录补写年份/月度`：只允许或强提示用于单月目录。
- `合并多月问卷`：输出目录默认应为 `data/raw/{year}/{batch}`。

补写年月工具需要新增保护文案：

- 单月批次可用。
- 合并批次不要整批写同一月份。
- 若当前批次不是单月，按钮默认可禁用，或弹出高风险确认。

### 3.6 分步高级页

原分项统计、汇总统计、PPT 生成页面可保留为“高级分步执行”，但默认目录必须来自 `build_pipeline_paths()`：

- 分项统计读取 `data/raw/{year}/{batch}`，输出 `data/satisfaction_detail/{year}/{batch}`。
- 汇总统计读取 `data/satisfaction_detail/{year}/{batch}`，输出 `data/satisfaction_summary/{year}/{batch}`。
- 样本统计新增状态展示，读取 `data/raw/{year}/{batch}`，输出 `data/sample_summary/{year}/{batch}`。
- PPT 读取 `data/satisfaction_detail/{year}/{batch}`，输出 `data/ppt/{year}/{batch}/{batch}满意度报告.pptx`。

高级页用于重跑单一步骤，不作为默认推荐路径。

## 4. 文件改动计划

### 4.1 `hangbo_gui.py`

计划修改：

1. 引入 `build_pipeline_paths()` 与 `parse_single_month_batch()`，新增 GUI 层的 `year` / `batch` 配置字段。
2. 将 `GuiBatchConfig` 默认值切换到新目录：
   - `year_value = "2026"`
   - `batch_value = "3月"`
   - `single_input_dir = PROJECT_ROOT / "data" / "raw" / "2026" / "3月"`
   - `stats_output_dir = PROJECT_ROOT / "data" / "satisfaction_detail" / "2026" / "3月"`
   - `summary_output_dir = PROJECT_ROOT / "data" / "satisfaction_summary" / "2026" / "3月"`
   - `sample_summary_output_dir = PROJECT_ROOT / "data" / "sample_summary" / "2026" / "3月"`
   - `output_ppt_path = PROJECT_ROOT / "data" / "ppt" / "2026" / "3月" / "3月满意度报告.pptx"`
3. 新增根据 `year` / `batch` 同步路径的 helper，例如 `build_gui_pipeline_defaults(year, batch)`。
4. 新增 `build_main_pipeline_command(config)`，命令为 `main_pipeline.py --year --batch --config pipeline.defaults.toml`。
5. 将“一键执行主流程”从 `build_task_commands()` 分步模式切换为 `build_main_pipeline_command()`。
6. 调整步骤状态 key，加入 `precheck` 与 `sample_summary`。
7. 调整页面标题和文案，突出 `data/raw/{year}/{batch}` 与 `logs/pipeline/{year}/{batch}`。
8. 保留 `build_survey_stats_command()`、`build_summary_command()`、`build_ppt_command()` 作为高级分步入口。
9. 为主流程交互等待增加 GUI 处理策略：
   - 最小方案：在日志中提示用户当前任务等待输入，并提供“继续/终止”按钮向子进程 stdin 写入 `y` 或 `stop`。
   - 若现有 `BackgroundTaskRunner` 不支持 stdin，则计划先扩展 runner，再接入按钮。

### 4.2 `tests/test_hangbo_gui.py`

计划新增或更新测试：

1. `GuiBatchConfig` 默认目录应使用 `data/...` 新目录。
2. `build_gui_pipeline_defaults("2026", "3月")` 应生成与 `build_pipeline_paths()` 一致的所有路径。
3. `build_main_pipeline_command()` 应生成：

   ```text
   python main_pipeline.py --year 2026 --batch 3月 --config pipeline.defaults.toml
   ```

4. 单月批次步骤文案应包含“预查错、分项、汇总、样本、PPT”。
5. 合并批次 `Q1` 应显示“不会自动补写年月”的说明。
6. 补写年月按钮在合并批次下应被禁用或触发高风险确认。
7. 分步高级命令仍使用新目录，而不是 `datas/`、`输出结果/`、`汇总结果/`。
8. 状态统计应包含样本统计表输出状态。

### 4.3 `docs/README.md` 或 `README.md`

实现完成后同步更新 GUI 说明：

- GUI 默认运行 `main_pipeline.py`。
- GUI 的主输入是 `年份 + 批次`。
- GUI 输出目录遵循 `data/...`。
- 分步页面是高级重跑工具。

## 5. TDD 实施顺序

### 阶段 A：配置与路径

1. 先写失败测试：GUI 默认配置不再使用旧目录。
2. 实现 `build_gui_pipeline_defaults()`。
3. 修改 `GuiBatchConfig` 默认值与 `_create_variables()` 默认值。
4. 运行：

   ```bash
   uv run python -m unittest tests/test_hangbo_gui.py
   ```

### 阶段 B：主流程命令

1. 先写失败测试：`build_main_pipeline_command()` 输出 `main_pipeline.py --year --batch --config`。
2. 实现命令构造函数。
3. 将主流程按钮改为调用主流程命令。
4. 保留原分步命令测试，并更新为高级路径。
5. 运行 GUI 单测。

### 阶段 C：页面与文案

1. 先写失败测试：页面文案包含新流程关键词与日志路径。
2. 更新工作台、批次设置、预查错/辅助工具页面。
3. 更新侧边栏顺序：
   - 工作台总览
   - 批次设置
   - 预查错
   - 主流程执行
   - 辅助工具
   - 高级分步
   - 任务日志
4. 运行 GUI 单测。

### 阶段 D：主流程交互等待

1. 先写失败测试：后台 runner 可向运行中进程 stdin 写入确认文本。
2. 扩展 `BackgroundTaskRunner` 或对应控制器，支持 `send_input("y\n")` 与 `send_input("stop\n")`。
3. 日志识别到 `修改完成后输入 y / yes / 继续` 时，启用“继续检查”按钮。
4. 运行 GUI 单测。

### 阶段 E：样本统计状态

1. 先写失败测试：工作台显示样本统计表状态。
2. 新增样本统计路径变量与文件存在性检查。
3. 主流程完成后刷新样本统计状态。
4. 运行 GUI 单测。

### 阶段 F：回归验证

1. 运行：

   ```bash
   uv run python -m unittest tests/test_hangbo_gui.py
   ```

2. 如改动涉及主流程路径或命令，再运行：

   ```bash
   uv run python -m unittest tests/test_main_pipeline.py tests/test_pipeline_paths.py tests/test_pipeline_config.py
   ```

3. 手动启动 GUI 做冒烟检查：

   ```bash
   uv run python hangbo_gui.py
   ```

## 6. 风险与处理

| 风险 | 影响 | 处理 |
| --- | --- | --- |
| 主流程遇到阻断问题时需要 stdin 交互 | GUI 后台进程可能卡住 | 扩展 runner 支持继续/终止输入 |
| 用户仍依赖旧 `datas/` 目录 | 升级后找不到数据 | 批次设置页明确提示迁移到 `data/raw/{year}/{batch}` |
| 合并批次被误用补年月 | 月份数据被覆盖 | 非单月批次禁用或强确认补写年月 |
| 高级分步入口与主流程默认配置不一致 | 输出分散、难以排查 | 分步路径全部由 `build_pipeline_paths()` 派生 |
| PPT 高级配置与 `pipeline.defaults.toml` 重复 | 用户不清楚哪个生效 | 主流程默认读 `pipeline.defaults.toml`，GUI 高级项仅用于分步重跑或后续另行设计写入配置 |

## 7. 建议确认点

进入实现前建议确认两点：

1. GUI 主流程是否只保留“一键执行 `main_pipeline.py`”作为默认推荐入口，原分步执行全部移动到高级区。
2. 主流程阻断后的 GUI 交互，是否采用“继续检查/终止任务”按钮向进程 stdin 写入的最小方案。

