# 统一 CLI 主流程与废弃 GUI 设计

## 背景

当前仓库已经有两套并行使用方式：

- 新主流程：`main_pipeline.py --year --batch` 按 `data/...` 固定目录约定运行。
- 旧分步入口：`job*.toml`、`report_jobs.*.toml`、`ppt_job.example.toml` 和 GUI 运行时配置仍按不同数据源复制路径参数。

这些旧入口让用户需要为 `1-2月`、`3月`、`Q1` 等数据源分别维护配置，和“固定目录 + 批次参数”的目标冲突。

## 目标

把推荐使用方式统一为单一 CLI 主流程：

```bash
uv run python main_pipeline.py --year 2026 --batch 3月
```

`pipeline.defaults.toml` 是默认全局配置文件，正常运行不需要显式传入 `--config pipeline.defaults.toml`。

## 非目标

- 不保留 GUI 兼容入口。
- 不继续维护按批次复制的统计/PPT TOML 配置。
- 不改变问卷统计、汇总、样本统计、PPT 生成的业务计算口径。
- 不迁移或删除已有业务数据文件。

## 目录与配置边界

### 固定目录

主流程只通过 `year` 和 `batch` 推导输入输出路径：

- 原始数据：`data/raw/{year}/{batch}`
- 分项结果：`data/satisfaction_detail/{year}/{batch}`
- 满意度汇总：`data/satisfaction_summary/{year}/{batch}`
- 样本汇总：`data/sample_summary/{year}/{batch}`
- PPT 输出：`data/ppt/{year}/{batch}/{batch}满意度报告.pptx`
- 日志：`logs/pipeline/{year}/{batch}`

### 全局配置

`pipeline.defaults.toml` 只保存跨批次共享的业务参数：

- `sheet_name`
- `calculation_mode`
- `sample_config_path`
- `[ppt]` 下的模板、版式、图表页、备注页等默认参数

路径类参数不再按批次写入 TOML。

## GUI 废弃策略

彻底删除 GUI 功能：

- 删除 `hangbo_gui.py`
- 删除 `tests/test_hangbo_gui.py`
- 删除或改写文档中 GUI 工作台、GUI 运行时配置、GUI 高级重跑工具相关内容

删除后，仓库不再提供图形界面入口。

## 旧配置废弃策略

删除旧分步配置文件，避免用户误用：

- `job.toml`
- `job01-02.toml`
- `job03.toml`
- `job_Q1.toml`
- `report_jobs.1-2月.toml`
- `report_jobs.3月.toml`
- `report_jobs.Q1.toml`
- `report_jobs.example.toml`
- `report_jobs.directory.example.toml`
- `ppt_job.example.toml`

如果仍需单独运行底层脚本，可直接使用脚本 CLI 参数，但不再提供按批次 TOML 样例作为推荐入口。

## CLI 行为

`main_pipeline.py` 保留 `--config` 参数作为高级覆盖选项，但默认值就是 `pipeline.defaults.toml`。

推荐文档只展示：

```bash
uv run python main_pipeline.py --year 2026 --batch Q1
```

只有在复制一份全局默认配置做临时覆盖时，才需要显式传入：

```bash
uv run python main_pipeline.py --year 2026 --batch Q1 --config pipeline.no-llm.toml
```

## 测试策略

采用测试驱动方式实施：

1. 先更新 CLI 与路径行为测试，确保默认配置不需要显式传入。
2. 再删除 GUI 测试和旧配置依赖测试。
3. 更新 PPT 配置测试，不再要求旧 `report_jobs.*.toml` 存在。
4. 更新文档约束测试或新增轻量测试，确保 README 推荐命令不再使用旧配置文件。
5. 使用 `uv run` 运行相关测试。

## 风险与处理

- 删除 GUI 会影响原 GUI 用户；通过 README 和 `/docs` 明确唯一 CLI 入口。
- 删除旧 TOML 会影响旧命令；通过文档给出主流程替代命令。
- `--config` 虽保留，但只作为高级覆盖，不作为默认运行方式展示。

## 验收标准

- `uv run python main_pipeline.py --year 2026 --batch 3月` 是文档中的唯一推荐主流程命令。
- 仓库中不再有 GUI 代码和 GUI 测试。
- 仓库中不再有按批次复制的旧任务 TOML。
- 测试不再依赖 `report_jobs.*.toml`、`job*.toml` 或 `hangbo_gui.py`。
- 现有主流程测试仍通过。
