# 从原始数据到 PPT 全流程命令

本文档记录当前仓库内可直接执行的全流程命令，覆盖以下三个批次：

- `1-2月`
- `3月`
- `Q1`

执行目标：

- 新推荐主流程从 `data/raw/{year}/{batch}` 原始数据生成分项统计结果
- 自动生成满意度汇总表、样本统计表和 PPT
- 兼容旧脚本链路时，仍会使用 `datas/ / 输出结果 / 汇总结果`

说明：

- 以下命令均可直接覆盖现有输出文件
- PPT 生成使用 OpenAI 兼容接口，当前配置文件里 `llm_notes.enabled = true`
- 运行前需确认项目根目录下 `.env` 与 `system_role.md` 已就绪
- 推荐主流程的目录约定为 `data/raw/{year}/{batch}`、`data/satisfaction_detail/{year}/{batch}`、`data/satisfaction_summary/{year}/{batch}`、`data/sample_summary/{year}/{batch}`、`data/ppt/{year}/{batch}`

## 推荐主流程命令

```bash
uv run python main_pipeline.py --year 2026 --batch 1-2月
uv run python main_pipeline.py --year 2026 --batch 3月
uv run python main_pipeline.py --year 2026 --batch Q1
```

说明：

- 若预查错发现阻断问题，程序会暂停
- 修正原始数据后，在终端确认继续
- 通过后自动完成分项统计、汇总表、样本表和 PPT

## 旧链路兼容校验

如果只想对旧 `generate_ppt.py` 配置做兼容性 dry-run 校验、不生成 PPT 文件，可以执行：

```bash
uv run python generate_ppt.py --config report_jobs.1-2月.toml --dry-run
uv run python generate_ppt.py --config report_jobs.3月.toml --dry-run
uv run python generate_ppt.py --config report_jobs.Q1.toml --dry-run
```

这些 dry-run 命令只校验旧脚本链路的 PPT 配置，不会验证新的 `data/...` 主流程目录约定，也不会覆盖预查错与人工确认流程。
