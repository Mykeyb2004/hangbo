# 从原始数据到 PPT 全流程命令

本文档记录当前仓库内可直接执行的全流程命令，覆盖以下三个批次：

- `1-2月`
- `3月`
- `Q1`

执行目标：

- 从 `datas/*` 原始数据生成分项统计结果
- 从分项统计结果生成满意度汇总表
- 生成开启大模型备注分析的 PPT

说明：

- 以下命令均可直接覆盖现有输出文件
- PPT 生成使用 OpenAI 兼容接口，当前配置文件里 `llm_notes.enabled = true`
- 运行前需确认项目根目录下 `.env` 与 `system_role.md` 已就绪

## 1-2月

```bash
uv run python survey_stats.py --config job01-02.toml
uv run python summary_table.py --input-dir 输出结果/1-2月 --output-dir 汇总结果/1-2月 --output-name 1-2月客户类型满意度汇总表.xlsx
uv run python generate_ppt.py --config report_jobs.1-2月.toml
```

输出：

- `输出结果/1-2月/*.xlsx`
- `汇总结果/1-2月/1-2月客户类型满意度汇总表.xlsx`
- `输出结果/1-2月满意度报告.pptx`

## 3月

```bash
uv run python survey_stats.py --config job03.toml
uv run python summary_table.py --input-dir 输出结果/3月 --output-dir 汇总结果/3月 --output-name 3月客户类型满意度汇总表.xlsx
uv run python generate_ppt.py --config report_jobs.3月.toml
```

输出：

- `输出结果/3月/*.xlsx`
- `汇总结果/3月/3月客户类型满意度汇总表.xlsx`
- `输出结果/3月满意度报告.pptx`

## Q1

```bash
uv run python survey_stats.py --config job_Q1.toml
uv run python summary_table.py --input-dir 输出结果/Q1 --output-dir 汇总结果/Q1 --output-name Q1客户类型满意度汇总表.xlsx
uv run python generate_ppt.py --config report_jobs.Q1.toml
```

输出：

- `输出结果/Q1/*.xlsx`
- `汇总结果/Q1/Q1客户类型满意度汇总表.xlsx`
- `输出结果/Q1满意度报告.pptx`

## 顺序建议

建议严格按下面顺序执行：

1. `survey_stats.py`
2. `summary_table.py`
3. `generate_ppt.py`

## 可选校验

如果只想先校验配置与输入，不生成 PPT 文件，可以执行：

```bash
uv run python generate_ppt.py --config report_jobs.1-2月.toml --dry-run
uv run python generate_ppt.py --config report_jobs.3月.toml --dry-run
uv run python generate_ppt.py --config report_jobs.Q1.toml --dry-run
```
