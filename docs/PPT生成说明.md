# PPT 生成说明

本文档说明新流程中 PPT 如何生成、默认配置从哪里来，以及需要单独运行 `generate_ppt.py` 时如何传参。

主流程推荐入口：

```bash
uv run python main_pipeline.py --year 2026 --batch 3月
```

主流程会在满意度分项统计、满意度汇总表、样本统计表生成后，自动生成 PPT。

---

## 1. 输入与输出

PPT 生成脚本：

- `generate_ppt.py`

主流程调用位置：

- `pipeline_runtime.py` 中的 `generate_presentation()`

默认输入：

- `data/satisfaction_detail/{year}/{batch}`

默认输出：

- `data/ppt/{year}/{batch}/{batch}满意度报告.pptx`

示例：

- `data/ppt/2026/3月/3月满意度报告.pptx`
- `data/ppt/2026/1-2月/1-2月满意度报告.pptx`
- `data/ppt/2026/Q1/Q1满意度报告.pptx`

默认模板：

- `templates/template.pptx`

---

## 2. 主流程默认配置

主流程读取：

- `pipeline.defaults.toml`

与 PPT 相关的默认配置位于 `[ppt]`、`[ppt.chart_page]`、`[ppt.llm_notes]`。

关键参数：

- `template_path`：PPT 模板路径
- `sheet_name_mode`：读取 Excel sheet 的方式，默认 `first`
- `section_mode`：二级标题识别口径，默认 `auto`
- `blank_display`：空值展示文本，默认空字符串
- `max_single_table_rows`：单表最大明细行数
- `max_split_table_rows`：左右双表每侧最大明细行数
- `template_slide_index`：模板页索引，代码中按 0 开始
- `chart_page.enabled`：是否在数据页后生成图表页
- `llm_notes.enabled`：是否生成备注页分析

如需临时使用另一套默认配置：

```bash
uv run python main_pipeline.py \
  --year 2026 \
  --batch 3月 \
  --config pipeline.defaults.toml
```

---

## 3. PPT 页面生成规则

PPT 默认适配 `survey_stats.py` 产出的分项统计 Excel。

输入 Excel 通常包含：

- `指标`
- `满意度`
- `重要性`

页面规则：

- 每个客户类型 Excel 生成 1 页数据页
- 页面顺序优先按 `survey_customer_category_rules.py` 中的展示顺序
- 页面标题使用 `客户大类——客户类型`
- 总体行放在标题下方摘要区域
- 明细指标放在正文表格区域
- 明细行较少时渲染为一张全宽表
- 明细行较多时按二级指标块拆成左右两张表
- 空值默认不显示
- 某个二级指标下所有三级指标满意度都为空时，该二级指标块不显示

标题转换规则：

- 如果 Excel 文件名命中客户类别规则，会自动转换为展示名
- 例如 `展览主承办.xlsx` 显示为 `会展客户——展览活动主（承）办`
- 聚合类或兼容别名会统一到业务展示名称

---

## 4. 二级标题识别口径

`generate_ppt.py` 支持三种 `section_mode`：

- `auto`
- `template`
- `summary`

说明：

- `template`：按原始分项模板识别，例如 `会展服务`、`硬件设施`、`配套服务`
- `summary`：按汇总表展示口径识别，例如 `产品服务`、`智慧场馆/服务`
- `auto`：自动选择与当前 Excel 内容更匹配的口径

主流程默认：

```toml
section_mode = "auto"
```

---

## 5. 图表页

图表页由 `pipeline.defaults.toml` 中的 `[ppt.chart_page]` 控制。

默认规则：

- `enabled = true` 时，每个客户分组数据页后追加 1 页图表页
- 图表页标题与前一页数据页保持一致
- 左侧渲染图表
- 右侧渲染分析文字框
- 二级指标数量为 2 时使用柱状图
- 二级指标数量为 3 个及以上时使用雷达图
- 过滤后不足 2 个二级指标时跳过图表页

常用参数：

- `enabled`：是否启用图表页
- `placeholder_text`：没有备注分析时的右侧占位文字
- `image_dpi`：图表 PNG 清晰度

---

## 6. 备注页 LLM 分析

备注页分析由 `[ppt.llm_notes]` 控制。

默认配置项：

- `enabled`
- `env_path`
- `system_role_path`
- `target_chars`
- `temperature`
- `max_tokens`
- `checkpoint_chars`

启用前需要准备：

- `.env`
- `system_role.md`

`.env` 可参考：

- `.env.example`

运行时会读取以下环境变量：

- `OPENAI_API_KEY` 或 `LLM_API_KEY`
- `OPENAI_BASE_URL` 或 `LLM_BASE_URL`
- `OPENAI_MODEL` 或 `LLM_MODEL`
- `OPENAI_TEMPERATURE` 或 `LLM_TEMPERATURE`
- `OPENAI_TIMEOUT` 或 `LLM_TIMEOUT`

生成过程中会按页写入备注页，并保存检查点：

- 生成中断时，已完成页面和当前页已流式接收的内容会尽量保存到 `*.partial.pptx`
- 全部成功后，会删除检查点，只保留正式 PPT

如果不希望运行 PPT 阶段调用模型，可在 `pipeline.defaults.toml` 中将：

```toml
[ppt.llm_notes]
enabled = false
```

---

## 7. 单独运行 PPT 生成

主流程之外，也可以单独执行：

```bash
uv run python generate_ppt.py \
  --template-path templates/template.pptx \
  --input-dir data/satisfaction_detail/2026/3月 \
  --output-ppt data/ppt/2026/3月/3月满意度报告.pptx
```

使用配置文件：

```bash
uv run python generate_ppt.py --config ppt_job.example.toml
```

临时覆盖配置：

```bash
uv run python generate_ppt.py \
  --config ppt_job.example.toml \
  --section-mode summary \
  --blank-display '--'
```

只校验输入和布局，不写出 PPT：

```bash
uv run python generate_ppt.py \
  --config ppt_job.example.toml \
  --dry-run
```

常用命令行参数：

- `--config`：TOML 配置文件
- `--template-path`：模板 PPT 路径
- `--input-dir`：分项统计 Excel 输入目录
- `--output-ppt`：输出 PPT 路径
- `--section-mode`：二级标题识别口径
- `--blank-display`：空值展示文本
- `--dry-run`：只校验，不写出

---

## 8. 与完整业务流的关系

PPT 是新流程最后一步，依赖前面产出的满意度分项统计：

```text
data/raw/{year}/{batch}
  -> data/satisfaction_detail/{year}/{batch}
  -> data/satisfaction_summary/{year}/{batch}
  -> data/sample_summary/{year}/{batch}
  -> data/ppt/{year}/{batch}/{batch}满意度报告.pptx
```

如果 PPT 内容异常，优先检查：

- `data/satisfaction_detail/{year}/{batch}` 中对应客户类型 Excel 是否存在
- 分项 Excel 是否包含 `指标`、`满意度`、`重要性`
- 客户类型文件名是否能被 `survey_customer_category_rules.py` 识别
- `pipeline.defaults.toml` 中的模板路径、图表页和备注页配置是否正确
