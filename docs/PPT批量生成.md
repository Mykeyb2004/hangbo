# PPT 批量生成

## 目标

根据 [template.pptx](/Users/zhangqijin/PycharmProjects/hangbo/templates/template.pptx) 作为模板，
读取某个目录下的统计结果 Excel，
按“每个 Excel 对应一页 PPT”的方式批量生成汇总演示文稿。

当前实现默认适配 `survey_stats.py` 产出的三列表格：

- `指标`
- `满意度`
- `重要性`

## 当前规则

- 页面顺序默认按 [客户类型汇总表.md](/Users/zhangqijin/PycharmProjects/hangbo/docs/客户类型汇总表.md) 中“## 大类与样本类型”的顺序排列
- 若配置了 `category_intro_slides`，会在某个客户大类第一次出现前插入指定章节页
- 若配置了 `chart_page.enabled=true`，会在每个客户分组的数据页后插入 1 页图表页
- 页面标题默认使用 `客户大类——客户分组`
- 每个 Excel 生成 1 页 PPT
- 模板页顶部标题占位符会被替换为展示标题
- 总体行单独放在标题下方摘要表
- 明细区域默认优先渲染为 1 张全宽表
- 如果明细行超过单表容量，会按“二级标题”整体拆成左右 2 张表
- 空值默认不显示
- 若某个二级指标分组下所有三级指标的“满意度”都为空，则该二级指标分组整块不显示

补充说明：

- 如果 Excel 文件名是汇总表里的来源别名，例如 `展览主承办.xlsx`、`会议主承办.xlsx`、`参会人员.xlsx`，标题会自动转换成汇总展示名，如 `会展客户——展览活动主（承）办`
- 若同一个汇总行对应多个来源文件，也会强制统一成同一个汇总展示标题；例如 `酒店宴会.xlsx`、`酒店自助餐.xlsx` 都显示为 `酒店客户——餐饮客户`
- 章节页按客户大类触发，每个客户大类在一次批量生成中最多只插入 1 次；若该大类本次没有匹配到 Excel，则不会插入章节页
- 图表页标题与前一页数据页保持一致
- 图表页默认使用“左侧图表 + 右侧文字框”布局，右侧文字框可先放占位文案
- 二级指标数量为 `2` 时使用柱状图，`3` 个及以上使用雷达图；若过滤后不足 `2` 个二级指标，则跳过图表页

## 二级标题识别

二级标题优先复用 [survey_stats.py](/Users/zhangqijin/PycharmProjects/hangbo/survey_stats.py) 里的客户群体模板定义，
这样 PPT 的拆分口径和统计脚本保持一致。

支持两种口径：

- `template`
  - 按原始分表导出口径识别，例如 `会展服务`、`硬件设施`、`配套服务`
- `summary`
  - 按汇总口径识别，例如 `产品服务`、`智慧场馆/服务`
- `auto`
  - 自动选择与当前 Excel 内容更匹配的一套

## 配置文件

示例文件见 [ppt_job.example.toml](/Users/zhangqijin/PycharmProjects/hangbo/ppt_job.example.toml)。

常用参数：

- `template_path`
  - 模板 PPT 路径
- `input_dir`
  - Excel 输入目录
- `output_ppt`
  - 输出 PPT 路径
- `sheet_name_mode`
  - `first` 表示读取首个 sheet
  - `named` 表示读取指定 sheet
- `blank_display`
  - 空值显示文本，留空表示不显示
- `max_single_table_rows`
  - 明细行数不超过该值时用单表
- `max_split_table_rows`
  - 超出单表后，左右双表每侧能容纳的最大明细行数
- `layout.*`
  - 表格位置和尺寸，单位为英寸
- `category_intro_slides`
  - 按客户大类配置章节页来源
  - `ppt_path` 为章节页 PPT 路径
  - `slide_number` 为章节页页码，按 PowerPoint 可见页码习惯从 `1` 开始
- `chart_page`
  - 控制是否生成图表页
  - `enabled` 为 `true` 时，在每个客户分组数据页后追加图表页
  - `placeholder_text` 为图表页右侧文字框的回退内容；若当前页已生成备注分析，则优先复用备注文本
  - `image_dpi` 控制图表 PNG 清晰度
- `layout.chart_image`
  - 图表图片区域位置和尺寸，单位为英寸
- `layout.chart_textbox`
  - 右侧文字框区域位置和尺寸，单位为英寸

章节页配置示例：

```toml
[category_intro_slides."一、会展客户"]
ppt_path = "templates/chapter.pptx"
slide_number = 3

[category_intro_slides."二、餐饮客户"]
ppt_path = "templates/chapter.pptx"
slide_number = 4

[category_intro_slides."三、G20峰会体验馆"]
ppt_path = "templates/chapter.pptx"
slide_number = 6

[category_intro_slides."五、酒店客户"]
ppt_path = "templates/chapter.pptx"
slide_number = 5

[chart_page]
enabled = false
placeholder_text = "图表分析内容待补充。后续将在此处补充该客户分组二级指标的整体解读、优势项与待提升项。"
image_dpi = 220
```

### 备注页 LLM 分析

支持使用 OpenAI 兼容的 `openai` Python SDK，为每一页 PPT 自动生成约 300 字的数据分析，并写入备注页。

配置项位于：

```toml
[llm_notes]
enabled = false
env_path = ".env"
system_role_path = "system_role.md"
target_chars = 300
temperature = 0.4
max_tokens = 500
checkpoint_chars = 80
```

说明：

- `enabled`
  - 是否启用备注页分析
- `env_path`
  - `.env` 文件路径
- `system_role_path`
  - system role 提示词文件路径
- `target_chars`
  - 目标字数，默认约 300 字
- `temperature`
  - LLM 生成温度
- `max_tokens`
  - 单次生成的最大 token 数
- `checkpoint_chars`
  - 流式生成时累计到多少字符，就把当前备注写回备注页并保存一次检查点

`.env` 参考 [/.env.example](/Users/zhangqijin/PycharmProjects/hangbo/.env.example)：

```env
OPENAI_API_KEY=your-api-key
OPENAI_BASE_URL=https://your-openai-compatible-endpoint/v1
OPENAI_MODEL=gpt-4.1-mini
OPENAI_TEMPERATURE=0.4
```

当前代码会从 `.env` 中读取这些基础连接配置：

- `OPENAI_API_KEY` 或 `LLM_API_KEY`
- `OPENAI_BASE_URL` 或 `LLM_BASE_URL`
- `OPENAI_MODEL` 或 `LLM_MODEL`
- `OPENAI_TEMPERATURE` 或 `LLM_TEMPERATURE`（可选，会覆盖配置文件中的 `temperature`）
- `OPENAI_TIMEOUT` 或 `LLM_TIMEOUT`（可选）

## 运行方式

使用配置文件：

```bash
uv run python generate_ppt.py --config ppt_job.example.toml
```

启用备注页分析前，先复制环境变量模板并填写真实值：

```bash
cp .env.example .env
```

然后把 `ppt_job.example.toml` 中的 `[llm_notes].enabled` 改成 `true`，再执行生成命令。

启用后，终端会按页输出备注页分析进度，例如：

```text
[1/15] 正在生成备注页分析：专业观众
[1/15] 流式输出：本页数据显示，整体满意度和重要性均处于较高水平……
[1/15] 已保存检查点：9月满意度报告.partial.pptx
[1/15] 备注页分析完成：专业观众（298字）
```

如果 LLM 调用失败、超时、卡住后被手动中断，或用户使用 `Ctrl+C` 终止程序：

- 已成功完成的页面备注会保存在 `*.partial.pptx` 检查点文件中
- 当前页已流式接收到的部分文本，也会尽量写入备注并保存到检查点
- 成功全部生成后，检查点文件会自动删除，只保留正式输出文件

直接传参数：

```bash
uv run python generate_ppt.py \
  --template-path templates/template.pptx \
  --input-dir 输出结果/9月 \
  --output-ppt 输出结果/9月满意度报告.pptx
```

只校验输入和布局，不写出文件：

```bash
uv run python generate_ppt.py --config ppt_job.example.toml --dry-run
```

## 输出说明

- 输出文件为单个 `.pptx`
- 若配置了章节页，会先输出客户大类章节页，再输出该大类下的数据页
- 若启用了图表页，会在每个客户分组数据页后追加 1 页图表页
- 每页顶部是 `客户大类——客户分组` 标题
- 标题下方为总体满意度/重要性摘要
- 下方为明细指标表
- 图表页左侧为二级指标图表，右侧为说明文字框
- 若启用 `llm_notes`，图表页右侧会优先复用对应数据页的备注分析文本
- 若未启用 `llm_notes`，或当前页没有可用备注文本，则图表页右侧回退为 `chart_page.placeholder_text`
- 超长明细会拆成左右双表，但不会把同一个二级标题拆到两边
- 表格主题默认采用截图风格近似配色：
  - 表头为深酒红
  - 总体/二级标题行为玫粉色
  - 正文行为浅粉灰色
  - 表格分隔线为白色
- 若启用 `llm_notes`，每页备注页会额外写入一段基于当页表格数据的中文分析描述，图表页右侧说明文字框也会复用这段文本

## 测试

```bash
uv run python -m unittest discover -s tests
```

本功能对应测试文件：

- [test_generate_ppt.py](/Users/zhangqijin/PycharmProjects/hangbo/tests/test_generate_ppt.py)
