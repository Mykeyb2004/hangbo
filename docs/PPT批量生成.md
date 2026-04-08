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

- 页面标题使用 Excel 文件名（去掉 `.xlsx`）
- 每个 Excel 生成 1 页 PPT
- 模板页顶部标题占位符会被替换为文件名
- 总体行单独放在标题下方摘要表
- 明细区域默认优先渲染为 1 张全宽表
- 如果明细行超过单表容量，会按“二级标题”整体拆成左右 2 张表
- 空值默认不显示

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

## 运行方式

使用配置文件：

```bash
uv run python generate_ppt.py --config ppt_job.example.toml
```

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
- 每页顶部是客户群体标题
- 标题下方为总体满意度/重要性摘要
- 下方为明细指标表
- 超长明细会拆成左右双表，但不会把同一个二级标题拆到两边

## 测试

```bash
uv run python -m unittest discover -s tests
```

本功能对应测试文件：

- [test_generate_ppt.py](/Users/zhangqijin/PycharmProjects/hangbo/tests/test_generate_ppt.py)
