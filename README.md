# Hangbo Survey Stats

基于会展问卷 Excel 的统计脚本。

目前支持按模板原公式计算以下群体：
- `展览主承办`
- `参展商`
- `专业观众`
- `会展服务商`

脚本入口文件：
- [survey_stats.py](/Users/zhangqijin/PycharmProjects/hangbo/survey_stats.py)

依赖通过 `uv` 管理，所有命令都建议使用 `uv run`。

## 安装依赖

在项目根目录执行：

```bash
uv sync
```

如果你只是直接运行脚本，也可以直接用 `uv run`，`uv` 会自动准备环境。

## 最常用用法

### 1. 配置文件批量模式

适合多个来源文件、多个客户群体、批量输出 10 个以上统计文件。

先参考示例配置：
- [report_jobs.example.toml](/Users/zhangqijin/PycharmProjects/hangbo/report_jobs.example.toml)

运行：

```bash
uv run python survey_stats.py --config report_jobs.example.toml
```

只校验和预览，不实际写文件：

```bash
uv run python survey_stats.py --config report_jobs.example.toml --dry-run
```

只跑某几个 job：

```bash
uv run python survey_stats.py \
  --config report_jobs.example.toml \
  --job 展览主承办 \
  --job 参展商
```

覆盖输出目录：

```bash
uv run python survey_stats.py \
  --config report_jobs.example.toml \
  --output-dir 输出结果
```

### 2. 单任务模式

适合临时只算一个群体。

```bash
uv run python survey_stats.py \
  --input '1-2月原始数据/展览-2.xlsx' \
  --template exhibitor \
  --role-name '参展商' \
  --output '输出结果/参展商.xlsx'
```

可选模板：
- `organizer`
- `exhibitor`
- `visitor`
- `service_provider`

### 3. 兼容旧版三输入模式

当三个群体分别来自不同文件时可用：

```bash
uv run python survey_stats.py \
  --organizer-input '文件1.xlsx' \
  --exhibitor-input '文件2.xlsx' \
  --visitor-input '文件3.xlsx' \
  --output-dir '输出结果'
```

## 配置文件说明

推荐把批量任务写进 TOML 配置文件。

最小示例：

```toml
output_dir = "输出结果"
output_format = "xlsx"

[[jobs]]
name = "展览主承办"
path = "1-2月原始数据/展览-2.xlsx"
sheet = "问卷数据"
template = "organizer"
role_name = "展览主承办"

[[jobs]]
name = "参展商"
path = "1-2月原始数据/展览-2.xlsx"
sheet = "问卷数据"
template = "exhibitor"
role_name = "参展商"

[[jobs]]
name = "专业观众"
path = "1-2月原始数据/展览-2.xlsx"
sheet = "问卷数据"
template = "visitor"
role_name = "专业观众"

[[jobs]]
name = "会展服务商"
path = "1-2月原始数据/7月会展服务商.xlsx"
sheet = "问卷数据"
template = "service_provider"
role_name = "会展服务商"
```

字段说明：
- `output_dir`：输出目录
- `output_format`：默认输出格式，支持 `xlsx`、`csv`、`md`
- `jobs[].name`：任务名，同时默认作为输出文件名和 sheet 名
- `jobs[].path`：该统计表对应的来源 Excel 文件
- `jobs[].sheet`：该统计表对应的来源 sheet
- `jobs[].template`：使用哪套统计模板
- `jobs[].role_name`：按来源 sheet 中哪个身份值筛选
- `jobs[].output_name`：可选，单独指定输出文件名
- `jobs[].output_format`：可选，覆盖全局输出格式

## 输出结果

默认支持以下格式：
- `xlsx`
- `csv`
- `md`

导出为 `xlsx` 时会自动添加基础样式：
- 总体行填充橙色
- 一级维度行填充浅绿色
- 文本加粗并居中

如果把 `--output-dir` 误传成了一个像文件名的值，例如 `输出文件.xlsx`，脚本会自动转成目录：

```text
输出文件_outputs/
```

## 计算规则

统计逻辑按模板原公式实现：
- 从来源文件的 `问卷数据` sheet 读取数据
- 按身份字段分组
- 只统计分值 `>0` 且 `<11` 的记录
- 明细指标先算平均值
- 一级维度再对明细指标求平均
- 总体满意度和重要性对一级维度求平均

注意：
- `参展商` 和 `专业观众` 当前是按模板原公式原样实现
- `会展服务商` 也按模板原公式原样实现
- 其中包含模板本身已有的一些特殊列引用，没有做纠正

## 测试

运行测试：

```bash
uv run python -m unittest discover -s tests
```

测试文件：
- [tests/test_survey_stats.py](/Users/zhangqijin/PycharmProjects/hangbo/tests/test_survey_stats.py)
