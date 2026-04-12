# survey_stats 配置结构说明

## 目的

`survey_stats.py` 当前支持两种批量配置结构：

- 旧结构：显式 `[[jobs]]` 模式
- 新结构：`input_dir` 目录模式

两种模式互斥，同一个配置文件中只能选择一种。

## 通用字段

两种模式都支持这些顶层字段：

| 字段 | 是否必填 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `output_dir` | 否 | 配置文件同目录下自动生成 `xxx_outputs` | 输出目录 |
| `output_format` | 否 | `xlsx` | 支持 `xlsx`、`csv`、`md` |
| `calculation_mode` | 否 | `template` | 支持 `template`、`summary` |
| `sheet_name` | 否 | `问卷数据` | 全局来源 sheet 名 |

## 旧结构：`[[jobs]]` 模式

### 适用场景

- 每个客户类型的数据源都要手动指定
- 来源文件名不稳定，且不想维护目录模式映射
- 想显式控制每个输出文件名、模板和身份值

### 结构示例

```toml
output_dir = "输出结果"
output_format = "xlsx"
calculation_mode = "template"
sheet_name = "问卷数据"

[[jobs]]
name = "展览主承办"
path = "datas/1-2月/展览.xlsx"
template = "organizer"
role_name = "展览主承办"

[[jobs]]
name = "专业买家-3月"
path = "datas/3月/展览.xlsx"
template = "visitor"
role_name = "专业买家"
output_name = "专业买家-3月"
output_format = "csv"
```

### `[[jobs]]` 字段说明

| 字段 | 是否必填 | 说明 |
| --- | --- | --- |
| `name` | 是 | job 名称，同时默认作为工作日志和 sheet 标题 |
| `path` | 是 | 来源 Excel 路径，相对配置文件解析 |
| `template` | 是 | 模板名，例如 `organizer`、`visitor` |
| `role_name` | 否 | 问卷中的身份值；默认取模板内置角色名 |
| `sheet` | 否 | 当前 job 的 sheet 名；默认取全局 `sheet_name` |
| `output_name` | 否 | 输出文件名；默认等于 `name` |
| `output_format` | 否 | 当前 job 单独覆盖输出格式 |

### 旧结构行为

- 每个 `job` 都会被直接执行。
- 如果来源文件存在，但没有匹配的身份值，仍会生成空白统计结果文件。
- 结尾会统一提示哪些客户分组在来源数据中未找到匹配记录。

## 新结构：`input_dir` 目录模式

### 适用场景

- 同一批数据目录会反复按固定模板客户跑统计
- 来源文件名基本稳定，只需少量覆盖
- 想自动知道当前目录里缺了哪些客户类型

### 结构示例

```toml
output_dir = "输出结果/3月"
output_format = "xlsx"
calculation_mode = "template"
input_dir = "datas/3月"
sheet_name = "问卷数据"
```

带文件名覆盖的示例：

```toml
output_dir = "输出结果/历史批次"
output_format = "xlsx"
input_dir = "datas/历史批次"

[source_file_overrides]
"会展服务商.xlsx" = "7月会展服务商.xlsx"
"酒店.xlsx" = "8月酒店过程分析.xlsx"
```

### 目录模式字段说明

| 字段 | 是否必填 | 说明 |
| --- | --- | --- |
| `input_dir` | 是 | 来源目录，相对配置文件解析 |
| `source_file_overrides` | 否 | 标准来源文件名到真实文件名的覆盖表 |

### 目录模式的标准来源文件

目录模式会按标准映射自动尝试查找这些来源文件：

- `展览.xlsx`
- `会展服务商.xlsx`
- `会议.xlsx`
- `酒店.xlsx`
- `餐饮.xlsx`

完整映射表见：

- [模板客户标准映射表.md](/Users/zhangqijin/PycharmProjects/hangbo/docs/模板客户标准映射表.md)

### 新结构行为

- 程序会自动展开当前已支持的 18 类模板客户。
- 如果来源文件不存在，该客户类型会被跳过。
- 如果来源文件存在，但没有匹配身份值，该客户类型也会被跳过。
- 以上两种情况都不会生成空白输出文件。
- 结尾会按原因分组提示：
  - 缺少来源文件
  - 来源文件存在但未找到匹配身份值

## 两种结构不能混用

以下写法不允许：

```toml
input_dir = "datas/3月"

[[jobs]]
name = "展览主承办"
path = "datas/3月/展览.xlsx"
template = "organizer"
```

原因：

- `input_dir` 表示自动发现任务
- `[[jobs]]` 表示手动声明任务
- 两者同时存在会造成配置语义冲突

## 字段选择建议

- 目录已经按标准文件名整理好：优先用 `input_dir`
- 目录里存在大量历史文件名：优先用 `input_dir + source_file_overrides`
- 单个客户类型需要特殊 `role_name` 或特殊输出文件名：继续用 `[[jobs]]`
- 希望缺数据时仍输出空白报表：继续用 `[[jobs]]`

## 从旧结构迁移到新结构

如果一批 `[[jobs]]` 满足下面两个条件，就很适合迁移成目录模式：

- 多个 job 只是重复引用同一个目录下的标准来源文件
- 不需要为某些 job 单独改 `output_name`、`role_name`、`output_format`

### 迁移前

```toml
output_dir = "输出结果/Q1"
output_format = "xlsx"
calculation_mode = "template"

[[jobs]]
name = "展览主承办"
path = "datas/合并结果/展览.xlsx"
template = "organizer"
role_name = "展览主承办"

[[jobs]]
name = "参展商"
path = "datas/合并结果/展览.xlsx"
template = "exhibitor"
role_name = "参展商"
```

### 迁移后

```toml
output_dir = "输出结果/Q1"
output_format = "xlsx"
calculation_mode = "template"
input_dir = "datas/合并结果"
```

## 相关文档

- [survey_stats目录模式.md](/Users/zhangqijin/PycharmProjects/hangbo/docs/survey_stats目录模式.md)
- [模板客户标准映射表.md](/Users/zhangqijin/PycharmProjects/hangbo/docs/模板客户标准映射表.md)
- [report_jobs.example.toml](/Users/zhangqijin/PycharmProjects/hangbo/report_jobs.example.toml)
- [report_jobs.directory.example.toml](/Users/zhangqijin/PycharmProjects/hangbo/report_jobs.directory.example.toml)
