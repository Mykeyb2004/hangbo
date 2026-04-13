# survey_stats 目录模式

## 目的

`survey_stats.py` 现在支持“统一指定目录”的批量模式。

和旧的 `[[jobs]]` 显式配置相比，目录模式只需要提供一个 `input_dir`，程序会根据标准映射自动尝试发现当前已支持的 20 类模板客户，并只为真实存在数据的客户类型生成结果。

标准映射表见：

- [模板客户标准映射表.md](/Users/zhangqijin/PycharmProjects/hangbo/docs/模板客户标准映射表.md)
- [survey_stats配置结构说明.md](/Users/zhangqijin/PycharmProjects/hangbo/docs/survey_stats配置结构说明.md)

## 基本配置

```toml
output_dir = "输出结果/3月"
output_format = "xlsx"
calculation_mode = "template"
input_dir = "datas/3月"
```

运行方式：

```bash
uv run python survey_stats.py --config report_jobs.directory.example.toml
```

## 可选配置

### `sheet_name`

全局来源 sheet 名，默认是 `问卷数据`。

```toml
sheet_name = "问卷数据"
```

### `source_file_overrides`

当目录中的真实文件名不是标准名时，可以显式覆盖。

标准来源文件名只有这 6 个：

- `展览.xlsx`
- `会展服务商.xlsx`
- `会议.xlsx`
- `旅游.xlsx`
- `酒店.xlsx`
- `餐饮.xlsx`

示例：

```toml
output_dir = "输出结果/历史批次"
output_format = "xlsx"
input_dir = "datas/历史批次"

[source_file_overrides]
"会展服务商.xlsx" = "7月会展服务商.xlsx"
"酒店.xlsx" = "8月酒店过程分析.xlsx"
```

## 行为说明

- 旧的 `[[jobs]]` 配置模式仍然保留，目录模式是新增能力，不影响旧流程。
- 目录模式下，如果标准来源文件不存在，对应客户类型会被跳过。
- 目录模式下，如果来源文件存在，但其中没有该客户类型对应的身份值，也会被跳过。
- 以上两类跳过都不会生成空白输出文件，只会在程序结尾统一提示。
- 结尾提示会区分两种原因：
  - 缺少来源文件
  - 来源文件存在但未找到匹配身份值

## 与旧 `[[jobs]]` 模式的差异

- 旧模式：每个 job 手动指定 `path/template/role_name`
- 新模式：只指定 `input_dir`，再按标准映射自动展开
- 旧模式：客户分组无匹配时仍输出空白统计结果
- 新模式：客户类型无数据时不输出空白文件，只在结尾提示

## 推荐场景

- 目录里文件命名已经接近标准口径，或只需要少量覆盖
- 每月、季度都要重复跑同一批模板客户
- 希望自动发现“这个目录里缺了哪些客户类型”
