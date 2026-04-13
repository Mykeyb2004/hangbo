# survey_stats 计算模式

## 配置开关

`survey_stats.py` 现在支持两套计算逻辑，通过 TOML 配置中的 `calculation_mode` 切换：

```toml
output_dir = "输出结果"
output_format = "xlsx"
calculation_mode = "template"
```

可选值：

- `template`
  - 默认模式
  - 按 [survey_stats.py](/Users/zhangqijin/PycharmProjects/hangbo/survey_stats.py) 模板中定义的原始二级指标直接计算
- `summary`
  - 汇总口径模式
  - 参考 [summary_table.py](/Users/zhangqijin/PycharmProjects/hangbo/summary_table.py) 的灰格/非灰格，按汇总表真实参与统计的二级指标重新计算

## `summary` 模式会做什么

### 会展类、会议类、酒店会议类、酒店参会类

- `会展服务` 或 `会场服务` 统一重组为 `产品服务`
- `智慧场馆` 统一重组为 `智慧场馆/服务`
- `配套服务` 会剔除其中的 `餐饮服务` 明细，避免和独立的 `餐饮服务` 二级指标重复计入
- `餐饮服务` 会单独作为一个二级指标输出

### 酒店散客、住宿团队

- `入住服务` 统一重组为 `产品服务`
- `智慧场馆` 统一重组为 `智慧场馆/服务`
- 灰格对应的 `配套服务` 不参与计算，也不输出

### 餐饮类客户

- 仅保留汇总表非灰格维度：
  - `硬件设施`
  - `智慧场馆/服务`
  - `餐饮服务`

## 运行示例

```bash
uv run python survey_stats.py --config job.toml
```

如需临时覆盖配置，也可以直接在命令行传：

```bash
uv run python survey_stats.py --config job.toml --calculation-mode summary
```

## 说明

- `template` 模式保持现有结果不变，兼容旧流程。
- `summary` 模式会改变二级指标结构，因此总体分、二级指标分和导出行顺序都可能与原模式不同。
- `旅游团餐` 虽然不在汇总表成品里单独展示，但在 `summary` 模式下按餐饮类客户口径处理。
- `会议.xlsx` 里的会议类/酒店会议类客户分流，不能只看身份列；当前实现会联合使用 `Q1-调研类别` 和 `Q3-您在会议中的身份`：
  - `会议` + `会议主承办` -> `会议活动主（承）办`
  - `会议` + `参会人员` -> `参会客户`
  - `酒店会议` + `会议主承办` -> `酒店会议活动主（承）办`
  - `酒店会议` + `参会人员` -> `酒店参会客户`
- [summary_table.py](/Users/zhangqijin/PycharmProjects/hangbo/summary_table.py) 的 `总分` 是直接读取输入分表中的总体满意度，所以“汇总表总分是否与客户分组分表一致”取决于你是否先用同一批分表再去汇总。
- `template` 模式是当前默认和推荐模式，因为它保持原模板公式口径不变；`summary` 模式并不是一定会导致“不一致”，但它会改变总分本身的计算口径。
