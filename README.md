# Hangbo Survey Stats

基于会展问卷 Excel 的数据分析流水线。

当前仓库的**唯一推荐主流程**是：

`原始数据目录 -> 预查错 -> 人工修正确认 -> 满意度分项统计 -> 满意度汇总 -> 样本汇总 -> PPT`

主流程入口：

```bash
uv run python main_pipeline.py --year 2026 --batch 3月
```

---

## 核心说明

- 现在的主流程围绕 `data/...` 目录约定运行，不再以 `datas/ / 输出结果 / 汇总结果` 作为主入口。
- 主流程会先做预查错；如果发现阻断问题，会停下来等你修改原始数据后再继续。
- 某个月份缺少某几个客户分组来源文件是正常情况，预查错**不会**因为这类情况而阻断。
- 对于单月批次，若缺少 `年份` / `月份` 列，主流程可以自动补齐。
- 对于合并批次，例如 `1-2月`、`Q1`，若缺少 `年份` / `月份` 列，主流程会阻断，需要先人工准备好。
- `Q1` 在当前新流程中是一个**已准备好的正式批次目录**，不是运行时自动由 `1-2月 + 3月` 现拼出来的。

---

## 安装依赖

项目使用 `uv` 作为包管理器。

初始化环境：

```bash
uv sync
```

如果只是直接运行脚本，也可以直接使用 `uv run`，`uv` 会自动准备运行环境。

---

## 新流程目录约定

### 输入目录

主流程固定从以下目录读取原始数据：

- `data/raw/{year}/{batch}`

例如：

- `data/raw/2026/3月`
- `data/raw/2026/1-2月`
- `data/raw/2026/Q1`

### 输出目录

主流程固定输出到：

- `data/satisfaction_detail/{year}/{batch}`
- `data/satisfaction_summary/{year}/{batch}`
- `data/sample_summary/{year}/{batch}`
- `data/ppt/{year}/{batch}`

例如：

- `data/satisfaction_detail/2026/3月`
- `data/satisfaction_summary/2026/3月`
- `data/sample_summary/2026/3月`
- `data/ppt/2026/3月/3月满意度报告.pptx`

### 标准来源文件

主流程按标准来源文件约定扫描批次目录：

- `展览.xlsx`
- `会议.xlsx`
- `酒店.xlsx`
- `餐饮.xlsx`
- `会展服务商.xlsx`
- `旅游.xlsx`

说明：

- 某个批次目录里只存在其中一部分文件是允许的。
- 如果整个批次目录里一个标准来源文件都没有，主流程会阻断。

---

## 主流程怎么跑

### 单月批次

例如处理 `3月`：

```bash
uv run python main_pipeline.py --year 2026 --batch 3月
```

例如处理 `4月`：

```bash
uv run python main_pipeline.py --year 2026 --batch 4月
```

### 合并批次

例如处理已经准备好的 `1-2月`：

```bash
uv run python main_pipeline.py --year 2026 --batch 1-2月
```

例如处理已经准备好的 `Q1`：

```bash
uv run python main_pipeline.py --year 2026 --batch Q1
```

### 主流程内部执行顺序

主流程会按以下顺序自动执行：

1. 预查错
2. 人工修正后确认继续
3. 生成满意度分项统计
4. 生成满意度汇总表
5. 生成样本统计表
6. 生成 PPT

---

## 预查错会检查什么

主流程的第一步是预查错。

### 会阻断的情况

- 原始批次目录不存在
- 批次目录里完全没有标准来源文件
- 合并批次缺少 `年份` / `月份` 列
- 存在未映射标签记录
- 缺少指定 `sheet`
- Excel 文件读取失败

### 不会因为这些情况阻断

- 只缺少某几个来源文件
- 某个月份没有某些客户分组数据

这类情况在当前业务里是正常的。

### 发现阻断问题后如何继续

主流程会提示你先修改原始数据目录。

修改完成后，在终端输入以下任一内容即可继续：

- `y`
- `yes`
- `继续`

如果输入：

- `stop`
- `quit`
- `exit`

主流程会终止。

### 预查错日志

主流程会在以下目录生成日志：

- `logs/pipeline/{year}/{batch}/precheck.log`
- `logs/pipeline/{year}/{batch}/unmapped_customer_records.log`

---

## 年份 / 月份是怎么处理的

### 单月批次

对于 `1月` 到 `12月` 这种单月批次：

- 如果原始文件缺少 `年份` / `月份` 列
- 主流程会自动调用 `fill_year_month_columns.py` 补齐

例如：

```bash
uv run python main_pipeline.py --year 2026 --batch 3月
```

如果 `data/raw/2026/3月` 缺少这两列，程序会自动把所有数据行补成：

- `年份 = 2026`
- `月份 = 3`

### 合并批次

对于：

- `1-2月`
- `Q1`

这类合并批次，程序不能自动判断每一行属于哪一个月。

因此：

- 缺少 `年份` / `月份` 时不会自动补
- 必须先人工准备好后，再进入主流程

### 手工补写年份 / 月份

如果你要先对单月原始目录做增强，可以执行：

```bash
uv run python fill_year_month_columns.py \
  --input-dir 'data/raw/2026/3月' \
  --year '2026' \
  --month '3'
```

说明：

- 这个脚本是**整文件批量写入**
- 它会把该文件中的所有数据行写成同一个年份、同一个月份
- 所以它适合单月目录，不适合直接拿来处理混合月份目录

---

## 多月如何合并

### 推荐处理方式

如果要把多个月份整理成季度或合并批次，推荐顺序是：

1. 先给每个单月目录补 `年份` / `月份`
2. 再执行多月合并
3. 把合并结果放到 `data/raw/{year}/{batch}`
4. 最后运行主流程

### 合并脚本

多月合并由 `merge_questionnaire_workbooks.py` 完成。

示例：

```bash
uv run python merge_questionnaire_workbooks.py \
  --input-dir 'data/raw/2026/1月' \
  --input-dir 'data/raw/2026/2月' \
  --input-dir 'data/raw/2026/3月' \
  --output-dir 'data/raw/2026/Q1'
```

### 合并规则

这个脚本的逻辑是：

- 按文件名分组，例如把多个目录里的 `餐饮.xlsx` 放在一起处理
- 只读取 `问卷数据` sheet
- 表头按“语义”对齐，会忽略 `Q1-`、`Q2-` 这类题号前缀
- 数据行在统一表头下直接追加到结果文件中

### 它不会自动做的事情

多月合并脚本**不会**自动做这些事：

- 不按 `提交序号` 去重
- 不自动保留“最新”记录
- 不自动修正客户标签
- 不自动处理标签冲突

所以它本质上是：

- 同名问卷文件的表头对齐 + 行拼接工具

### 关于 Q1 的当前理解

在当前新流程里：

- `Q1` 不是运行时动态计算概念
- `Q1` 是一个你已经准备好的原始数据目录

也就是说，只有当：

- `data/raw/2026/Q1`

已经准备好之后，才推荐执行：

```bash
uv run python main_pipeline.py --year 2026 --batch Q1
```

---

## 最终会生成什么

主流程跑通后，会自动生成以下结果：

### 1. 满意度分项统计

目录：

- `data/satisfaction_detail/{year}/{batch}`

内容：

- 每个客户类型一份明细统计结果 `xlsx`

### 2. 满意度汇总表

目录：

- `data/satisfaction_summary/{year}/{batch}`

内容：

- `{batch}客户类型满意度汇总表.xlsx`

### 3. 样本统计表

目录：

- `data/sample_summary/{year}/{batch}`

内容：

- `{batch}客户类型样本统计表.xlsx`

### 4. PPT 报告

目录：

- `data/ppt/{year}/{batch}`

内容：

- `{batch}满意度报告.pptx`

---

## PPT 生成说明

主流程最后一步会自动调用 PPT 生成。

默认输出位置：

- `data/ppt/{year}/{batch}/{batch}满意度报告.pptx`

例如：

- `data/ppt/2026/3月/3月满意度报告.pptx`
- `data/ppt/2026/Q1/Q1满意度报告.pptx`

如果 `pipeline.defaults.toml` 中开启了 `llm_notes`，则：

- PPT 生成阶段会调用模型生成备注页分析
- 运行前需确认 `.env` 和 `system_role.md` 可用

---

## GUI 入口

如果你希望通过桌面界面操作，可以运行：

```bash
uv run python hangbo_gui.py
```

GUI 当前也围绕同一套新流程工作，主线仍然是：

- 数据源准备
- 可选预处理
- 预查错
- 分项统计
- 汇总统计
- 样本统计
- PPT 生成

其中：

- 多月合并只在“多月模式”下按需执行
- 不属于单月模式默认步骤

---

## 关键脚本

面向新流程，最重要的脚本如下：

- `main_pipeline.py`
  - 新主流程入口
- `pipeline_precheck.py`
  - 主流程预查错
- `fill_year_month_columns.py`
  - 单月目录补写 `年份` / `月份`
- `merge_questionnaire_workbooks.py`
  - 多月问卷合并
- `survey_stats.py`
  - 满意度分项统计引擎
- `summary_table.py`
  - 满意度汇总表
- `sample_table.py`
  - 样本统计表
- `generate_ppt.py`
  - PPT 生成

---

## 推荐操作模板

### 模板 A：单月直接跑

```bash
uv run python main_pipeline.py --year 2026 --batch 3月
```

适合：

- 原始数据已经放在 `data/raw/2026/3月`
- 只需要完成这个月的全流程输出

### 模板 B：先合并季度，再跑季度

```bash
uv run python fill_year_month_columns.py --input-dir 'data/raw/2026/1月' --year '2026' --month '1'
uv run python fill_year_month_columns.py --input-dir 'data/raw/2026/2月' --year '2026' --month '2'
uv run python fill_year_month_columns.py --input-dir 'data/raw/2026/3月' --year '2026' --month '3'

uv run python merge_questionnaire_workbooks.py \
  --input-dir 'data/raw/2026/1月' \
  --input-dir 'data/raw/2026/2月' \
  --input-dir 'data/raw/2026/3月' \
  --output-dir 'data/raw/2026/Q1'

uv run python main_pipeline.py --year 2026 --batch Q1
```

适合：

- 要按规范重新准备季度批次
- 需要确保季度数据保留正确年月信息

### 模板 C：已存在历史合并批次

```bash
uv run python main_pipeline.py --year 2026 --batch 1-2月
uv run python main_pipeline.py --year 2026 --batch Q1
```

前提：

- 对应目录已经准备好
- 合并批次中的 `年份` / `月份` 信息已经正确存在

---

## 相关文档

- [docs/README.md](/Users/zhangqijin/PycharmProjects/hangbo/docs/README.md)
- [docs/新数据分析流程说明.md](/Users/zhangqijin/PycharmProjects/hangbo/docs/新数据分析流程说明.md)
- [docs/数据准备与预查错.md](/Users/zhangqijin/PycharmProjects/hangbo/docs/数据准备与预查错.md)
- [docs/统计口径与结果说明.md](/Users/zhangqijin/PycharmProjects/hangbo/docs/统计口径与结果说明.md)
- [docs/PPT生成说明.md](/Users/zhangqijin/PycharmProjects/hangbo/docs/PPT生成说明.md)

如果只记一句话，请记这句：

**单月直接跑 `main_pipeline.py`；多月先补年月、再合并、最后再跑 `main_pipeline.py`。**
