# Hangbo Survey Stats

### 完整输出流程，从客户分组分项文件，到满意度汇总
```bash
uv run python survey_stats.py --config job.toml --calculation-mode template
uv run python summary_table.py --input-dir '输出结果' --output-dir '汇总结果'
```



基于会展问卷 Excel 的统计脚本。

目前支持按模板原公式计算以下群体：
- `展览主承办`
- `参展商`
- `专业观众`
- `会展服务商`
- `会议主承办`
- `酒店会议主承办`
- `酒店参会客户`
- `参会人员`
- `散客`
- `住宿团队`
- `特色美食廊`
- `商务简餐`
- `旅游团餐`
- `宴会`
- `婚宴`
- `自助餐`
- `酒店宴会`
- `酒店自助餐`

脚本入口文件：
- [phase_column_preprocess.py](/Users/zhangqijin/PycharmProjects/hangbo/phase_column_preprocess.py)
- [survey_stats.py](/Users/zhangqijin/PycharmProjects/hangbo/survey_stats.py)
- [summary_table.py](/Users/zhangqijin/PycharmProjects/hangbo/summary_table.py)
- [check_start_time_month.py](/Users/zhangqijin/PycharmProjects/hangbo/check_start_time_month.py)
- [fill_year_month_columns.py](/Users/zhangqijin/PycharmProjects/hangbo/fill_year_month_columns.py)
- [merge_questionnaire_workbooks.py](/Users/zhangqijin/PycharmProjects/hangbo/merge_questionnaire_workbooks.py)

依赖通过 `uv` 管理，所有命令都建议使用 `uv run`。

## 安装依赖

在项目根目录执行：

```bash
uv sync
```

如果你只是直接运行脚本，也可以直接用 `uv run`，`uv` 会自动准备环境。

## 推荐流程

典型使用顺序是先跑 [survey_stats.py](/Users/zhangqijin/PycharmProjects/hangbo/survey_stats.py)，把原始问卷 `xlsx` 计算成“单群体统计结果”；再跑 [summary_table.py](/Users/zhangqijin/PycharmProjects/hangbo/summary_table.py)，把这些单群体统计结果汇总成客户类型总表。

```bash
uv run python survey_stats.py --config job.toml
uv run python summary_table.py --input-dir 输出结果 --output-dir 汇总结果
```

默认推荐使用 `template` 计算模式。

- `template`：按各客户群体模板原公式计算，适合当前正式流程，也是 [job.toml](/Users/zhangqijin/PycharmProjects/hangbo/job.toml) 的默认配置
- `summary`：按汇总表实际展示维度重组后再计算，会改变部分客户群体的二级指标结构和总体分口径

说明：
- [summary_table.py](/Users/zhangqijin/PycharmProjects/hangbo/summary_table.py) 的 `总分` 是直接读取输入分表中的总体满意度
- 所以只要先用同一批分表生成汇总表，汇总表 `总分` 与对应客户分组分表的 `总分` 就会一致
- `template` 不是“唯一能保证一致”的模式，但它是当前默认模式，也是保持“原模板公式口径”不变的推荐模式
- 如果改用 `summary` 模式，请先重新生成 [输出结果](/Users/zhangqijin/PycharmProjects/hangbo/输出结果) 再跑汇总；这时汇总表也会与分表一致，但总分的业务口径会和 `template` 模式不同

## `check_start_time_month.py` 用法

用途：
- 遍历指定目录下的 `xlsx`
- 只读取 `问卷数据` sheet
- 检查 `开始填表时间` 是否都属于同一个月
- 列出每个文件对应的月份，以及字段缺失、空值、跨月情况

基本用法：

```bash
uv run python check_start_time_month.py --input-dir '1-2月原始数据'
```

递归扫描子目录：

```bash
uv run python check_start_time_month.py --input-dir '1-2月原始数据' --recursive
```

输出会包含：
- 整体结论：是否都属于同一个月、检测到的月份
- 文件明细：逐个文件属于哪个月，是否跨月，是否缺少字段或 sheet

详细说明见：
- [docs/开始填表时间月份检查.md](/Users/zhangqijin/PycharmProjects/hangbo/docs/开始填表时间月份检查.md)

## `fill_year_month_columns.py` 用法

用途：
- 遍历指定目录下的 `xlsx`
- 只处理 `问卷数据` sheet
- 写入 `年份`、`月份` 两列
- 两列都按文本值写入；如果列已存在，则覆盖原值

基本用法：

```bash
uv run python fill_year_month_columns.py \
  --input-dir './datas/1-2月' \
  --year '2026' \
  --month '02'
```

递归扫描子目录：

```bash
uv run python fill_year_month_columns.py \
  --input-dir './datas/1-2月' \
  --year '2026' \
  --month '02' \
  --recursive
```

输出会包含：
- 扫描文件数
- 更新成功数量、跳过/失败数量
- 每个文件是否已更新，或是否缺少 `问卷数据` sheet
- 如果存在跳过文件，会在结尾单独列出被跳过的文件和原因

详细说明见：
- [docs/问卷数据年月填充.md](/Users/zhangqijin/PycharmProjects/hangbo/docs/问卷数据年月填充.md)

## `merge_questionnaire_workbooks.py` 用法

用途：
- 接收多个输入目录
- 按文件名合并这些目录中的 Excel 文件
- 只读取并输出 `问卷数据` sheet
- 同名列会合并到同一列
- 不同名列会追加到结果文件末尾的新列中

基本用法：

```bash
uv run python merge_questionnaire_workbooks.py \
  --input-dir './datas/1月' \
  --input-dir './datas/2月' \
  --output-dir './datas/合并结果'
```

递归扫描子目录：

```bash
uv run python merge_questionnaire_workbooks.py \
  --input-dir './datas/1月' \
  --input-dir './datas/2月' \
  --output-dir './datas/合并结果' \
  --recursive
```

输出会包含：
- 文件名分组数
- 合并成功数量、跳过/失败数量
- 每个同名文件是否已合并，或是否因为缺少 `问卷数据` sheet、存在重复列名而被跳过
- 如果后续文件出现新列，这些列会按发现顺序追加到输出表头最后

详细说明见：
- [docs/同名Excel问卷数据合并.md](/Users/zhangqijin/PycharmProjects/hangbo/docs/同名Excel问卷数据合并.md)

## `phase_column_preprocess.py` 用法

用途：
- 检查 Excel 指定 sheet 的第三列是否存在 `一期`、`二期` 这类期次标记
- 如果命中，就把第三列移动到最后一列并原地保存
- 如果第三列没命中，但别的列里发现了同类期次值，会提示“可能已经处理过”
- 在终端输出逐文件处理进度和结果提示

基本用法：

```bash
uv run python phase_column_preprocess.py 'datas/3月/展览.xlsx'
```

一次处理多个文件：

```bash
uv run python phase_column_preprocess.py datas/3月/*.xlsx
```

指定其他 sheet：

```bash
uv run python phase_column_preprocess.py 'datas/3月/展览.xlsx' --sheet-name '问卷数据'
```

输出会覆盖这些情况：
- 开始检查文件
- 文件不存在
- 缺少指定 sheet
- 文件列数不足，未发现第三列
- 第三列未检测到期次标记，但别的列发现了符合特征的值，提示可能已经处理过
- 整张表未发现期次特征列，无需处理
- 已完成预处理并保存
- 处理结束汇总，会列出成功处理、疑似已处理过、不含期次特征列、列数不足、失败等分类

详细说明见：
- [docs/问卷期次列预处理.md](/Users/zhangqijin/PycharmProjects/hangbo/docs/问卷期次列预处理.md)

## `survey_stats.py` 用法

用途：
- 输入原始问卷 `xlsx`
- 按客户群体模板计算满意度、重要性
- 输出单群体统计结果，支持 `xlsx`、`csv`、`md`
- 如果检测到 `问卷数据` sheet 的第三列含有 `一期`、`二期` 这类期次标记，会先把这一列移动到最后一列并原地保存，再继续统计

### 1. 配置文件批量模式

适合多个来源文件、多个客户群体、批量输出 10 个以上统计文件。

先参考示例配置：
- [report_jobs.example.toml](/Users/zhangqijin/PycharmProjects/hangbo/report_jobs.example.toml)

运行：

```bash
uv run python survey_stats.py --config report_jobs.example.toml
```

只校验并查看处理进度，不实际写文件：

```bash
uv run python survey_stats.py --config report_jobs.example.toml --dry-run
```

终端默认只输出文件处理进度；如果运行过程中触发了上面的期次列预处理，也会额外打印一条提示，说明哪个来源文件已经被自动调整。

只跑某几个 job：

```bash
uv run python survey_stats.py \
  --config report_jobs.example.toml \
  --job 展览主承办 \
  --job 参展商
```

覆盖输出目录或输出格式：

```bash
uv run python survey_stats.py \
  --config report_jobs.example.toml \
  --output-dir 输出结果 \
  --output-format xlsx
```

覆盖计算模式：

```bash
uv run python survey_stats.py \
  --config report_jobs.example.toml \
  --calculation-mode summary
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

常用参数：
- `--input`：原始问卷 Excel 文件
- `--template`：模板类型
- `--role-name`：来源 sheet 中用于筛选的客户分组名称
- `--output`：输出文件路径
- `--sheet-name`：来源 sheet 名，默认 `问卷数据`
- `--calculation-mode`：计算口径，支持 `template`、`summary`
- `--dry-run`：只校验并显示处理进度，不写文件

可选模板：
- `organizer`
- `exhibitor`
- `visitor`
- `service_provider`
- `meeting_organizer`
- `hotel_meeting_organizer`
- `hotel_meeting_attendee`
- `meeting_attendee`
- `hotel_individual_guest`
- `hotel_group_guest`
- `catering_food_hall`
- `catering_business_meal`
- `catering_tour_meal`
- `catering_banquet`
- `catering_wedding_banquet`
- `catering_buffet`
- `catering_hotel_banquet`
- `catering_hotel_buffet`

### 3. 兼容旧版三输入模式

当三个群体分别来自不同文件时可用：

```bash
uv run python survey_stats.py \
  --organizer-input '文件1.xlsx' \
  --exhibitor-input '文件2.xlsx' \
  --visitor-input '文件3.xlsx' \
  --output-dir '输出结果'
```

### 4. 输出内容

单群体统计结果默认包含：
- 总体行
- 一级维度行
- 明细指标行

导出为 `xlsx` 时会自动加基础样式：
- 总体行填充橙色
- 一级维度行填充浅绿色
- 文本加粗并居中

如果把 `--output-dir` 误传成了一个像文件名的值，例如 `输出文件.xlsx`，脚本会自动转成目录：

```text
输出文件_outputs/
```

相关说明见：
- [docs/问卷期次列预处理.md](/Users/zhangqijin/PycharmProjects/hangbo/docs/问卷期次列预处理.md)

## `summary_table.py` 用法

用途：
- 输入 `survey_stats.py` 导出的单群体统计结果 `xlsx`
- 按截图中的客户大类、样本类型、列映射汇总
- 输出 `客户类型满意度汇总表.xlsx`

### 1. 基本用法

```bash
uv run python summary_table.py \
  --input-dir '输出结果' \
  --output-dir '汇总结果'
```

### 2. 自定义输出文件名

```bash
uv run python summary_table.py \
  --input-dir '输出结果' \
  --output-dir '汇总结果' \
  --output-name '2026年1-2月客户类型汇总表.xlsx'
```

### 3. 递归扫描子目录

```bash
uv run python summary_table.py \
  --input-dir '输出结果' \
  --output-dir '汇总结果' \
  --recursive
```

### 4. 输入要求

- 输入目录中的文件需要是单群体统计结果 `xlsx`
- 第一行表头至少包含 `指标`、`满意度`
- 最稳妥的输入来源，就是直接使用 `survey_stats.py` 导出的 `xlsx`
- 汇总脚本会自动跳过不符合该结构的 `xlsx`

### 5. 输出规则

- 输出文件默认名为 `客户类型满意度汇总表.xlsx`
- 工作表名为 `汇总表`
- 顶部标题为 `杭博客户类型满意度情况表`
- 数值单元格固定显示 2 位小数
- `专项调研` 保留空行，当前不做数据匹配

映射规则见：
- [docs/客户类型汇总表.md](/Users/zhangqijin/PycharmProjects/hangbo/docs/客户类型汇总表.md)
- [docs/汇总文件对应关系.md](/Users/zhangqijin/PycharmProjects/hangbo/docs/汇总文件对应关系.md)

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
- `calculation_mode`：计算口径，支持 `template`、`summary`；默认 `template`
- `jobs[].name`：任务名，同时默认作为输出文件名和 sheet 名
- `jobs[].path`：该统计表对应的来源 Excel 文件
- `jobs[].sheet`：该统计表对应的来源 sheet
- `jobs[].template`：使用哪套统计模板
- `jobs[].role_name`：按来源 sheet 中哪个身份值筛选
- `jobs[].output_name`：可选，单独指定输出文件名
- `jobs[].output_format`：可选，覆盖全局输出格式

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
- `会议` 4 个客户分组按你确认后的修正版公式映射实现
- `酒店住宿` 2 个客户分组按最新版 `酒店过程分析.xlsx` 的公式映射实现
- `餐饮` 8 个客户分组按最新版 `餐饮过程分析.xlsx` 的公式映射实现
- 如果配置里指定了某个客户分组，但来源 `问卷数据` 中完全没有该分组记录，脚本会照常输出空白结果，并在全部任务结束后统一提示
- 其中包含模板本身已有的一些特殊列引用，没有做纠正

两种计算模式的差异见：
- [docs/survey_stats计算模式.md](/Users/zhangqijin/PycharmProjects/hangbo/docs/survey_stats计算模式.md)

## 测试

运行测试：

```bash
uv run python -m unittest discover -s tests
```

测试文件：
- [tests/test_survey_stats.py](/Users/zhangqijin/PycharmProjects/hangbo/tests/test_survey_stats.py)
- [tests/test_summary_table.py](/Users/zhangqijin/PycharmProjects/hangbo/tests/test_summary_table.py)
