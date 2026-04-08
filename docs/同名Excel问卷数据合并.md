# 同名 Excel 问卷数据合并

## 用途

- 接收多个输入目录
- 按 `xlsx` 文件名分组
- 读取每个文件的 `问卷数据` sheet
- 将同名文件的 `问卷数据` 按行合并
- 输出到指定目录，输出文件名保持原名

适合把多个批次、多个时间段或多个来源目录中的同名问卷文件汇总到一个目录中。

## 合并规则

- 只处理 `xlsx` 文件
- 默认只扫描输入目录当前层级；加 `--recursive` 后会递归扫描子目录
- 忽略 Excel 临时文件，例如 `~$xxx.xlsx`、`._xxx.xlsx`
- 输出文件只包含一个 `问卷数据` sheet
- 如果列名完全一致，会直接合并
- 如果列名相同但顺序不同，会按列名对齐后再合并
- 如果列名不一致，会跳过该同名文件，并输出列差异
- 如果同名文件中任意一个缺少 `问卷数据` sheet，也会跳过该同名文件

## 基本用法

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

自定义 sheet 名：

```bash
uv run python merge_questionnaire_workbooks.py \
  --input-dir './datas/目录A' \
  --input-dir './datas/目录B' \
  --output-dir './datas/合并结果' \
  --sheet-name '问卷数据'
```

## 输出说明

脚本会输出：

- 输出目录
- 输入目录数量
- 发现的文件名分组数量
- 合并成功数量
- 跳过/失败数量
- 每个同名文件的处理结果

如果列名不一致，会额外列出：

- 哪两个文件在对比
- 左侧文件独有的列名
- 右侧文件独有的列名
- 两边完整的列名列表

如果缺少 `问卷数据` sheet，会列出具体缺少的文件路径。
