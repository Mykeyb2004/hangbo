# 预处理新版数据结构（把第几期的字段放最后）
`uv run python phase_column_preprocess.py datas/3月/*.xlsx`

# 合并数据
```bash
uv run python /Users/zhangqijin/PycharmProjects/hangbo/merge_questionnaire_workbooks.py \
  --input-dir 'datas/1-2月' \
  --input-dir 'datas/3月' \
  --output-dir 'datas/合并结果'
```

# 为数据加上年份+月份标记
uv run python fill_year_month_columns.py \
  --input-dir './datas/1-2月' \
  --year '2026' \
  --month '01-02'

uv run python fill_year_month_columns.py \
  --input-dir './datas/3月' \
  --year '2026' \
  --month '03'

# 统计客群分组
uv run python survey_stats.py --config job_Q1.toml

# 统计汇总
uv run python summary_table.py --input-dir "输出结果/Q1" --output-dir "汇总结果/Q1" --output-name "Q1客户类型满意度汇总表.xlsx"

# 生成PPT
uv run python generate_ppt.py --config ppt_job.example.toml
uv run python generate_ppt.py --config report_jobs.Q1.toml
