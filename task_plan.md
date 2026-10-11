# Task Plan: Excel 批量生成 PPT 增强规划

## 当前任务：TOML 配置写入 Excel（2026-10-09）

- 目标：新增独立 `insert_excel_values.py`，按 TOML 写入单元格并保留原排版。
- [x] 用户确认方案；确认 Python >=3.11、已有 openpyxl，无需新增依赖。
- [x] 检查工作区：现有图谱及业务 Excel 改动不属于本任务，保持原状。
- [x] GitNexus impact：新脚本未索引，目标不存在（UNKNOWN），不修改现有业务符号。
- [x] TDD：命令行作为已确认的测试入口，每轮先失败、后实现；16 项定向测试通过。
- [x] 示例 TOML 与 docs 文档、入口索引。
- [x] 定向测试、全量回归、差异检查及交付；全量 229 项中 3 项已有失败、1 项跳过，已独立复现并记录。
- 路径相对 TOML 所在目录解析；默认另存，覆盖需 overwrite=true。
- 保留已有规划记录，在本节及 findings/progress 追加本任务进展。
- 工具错误：误读技能路径已纠正；uv 默认缓存不可写，改用 /private/tmp/hangbo-uv-cache；代码图查询超时不阻碍独立新增功能。

## Goal
在现有 `generate_ppt.py` 基础上，为每个客户分组新增图表页能力：基于二级指标绘制柱状图或雷达图，将图表与占位文字框并列展示，并保持独立模块封装、可配置、可测试。

## Current Phase
Phase 5

## Phases

### Phase 1: Requirements & Scope Freeze
- [x] 明确图表页触发条件、页内布局和占位文案策略
- [x] 明确图表数据口径与空值过滤规则
- [x] 记录依赖与实现约束
- **Status:** complete

### Phase 2: Module Design
- [x] 设计独立图表模块 API
- [x] 设计图表页配置结构与布局参数
- [x] 定义生成中间产物与临时文件管理方式
- **Status:** complete

### Phase 3: TDD & Implementation
- [x] 先补图表数据提取/图表类型判断测试
- [x] 实现图表图片生成模块
- [x] 将图表页接入 PPT 主生成流程
- **Status:** complete

### Phase 4: Verification
- [x] 校验图表页顺序、标题复用和图片插入
- [x] 用真实 Q1 数据生成并人工验收
- [x] 更新 progress.md 中的测试记录
- **Status:** complete

### Phase 5: Documentation & Delivery
- [x] 更新 `docs/PPT批量生成.md`
- [x] 更新示例 TOML 配置
- [x] 向用户交付运行命令与注意事项
- **Status:** complete

## Key Questions
1. 图表是否默认展示“满意度 + 重要性”双系列？
2. 图表页的右侧文字框是否先用固定占位文案，后续再接 LLM？
3. 图表绘制依赖是否引入 `matplotlib` 作为新依赖？

## Decisions Made
| Decision | Rationale |
|----------|-----------|
| 图表能力先独立封装到新文件 | 降低 `generate_ppt.py` 复杂度，便于单测 |
| 图表页放在每个客户分组表格页之后 | 用户已明确要求 |
| 图表页标题与前一页保持一致 | 用户已明确要求 |
| 二级指标数为 2 时使用柱状图，3 个及以上用雷达图 | 用户已明确要求 |
| 图表与文字框左右并列 | 参考用户给出的版式截图 |
| 右侧文字框内容先占位 | 用户明确暂时无需生成正式分析文案 |
| 图表数据直接复用当前二级指标口径 | 保证表格页与图表页一致 |

## Errors Encountered
| Error | Attempt | Resolution |
|-------|---------|------------|
|       | 1       |            |

## Notes
- 遵循 AGENTS.md：使用 `uv` 作为包管理器，测试通过 `uv run` 执行。
- 若开始实现，需要同步补充 `/docs` 技术文档和 `/tests` 测试脚本。
- 图表页预计需要新增绘图库依赖，优先考虑 `matplotlib`。
- 当前已完成独立图表模块、主流程接入、文档更新和 Q1 实际产物验证。
