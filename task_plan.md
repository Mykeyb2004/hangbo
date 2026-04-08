# Task Plan: Excel 批量生成 PPT 方案设计

## Goal
基于 `templates/template.pptx` 和 `输出结果/9月/` 下的 Excel 文件，设计一套参数可配置的 PPT 生成方案，并在用户确认后进入实现。

## Current Phase
Phase 5

## Phases

### Phase 1: Requirements & Discovery
- [x] Understand user intent
- [x] Identify constraints and requirements
- [x] Document findings in findings.md
- **Status:** complete

### Phase 2: Planning & Structure
- [x] Define technical approach
- [x] Create project structure if needed
- [x] Document decisions with rationale
- **Status:** complete

### Phase 3: Implementation
- [x] Execute the plan step by step
- [x] Write code to files before executing
- [x] Test incrementally
- **Status:** complete

### Phase 4: Testing & Verification
- [x] Verify all requirements met
- [x] Document test results in progress.md
- [x] Fix any issues found
- **Status:** complete

### Phase 5: Delivery
- [x] Review all output files
- [x] Ensure deliverables are complete
- [x] Deliver to user
- **Status:** complete

## Key Questions
1. `template.pptx` 中用于标题、表格、图表的占位方式是什么？
2. Excel 的读取规则应是“整表搬运”还是“按命名区域/指定 sheet 映射”？
3. 参数配置采用 TOML、命令行参数，还是两者结合？

## Decisions Made
| Decision | Rationale |
|----------|-----------|
| 先做方案确认，再动代码 | 用户明确要求先给方案确认 |
| 优先沿用 TOML 配置风格 | 仓库中已有 `job.toml` / `report_jobs.example.toml`，一致性更好 |
| PPT 首版方案采用“模板页复制 + 动态插表” | 模板中只有标题占位符，没有现成表格或图表对象 |
| Excel 首版按“首个工作表整体读入”处理 | 当前样本文件结构统一，第一页就是目标数据 |
| 总体行单独放在摘要表 | 用户要求超长时按二级标题拆左右双表，总体行单独放置更清晰 |
| 左右双表每侧默认上限设为 19 行 | 结合 9 月真实数据统计，多个客户群体最优切分至少需要 19 行容量 |
| 备注页分析采用 OpenAI 兼容 SDK + `.env` 配置 | 满足用户对兼容式 LLM 接入和连接参数外置的要求 |

## Errors Encountered
| Error | Attempt | Resolution |
|-------|---------|------------|
|       | 1       |            |

## Notes
- 遵循 AGENTS.md：使用 `uv` 作为包管理器，测试通过 `uv run` 执行。
- 若开始实现，需要同步补充 `/docs` 技术文档和 `/tests` 测试脚本。
- 已生成实际产物：`输出结果/9月满意度报告.pptx`
