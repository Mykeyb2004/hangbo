# Task Plan: Excel 批量生成 PPT 增强规划

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
