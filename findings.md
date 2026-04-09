# Findings & Decisions

## Requirements
- 根据 `templates/template.pptx` 作为 PPT 模板生成最终 PPT。
- 数据来源是 `输出结果/9月/` 目录下的 Excel 文件。
- 每个 Excel 文件对应一个 PPT 页面，页面标题使用 Excel 文件名。
- 参数需要可配置。
- 当前阶段先输出实施方案，等待用户确认。

## Research Findings
- 项目当前依赖只有 `openpyxl` 和 `pandas`，尚未包含 `python-pptx`。
- 仓库根目录存在 `job.toml` 和 `report_jobs.example.toml`，说明配置文件方案已有先例。
- `输出结果/9月/` 下当前有 15 个 Excel 文件，可自然映射为 15 页或 15 个基于模板复制的页面。
- `template.pptx` 当前只有 1 个 slide。
- 该模板页中只有 1 个标准标题占位符，文本内容为 `项目执行情况 — 样本收集情况`。
- 模板页中未发现现成的表格、图表、图片占位框；首版更适合在运行时动态插入表格。
- 已抽查多个 Excel，结构高度一致：首个 sheet 名与文件名一致，列头固定为 `指标`、`满意度`、`重要性`。
- 不同 Excel 的行数不同，约在 19 到 38 行之间，需要做自适应行高或分页策略。
- 对 9 月实际数据按二级标题切分后，左右双表最优容量需求在 9 到 19 行之间。
- 其中 `会展服务商`、`会议主承办`、`酒店会议主承办` 的最优单侧容量都需要 19 行。
- 表格样式已调整为接近截图的配色主题：深酒红表头、玫粉分组行、浅粉灰正文、白色分隔线。
- `python-pptx` 的 `slide.notes_slide.notes_text_frame` 可直接写入备注页文本。
- 仓库当前不存在 `.env`，因此补充了 `.env.example` 作为连接配置模板。
- LLM 备注页现已支持流式输出到终端，并在生成过程中持续把收到的文本写回备注页。
- 当 LLM 中断或用户 `Ctrl+C` 终止时，会保留 `*.partial.pptx` 检查点，已完成页和当前页已收到的片段文本都可保住。
- 当前项目依赖中还没有图表绘制库；若要输出 PNG 图表，较稳妥的选择是新增 `matplotlib`。
- `generate_ppt.py` 当前已经具备“章节页插入 -> 数据页插入”的页级编排能力，继续扩展“数据页后追加图表页”较自然。
- 真实导出的 Excel 中，二级指标就是“总体行之后、三级指标之前”的聚合行，适合作为图表直接数据源。
- 当前数据中确实存在二级指标整体为空的情况，图表模块应复用现有过滤规则，避免画出全空维度。
- 用户补充的目标版式是“左侧图表 + 右侧文字框”，文字框内容暂时只需要占位。
- `python-pptx` 支持直接用 `BytesIO` 将 PNG 字节流插入到幻灯片中，因此图表模块无需在项目目录落地临时图片文件。
- `输出结果/Q1` 实际生成验证通过后，总页数从原来的 18 页增加到 33 页，结构为“章节页 -> 数据页 -> 图表页”。

## Technical Decisions
| Decision | Rationale |
|----------|-----------|
| 方案优先考虑 `python-pptx` | Python 生态中最常见，适合模板化生成 PPT |
| 参数设计优先考虑“配置文件 + CLI 覆盖” | 兼顾批处理和临时执行 |
| 首版优先生成“每个 Excel 1 页” | 与用户当前描述最一致，实现路径最短 |
| 对空值单元格做可配置渲染 | 样本中存在 `None`，例如 `客房服务`、`室内导航系统` 等行 |
| 采用“摘要表 + 明细表”布局 | 让总体行不参与左右双表拆分 |
| 实际默认值采用 `max_single_table_rows=18`、`max_split_table_rows=19` | 与当前模板布局和 9 月数据最匹配 |
| LLM 备注页做成可开关配置 | 不影响现有纯本地 PPT 生成流程 |
| 空值二级/三级指标不再单独喂给 LLM | 用户明确要求空值不特别提及 |
| LLM 流式过程中按字符阈值保存检查点 | 在失败、中断场景下尽量保住已生成文本 |
| 图表页能力拆到独立模块文件 | 降低主脚本耦合，便于单测和后续替换图表引擎 |
| 图表默认使用二级指标满意度/重要性双系列 | 与现有表格结构天然对应 |
| 图表页标题直接复用前一页标题 | 避免额外命名规则和展示歧义 |
| 二级指标数为 2 时画柱状图，>=3 时画雷达图 | 用户明确指定的切换规则 |
| 右侧文字框先用占位文本 | 先完成结构与版式，后续再接自动文案 |
| 图表图片以内存字节流形式插入 PPT | 不污染工作目录，也便于测试 |
| Q1 配置默认开启图表页 | 便于用户直接用真实数据验证新功能 |

## Issues Encountered
| Issue | Resolution |
|-------|------------|
| 尚未确认模板页内部占位结构 | 已通过解析 PPT XML 确认仅有标题占位符 |
| 初始双表容量设为 12 行时无法覆盖真实数据 | 统计所有 9 月文件最优切分结果后，将默认值调整为 19 |
| 现有依赖无法直接输出图表 PNG | 计划在实现阶段通过 `uv` 补充图表依赖 |
| Matplotlib 初始字体回退会打印大量 `findfont` 警告 | 改为运行时探测本机可用字体，再设置图表字体 |

## Resources
- `/Users/zhangqijin/PycharmProjects/hangbo/templates/template.pptx`
- `/Users/zhangqijin/PycharmProjects/hangbo/输出结果/9月/专业观众.xlsx`
- `/Users/zhangqijin/PycharmProjects/hangbo/pyproject.toml`
- `/Users/zhangqijin/PycharmProjects/hangbo/job.toml`
- `/Users/zhangqijin/PycharmProjects/hangbo/report_jobs.example.toml`
- `/Users/zhangqijin/PycharmProjects/hangbo/generate_ppt.py`
- `/Users/zhangqijin/PycharmProjects/hangbo/ppt_job.example.toml`
- `/Users/zhangqijin/PycharmProjects/hangbo/docs/PPT批量生成.md`
- `/Users/zhangqijin/PycharmProjects/hangbo/tests/test_generate_ppt.py`
- `/Users/zhangqijin/PycharmProjects/hangbo/.env.example`
- `/Users/zhangqijin/PycharmProjects/hangbo/system_role.md`

## Visual/Browser Findings
- `template.pptx` 文件存在，大小约 689 KB，说明不是空模板。
- `输出结果/9月/` 目录下 Excel 文件名包括 `专业观众`、`会展服务商`、`会议主承办`、`参会人员`、`参展商` 等，适合作为标题直接展示。
- `专业观众.xlsx` 的完整数据表为三列纵向指标表，含部分空值行。
- `会议主承办.xlsx` 与其他样本保持同一表结构，只是指标项数量不同。
- 生成后的 `输出结果/9月满意度报告.pptx` 共 15 页。
- `特色美食廊`、`自助餐`、`酒店自助餐` 使用 2 张表（摘要 + 单表）。
- 其余 12 个客户群体使用 3 张表（摘要 + 左右双表）。
- 当前生成版表格配色已从默认蓝绿橙切换为截图风格的红粉色系。
- 备注页占位文本框存在，可直接用于写入 LLM 生成的分析段落。

---
*Update this file after every 2 view/browser/search operations*
*This prevents visual information from being lost*
