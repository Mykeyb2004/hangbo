使用uv作为包管理器，使用`uv run`运行和测试代码
开发过程中有技术文档要生成的，目录为 /docs。及时更新维护。
有测试脚本要保存的，目录为 /tests
日志文件：/logs
测试驱动开发

数据文件背景：
datas/1-2月  为1-2月份原始数据
datas/3月    为3月份原始数据
datas/Q1     为Q1季度原始数据（合并了1-2月份+3月份的原始数据）

输出结果/1-2月  为1-2月份分项数据
输出结果/3月    为3月份分项数据
输出结果/Q1     为Q1季度分项数据

汇总结果/1-2月  为1-2月份汇总数据
汇总结果/3月    为3月份汇总数据
汇总结果/Q1     为Q1季度汇总数据

`survey_stats.py`为统计分项数据的代码
`summary_table.py`为汇总分项数据的代码
`generate_ppt.py`为根据汇总数据生成PPT的代码

`datas/customer_mapping.xlsx`为客户大类、客户类型、数据分类标签的关系映射表

<!-- gitnexus:start -->
# GitNexus — Code Intelligence

This project is indexed by GitNexus as **hangbo** (3734 symbols, 5307 relationships, 95 execution flows). Use the GitNexus MCP tools to understand code, assess impact, and navigate safely.

> If any GitNexus tool warns the index is stale, run `npx gitnexus analyze` in terminal first.

## Always Do

- **MUST run impact analysis before editing any symbol.** Before modifying a function, class, or method, run `gitnexus_impact({target: "symbolName", direction: "upstream"})` and report the blast radius (direct callers, affected processes, risk level) to the user.
- **MUST run `gitnexus_detect_changes()` before committing** to verify your changes only affect expected symbols and execution flows.
- **MUST warn the user** if impact analysis returns HIGH or CRITICAL risk before proceeding with edits.
- When exploring unfamiliar code, use `gitnexus_query({query: "concept"})` to find execution flows instead of grepping. It returns process-grouped results ranked by relevance.
- When you need full context on a specific symbol — callers, callees, which execution flows it participates in — use `gitnexus_context({name: "symbolName"})`.

## Never Do

- NEVER edit a function, class, or method without first running `gitnexus_impact` on it.
- NEVER ignore HIGH or CRITICAL risk warnings from impact analysis.
- NEVER rename symbols with find-and-replace — use `gitnexus_rename` which understands the call graph.
- NEVER commit changes without running `gitnexus_detect_changes()` to check affected scope.

## Resources

| Resource | Use for |
|----------|---------|
| `gitnexus://repo/hangbo/context` | Codebase overview, check index freshness |
| `gitnexus://repo/hangbo/clusters` | All functional areas |
| `gitnexus://repo/hangbo/processes` | All execution flows |
| `gitnexus://repo/hangbo/process/{name}` | Step-by-step execution trace |

## CLI

| Task | Read this skill file |
|------|---------------------|
| Understand architecture / "How does X work?" | `.claude/skills/gitnexus/gitnexus-exploring/SKILL.md` |
| Blast radius / "What breaks if I change X?" | `.claude/skills/gitnexus/gitnexus-impact-analysis/SKILL.md` |
| Trace bugs / "Why is X failing?" | `.claude/skills/gitnexus/gitnexus-debugging/SKILL.md` |
| Rename / extract / split / refactor | `.claude/skills/gitnexus/gitnexus-refactoring/SKILL.md` |
| Tools, resources, schema reference | `.claude/skills/gitnexus/gitnexus-guide/SKILL.md` |
| Index, status, clean, wiki CLI commands | `.claude/skills/gitnexus/gitnexus-cli/SKILL.md` |

<!-- gitnexus:end -->
