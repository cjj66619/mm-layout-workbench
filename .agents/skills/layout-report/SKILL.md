---
name: layout-report
description: 汇总 compose_report.json + lint_report.json + NORMALIZE_LOG.md 生成 layout/LAYOUT_REPORT.md：输入源、页数与表图公式计数、每张表的字号/列数/状态、WARN 清单与建议、目录、人工待办（封面年份/超宽表/摘要超页/Word F9 更新目录）。管线最后一步，交付前读它。
---

# layout-report

## 用法

```bash
python .agents/skills/layout-report/scripts/report_layout.py <proj> [--out layout/LAYOUT_REPORT.md]
```

只读 `layout/compose_report.json`（必需）、`layout/lint_report.json`、`layout/NORMALIZE_LOG.md`，
写 `layout/LAYOUT_REPORT.md`。不碰 docx/pdf，不碰论文。

## 报告结构

1. 结论行：lint FAIL/WARN/PASS 计数、产物路径、模板名。
2. 输入源：normalize 的源目录、事实零漂移、未定义引用。
3. 计数表：PDF 页数、摘要页数、标题三级、三线表（有表注）、图、编号公式、代码清单、参考文献、列表项、目录条目。
4. 表格状态：汇总 + 每表一行（表注 / 行×列 / 字号 / 状态 / 备注）。
5. FAIL（若有）、WARN 与建议、compose 警告、INFO。
6. 目录（带页码）。
7. 人工待办：封面年份/学校/队号；超宽表清单；摘要超 1 页；Word 中 F9 更新目录；缺表注的表；交稿用 Word 另存 PDF。

## 给用户汇报时

- 一句话说结论 + FAIL/WARN 数 + 人工待办条数，附 `layout/LAYOUT_REPORT.md` 路径；不要把报告整段贴出来。
- 内容侧问题（超宽表、摘要超页、缺表注）明确说“工作流不改文字，需回 polish 改”。
