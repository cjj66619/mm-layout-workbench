---
name: layout-lint
description: 对 layout/main_layout.docx（及同名 PDF）做排版体检，输出 LAYOUT_LINT.md + lint_report.json：页面/标题编号/列表圆点/三线表/图注表注/公式编号/双前缀/内部路径泄漏/目录页码/引号，PDF 层查页数、越界、页底空白、摘要页数。判 FAIL/WARN/INFO，--strict 时 FAIL 返回 1。
---

# layout-lint

## 用法

```bash
python .agents/skills/layout-lint/scripts/lint_layout.py <proj> [--strict]
```

读 `layout/main_layout.docx`、`layout/main_layout.pdf`（可选）、`layout/compose_report.json`（可选），
写 `layout/LAYOUT_LINT.md` 与 `layout/lint_report.json`。

## 检查项与口径

OOXML 层（必需）：
- 节均为 A4、四边 1418 twips。
- 标题：无手写序号；样式 `1/2/3` 绑定模板多级编号 numId 1（**段落级没有 numPr 是正常的**，编号在样式里）。
- 列表：无 Word 圆点项目符号、正文无 `•●·` 等字符（代码清单内除外）。
- 三线表：只查 `tblCaption="hwb-table"` 的表（封面表格、公式表、代码表不算）；居中、≤ 版心、表头重复、
  表注在上且“表N”连续；无表注 → WARN 提示回 polish 补 `Table: 标题 {#tbl:id}`。
- 图：≤ 版心、居中、图注在下、“图N”连续、宽度档位不宜过多。
- 公式：右侧 `(n)` 编号连续；无 `¿`/替换字符。
- 文本：逐段查“表 表N/图 图N/式 式N”双前缀；无内部路径/工作流名/未解析 `@ref`；无域代码文本；
  中文引号配对；无直引号（代码清单、行内代码除外）。
- 目录域存在、条目均有页码；页脚有 PAGE 域。

PDF 层（有 pymupdf 且有 PDF 时）：
- 全部 A4；正文越出左右边距（容差 14 pt，LibreOffice 两端对齐悬挂标点）；正文页无大片页底空白；
- 摘要页数 = 目录页索引 − 1（封面 1 页），> 1 页 WARN；
- `¿` 只给 INFO（LibreOffice 导入以 `=` 开头 OMML 的已知缺陷，Word 中正常）。

## 等级

- FAIL：必须修（多半是工作流 bug）。
- WARN：建议修或人工确认（超宽表、摘要超页、缺表注等**内容侧**问题只能回 polish 改）。
- INFO：已知渲染差异，可忽略。

## 加新检查时

- 用 `lint.check(cond, pass_msg, fail_msg, level)`；计数写进 `lint.counts`（report 会用）。
- 常量、样式 ID 从 `oxml_utils` 取；表格范围用 `TABLE_TAG`，代码表用 `_in_table(p, CODE_TAG)` 排除。
- 先在 2025-E 一类真实项目上跑一遍，确认没有新增误报再提交。
