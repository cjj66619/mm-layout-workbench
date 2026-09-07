# AGENTS.md — mm-layout-workbench

数模论文**排版专用**下游工作流（`mm-draft-workbench` → `mm-polish-workbench` → 本仓库）。
输入一个项目目录 `<proj>`，输出 `<proj>/layout/`。入口技能：`.agents/skills/layout-kickoff/SKILL.md`。

## 硬规则（违反即失败）

1. **只写 `<proj>/layout/`**。`paper/`、`polish/`、`data/`、`code/`、`figures/`、`results/` 一律只读。
2. **不改内容**：不改一个字、一个数字、一个公式、一个图路径；不合并/拆分段落；不改标题层级。
   排版只动样式、编号、字号、列宽、分页。text-normalize 唯一允许的“文字”改动是引用双前缀与引号形态，
   且必须通过事实零漂移校验。
3. 内容侧问题（超宽表、摘要超页、缺表注、图宽不一）只在 `LAYOUT_REPORT.md` 里给建议，由用户回 polish 改。
4. 不自动改封面年份/学校/队号；不自动把超宽表横排。
5. 正文里不得出现内部路径、工作流名、未解析的 `@fig:/@tbl:/@eq:`（lint 会查）。
6. 不提交比赛数据、`layout/` 产物、队号、API Key、任何凭据到仓库。
7. 代码：Python 3.10+，`pathlib`，所有文件读写显式 `encoding="utf-8"`，跨平台（Windows/Linux），
   不调用 bash / shell 字符串命令（子进程用 `subprocess.run([...])` 列表形式）。不新增第三方依赖
   （现有：`python-docx`、`pyyaml`、可选 `pymupdf`；外部：pandoc、可选 LibreOffice）。
8. OOXML 插入一律走 `oxml_utils.insert_ordered / child`；样式 ID 只从 `oxml_utils.STYLE` 取；
   模板契约见 `_references/template/template_spec.json`，改模板两处同步。

## 管线

```text
layout-kickoff/scripts/run_layout.py <proj>
  → text-normalize/scripts/normalize_md.py   polish/sections → layout/_src/sections + NORMALIZE_LOG.md
  → docx-compose/scripts/compose_docx.py     pandoc + 模板移植 + table_fit → main_layout.docx/.pdf + compose_report.json
  → layout-lint/scripts/lint_layout.py       LAYOUT_LINT.md + lint_report.json（--strict 时 FAIL 返回 1）
  → layout-report/scripts/report_layout.py   LAYOUT_REPORT.md（结论 + 人工待办）
```

任一步非零退出即停止。每个技能目录下的 `SKILL.md` 说明该步的规则与排错。

## 改代码后必须跑

```bash
python -m compileall -q .agents/skills scripts
python scripts/smoke_test.py            # 最小项目：normalize → compose --no-pdf → lint --strict
```

有真实项目时再跑一遍完整管线，确认 `LAYOUT_LINT.md` 没有新增 FAIL/误报。

## 已知的 LibreOffice 渲染差异（不是 bug）

以 `=` 开头的 OMML 公式渲成 `¿`；两端对齐段落右侧多出 ≈12 pt 悬挂标点宽度；多级编号显示为 isLgl。
lint 对这些只给 INFO / 放宽容差。交稿 PDF 请用 Word 另存。
