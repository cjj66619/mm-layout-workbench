---
name: layout-kickoff
description: 数模论文排版工作流入口。接到“排版 / 生成 Word / 出终稿 docx”类需求时先用它：核对项目输入与依赖，然后按 text-normalize → docx-compose → layout-lint → layout-report 顺序一键运行，产物只写到 <proj>/layout/。
---

# layout-kickoff

排版管线的编排入口。**只写 `<proj>/layout/`**；`paper/`、`polish/`、`data/`、`code/`、`figures/` 全部只读，
不改任何文字、公式、数值、图路径。封面年份/学校/队号由用户自己填。

## 何时用

用户说“排版”“生成 Word/PDF 终稿”“套华为杯模板”“表格太宽/列表圆点/不居中”等，而且项目里已有
`polish/sections/*.md`（或 `paper/sections/*.md`）和 `paper/paper.yaml`。

## 步骤

1. 核对输入（缺哪项就停下问用户，不要自己造）：
   - `<proj>/paper/paper.yaml`（题目、关键词等元数据）
   - `<proj>/polish/sections/*.md`，没有则 `<proj>/paper/sections/*.md`
   - 图片路径能从章节 Markdown 解析到（通常在 `<proj>/figures/`）
2. 核对依赖：Python 3.10+、`python-docx`、`pyyaml`、`pymupdf`（PDF 层检查/目录页码）、pandoc ≥ 3
   （可用 `PANDOC_BIN` 指定）、LibreOffice `soffice`（可选；没有就加 `--no-pdf`）。
3. 一键运行：

   ```bash
   python .agents/skills/layout-kickoff/scripts/run_layout.py <proj> [--no-pdf] [--fig-max-cm 15] [--strict]
   ```

   等价于依次执行四个脚本，任一步返回非零就停止：

   ```bash
   python .agents/skills/text-normalize/scripts/normalize_md.py <proj>
   python .agents/skills/docx-compose/scripts/compose_docx.py <proj>
   python .agents/skills/layout-lint/scripts/lint_layout.py <proj>
   python .agents/skills/layout-report/scripts/report_layout.py <proj>
   ```

4. 读 `<proj>/layout/LAYOUT_REPORT.md`，把 FAIL / WARN / 人工待办用一句话汇报给用户；不要复述整份报告。
5. 若某步失败：看该步的日志（`NORMALIZE_LOG.md` / compose 终端输出 / `LAYOUT_LINT.md`），
   修的是**工作流脚本**或**排版参数**，不是论文内容。内容问题（表太宽、摘要超页、缺表注）只在报告里给建议，
   由用户回 polish 改。

## 产物

```text
<proj>/layout/
├── _src/sections/*.md      规范化后的章节副本（排版实际输入）
├── NORMALIZE_LOG.md        规范化统计 + 事实零漂移校验
├── main_layout.docx        终稿 Word（华为杯模板）
├── main_layout.pdf         LibreOffice 渲染的预览（交稿请用 Word 另存 PDF）
├── compose_report.json     计数 / 表格状态 / 目录 / 警告
├── LAYOUT_LINT.md + lint_report.json
└── LAYOUT_REPORT.md        汇总报告 + 人工待办
```

## 不要做

- 不改 `paper/`、`polish/`、数据、代码、图；不重跑 draft/polish。
- 不自动改封面年份；不自动把超宽表横排。
- 不把 `layout/` 产物、比赛数据、队号、API Key 提交进任何仓库。
