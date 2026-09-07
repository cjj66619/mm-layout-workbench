# mm-layout-workbench

数学建模论文（华为杯研究生数模）**排版工作流**：把 polish 阶段的 Markdown 章节按官方 Word 模板重新合成
`main_layout.docx` / `main_layout.pdf`，并出体检报告。只写 `<proj>/layout/`，不改论文内容。

## 依赖

- Python 3.10+；`pip install python-docx pyyaml pymupdf`（`pymupdf` 可选：PDF 层检查、目录页码、PDF 图转 PNG）
- pandoc ≥ 3（或设 `PANDOC_BIN`）
- LibreOffice `soffice`（可选；没有就加 `--no-pdf`，目录页码留空，Word 中 F9 更新）

## 输入

```text
<proj>/
├── paper/paper.yaml          title / keywords / sections 顺序
├── polish/sections/*.md      首选源（缺省回退 paper/sections/*.md）；00_abstract.md 含 **关键词：** 行
└── figures/                  章节里引用的图（png/pdf）
```

## 用法

```bash
python .agents/skills/layout-kickoff/scripts/run_layout.py <proj> [--no-pdf] [--fig-max-cm 15] [--strict]
```

或分步：

```bash
python .agents/skills/text-normalize/scripts/normalize_md.py <proj>
python .agents/skills/docx-compose/scripts/compose_docx.py <proj>
python .agents/skills/layout-lint/scripts/lint_layout.py <proj>
python .agents/skills/layout-report/scripts/report_layout.py <proj>
```

## 产物（`<proj>/layout/`）

| 文件 | 说明 |
| --- | --- |
| `_src/sections/*.md` | 规范化后的章节副本（排版实际输入） |
| `NORMALIZE_LOG.md` | 规范化统计 + 事实零漂移校验 |
| `main_layout.docx` | 终稿 Word（模板封面/摘要/目录/页眉页脚原样） |
| `main_layout.pdf` | LibreOffice 预览（交稿请用 Word 另存 PDF） |
| `compose_report.json` | 计数 / 每表状态 / 目录 / 警告 |
| `LAYOUT_LINT.md`、`lint_report.json` | 逐项 PASS / WARN / FAIL |
| `LAYOUT_REPORT.md` | 汇总 + 人工待办 |

## 排版规则摘要

- 标题 `一、/1.1/1.1.1` 由模板多级列表编号；列表用 `(1)(2)`，不用圆点。
- 图注在下、表注在上，`图N/表N/式(N)` 由 `@fig:/@tbl:/@eq:` 交叉引用统一生成（无题注的表不占号）。
- 三线表按内容定列宽、居中；超版心时 12 → 10.5 → 9 pt 自动降字号，9 pt 仍不够则压列宽并报 WARN
  “建议转置或拆表”，**不自动横排**。
- 行间公式三栏居中、右侧 `(n)`；代码块进等宽代码清单表并允许折行；参考文献 `[1]`。
- 封面年份/学校/队号由用户在 Word 中填写。

## 开发

```bash
python -m compileall -q .agents/skills scripts
python scripts/smoke_test.py      # 最小项目冒烟（不需要 LibreOffice）
```

技能说明见 `.agents/skills/*/SKILL.md`，硬规则见 `AGENTS.md`，模板契约见 `_references/template/template_spec.json`。
