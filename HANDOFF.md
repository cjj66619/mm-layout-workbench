# HANDOFF — mm-layout-workbench 交接（给下一个 AI）

> 状态：**核心管线已跑通，仓库尚未推送**。本目录是 Devin 侧 `/home/ubuntu/repos/mm-layout-workbench` 的完整快照。
> 目标仓库：`cjj66619/mm-layout-workbench`（用户已建/将建空仓库），把本目录内容作为初始提交推上去即可。

## 1. 这是什么

`mm-layout-workbench` 是 `mm-draft-workbench` → `mm-polish-workbench` 之后的**排版专用**下游工作流，
供云端 Devin 使用。输入是一个项目文件夹（含 `paper/paper.yaml`、`polish/sections/*.md` 或 `paper/sections/*.md`、
`figures/`），输出写到 `<proj>/layout/`，**不改 `paper/`、`polish/`、数据、代码、文字内容**。

用户已确认的三项设计决策：

1. **从 Markdown 重新合成 Word**（不是原地修 `polish/main_polished.docx`）；源优先 `polish/sections`，回退 `paper/sections`。
2. **模板固化**为 `_references/template/hwb_template.docx`（用户提供的华为杯模板，封面/摘要/目录/页眉页脚原样保留；封面年份用户自己改，不用管）。
3. **超宽表**：12 → 10.5 → 9 pt 自动降字号，9 pt 仍装不下 → 保持 9 pt、压列宽并报 WARN “建议转置或拆表”；**不**自动横排。

用户最初指出的三个排版问题（都已解决）：表格内文字畸形换行、表格宽度不一/不居中、列表用圆点（要用 `(1)(2)` 括号编号）。

## 2. 管线与文件

```text
layout-kickoff（编排，待写 SKILL.md）
  → text-normalize   .agents/skills/text-normalize/scripts/normalize_md.py   ✅ 完成
  → docx-compose     .agents/skills/docx-compose/scripts/compose_docx.py      ✅ 完成（含 oxml_utils.py）
  → table-fit        .agents/skills/table-fit/scripts/table_fit.py            ✅ 完成（由 compose 调用，也可单独跑）
  → layout-lint      .agents/skills/layout-lint/scripts/lint_layout.py        🟡 可跑，有几处误报待修（见 §5）
  → layout-report    .agents/skills/layout-report/scripts/report_layout.py    ❌ 未写
```

### 命令（在项目根目录 `<proj>` 下）

```bash
# 1) 规范化：选源目录 → layout/_src/sections/*.md + layout/NORMALIZE_LOG.md
python <repo>/.agents/skills/text-normalize/scripts/normalize_md.py <proj>

# 2) 合成：pandoc → 模板移植 → 表格自适应 → LibreOffice 渲 PDF → 回填目录页码
#    产物 layout/main_layout.docx / main_layout.pdf / compose_report.json
python <repo>/.agents/skills/docx-compose/scripts/compose_docx.py <proj>   [--no-pdf] [--fig-max-cm 15] [--keep-entry]

# 3) 体检：layout/LAYOUT_LINT.md + lint_report.json
python <repo>/.agents/skills/layout-lint/scripts/lint_layout.py <proj>   [--strict]
```

依赖：Python 3.10+、`python-docx`、`pymupdf`（可选，PDF 层检查/目录页码）、pandoc ≥ 3（`PANDOC_BIN` 可指定）、
LibreOffice `soffice`（可选，无则 `--no-pdf`，目录页码留空由 Word F9 更新）。

### 2025-E 项目上的最新结果

`run_layout.py <proj> --strict`：73 页，50 张三线表、16 个编号公式、24 张图，lint FAIL 0 / WARN 13（10 张 tight/overflow 表、
1 张无表注、图宽不统一、摘要 3 页）→ `layout/LAYOUT_REPORT.md`。
抽检 PDF：封面、摘要、目录、`一、`/`1.1`/`1.1.1` 标题编号、`(1)(2)` 列表、居中三线表（表 12/13/16/48）、表注在上图注在下、
公式右侧 `(n)`、参考文献 `[1]`、附录代码框折行——均正常。`scripts/smoke_test.py` PASS。

## 3. 关键实现说明（改代码前必读）

- **模板样式 ID**（`oxml_utils.STYLE`）：Normal=`a4`、标题 `1/2/3`、表格文字 `a9`、表注 `a0`、图注 `a1`、图片 `aff1`、
  有序列表 `a3`、参考文献 `a`、三线表 `afb`、公式表 `afe`、代码表 `aff0`、TOC/TOC1-3、页眉 `af0`、页脚 `ae`。
  numbering：numId 1 = 标题多级（一、/1.1/1.1.1），abstractNum 4 = `(%1)` 有序列表（每个列表新建一个 num 实例），numId 4 = 参考文献 `[%1]`。
- **compose 流程**：`pick_sections_dir` → `load_meta(paper/paper.yaml)` → `resolve_crossrefs`（自己解析 `@fig:/@tbl:/@eq:`，
  不依赖 pandoc-crossref）→ pandoc（Lua 过滤器 `LUA_FILTER` 去掉手写标题序号、把 `Table:` 标题变表注段等）→
  `process_headings/lists/figures/equations/code_blocks/references` → `fit_all_tables` → `assemble()` 把内容装进模板
  （保留封面、摘要框架、目录 sdt、节属性）→ `docx_to_pdf` → `locate_heading_pages` → `fill_toc_pages` → 再渲一次 PDF。
- **OOXML 子元素顺序**：一切插入走 `oxml_utils.insert_ordered/child`，`_ORDER` 表 + `_PROPS_FIRST`；不要用裸 `append`。
- **表格算法**（`table_fit.plan_widths`）：`text_width_twips` 估宽（全角 1em、大写 0.68、数字 0.5、窄字符 0.3，×1.06 安全系数，
  单元格边距 `CELL_PAD=216`）；对每档字号：全部不换行 → `fit`；可换行列压到 ≥`COMFY_EM=8em` → `wrap`；
  9 pt 下压到最长不可断单元 → `tight`(WARN)；仍超 → 边距收到 `TIGHT_PAD=80` 等比压缩 → `overflow`(WARN)。
  ≤8 行的表整体 keepNext 不分页；长表允许跨页、表头 `tblHeader` 重复、行 `cantSplit`。含合并单元格的表只套样式（`merged`）。
  处理过的表在 `tblCaption` 打 `hwb-table` 标记；公式表 `hwb-equation`、代码表 `hwb-code`。
- **normalize**：双前缀“表 @tbl:”→“@tbl:”、直引号→中文弯引号、事实零漂移校验（数字/公式/图路径/引用集合不变，FAIL 0 才写出）。
- **已知 LibreOffice 渲染差异**（不是 bug，Word 里正常）：以 `=` 开头的公式渲出 `¿`；多级标题 isLgl 显示；两端对齐段落 bbox 右侧
  多出一个悬挂标点宽度（≈12 pt）。lint 的 PDF 层要按此放宽。

## 4. 目录内容

```text
mm-layout-workbench/
├── HANDOFF.md                      ← 本文件
├── AGENTS.md / README.md / .gitignore
├── .github/workflows/ci.yml        ← compileall + smoke（无 LibreOffice）
├── scripts/smoke_test.py
├── _references/template/{hwb_template.docx, template_spec.json}
└── .agents/skills/                 ← 每个目录有 SKILL.md
    ├── layout-kickoff/scripts/run_layout.py   ← 一键：normalize → compose → lint → report
    ├── text-normalize/scripts/normalize_md.py
    ├── docx-compose/scripts/{compose_docx.py, oxml_utils.py}
    ├── table-fit/scripts/table_fit.py
    ├── layout-lint/scripts/lint_layout.py
    └── layout-report/scripts/report_layout.py
```

（如果隧道写二进制失败，`hwb_template.docx` 请直接用用户提供的 `dDesktop22.docx` 复制过去，文件内容完全相同。）

## 5. 待办（按优先级）

> 第 1–7 项已完成（2026-09-07）。剩第 8 项：待用户确认后再 `git init` 并推送。

1. **修 lint 误报**（`lint_layout.py`）：
   - `lint_lists`：`txt[:1] in BULLET_CHARS` 对空串为 True → 改成 `txt[:1] and txt[:1] in BULLET_CHARS`。
   - `lint_tables`：只检查 `tblCaption == "hwb-table"` 的表（封面表格不是三线表，当前被误报“表1 样式不对/无表注”）。
   - `lint_headings`：标题编号在样式 `1/2/3` 的 `pPr/numPr` 里，段落级没有 numPr 是正常的 → 改为检查 styles.xml 的样式定义或删掉该项。
   - `lint_text` 双前缀用 `joined` 跨段匹配 → 逐段匹配。
   - compose_report 里表字段是 `cols`（不是 `ncols`）。
   - PDF 越界容差放到 ≥14 pt；PDF 里的 `¿` 降为 INFO 并说明是 LibreOffice 导入缺陷；摘要页数 = 目录页索引 − 1（封面 1 页）。
   - 直引号检查排除 `hwb-code` 表内文本；无表注的表（2025-E 有 1 张：07_q4 “解释按三层展开”下表）报 WARN “回 polish 补 `Table: … {#tbl:id}`”。
2. **写 `layout-report/scripts/report_layout.py`**：汇总 `compose_report.json` + `lint_report.json` + `NORMALIZE_LOG.md` →
   `layout/LAYOUT_REPORT.md`（输入源、页数/表图公式计数、表格状态表、WARN 表清单与建议、目录、人工待办：封面年份/超宽表/摘要超 1 页/Word 中 F9 更新目录）。
3. **写 6 个 `SKILL.md`**（frontmatter 只留 `name`、`description`）+ `layout-kickoff/scripts/run_layout.py`（顺序跑三步，任一步失败停）。
4. **`AGENTS.md`、`README.md`**：硬规则——只写 `layout/`；`paper/`、`polish/`、数据/代码/图只读；不改公式/数值/图路径/文字；
   正文不得出现内部路径/工作流名；Python 3.10+、`pathlib`、显式 UTF-8、不调 bash；不提交比赛数据/队号/API Key。
5. **`_references/template/template_spec.json`**（页面/字体/样式 ID/编号契约，`oxml_utils` 常量的文档镜像）。
6. **`scripts/smoke_test.py`**：造一个 3 节的最小项目（1 表 1 图 1 公式 1 列表）→ normalize → compose `--no-pdf` → lint `--strict`；
   `.github/workflows/ci.yml`：`python -m compileall -q .agents/skills` + smoke（ubuntu，装 pandoc，跳过 soffice）。
7. 全量复核 2025-E：`python -m compileall -q .`、三步跑完、抽检 PDF 页（表 12/13/16/49、附录代码、参考文献）。
8. `git init` → 首次提交 → 推 `cjj66619/mm-layout-workbench`（不要提交 `layout/` 产物或任何比赛数据）。

## 6. 不要做的事

- 不改模型/数据/实验/文字内容；不重跑 draft/polish；不动 `paper/`、`polish/`。
- 不自动改封面年份。
- 不把本地工作区（隧道）里的改动 push 到 GitHub；工作流仓库本身是独立仓库，正常推。
- 不用 `gh pr create`，用平台内置的 PR 工具。
