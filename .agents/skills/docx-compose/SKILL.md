---
name: docx-compose
description: 从 Markdown 章节按华为杯 Word 模板重新合成 layout/main_layout.docx（pandoc → 样式移植 → 表格自适应 → LibreOffice 渲 PDF → 回填目录页码）。修改 Word 排版逻辑、模板样式映射、图/表/公式/代码/参考文献格式时看这里。
---

# docx-compose

把 `layout/_src/sections/*.md`（回退 `polish/sections` → `paper/sections`）和 `paper/paper.yaml`
合成为模板口径的 `layout/main_layout.docx`。**从 Markdown 重建**，不是原地修 `polish/main_polished.docx`。

## 用法

```bash
python .agents/skills/docx-compose/scripts/compose_docx.py <proj> [--no-pdf] [--fig-max-cm 15] [--toc-depth 3] [--keep-entry]
```

- 模板固定为 `_references/template/hwb_template.docx`（契约见 `_references/template/template_spec.json`）。
- 依赖：pandoc ≥ 3（`PANDOC_BIN` 可指定）、`python-docx`、`pyyaml`；可选 `soffice`（PDF + 目录页码）、`pymupdf`（PDF 图转 PNG、目录定位）。
- 产物：`layout/main_layout.docx`、`layout/main_layout.pdf`、`layout/compose_report.json`（counts / warnings / tables / toc / files）。

## 流程（改代码前先对着 compose_docx.py 读一遍）

1. `pick_sections_dir` → `load_meta(paper/paper.yaml)` → `resolve_crossrefs`：自己解析 `@fig:/@tbl:/@eq:`，
   **只有带 `{#tbl:id}` 题注的表才占号**，Lua 过滤器同规则，二者必须一致。
2. pandoc `--reference-doc` 模板 → DOCX，公式为原生 OMML；Lua 过滤器去掉手写标题序号、给图/表题加“图N/表N”、
   `::: {.page-break}` → 分页。
3. 后处理：标题套模板样式 `1/2/3`（编号由样式自带的多级列表 numId 1 生成，参考文献/附录关掉）；
   列表 → 模板 `(1)(2)` 编号（每个列表新建 num 实例）；图注下/表注上；行间公式三栏表居中 + 右侧 `(n)`；
   代码块装进 `hwb-code` 表（Consolas，允许折行）；参考文献 `[1]` 样式；`fit_all_tables`（见 table-fit）。
4. `assemble()`：模板封面原样保留 → 摘要页回填题目/摘要/关键词 → 目录 sdt → 正文/参考文献/附录三节。
5. `docx_to_pdf` → `locate_heading_pages` → `fill_toc_pages` → 再渲一次 PDF。

## 硬规则

- 一切 OOXML 插入走 `oxml_utils.insert_ordered / child`（子元素顺序表 `_ORDER`），不要裸 `append`。
- 样式 ID 只从 `oxml_utils.STYLE` 取，不要写魔法字串。
- 不改封面年份；不自动横排；不改任何正文文字/数字/公式/图路径。
- 已知 LibreOffice 差异（Word 里正常）：以 `=` 开头的公式渲成 `¿`；两端对齐段落 bbox 右侧多 ≈12 pt。

## 排错

- `未找到 pandoc`：装 pandoc 或设 `PANDOC_BIN`。
- 没有 `soffice`：加 `--no-pdf`，目录页码留空，Word 里 F9 更新。
- 想看 pandoc 中间产物：`--keep-entry`，会在 `layout/` 下留下 `_entry.md` 与 `_pandoc_raw.docx`。
