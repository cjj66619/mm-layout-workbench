"""lint_layout.py — 对 layout/main_layout.docx（及同名 PDF）做排版体检，输出 LAYOUT_LINT.md + lint_report.json。

用法：
    python lint_layout.py <proj> [--docx layout/main_layout.docx] [--strict]

检查分两层：
  OOXML 层（必需，python-docx）：页面尺寸/页边距、标题编号（模板样式 1/2/3 的多级列表）、列表无圆点、
      三线表（tblCaption="hwb-table"）居中/≤版心/重复表头/表注在上、图注在下/图宽≤版心、公式无 ¿ 残留、
      双前缀“表 表”、内部路径泄漏、域代码残留、目录/页码域、引号配对；
  PDF 层（可选，pymupdf）：页数、A4 版面、正文越界、页底大片空白、摘要页数。
  PDF 由 LibreOffice 渲染，与 Word 有已知差异（以 `=` 开头的公式渲成 ¿、两端对齐段落右侧多出一个悬挂标点宽度），
  相关检查只给 INFO / 放宽容差。

结论等级：FAIL（必须修）/ WARN（建议修或人工确认）/ INFO。--strict 时有 FAIL 返回码 1。
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from collections import Counter
from dataclasses import dataclass, field
from pathlib import Path

import docx
from docx.oxml.ns import qn

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parents[1] / "docx-compose" / "scripts"))
from oxml_utils import MARGIN, NUM_ID_HEADINGS, PAGE_H, PAGE_W, STYLE, TEXT_W, find_style, p_style, p_text  # noqa: E402

WP_EXTENT = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}extent"
M_OMATH = "{http://schemas.openxmlformats.org/officeDocument/2006/math}oMath"
EMU_PER_TWIP = 635
TABLE_TAG = "hwb-table"
EQ_TAG, CODE_TAG = "hwb-equation", "hwb-code"
SPECIAL_CAPTIONS = (EQ_TAG, CODE_TAG)
LEAK_PATTERNS = ("reports/", "figures/", "results/", "polish/", "layout/", "paper/sections", "AGENTS.md", "plan.md",
                 "todo.md", "run_all", "mm-layout-workbench", "mm-draft-workbench", "mm-polish-workbench",
                 "mm-workbench", "@fig:", "@tbl:", "@eq:", "{#", "\\label{")
DOUBLE_PREFIX = re.compile(r"(表\s*表|图\s*图|式\s*式)\s*\d")
MANUAL_HEADING_NUM = re.compile(r"^\s*(\d+(\.\d+)*\.?|[一二三四五六七八九十]+、)\s*\S")
BULLET_CHARS = "•●·◦▪■○"
A4_PT = (595.3, 841.9)
MARGIN_PT = MARGIN / 20
PDF_OVERFLOW_TOL_PT = 14      # LibreOffice 两端对齐段落 bbox 右侧会多出一个悬挂标点宽度（≈12 pt）


@dataclass
class Lint:
    fails: list[str] = field(default_factory=list)
    warns: list[str] = field(default_factory=list)
    infos: list[str] = field(default_factory=list)
    passes: list[str] = field(default_factory=list)
    counts: dict[str, int | float | None] = field(default_factory=dict)

    def fail(self, m: str) -> None:
        self.fails.append(m)

    def warn(self, m: str) -> None:
        self.warns.append(m)

    def info(self, m: str) -> None:
        self.infos.append(m)

    def ok(self, m: str) -> None:
        self.passes.append(m)

    def check(self, cond: bool, ok_msg: str, bad_msg: str, level: str = "FAIL") -> None:
        if cond:
            self.ok(ok_msg)
        elif level == "FAIL":
            self.fail(bad_msg)
        else:
            self.warn(bad_msg)


def _attr(el, name: str, default: str = "") -> str:
    return el.get(qn(name), default) if el is not None else default


def _tbl_caption(tbl) -> str:
    return _attr(tbl.find(f"{qn('w:tblPr')}/{qn('w:tblCaption')}"), "w:val")


def _body_paragraphs(body):
    """正文层级（不含表格内）的段落，保持文档顺序。"""
    return [el for el in body if el.tag == qn("w:p")]


def _prev_para(el):
    """前一个非空段落（跳过 compose 插在相邻表之间的 1 pt 空段）。"""
    prev = el.getprevious()
    while prev is not None and (prev.tag != qn("w:p") or not p_text(prev).strip()):
        if prev.tag == qn("w:tbl"):
            return None
        prev = prev.getprevious()
    return prev


def _prose_text(p) -> str:
    """段落文本，去掉行内代码（VerbatimChar 字符样式）的 run。"""
    out = []
    for r in p.iter(qn("w:r")):
        rs = r.find(f"{qn('w:rPr')}/{qn('w:rStyle')}")
        if rs is not None and _attr(rs, "w:val") == "VerbatimChar":
            continue
        out.extend(t.text or "" for t in r.iter(qn("w:t")))
    return "".join(out)


def _next_para(el):
    nxt = el.getnext()
    while nxt is not None and nxt.tag != qn("w:p"):
        nxt = nxt.getnext()
    return nxt


def _bullet_num_ids(numbering) -> set[str]:
    if numbering is None:
        return set()
    bullet_abs = set()
    for an in numbering.findall(qn("w:abstractNum")):
        lvl0 = an.find(qn("w:lvl"))
        if lvl0 is not None and _attr(lvl0.find(qn("w:numFmt")), "w:val") == "bullet":
            bullet_abs.add(an.get(qn("w:abstractNumId")))
    return {n.get(qn("w:numId")) for n in numbering.findall(qn("w:num"))
            if _attr(n.find(qn("w:abstractNumId")), "w:val") in bullet_abs}


def _num_id(p) -> str:
    return _attr(p.find(f"{qn('w:pPr')}/{qn('w:numPr')}/{qn('w:numId')}"), "w:val")


def _style_num_id(styles_el, style_id: str) -> str:
    st = find_style(styles_el, style_id)
    if st is None:
        return ""
    return _attr(st.find(f"{qn('w:pPr')}/{qn('w:numPr')}/{qn('w:numId')}"), "w:val")


def _in_table(p, caption: str | None = None) -> bool:
    tc = p.getparent()
    if tc is None or tc.tag != qn("w:tc"):
        return False
    if caption is None:
        return True
    tbl = tc.getparent().getparent()
    return tbl is not None and tbl.tag == qn("w:tbl") and _tbl_caption(tbl) == caption


# --------------------------------------------------------------------------- OOXML 检查
def lint_page_setup(body, lint: Lint) -> None:
    sects = list(body.iter(qn("w:sectPr")))
    bad = []
    for i, s in enumerate(sects):
        sz, mar = s.find(qn("w:pgSz")), s.find(qn("w:pgMar"))
        if _attr(sz, "w:w") != str(PAGE_W) or _attr(sz, "w:h") != str(PAGE_H):
            bad.append(f"节{i + 1} 页面 {_attr(sz, 'w:w')}×{_attr(sz, 'w:h')}")
        for side in ("top", "bottom", "left", "right"):
            if _attr(mar, f"w:{side}") != str(MARGIN):
                bad.append(f"节{i + 1} {side} 边距 {_attr(mar, f'w:{side}')}")
    lint.counts["sections"] = len(sects)
    lint.check(not bad, f"{len(sects)} 个节均为 A4、四边 {MARGIN} twips", "页面设置偏离模板：" + "；".join(bad[:6]))


def lint_headings(body, styles_el, lint: Lint) -> None:
    """标题编号来自模板样式 1/2/3 自带的 numPr（numId 1 多级列表）；段落级只在“参考文献/附录”块显式关掉（numId 0）。"""
    hs = {STYLE["h1"]: [], STYLE["h2"]: [], STYLE["h3"]: []}
    manual = []
    for p in _body_paragraphs(body):
        st = p_style(p)
        if st in hs:
            txt = p_text(p).strip()
            hs[st].append(txt)
            if MANUAL_HEADING_NUM.match(txt) and not txt.startswith("附录"):
                manual.append(txt[:30])
    n1, n2, n3 = (len(hs[STYLE[k]]) for k in ("h1", "h2", "h3"))
    lint.counts.update({"h1": n1, "h2": n2, "h3": n3})
    lint.check(n1 > 0, f"标题 {n1}/{n2}/{n3}（一/二/三级）", "没有一级标题")
    lint.check(not manual, "标题无手写序号（编号由模板多级列表生成）",
               f"{len(manual)} 个标题含手写序号，将与自动编号重复：{manual[:3]}")
    style_nums = {k: _style_num_id(styles_el, STYLE[k]) for k in ("h1", "h2", "h3")}
    bad = [f"{k}={v or '无'}" for k, v in style_nums.items() if v != str(NUM_ID_HEADINGS)]
    lint.check(not bad, f"标题样式 1/2/3 均绑定模板多级编号 numId {NUM_ID_HEADINGS}",
               f"标题样式未绑定模板编号：{bad}", "WARN")


def lint_lists(body, numbering, lint: Lint) -> None:
    bullet_ids = _bullet_num_ids(numbering)
    bulleted, glyph = [], []
    n_list = 0
    for p in body.iter(qn("w:p")):
        nid = _num_id(p)
        if nid and nid != "0":
            n_list += 1
            if nid in bullet_ids:
                bulleted.append(p_text(p)[:30])
        txt = p_text(p).lstrip()
        if txt[:1] and txt[:1] in BULLET_CHARS and not _in_table(p, CODE_TAG):
            glyph.append(txt[:30])
    lint.counts["list_paragraphs"] = n_list
    lint.check(not bulleted, "列表无 Word 圆点项目符号（均为 (1)/(2) 编号）", f"{len(bulleted)} 段仍是圆点列表：{bulleted[:3]}")
    lint.check(not glyph, "正文无手写圆点字符", f"{len(glyph)} 段以 •/· 开头：{glyph[:3]}")


def lint_tables(body, lint: Lint, compose_tables: list[dict]) -> None:
    """只体检 table-fit 处理过的三线表（tblCaption="hwb-table"）；模板封面表、公式表、代码表不在此列。"""
    all_tables = list(body.iter(qn("w:tbl")))
    tables = [t for t in all_tables if _tbl_caption(t) == TABLE_TAG]
    lint.counts["tables"] = len(tables)
    lint.counts["equation_tables"] = sum(1 for t in all_tables if _tbl_caption(t) == EQ_TAG)
    lint.counts["code_tables"] = sum(1 for t in all_tables if _tbl_caption(t) == CODE_TAG)
    lint.counts["other_tables"] = len(all_tables) - lint.counts["tables"] - lint.counts["equation_tables"] - lint.counts["code_tables"]
    too_wide, not_center, no_header, no_caption, bad_style, grid_mismatch, keep_all = [], [], [], [], [], [], []
    cap_nums = []
    for i, t in enumerate(tables, 1):
        pr = t.find(qn("w:tblPr"))
        w = int(_attr(pr.find(qn("w:tblW")), "w:w", "0"))
        grid = sum(int(_attr(g, "w:w", "0")) for g in t.findall(f"{qn('w:tblGrid')}/{qn('w:gridCol')}"))
        if w > TEXT_W + 10 or grid > TEXT_W + 10:
            too_wide.append(f"表{i}:{max(w, grid)}")
        if abs(w - grid) > 20:
            grid_mismatch.append(f"表{i}:{w}/{grid}")
        if _attr(pr.find(qn("w:jc")), "w:val") != "center":
            not_center.append(f"表{i}")
        if _attr(pr.find(qn("w:tblStyle")), "w:val") != STYLE["three_line_table"]:
            bad_style.append(f"表{i}")
        tr0 = t.find(qn("w:tr"))
        if tr0 is None or tr0.find(f"{qn('w:trPr')}/{qn('w:tblHeader')}") is None:
            no_header.append(f"表{i}")
        cap = _prev_para(t)
        cap_txt = p_text(cap).strip() if cap is not None else ""
        m = re.match(r"^表\s*(\d+)", cap_txt)
        if cap is None or p_style(cap) != STYLE["table_caption"] or not m:
            no_caption.append(f"第{i}张表（前一段：{cap_txt[:20]!r}）")
        else:
            cap_nums.append(int(m.group(1)))
        rows = t.findall(qn("w:tr"))
        if len(rows) > 8:
            cell_ps = [p for tr in rows for p in tr.iter(qn("w:p"))]
            kn = sum(1 for p in cell_ps if p.find(f"{qn('w:pPr')}/{qn('w:keepNext')}") is not None)
            if kn == len(cell_ps):
                keep_all.append(f"表{i}")
    lint.check(not too_wide, f"{len(tables)} 张三线表宽度均 ≤ 版心 {TEXT_W}", f"超出版心：{too_wide[:5]}")
    lint.check(not not_center, "三线表全部整表居中", f"未居中：{not_center[:5]}")
    lint.check(not bad_style, "三线表全部使用模板“三线表”样式", f"样式不对：{bad_style[:5]}")
    lint.check(not grid_mismatch, "表宽与列宽之和一致", f"表宽≠列宽和：{grid_mismatch[:5]}", "WARN")
    lint.check(not no_header, "表头行全部设为跨页重复", f"未重复表头：{no_header[:5]}", "WARN")
    lint.check(not no_caption, "表注均在表上方且以“表N”开头",
               f"表注缺失/位置不对：{no_caption[:5]}——回 polish 在表格下一行补 `Table: 标题 {{#tbl:id}}`", "WARN")
    lint.check(not keep_all, "长表未整表 keepNext（允许跨页，不会把整表推到下页留白）", f"长表所有单元格 keepNext：{keep_all[:5]}", "WARN")
    if cap_nums:
        expect = list(range(1, len(cap_nums) + 1))
        lint.check(cap_nums == expect, f"表号连续 1–{len(cap_nums)}", f"表号不连续：{_gaps(cap_nums)}", "WARN")
    by_status = Counter(t.get("status", "?") for t in compose_tables)
    lint.counts["table_status"] = dict(by_status)  # type: ignore[assignment]
    for t in compose_tables:
        if t.get("status") in ("tight", "overflow"):
            label = (t.get("caption") or f"第 {t.get('index')} 张表")[:40]
            lint.warn(f"{label}：{t.get('status')}（{t.get('font_pt')} pt，{t.get('cols')} 列）"
                      "——建议转置或拆表，或在源 Markdown 缩短表头")


def _gaps(nums: list[int]) -> str:
    """[1,2,3,5,6,8] → '3→5, 6→8'（只列出跳号处）。"""
    out = [f"{a}→{b}" for a, b in zip(nums, nums[1:]) if b != a + 1]
    if nums and nums[0] != 1:
        out.insert(0, f"起始 {nums[0]}")
    return ", ".join(out[:6]) or str(nums[:12])


def lint_figures(body, lint: Lint) -> None:
    figs = []
    for p in _body_paragraphs(body):
        ext = p.find(f".//{WP_EXTENT}")
        if ext is None:
            continue
        cx = int(ext.get("cx", "0"))
        jc = _attr(p.find(f"{qn('w:pPr')}/{qn('w:jc')}"), "w:val")
        nxt = _next_para(p)
        cap = p_text(nxt).strip() if nxt is not None else ""
        figs.append((cx, jc, p_style(nxt) if nxt is not None else "", cap))
    body_figs = [f for f in figs if f[2] == STYLE["figure_caption"] or f[3].startswith("图")]
    lint.counts["figures"] = len(body_figs)
    if not body_figs:
        lint.info("正文没有插图")
        return
    wide = [f[3][:20] for f in body_figs if f[0] > TEXT_W * EMU_PER_TWIP + 1000]
    off = [f[3][:20] for f in body_figs if f[1] != "center"]
    nocap = [f[3][:20] for f in body_figs if not re.match(r"^图\s*\d+", f[3])]
    lint.check(not wide, "图宽均 ≤ 版心", f"图超出版心：{wide[:5]}")
    lint.check(not off, "图片全部居中", f"未居中：{off[:5]}")
    lint.check(not nocap, "图注均在图下方且以“图N”开头", f"图注缺失：{nocap[:5]}")
    nums = [int(m.group(1)) for f in body_figs if (m := re.match(r"^图\s*(\d+)", f[3]))]
    lint.check(nums == list(range(1, len(nums) + 1)), f"图号连续 1–{len(nums)}", f"图号不连续：{nums[:12]}", "WARN")
    widths_cm = Counter(round(f[0] / 360000, 1) for f in body_figs)
    lint.counts["figure_widths_cm"] = len(widths_cm)
    lint.check(len(widths_cm) <= 3, f"图宽档位 {len(widths_cm)} 种：{dict(widths_cm)}",
               f"图宽档位过多（{len(widths_cm)} 种）：{dict(widths_cm)}——建议统一 --fig-max-cm 或源图尺寸", "WARN")


def lint_equations(body, lint: Lint) -> None:
    n = sum(1 for _ in body.iter(M_OMATH))
    lint.counts["omath"] = n
    eq_tables = [t for t in body.iter(qn("w:tbl")) if _tbl_caption(t) == "hwb-equation"]
    labels = []
    for t in eq_tables:
        cells = t.findall(f"{qn('w:tr')}/{qn('w:tc')}")
        labels.append(p_text(cells[-1]).strip() if cells else "")
    bad = [x for x in labels if not re.fullmatch(r"\(\d+\)", x)]
    lint.check(not bad, f"{len(eq_tables)} 个行间公式均带右对齐编号 (n)", f"公式编号异常：{bad[:5]}", "WARN")
    nums = [int(x.strip("()")) for x in labels if re.fullmatch(r"\(\d+\)", x)]
    if nums:
        lint.check(nums == list(range(1, len(nums) + 1)), f"公式号连续 1–{len(nums)}", f"公式号不连续：{nums[:12]}", "WARN")


def lint_text(body, lint: Lint, doc) -> None:
    paras = list(body.iter(qn("w:p")))
    texts = [p_text(p) for p in paras]
    joined = "\n".join(texts)
    dbl = [m.group(0) for t in texts for m in DOUBLE_PREFIX.finditer(t)]
    lint.check(not dbl, "无“表 表 N / 图 图 N / 式 式 N”双前缀", f"双前缀 {len(dbl)} 处：{dbl[:3]}")
    leaks = [pat for pat in LEAK_PATTERNS if pat in joined]
    leak_lines = [t[:40] for t in texts if any(pat in t for pat in leaks)][:3]
    lint.check(not leaks, "正文无内部路径/工作流名/未解析引用泄漏", f"泄漏 {leaks}：{leak_lines}")
    inv = [t[:40] for t in texts if "¿" in t or "\ufffd" in t]
    lint.check(not inv, "无 ¿ / 替换字符（公式转换残留）", f"{len(inv)} 段含 ¿/�：{inv[:3]}")
    merge = [t[:40] for t in texts if "MERGEFORMAT" in t or re.search(r"\bTOC \\o", t)]
    lint.check(not merge, "无可见域代码文本", f"{len(merge)} 段可见域代码：{merge[:3]}")
    lq, rq = joined.count("“"), joined.count("”")
    lint.counts["quotes"] = (lq, rq)  # type: ignore[assignment]
    lint.check(abs(lq - rq) <= 2, f"中文引号配对（“{lq} / ”{rq}）", f"中文引号不配对：“{lq} / ”{rq}", "WARN")
    prose = [(p, _prose_text(p)) for p in paras if not _in_table(p, CODE_TAG)]
    straight_paras = [t[:40] for _, t in prose if '"' in t]
    straight = sum(t.count('"') for _, t in prose)
    lint.check(straight == 0, "正文无直引号（代码清单/行内代码除外）",
               f"正文含 {straight} 个直引号 \"（应为“”）：{straight_paras[:3]}", "WARN")
    toc_sdt = body.find(f".//{qn('w:sdt')}")
    toc_entries = [t for t in texts if t] if toc_sdt is None else [p_text(p) for p in toc_sdt.iter(qn("w:p")) if p_text(p).strip()]
    toc_ok = any("TOC" in (it.text or "") for it in body.iter(qn("w:instrText")))
    lint.check(toc_ok, "目录域存在（Word 中可 F9 更新）", "缺少目录域", "WARN")
    if toc_sdt is not None:
        no_page = [t[:30] for t in toc_entries[1:] if not re.search(r"\d+\s*$", t)]
        lint.counts["toc_entries"] = max(0, len(toc_entries) - 1)
        lint.check(not no_page, f"目录 {len(toc_entries) - 1} 条均已填页码", f"{len(no_page)} 条目录无页码：{no_page[:3]}", "WARN")
    page_field = False
    for part in doc.part.package.iter_parts():
        if "footer" in str(part.partname):
            page_field |= b"PAGE" in part.blob
    lint.check(page_field, "页脚含 PAGE 页码域", "页脚无页码域")


def lint_code(body, lint: Lint) -> None:
    code_tables = [t for t in body.iter(qn("w:tbl")) if _tbl_caption(t) == "hwb-code"]
    nowrap = 0
    for t in code_tables:
        for p in t.iter(qn("w:p")):
            ww = p.find(f"{qn('w:pPr')}/{qn('w:wordWrap')}")
            if ww is not None and _attr(ww, "w:val") in ("0", "false", "off"):
                nowrap += 1
    lint.check(nowrap == 0, f"{len(code_tables)} 个代码块均允许折行（不会横向溢出）", f"{nowrap} 段代码 wordWrap=off", "WARN")


# --------------------------------------------------------------------------- PDF 检查
def lint_pdf(pdf_path: Path, docx_path: Path, lint: Lint) -> None:
    try:
        import pymupdf  # type: ignore
    except ImportError:
        lint.info("未安装 pymupdf，跳过 PDF 层检查")
        return
    if not pdf_path.exists():
        lint.info(f"无 PDF（{pdf_path.name}），跳过 PDF 层检查")
        return
    if pdf_path.stat().st_mtime < docx_path.stat().st_mtime:
        lint.info(f"{pdf_path.name} 旧于 {docx_path.name}（compose 用了 --no-pdf？），跳过 PDF 层检查")
        return
    with pymupdf.open(pdf_path) as doc:
        n = doc.page_count
        lint.counts["pdf_pages"] = n
        bad_size, overflow, blank_pages = [], [], []
        toc_page = None
        for i, page in enumerate(doc):
            r = page.rect
            if abs(r.width - A4_PT[0]) > 2 or abs(r.height - A4_PT[1]) > 2:
                bad_size.append(i + 1)
            blocks = [b for b in page.get_text("blocks") if b[6] == 0 and b[4].strip()]
            body_blocks = [b for b in blocks if MARGIN_PT * 0.6 < b[1] < r.height - MARGIN_PT * 0.6]
            for b in body_blocks:
                if b[2] > r.width - MARGIN_PT + PDF_OVERFLOW_TOL_PT or b[0] < MARGIN_PT - PDF_OVERFLOW_TOL_PT:
                    overflow.append(i + 1)
                    break
            text = page.get_text()
            if toc_page is None and re.match(r"^目\s*录", text.lstrip()):
                toc_page = i
            if body_blocks and 0 < i < n - 1:
                last_bottom = max(b[3] for b in body_blocks)
                if last_bottom < r.height * 0.55 and toc_page is not None and i > toc_page:
                    nxt = doc[i + 1].get_text().lstrip()[:12]
                    if not re.match(r"^(附录|参考文献|[一二三四五六七八九十]+、)", nxt):
                        blank_pages.append(i + 1)
        # 封面 1 页，摘要从第 2 页起到目录页之前
        abstract_pages = (toc_page - 1) if toc_page is not None and toc_page >= 1 else None
        lint.counts["abstract_pages"] = abstract_pages
        lint.check(not bad_size, f"PDF {n} 页均为 A4", f"非 A4 页：{bad_size[:5]}")
        lint.check(not overflow, f"PDF 无正文越出左右页边距（容差 {PDF_OVERFLOW_TOL_PT} pt）",
                   f"{len(overflow)} 页有内容越出边距（多为超宽表/代码）：{overflow[:8]}", "WARN")
        lint.check(not blank_pages, "正文页无大片页底空白", f"{len(blank_pages)} 页下半页空白（检查前页是否被整表/整图推走）：{blank_pages[:8]}", "WARN")
        if abstract_pages is None:
            lint.info("PDF 中未定位到目录页，无法统计摘要页数")
        else:
            lint.check(abstract_pages <= 1, f"摘要 {abstract_pages} 页",
                       f"摘要占 {abstract_pages} 页（竞赛惯例 1 页，需回 polish 压缩）", "WARN")
        inv = [i + 1 for i, page in enumerate(doc) if "¿" in page.get_text()]
        if inv:
            lint.info(f"PDF 第 {inv[:5]} 页含 ¿：LibreOffice 导入以 `=` 开头的 OMML 公式的已知缺陷，Word 中正常，可忽略")
        else:
            lint.ok("PDF 无 ¿ 残留")


# --------------------------------------------------------------------------- 输出
def render_md(lint: Lint, docx_path: Path, strict: bool) -> str:
    verdict = "FAIL" if lint.fails else ("WARN" if lint.warns else "PASS")
    out = [f"# LAYOUT_LINT — {verdict}", "", f"- 文件：`{docx_path.name}`",
           f"- FAIL {len(lint.fails)} / WARN {len(lint.warns)} / PASS {len(lint.passes)} / INFO {len(lint.infos)}",
           f"- 模式：{'strict' if strict else 'normal'}", ""]
    for title, items in (("## FAIL（必须修）", lint.fails), ("## WARN（建议修 / 人工确认）", lint.warns),
                         ("## PASS", lint.passes), ("## INFO", lint.infos)):
        if items:
            out.append(title)
            out.append("")
            out.extend(f"- {m}" for m in items)
            out.append("")
    out.append("## 统计")
    out.append("")
    out.append("```json")
    out.append(json.dumps(lint.counts, ensure_ascii=False, indent=2))
    out.append("```")
    out.append("")
    out.append("> LibreOffice 渲染仅用于自检：多级标题 isLgl、公式字体与 Word 略有差异，以 Word 打开为准。")
    return "\n".join(out) + "\n"


def run_lint(proj: Path, docx_rel: str, compose_json: str) -> tuple[Lint, Path]:
    docx_path = (proj / docx_rel).resolve()
    lint = Lint()
    if not docx_path.exists():
        lint.fail(f"找不到 {docx_path}")
        return lint, docx_path
    d = docx.Document(str(docx_path))
    body = d.element.body
    numbering = d.part.numbering_part.element if d.part.numbering_part is not None else None
    compose_tables: list[dict] = []
    cj = proj / compose_json
    if cj.exists():
        compose_tables = json.loads(cj.read_text(encoding="utf-8")).get("tables", [])
    lint_page_setup(body, lint)
    lint_headings(body, d.styles.element, lint)
    lint_lists(body, numbering, lint)
    lint_tables(body, lint, compose_tables)
    lint_figures(body, lint)
    lint_equations(body, lint)
    lint_code(body, lint)
    lint_text(body, lint, d)
    lint_pdf(docx_path.with_suffix(".pdf"), docx_path, lint)
    return lint, docx_path


def main() -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("proj")
    ap.add_argument("--docx", default="layout/main_layout.docx")
    ap.add_argument("--compose-json", default="layout/compose_report.json")
    ap.add_argument("--out", default="layout/LAYOUT_LINT.md")
    ap.add_argument("--json", default="layout/lint_report.json")
    ap.add_argument("--strict", action="store_true", help="有 FAIL 时返回 1")
    args = ap.parse_args()
    proj = Path(args.proj).resolve()
    lint, docx_path = run_lint(proj, args.docx, args.compose_json)
    md = render_md(lint, docx_path, args.strict)
    out = proj / args.out
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(md, encoding="utf-8")
    (proj / args.json).write_text(json.dumps({"fails": lint.fails, "warns": lint.warns, "passes": lint.passes,
                                              "infos": lint.infos, "counts": lint.counts},
                                             ensure_ascii=False, indent=2), encoding="utf-8")
    verdict = "FAIL" if lint.fails else ("WARN" if lint.warns else "PASS")
    print(f"[lint] {verdict}  FAIL={len(lint.fails)} WARN={len(lint.warns)} PASS={len(lint.passes)} → {out}")
    for m in lint.fails:
        print(f"  FAIL {m}")
    return 1 if (args.strict and lint.fails) else 0


if __name__ == "__main__":
    raise SystemExit(main())
