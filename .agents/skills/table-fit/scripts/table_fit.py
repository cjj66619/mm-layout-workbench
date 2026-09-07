#!/usr/bin/env python3
"""table_fit.py — 把 DOCX 里的普通表格整理成模板三线表，并按内容自适应列宽与字号。

规则（用户确认）：
- 列宽按"不换行所需宽度"估算（全角 1 em、半角 0.5 em），窄表按内容宽度居中，不拉满版心；
- 总宽超过版心（9070 twips ≈ 16 cm）时：先让"长文本列"在不窄于 8 em 的前提下换行（数字/短字段列不换行），
  仍超则依次尝试 12 pt → 10.5 pt → 9 pt；
- 9 pt 仍超：长文本列压到最长不可断单元（tight，WARN）；还压不下去则整体等比压缩并收窄单元格边距（overflow，WARN），
  建议人工转置或拆表。不自动横排。
- 三线表：顶/底线 1.5 pt，表头下线 0.5 pt，表头加粗居中、跨页重复；单元格垂直居中；行不跨页拆分；
  ≤ 8 行的短表整体不分页，长表允许跨页（不给每个单元格加 keepNext，避免上一页留大片空白）。
- 公式三栏表（tblCaption="hwb-equation"）、代码清单表（tblCaption="hwb-code"）与嵌套表不处理。

既可被 compose_docx.py 导入，也可单独修一份已有 Word：
    python table_fit.py in.docx [--out out.docx]
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

import docx
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "docx-compose" / "scripts"))
from oxml_utils import (M_T, STYLE, TEXT_W, child, clear, insert_ordered, p_style, p_text, para_props,  # noqa: E402
                        ppr, set_bold, set_pstyle, set_size, text_width_twips)

FONT_LADDER: tuple[tuple[float, int], ...] = ((12.0, 360), (10.5, 320), (9.0, 280))   # (字号 pt, 固定行距 twips)
CELL_PAD = 2 * 108          # 模板 Normal Table 左右单元格边距
TIGHT_PAD = 2 * 40          # overflow 表收窄后的边距
COMFY_EM = 8                # 文本列换行时的“舒适”最小宽度（em）
COL_MIN = 560               # 列宽下限（约 2 个 12 pt 全角字符）
SHORT_TABLE_ROWS = 8        # ≤ 该行数的表整体不分页
SPECIAL_CAPTIONS = ("hwb-equation", "hwb-code")
TABLE_TAG = "hwb-table"     # 处理过的普通三线表在 tblCaption 上打的标记（lint 据此识别）
_WRAP_TOKEN_BREAKS = " /,;，；、-–—()（）[]"


def is_special_table(tbl) -> bool:
    cap = tbl.find(f"{qn('w:tblPr')}/{qn('w:tblCaption')}")
    return cap is not None and cap.get(qn("w:val")) in SPECIAL_CAPTIONS


def _cell_lines(tc) -> list[str]:
    lines: list[str] = []
    for p in tc.findall(qn("w:p")):
        cur = ""
        for el in p.iter():
            if el.tag in (qn("w:t"), M_T):
                cur += el.text or ""
            elif el.tag == qn("w:br"):
                lines.append(cur)
                cur = ""
        lines.append(cur)
    return [ln for ln in lines if ln] or [""]


def _natural(lines: list[str], pt: float, pad: int = CELL_PAD) -> int:
    return int(max(text_width_twips(ln, pt) for ln in lines) * 1.06) + pad


def _min_unit(lines: list[str], pt: float) -> int:
    """可换行的最小宽度：最长不可断单元（西文单词/数字串），中文至少 2 个字。"""
    em = pt * 20
    best = 2 * em
    for ln in lines:
        token = ""
        for ch in ln + " ":
            if ch in _WRAP_TOKEN_BREAKS or ord(ch) > 0x2E7F:
                if token:
                    best = max(best, text_width_twips(token, pt))
                    token = ""
                if ord(ch) > 0x2E7F:
                    best = max(best, em)
            else:
                token += ch
    return int(best) + CELL_PAD


def _column_text(tbl):
    rows = tbl.findall(qn("w:tr"))
    cells = [r.findall(qn("w:tc")) for r in rows]
    ncol = max((len(c) for c in cells), default=0)
    cols: list[list[list[str]]] = [[] for _ in range(ncol)]
    for row in cells:
        for i, tc in enumerate(row):
            cols[i].append(_cell_lines(tc))
    return rows, cells, cols


def _has_merges(cells) -> bool:
    for row in cells:
        for tc in row:
            tcpr = tc.find(qn("w:tcPr"))
            if tcpr is not None and (tcpr.find(qn("w:gridSpan")) is not None or tcpr.find(qn("w:vMerge")) is not None):
                return True
    return False


def _measure(cols, pt: float, pad: int = CELL_PAD) -> tuple[list[int], list[int], list[bool]]:
    """每列的 (不换行宽度, 最小宽度, 是否可换行的文本列)。"""
    nat = [max(COL_MIN, max(_natural(c, pt, pad) for c in col)) for col in cols]
    mins = [max(COL_MIN, max(_min_unit(c, pt) - CELL_PAD + pad for c in col)) for col in cols]
    wrapable = [nat[i] - mins[i] >= 2 * pt * 20 for i in range(len(cols))]
    return nat, mins, wrapable


def _shrink(nat: list[int], lower: list[int], wrapable: list[bool]) -> list[int] | None:
    """固定列保持 nat，可换行列在 [lower, nat] 之间等比压缩到恰好填满版心；装不下返回 None。"""
    ncol = len(nat)
    fixed = sum(nat[i] for i in range(ncol) if not wrapable[i])
    flex_nat = sum(nat[i] for i in range(ncol) if wrapable[i])
    flex_low = sum(lower[i] for i in range(ncol) if wrapable[i])
    avail = TEXT_W - fixed
    if not flex_nat or avail < flex_low:
        return None
    scale = avail / flex_nat
    widths = [nat[i] if not wrapable[i] else max(lower[i], int(nat[i] * scale)) for i in range(ncol)]
    excess = sum(widths) - TEXT_W
    for i in sorted((i for i in range(ncol) if wrapable[i]), key=lambda i: -widths[i]):
        if excess <= 0:
            break
        cut = min(excess, widths[i] - lower[i])
        widths[i] -= cut
        excess -= cut
    return widths


def plan_widths(cols: list[list[list[str]]]) -> tuple[list[int], float, int, str, str]:
    """返回 (列宽, 字号 pt, 行距, 状态 fit|wrap|tight|overflow, 说明)。

    fit    ：全部列不换行，按内容宽居中；
    wrap   ：长文本列换行，但每列不窄于 COMFY_EM（正常现象，不报警）；
    tight  ：9 pt 下文本列压到最长不可断单元（WARN）；
    overflow：连最小宽度都装不下，整体等比压缩（WARN）。
    """
    ncol = len(cols)
    for pt, line in FONT_LADDER:
        em = pt * 20
        nat, mins, wrapable = _measure(cols, pt)
        if sum(nat) <= TEXT_W:
            extra = min(TEXT_W - sum(nat), int(0.08 * sum(nat)))
            widths = [w + int(extra * w / sum(nat)) for w in nat]
            return widths, pt, line, "fit", f"{pt:g} pt，内容宽 {sum(widths)} twips"
        comfy = [max(mins[i], min(nat[i], int(COMFY_EM * em) + CELL_PAD)) for i in range(ncol)]
        widths = _shrink(nat, comfy, wrapable)
        if widths is not None:
            narrowest = min(widths[i] for i in range(ncol) if wrapable[i])
            return widths, pt, line, "wrap", f"{pt:g} pt，文本列换行（最窄文本列 {narrowest} twips）"
    pt, line = FONT_LADDER[-1]
    nat, mins, wrapable = _measure(cols, pt)
    widths = _shrink(nat, mins, wrapable)
    if widths is not None:
        narrowest = min(widths[i] for i in range(ncol) if wrapable[i])
        return widths, pt, line, "tight", f"9 pt 下文本列已压到最小宽度（最窄文本列 {narrowest} twips）"
    nat, mins, wrapable = _measure(cols, pt, TIGHT_PAD)
    widths = _shrink(nat, mins, wrapable)
    if widths is not None:
        return widths, pt, line, "overflow", f"9 pt 下已收窄单元格边距并压到最小宽度，需 {sum(nat)} twips（版心 {TEXT_W}）"
    k = TEXT_W / sum(nat)
    widths = [int(w * k) for w in nat]
    return widths, pt, line, "overflow", f"9 pt 仍需 {sum(nat)} twips（版心 {TEXT_W}），已等比压缩 {k:.2f}"


def _set_grid(tbl, widths: list[int]) -> None:
    grid = tbl.find(qn("w:tblGrid"))
    if grid is None:
        grid = OxmlElement("w:tblGrid")
        insert_ordered(tbl, grid)
    for gc in list(grid):
        grid.remove(gc)
    for w in widths:
        gc = OxmlElement("w:gridCol")
        gc.set(qn("w:w"), str(w))
        grid.append(gc)


def _borders(tblpr) -> None:
    clear(tblpr, "w:tblBorders")
    bd = child(tblpr, "w:tblBorders")
    for side, val, sz in (("top", "single", 12), ("left", "nil", 0), ("bottom", "single", 12),
                          ("right", "nil", 0), ("insideH", "nil", 0), ("insideV", "nil", 0)):
        el = child(bd, f"w:{side}", val=val)
        if val != "nil":
            el.set(qn("w:sz"), str(sz))
            el.set(qn("w:space"), "0")
            el.set(qn("w:color"), "auto")


def caption_of(tbl) -> str:
    prev = tbl.getprevious()
    while prev is not None and prev.tag == qn("w:p") and not p_text(prev).strip():
        prev = prev.getprevious()
    if prev is not None and prev.tag == qn("w:p") and p_style(prev) in ("TableCaption", STYLE["table_caption"], "a8"):
        return p_text(prev).strip()
    return ""


def fit_table(tbl, index: int = 0) -> dict:
    tblpr = child(tbl, "w:tblPr")
    rows, cells, cols = _column_text(tbl)
    ncol = len(cols)
    info: dict = {"index": index, "caption": caption_of(tbl), "rows": len(rows), "cols": ncol}
    if ncol == 0:
        info.update(status="empty", font_pt=12.0, width=0, note="空表")
        return info
    merged = _has_merges(cells)
    if merged:
        pt, line, status, note = 12.0, 360, "merged", "含合并单元格，只套样式不重算列宽"
        widths = None
    else:
        widths, pt, line, status, note = plan_widths(cols)
    info.update(status=status, font_pt=pt, note=note, width=sum(widths) if widths else 0)

    clear(tblpr, "w:tblStyle")
    child(tblpr, "w:tblStyle", val=STYLE["three_line_table"])
    child(tblpr, "w:jc", val="center")
    clear(tblpr, "w:tblInd", "w:tblpPr")
    if widths:
        child(tblpr, "w:tblW", w=sum(widths), type="dxa")
        child(tblpr, "w:tblLayout", type="fixed")
        _set_grid(tbl, widths)
    else:
        child(tblpr, "w:tblW", w=5000, type="pct")
        child(tblpr, "w:tblLayout", type="autofit")
    mar = child(tblpr, "w:tblCellMar")
    pad = TIGHT_PAD if status == "overflow" else CELL_PAD
    for side in ("left", "right"):
        child(mar, f"w:{side}", w=pad // 2, type="dxa")
    child(tblpr, "w:tblLook", val="04A0", firstRow=1, lastRow=0, firstColumn=0, lastColumn=0, noHBand=0, noVBand=1)
    child(tblpr, "w:tblCaption", val=TABLE_TAG)
    _borders(tblpr)

    header_rows = [r for r in rows if r.find(f"{qn('w:trPr')}/{qn('w:tblHeader')}") is not None] or rows[:1]
    last_header = header_rows[-1]
    keep_whole = len(rows) <= SHORT_TABLE_ROWS
    half = int(pt * 2)
    for ri, tr in enumerate(rows):
        is_header = tr in header_rows
        trpr = child(tr, "w:trPr")
        child(trpr, "w:cantSplit")
        if is_header:
            child(trpr, "w:tblHeader")
        for ci, tc in enumerate(tr.findall(qn("w:tc"))):
            tcpr = child(tc, "w:tcPr")
            if widths and ci < len(widths):
                child(tcpr, "w:tcW", w=widths[ci], type="dxa")
            clear(tcpr, "w:tcBorders", "w:shd")
            if tr is last_header:
                tb = child(tcpr, "w:tcBorders")
                child(tb, "w:bottom", val="single", sz=4, space=0, color="auto")
            child(tcpr, "w:vAlign", val="center")
            for p in tc.findall(qn("w:p")):
                pp = ppr(p)
                jc = pp.find(qn("w:jc"))
                jc_val = jc.get(qn("w:val")) if jc is not None else "center"
                if is_header:
                    jc_val = "center"
                clear(pp, "w:ind", "w:spacing", "w:jc", "w:keepNext")
                set_pstyle(p, STYLE["table_text"])
                para_props(pp, jc=jc_val, first_line_chars=0, before=0, after=0, line=line, line_rule="exact",
                           keep_next=is_header or (keep_whole and ri < len(rows) - 1))
                for r in p.iter(qn("w:r")):
                    rpr = child(r, "w:rPr")
                    set_size(rpr, half)
                    if is_header:
                        set_bold(rpr, True)
                prpr = child(pp, "w:rPr")
                set_size(prpr, half)
    return info


def fit_all_tables(body) -> list[dict]:
    out: list[dict] = []
    n = 0
    for tbl in body.iter(qn("w:tbl")):
        if is_special_table(tbl) or tbl.getparent().tag == qn("w:tc"):
            continue
        n += 1
        out.append(fit_table(tbl, n))
    return out


def summarize(infos: list[dict]) -> list[str]:
    """给 lint / report 用的 WARN 文本。"""
    warns: list[str] = []
    for t in infos:
        label = t["caption"] or f"第 {t['index']} 张表"
        if not t["caption"]:
            warns.append(f"WARN {label} 没有表注：请在源 Markdown 表格下一行加 `Table: 标题 {{#tbl:id}}`")
        if t["status"] in ("tight", "overflow"):
            warns.append(f"WARN 表格超宽：{label}（{t['cols']} 列）{t['note']}；建议转置或拆表")
        elif t["status"] == "merged":
            warns.append(f"INFO {label} 含合并单元格，未重算列宽，请人工检查")
    return warns


def main() -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("docx_in")
    ap.add_argument("--out", default=None, help="默认原地覆盖")
    args = ap.parse_args()
    src = Path(args.docx_in)
    dst = Path(args.out) if args.out else src
    d = docx.Document(str(src))
    infos = fit_all_tables(d.element.body)
    d.save(str(dst))
    by = {}
    for t in infos:
        by[t["status"]] = by.get(t["status"], 0) + 1
    print(f"[table-fit] {len(infos)} 张表 → {dst}；{by}")
    for w in summarize(infos):
        print("  " + w)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
