"""oxml_utils.py — docx-compose / table-fit 共用的 OOXML 小工具与模板常量。

模板契约（_references/template/template_spec.json 的代码镜像）：A4、四边 1418 twips（2.5 cm）、版心 9070 twips；
正文宋体 / Times New Roman 12 pt；模板自带样式 ID 见 STYLE。Word 对子元素顺序敏感，所有插入都走 _insert_ordered。
"""
from __future__ import annotations

from docx.oxml import OxmlElement
from docx.oxml.ns import qn

PAGE_W, PAGE_H = 11906, 16838
MARGIN = 1418
HEADER_DIST, FOOTER_DIST = 851, 992
TEXT_W = PAGE_W - 2 * MARGIN            # 9070
SONG, HEI, KAI, LI = "宋体", "黑体", "楷体", "隶书"
TNR = "Times New Roman"
CODE_FONT = "Consolas"
CN_DIGITS = "零一二三四五六七八九"

# 模板 styles.xml 中的样式 ID（名称 → ID）
STYLE = {
    "normal": "a4",           # Normal：首行缩进 2 字符、两端对齐、12 pt
    "h1": "1", "h2": "2", "h3": "3",
    "table_text": "a9",       # 表格：18 pt 固定行距
    "table_caption": "a0",    # 表注：加粗 11 pt 居中 keepNext
    "figure_caption": "a1",   # 图注
    "figure": "aff1",         # 图片：居中 keepNext
    "ordered_list": "a3",     # 有序列表 (1) (2)
    "references": "a",        # 参考文献 [1]
    "three_line_table": "afb",
    "equation_table": "afe",
    "code_table": "aff0",     # 代码清单（表格样式，四边框）
    "toc_heading": "TOC", "toc1": "TOC1", "toc2": "TOC2", "toc3": "TOC3",
    "header": "af0", "footer": "ae",
}
# numbering.xml：numId 3 → abstractNum 4 "(%1)/%2)/%3."；numId 1 → 标题 "一、/1.1/1.1.1"；numId 4 → 参考文献 "[%1]"
NUM_ID_HEADINGS = 1
NUM_ID_ORDERED = 3
NUM_ID_REFERENCES = 4
ABSTRACT_NUM_ORDERED = 4

M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
M_T = f"{{{M_NS}}}t"

_ORDER: dict[str, list[str]] = {
    "w:pPr": ["pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl", "numPr",
              "suppressLineNumbers", "pBdr", "shd", "tabs", "suppressAutoHyphens", "kinsoku", "wordWrap",
              "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd",
              "snapToGrid", "spacing", "ind", "contextualSpacing", "mirrorIndents", "suppressOverlap", "jc",
              "textDirection", "textAlignment", "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr",
              "sectPr", "pPrChange"],
    "w:rPr": ["rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike", "dstrike", "outline",
              "shadow", "emboss", "imprint", "noProof", "snapToGrid", "vanish", "webHidden", "color", "spacing",
              "w", "kern", "position", "sz", "szCs", "highlight", "u", "effect", "bdr", "shd", "fitText",
              "vertAlign", "rtl", "cs", "em", "lang", "eastAsianLayout", "specVanish", "oMath"],
    "w:tblPr": ["tblStyle", "tblpPr", "tblOverlap", "bidiVisual", "tblStyleRowBandSize", "tblStyleColBandSize",
                "tblW", "jc", "tblCellSpacing", "tblInd", "tblBorders", "shd", "tblLayout", "tblCellMar",
                "tblLook", "tblCaption", "tblDescription"],
    "w:tcPr": ["cnfStyle", "tcW", "gridSpan", "hMerge", "vMerge", "tcBorders", "shd", "noWrap", "tcMar",
               "textDirection", "tcFitText", "vAlign", "hideMark"],
    "w:trPr": ["cnfStyle", "divId", "gridBefore", "gridAfter", "wBefore", "wAfter", "cantSplit", "trHeight",
               "tblHeader", "tblCellSpacing", "jc", "hidden"],
    "w:tblBorders": ["top", "start", "left", "bottom", "end", "right", "insideH", "insideV"],
    "w:tcBorders": ["top", "start", "left", "bottom", "end", "right", "insideH", "insideV", "tl2br", "tr2bl"],
    "w:pBdr": ["top", "left", "bottom", "right", "between", "bar"],
    "w:tblCellMar": ["top", "start", "left", "bottom", "end", "right"],
    "w:sectPr": ["headerReference", "footerReference", "footnotePr", "endnotePr", "type", "pgSz", "pgMar",
                 "paperSrc", "pgBorders", "lnNumType", "pgNumType", "cols", "formProt", "vAlign", "noEndnote",
                 "titlePg", "textDirection", "bidi", "rtlGutter", "docGrid"],
    "w:style": ["name", "aliases", "basedOn", "next", "link", "autoRedefine", "hidden", "uiPriority", "semiHidden",
                "unhideWhenUsed", "qFormat", "locked", "personal", "personalCompose", "personalReply", "rsid",
                "pPr", "rPr", "tblPr", "trPr", "tcPr", "tblStylePr"],
    "w:numPr": ["ilvl", "numId", "numberingChange", "ins"],
    "w:p": ["pPr"],
    "w:r": ["rPr"],
    "w:tc": ["tcPr"],
    "w:tbl": ["tblPr", "tblGrid", "tr"],
    "w:tr": ["tblPrEx", "trPr", "tc"],
}
# 这些容器的属性元素必须排在所有内容子元素（w:r/w:t/w:p…）之前
_PROPS_FIRST = {"w:p", "w:r", "w:tc", "w:tbl", "w:tr"}


def local(tag: str) -> str:
    return tag.rsplit("}", 1)[-1] if "}" in tag else tag.split(":")[-1]


def insert_ordered(parent, el) -> None:
    order = _ORDER.get(f"w:{local(parent.tag)}")
    name = local(el.tag)
    if order is None or name not in order:
        parent.append(el)
        return
    rank = order.index(name)
    props_first = f"w:{local(parent.tag)}" in _PROPS_FIRST
    for sib in parent:
        sib_name = local(sib.tag)
        if sib_name in order:
            if order.index(sib_name) > rank:
                sib.addprevious(el)
                return
        elif props_first:
            sib.addprevious(el)
            return
    parent.append(el)


def child(parent, tag: str, **attrs):
    """取或按 schema 顺序新建子元素；属性名不带前缀（val="x" → w:val）。"""
    el = parent.find(qn(tag))
    if el is None:
        el = OxmlElement(tag)
        insert_ordered(parent, el)
    for k, v in attrs.items():
        el.set(qn(f"w:{k}"), str(v))
    return el


def clear(parent, *tags: str) -> None:
    for t in tags:
        for el in parent.findall(qn(t)):
            parent.remove(el)


def ppr(p):
    return child(p, "w:pPr")


def set_fonts(rpr, west: str = TNR, east: str = SONG) -> None:
    rf = child(rpr, "w:rFonts")
    for a in ("ascii", "hAnsi", "cs"):
        rf.set(qn(f"w:{a}"), west)
    rf.set(qn("w:eastAsia"), east)
    for a in ("asciiTheme", "hAnsiTheme", "eastAsiaTheme", "cstheme", "hint"):
        rf.attrib.pop(qn(f"w:{a}"), None)


def set_size(rpr, half_pts: int) -> None:
    child(rpr, "w:sz", val=half_pts)
    child(rpr, "w:szCs", val=half_pts)


def set_bold(rpr, bold: bool) -> None:
    for tag in ("w:b", "w:bCs"):
        el = rpr.find(qn(tag))
        if bold:
            if el is None:
                el = OxmlElement(tag)
                insert_ordered(rpr, el)
            el.attrib.pop(qn("w:val"), None)
        elif el is not None:
            rpr.remove(el)


def para_props(pp, *, jc: str | None = None, first_line_chars: int | None = None, left: int | None = None,
               hanging: int | None = None, before: int | None = None, after: int | None = None,
               line: int | None = None, line_rule: str | None = None, keep_next: bool | None = None,
               keep_lines: bool | None = None) -> None:
    if keep_next is True:
        child(pp, "w:keepNext").attrib.pop(qn("w:val"), None)
    elif keep_next is False:
        child(pp, "w:keepNext", val=0)  # 显式关闭，覆盖表格样式继承的 keepNext
    if keep_lines is True:
        child(pp, "w:keepLines")
    elif keep_lines is False:
        clear(pp, "w:keepLines")
    if before is not None or after is not None or line is not None:
        sp = child(pp, "w:spacing")
        if before is not None:
            sp.set(qn("w:before"), str(before))
            sp.attrib.pop(qn("w:beforeLines"), None)
        if after is not None:
            sp.set(qn("w:after"), str(after))
            sp.attrib.pop(qn("w:afterLines"), None)
        if line is not None:
            sp.set(qn("w:line"), str(line))
            sp.set(qn("w:lineRule"), line_rule or "auto")
    if first_line_chars is not None or left is not None or hanging is not None:
        ind = child(pp, "w:ind")
        for a in ("firstLine", "firstLineChars", "left", "leftChars", "hanging", "hangingChars", "start",
                  "startChars", "right", "rightChars", "end"):
            ind.attrib.pop(qn(f"w:{a}"), None)
        if first_line_chars is not None:
            ind.set(qn("w:firstLineChars"), str(first_line_chars))
            ind.set(qn("w:firstLine"), str(first_line_chars * 12 // 5))
        if left is not None:
            ind.set(qn("w:left"), str(left))
        if hanging is not None:
            ind.set(qn("w:hanging"), str(hanging))
    if jc is not None:
        child(pp, "w:jc", val=jc)


def run(text: str, *, west: str = TNR, east: str = SONG, size: int | None = None, bold: bool = False,
        underline: bool = False):
    r = OxmlElement("w:r")
    rpr = OxmlElement("w:rPr")
    r.append(rpr)
    set_fonts(rpr, west, east)
    if bold:
        set_bold(rpr, True)
    if size:
        set_size(rpr, size)
    if underline:
        child(rpr, "w:u", val="single")
    t = OxmlElement("w:t")
    t.set(qn("xml:space"), "preserve")
    t.text = text
    r.append(t)
    return r


def tab_run():
    r = OxmlElement("w:r")
    r.append(OxmlElement("w:tab"))
    return r


def para(style: str | None = None):
    p = OxmlElement("w:p")
    if style:
        child(ppr(p), "w:pStyle", val=style)
    return p


def p_text(p) -> str:
    return "".join(t.text or "" for t in p.iter() if t.tag in (qn("w:t"), M_T))


def p_style(p) -> str:
    ps = p.find(f"{qn('w:pPr')}/{qn('w:pStyle')}")
    return ps.get(qn("w:val")) if ps is not None else ""


def set_pstyle(p, style_id: str) -> None:
    pp = ppr(p)
    clear(pp, "w:pStyle")
    child(pp, "w:pStyle", val=style_id)


def set_numbering(p, num_id: int, ilvl: int = 0) -> None:
    pp = ppr(p)
    clear(pp, "w:numPr")
    np_ = child(pp, "w:numPr")
    child(np_, "w:ilvl", val=ilvl)
    child(np_, "w:numId", val=num_id)


def page_break_para():
    p = OxmlElement("w:p")
    r = OxmlElement("w:r")
    br = OxmlElement("w:br")
    br.set(qn("w:type"), "page")
    r.append(br)
    p.append(r)
    return p


def cn_number(n: int) -> str:
    """1 → 一，10 → 十，12 → 十二，21 → 二十一。"""
    if n <= 0 or n >= 100:
        return str(n)
    if n < 10:
        return CN_DIGITS[n]
    tens, ones = divmod(n, 10)
    return ("" if tens == 1 else CN_DIGITS[tens]) + "十" + (CN_DIGITS[ones] if ones else "")


def find_style(styles_el, style_id: str):
    for st in styles_el.findall(qn("w:style")):
        if st.get(qn("w:styleId")) == style_id:
            return st
    return None


def find_style_by_name(styles_el, *names: str):
    lowered = {n.lower() for n in names}
    for st in styles_el.findall(qn("w:style")):
        nm = st.find(qn("w:name"))
        if st.get(qn("w:styleId")) in names or (nm is not None and (nm.get(qn("w:val")) or "").lower() in lowered):
            return st
    return None


def ensure_style(styles_el, style_id: str, name: str, kind: str = "paragraph", based_on: str | None = None):
    st = find_style_by_name(styles_el, style_id, name)
    if st is not None:
        return st
    st = OxmlElement("w:style")
    st.set(qn("w:type"), kind)
    st.set(qn("w:styleId"), style_id)
    child(st, "w:name", val=name)
    if based_on:
        child(st, "w:basedOn", val=based_on)
    child(st, "w:qFormat")
    styles_el.append(st)
    return st


def text_width_twips(text: str, size_pt: float = 12.0) -> int:
    """不换行所需宽度估算：全角字 1 em，半角字 0.5 em（大写 0.68，窄字符 0.3），1 em = size_pt*20 twips。"""
    em = size_pt * 20
    w = 0.0
    for ch in text:
        o = ord(ch)
        if o > 0x2E7F:
            w += em
        elif ch.isupper():
            w += 0.68 * em
        elif ch.isdigit():
            w += 0.5 * em
        elif ch in "il.,;:'|!()[]":
            w += 0.3 * em
        else:
            w += 0.5 * em
    return int(w + 0.5)
