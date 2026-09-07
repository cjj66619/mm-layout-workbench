"""report_layout.py — 汇总 compose_report.json + lint_report.json + NORMALIZE_LOG.md → layout/LAYOUT_REPORT.md。

用法：
    python report_layout.py <proj> [--out layout/LAYOUT_REPORT.md]

报告内容：输入源与产物、页数/表图公式计数、表格状态一览（字号/列数/状态）、WARN 清单与建议、目录、
人工待办（封面年份 / 超宽表 / 摘要超 1 页 / Word 中 F9 更新目录）。只读取 layout/ 下的 JSON 与日志，不碰正文。
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path

STATUS_LABEL = {"fit": "适配", "wrap": "折行", "tight": "紧凑(WARN)", "overflow": "超宽(WARN)", "merged": "合并单元格"}
WIDE_STATUSES = ("tight", "overflow")
ABSTRACT_MAX_PAGES = 1


def load_json(path: Path) -> dict:
    if not path.is_file():
        return {}
    return json.loads(path.read_text(encoding="utf-8"))


def normalize_summary(log: Path) -> list[str]:
    """从 NORMALIZE_LOG.md 头部抽取“- 源目录 / 事实零漂移 / 未定义引用”几行。"""
    if not log.is_file():
        return ["- 未找到 NORMALIZE_LOG.md（未跑 text-normalize？）"]
    out = []
    for line in log.read_text(encoding="utf-8").splitlines():
        if line.startswith("- "):
            out.append(line)
        elif line.startswith("## "):
            break
    return out or ["- NORMALIZE_LOG.md 无摘要行"]


def md_escape(s: str) -> str:
    return s.replace("|", "\\|").replace("\n", " ")


def table_rows(tables: list[dict]) -> list[str]:
    rows = ["| # | 表注 | 行×列 | 字号 | 状态 | 备注 |", "| ---: | --- | :---: | ---: | --- | --- |"]
    for t in tables:
        cap = md_escape((t.get("caption") or "（无表注）")[:48])
        status = t.get("status", "?")
        rows.append(f"| {t.get('index')} | {cap} | {t.get('rows')}×{t.get('cols')} | {t.get('font_pt')} | "
                    f"{STATUS_LABEL.get(status, status)} | {md_escape(t.get('note') or '')} |")
    return rows


def toc_lines(toc: list[dict]) -> list[str]:
    out = []
    for e in toc:
        indent = "  " * (int(e.get("level", 1)) - 1)
        page = e.get("page")
        out.append(f"{indent}- {md_escape(e.get('text', ''))}{f'  ……{page}' if page else ''}")
    return out or ["- （compose 未输出目录，或未渲 PDF 回填页码）"]


def manual_todos(counts: dict, tables: list[dict], lint_counts: dict, has_pdf: bool) -> list[str]:
    todo = ["封面：核对年份 / 学校 / 参赛队号（工作流不会自动改封面，队号不要写进仓库）"]
    wide = [f"表{t['index']}" if not (t.get("caption") or "").startswith("表") else (t.get("caption") or "").split(" ")[0]
            for t in tables if t.get("status") in WIDE_STATUSES]
    if wide:
        todo.append(f"超宽表 {len(wide)} 张（{'、'.join(wide[:12])}{'…' if len(wide) > 12 else ''}）："
                    "已降到 9 pt 并压列宽，请在 Word 中目视确认；建议回 polish 转置/拆表/缩短表头")
    ab = lint_counts.get("abstract_pages")
    if isinstance(ab, int) and ab > ABSTRACT_MAX_PAGES:
        todo.append(f"摘要占 {ab} 页（竞赛惯例 {ABSTRACT_MAX_PAGES} 页）：回 polish 压缩 00_abstract.md，工作流不改文字")
    if has_pdf:
        todo.append("目录：页码已按 LibreOffice 渲染结果回填，Word 中字体/断行略有差异，交稿前在 Word 里 Ctrl+A → F9 更新一次目录")
    else:
        todo.append("目录：本次未渲 PDF（--no-pdf 或无 soffice），页码为空，Word 中 Ctrl+A → F9 更新目录")
    if counts.get("tables", 0) != counts.get("tables_defined", 0):
        todo.append(f"有 {counts.get('tables', 0) - counts.get('tables_defined', 0)} 张表没有表注（不占表号）："
                    "回 polish 在表格下一行补 `Table: 标题 {#tbl:id}`")
    todo.append("交稿前用 Word 另存 PDF（不要用 LibreOffice 的 PDF 交稿：以 `=` 开头的公式会渲成 ¿）")
    return todo


def build_report(proj: Path, compose: dict, lint: dict, norm_lines: list[str]) -> str:
    counts, tables, toc = compose.get("counts", {}), compose.get("tables", []), compose.get("toc", [])
    files = compose.get("files", {})
    lc = lint.get("counts", {})
    fails, warns, infos = lint.get("fails", []), lint.get("warns", []), lint.get("infos", [])
    verdict = "FAIL" if fails else ("WARN" if warns else "PASS")
    has_pdf = bool(files.get("pdf")) and (proj / "layout" / Path(files["pdf"]).name).is_file()
    status = lc.get("table_status") or {}
    status_str = " / ".join(f"{STATUS_LABEL.get(k, k)} {v}" for k, v in status.items()) or "—"

    def rel(p: str | None) -> str:
        if not p:
            return "—"
        try:
            return f"`{Path(p).resolve().relative_to(proj.resolve())}`"
        except ValueError:
            return f"`{p}`"

    lines = [f"# LAYOUT_REPORT — {verdict}", "",
             f"- 项目：`{proj.name}`",
             f"- 结论：lint **{verdict}**（FAIL {len(fails)} / WARN {len(warns)} / PASS {len(lint.get('passes', []))} / INFO {len(infos)}）",
             f"- 产物：{rel(files.get('docx'))}" + (f"、{rel(files.get('pdf'))}" if has_pdf else "（未渲 PDF）"),
             f"- 模板：`{Path(files.get('template', '')).name or '—'}`", "",
             "## 输入源（text-normalize）", "", *norm_lines,
             f"- 排版源目录：{rel(files.get('sections_dir'))}", "",
             "## 计数", "", "| 项目 | 数量 |", "| --- | ---: |",
             f"| PDF 页数 | {counts.get('pdf_pages', lc.get('pdf_pages', '—'))} |",
             f"| 摘要页数 | {lc.get('abstract_pages', '—')} |",
             f"| 标题 一/二/三级 | {counts.get('h1', 0)} / {counts.get('h2', 0)} / {counts.get('h3', 0)} |",
             f"| 三线表（有表注） | {counts.get('tables', 0)}（{counts.get('tables_defined', 0)}） |",
             f"| 图 | {counts.get('images', counts.get('figures_defined', 0))} |",
             f"| 编号公式 | {counts.get('display_equations', 0)} |",
             f"| 代码清单 | {counts.get('code_blocks', 0)} |",
             f"| 参考文献 | {counts.get('reference_entries', 0)} |",
             f"| 列表项 | {counts.get('list_items', 0)} |",
             f"| 目录条目 | {len(toc)} |", "",
             "## 表格状态", "", f"- 汇总：{status_str}",
             "- 规则：12 → 10.5 → 9 pt 自动降字号；9 pt 仍不够 → 压列宽并标 WARN（不自动横排）", "",
             *table_rows(tables), ""]
    if fails:
        lines += ["## FAIL（必须修）", "", *[f"- {m}" for m in fails], ""]
    lines += ["## WARN 与建议", ""]
    lines += [f"- {m}" for m in warns] or ["- 无"]
    lines += ["", "compose 警告：", ""] + ([f"- {m}" for m in compose.get("warnings", [])] or ["- 无"])
    if infos:
        lines += ["", "INFO：", "", *[f"- {m}" for m in infos]]
    lines += ["", "## 目录", "", *toc_lines(toc), "",
              "## 人工待办", "", *[f"{i}. {m}" for i, m in enumerate(manual_todos(counts, tables, lc, has_pdf), 1)], "",
              "## 详细报告", "", "- `layout/LAYOUT_LINT.md`（逐项 PASS/WARN/FAIL）", "- `layout/NORMALIZE_LOG.md`（规范化与事实零漂移）",
              "- `layout/compose_report.json` / `layout/lint_report.json`（机器可读）", ""]
    return "\n".join(lines)


def main() -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("proj")
    ap.add_argument("--compose-json", default="layout/compose_report.json")
    ap.add_argument("--lint-json", default="layout/lint_report.json")
    ap.add_argument("--normalize-log", default="layout/NORMALIZE_LOG.md")
    ap.add_argument("--out", default="layout/LAYOUT_REPORT.md")
    args = ap.parse_args()
    proj = Path(args.proj).resolve()
    compose = load_json(proj / args.compose_json)
    if not compose:
        print(f"[error] 未找到 {proj / args.compose_json}，先跑 docx-compose", file=sys.stderr)
        return 2
    lint = load_json(proj / args.lint_json)
    if not lint:
        print(f"[warn] 未找到 {proj / args.lint_json}，报告不含 lint 结论", file=sys.stderr)
    md = build_report(proj, compose, lint, normalize_summary(proj / args.normalize_log))
    out = proj / args.out
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(md, encoding="utf-8")
    verdict = re.match(r"# LAYOUT_REPORT — (\w+)", md).group(1)
    print(f"[report] {verdict}  pages={compose.get('counts', {}).get('pdf_pages')} "
          f"tables={compose.get('counts', {}).get('tables')} warn={len(lint.get('warns', []))} → {out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
