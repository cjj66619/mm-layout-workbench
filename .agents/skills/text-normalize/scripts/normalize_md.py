#!/usr/bin/env python3
"""normalize_md.py — 排版前的 Markdown 文本规范化（只改标点/引用写法，不改任何事实）。

把 <proj>/polish/sections/*.md（缺省回退 paper/sections）复制到 <proj>/layout/_src/sections/，
在副本上做与"排版"有关、与"内容"无关的机械修正：

1. 交叉引用双前缀：`表 @tbl:x` / `图 @fig:x` / `式 @eq:x` → `@tbl:x` …（合成时由 @ref 生成“表N/图N/式(N)”，
   否则 Word 里会出现“表 表30”）；`表 表3`、`图 图2`、`式 式(4)` 这类已经写死的双前缀也一并去重。
2. 直引号成对转弯引号：`"术语"` → `“术语”`；已经错配的 `”术语”` 修成 `“术语”`。代码块、行内代码、公式不动。
3. 段尾多余空白、连续 3 个以上空行压成 2 个；`\\newpage`/`\\clearpage`/`<div style="page-break…">` 统一成
   pandoc 认识的 `::: {.page-break}` 块。
4. 中文句子里的半角逗号/句号后跟中文时不改（polish 阶段已管；本脚本只处理引号，避免误伤小数点）。

零漂移：处理前后分别抽取“数字（含小数/百分号/单位）、行内公式、行间公式、@ref 引用、图片路径”多重集，
任何差异都记 FAIL 并写进 NORMALIZE_LOG.md；--strict 下有 FAIL 以非零退出。

用法：
    python normalize_md.py <proj> [--src polish/sections] [--out layout/_src/sections] [--log layout/NORMALIZE_LOG.md] [--strict]
"""
from __future__ import annotations

import argparse
import re
import shutil
import sys
from collections import Counter
from pathlib import Path

CODE_BLOCK = re.compile(r"```.*?```", re.S)
INLINE_CODE = re.compile(r"`[^`\n]+`")
DISPLAY_MATH = re.compile(r"\$\$(?:[^$]|\$(?!\$))*\$\$")
INLINE_MATH = re.compile(r"(?<![\\$])\$(?!\$)(?:[^$\n\\]|\\.)+\$(?!\$)")
REF = re.compile(r"(?<![A-Za-z0-9_@])@((?:fig|tbl|eq):[\w:.-]*[\w])")
IMAGE = re.compile(r"!\[[^\]]*\]\(([^)\s]+)")
NUMBER = re.compile(r"(?<![\w.])[-+]?\d+(?:[.,]\d+)*(?:\s?[%‰]|\s?(?:kHz|Hz|s|ms|rpm|hp|mm|cm|m|kg|MB|GB|dB|°C|个|条|段|次|张|篇|页|维|折|倍))?")

DOUBLE_PREFIX_REF = re.compile(r"(?:[表图式]\s*)+(?=@(?:tbl|fig|eq):)")
DOUBLE_PREFIX_LITERAL = re.compile(r"([表图])\s+\1(?=\d)|(式)\s+\2(?=\()")
PAGE_BREAK_RAW = re.compile(r"^(?:\\newpage|\\clearpage|\\pagebreak|<div[^>]*page-break[^>]*>\s*</div>)\s*$", re.M)
STRAIGHT_QUOTED = re.compile(r'(?<![A-Za-z0-9\\])"([^"\n]{1,80})"(?![A-Za-z0-9])')
MISMATCH_QUOTED = re.compile(r"(?<![\u201c])[\u201d]([^\u201c\u201d\n]{1,80})[\u201d]")


def _mask(text: str) -> tuple[str, list[str]]:
    """把代码块/行内代码/公式替换成占位符，避免规则误伤；返回 (masked, 原片段列表)。"""
    kept: list[str] = []

    def _keep(m: re.Match) -> str:
        kept.append(m.group(0))
        return f"\u0000{len(kept) - 1}\u0000"

    masked = CODE_BLOCK.sub(_keep, text)
    masked = DISPLAY_MATH.sub(_keep, masked)
    masked = INLINE_CODE.sub(_keep, masked)
    masked = INLINE_MATH.sub(_keep, masked)
    return masked, kept


def _unmask(masked: str, kept: list[str]) -> str:
    return re.sub("\u0000(\\d+)\u0000", lambda m: kept[int(m.group(1))], masked)


def normalize_text(text: str) -> tuple[str, Counter]:
    stats: Counter = Counter()
    masked, kept = _mask(text)

    masked, n = DOUBLE_PREFIX_REF.subn("", masked)
    stats["ref_double_prefix"] += n
    masked, n = DOUBLE_PREFIX_LITERAL.subn(lambda m: m.group(1) or m.group(2), masked)
    stats["literal_double_prefix"] += n

    masked, n = STRAIGHT_QUOTED.subn("\u201c\\1\u201d", masked)
    stats["straight_quotes_paired"] += n
    masked, n = MISMATCH_QUOTED.subn("\u201c\\1\u201d", masked)
    stats["mismatched_quotes_fixed"] += n

    masked, n = PAGE_BREAK_RAW.subn("::: {.page-break}\n:::", masked)
    stats["page_breaks_unified"] += n

    lines = [ln.rstrip() for ln in masked.split("\n")]
    if any(ln != raw for ln, raw in zip(lines, masked.split("\n"))):
        stats["trailing_ws_lines"] += sum(1 for ln, raw in zip(lines, masked.split("\n")) if ln != raw)
    masked = "\n".join(lines)
    masked, n = re.subn(r"\n{4,}", "\n\n\n", masked)
    stats["blank_runs_collapsed"] += n

    out = _unmask(masked, kept)
    if not out.endswith("\n"):
        out += "\n"
    return out, stats


def facts(text: str) -> dict[str, Counter]:
    code_free = CODE_BLOCK.sub(" ", text)
    display = Counter(m.group(0).strip() for m in DISPLAY_MATH.finditer(code_free))
    no_display = DISPLAY_MATH.sub(" ", code_free)
    inline = Counter(m.group(0) for m in INLINE_MATH.finditer(no_display))
    plain = INLINE_MATH.sub(" ", INLINE_CODE.sub(" ", no_display))
    refs = Counter(m.group(1) for m in REF.finditer(plain))
    images = Counter(m.group(1) for m in IMAGE.finditer(code_free))
    numbers = Counter(m.group(0).replace(" ", "") for m in NUMBER.finditer(plain) if m.group(0).strip())
    return {"display_math": display, "inline_math": inline, "refs": refs, "images": images, "numbers": numbers}


def diff_facts(before: dict[str, Counter], after: dict[str, Counter]) -> list[str]:
    fails: list[str] = []
    for kind in before:
        lost = before[kind] - after[kind]
        gained = after[kind] - before[kind]
        for item, n in lost.items():
            fails.append(f"{kind} 丢失 ×{n}: {item[:80]}")
        for item, n in gained.items():
            fails.append(f"{kind} 新增 ×{n}: {item[:80]}")
    return fails


def pick_source(proj: Path, src: str | None) -> Path:
    if src:
        p = proj / src
        if not p.is_dir():
            sys.exit(f"[error] 指定的章节目录不存在: {p}")
        return p
    for cand in ("polish/sections", "paper/sections"):
        p = proj / cand
        if p.is_dir() and any(p.glob("*.md")):
            return p
    sys.exit(f"[error] {proj} 下没有 polish/sections 或 paper/sections")


def main() -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("proj", help="项目根目录（含 polish/ 或 paper/）")
    ap.add_argument("--src", default=None, help="章节目录（相对项目根），默认 polish/sections，缺省回退 paper/sections")
    ap.add_argument("--out", default="layout/_src/sections", help="规范化副本目录（相对项目根）")
    ap.add_argument("--log", default="layout/NORMALIZE_LOG.md")
    ap.add_argument("--strict", action="store_true", help="事实漂移或未定义引用时以非零退出")
    args = ap.parse_args()

    proj = Path(args.proj).resolve()
    src = pick_source(proj, args.src)
    out = proj / args.out
    if out.exists():
        shutil.rmtree(out)
    out.mkdir(parents=True)

    files = sorted(p for p in src.glob("*.md") if not p.name.startswith("_"))
    total: Counter = Counter()
    per_file: list[tuple[str, Counter, list[str]]] = []
    all_fails: list[str] = []
    for f in files:
        raw = f.read_text(encoding="utf-8")
        new, stats = normalize_text(raw)
        fails = diff_facts(facts(raw), facts(new))
        (out / f.name).write_text(new, encoding="utf-8")
        total.update(stats)
        per_file.append((f.name, stats, fails))
        all_fails.extend(f"{f.name}: {x}" for x in fails)

    # 全局：@ref 定义/引用一致性（未定义引用会在 Word 里原样露出）
    joined = "\n".join((out / f.name).read_text(encoding="utf-8") for f in files)
    defined = set(re.findall(r"\{[^}]*#((?:fig|tbl|eq):[\w:.-]+)", joined))
    used = set(REF.findall(CODE_BLOCK.sub(" ", joined)))
    undefined = sorted(used - defined)
    unused = sorted(defined - used)

    log = proj / args.log
    log.parent.mkdir(parents=True, exist_ok=True)
    lines = ["# NORMALIZE_LOG — 文本规范化记录", "",
             f"- 源目录：`{src.relative_to(proj).as_posix()}`（{len(files)} 个文件）",
             f"- 输出：`{out.relative_to(proj).as_posix()}`",
             f"- 事实零漂移：{'PASS' if not all_fails else f'FAIL ×{len(all_fails)}'}",
             f"- 未定义的交叉引用：{len(undefined)}", "",
             "## 修正统计", "", "| 规则 | 次数 |", "| --- | ---: |"]
    for k in ("ref_double_prefix", "literal_double_prefix", "straight_quotes_paired", "mismatched_quotes_fixed",
              "page_breaks_unified", "trailing_ws_lines", "blank_runs_collapsed"):
        lines.append(f"| {k} | {total.get(k, 0)} |")
    lines += ["", "## 逐文件", "", "| 文件 | 双前缀 | 引号 | 漂移 |", "| --- | ---: | ---: | --- |"]
    for name, st, fails in per_file:
        lines.append(f"| {name} | {st.get('ref_double_prefix', 0) + st.get('literal_double_prefix', 0)} | "
                     f"{st.get('straight_quotes_paired', 0) + st.get('mismatched_quotes_fixed', 0)} | "
                     f"{'FAIL' if fails else 'PASS'} |")
    if all_fails:
        lines += ["", "## FAIL 明细", ""] + [f"- {x}" for x in all_fails]
    if undefined:
        lines += ["", "## 未定义的交叉引用（Word 中会原样出现，请回 polish 修正）", ""] + [f"- `@{x}`" for x in undefined]
    if unused:
        lines += ["", "## 定义了但正文未引用的标签（仅提示）", ""] + [f"- `{x}`" for x in unused]
    log.write_text("\n".join(lines) + "\n", encoding="utf-8")

    print(f"[normalize] {len(files)} 文件 → {out}；双前缀 {total['ref_double_prefix'] + total['literal_double_prefix']}，"
          f"引号 {total['straight_quotes_paired'] + total['mismatched_quotes_fixed']}，"
          f"漂移 FAIL {len(all_fails)}，未定义引用 {len(undefined)}；日志 {log}")
    if args.strict and (all_fails or undefined):
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
