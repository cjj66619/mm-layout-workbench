"""smoke_test.py — 用一个最小项目（3 节：摘要 + 正文 + 参考文献；含 1 表 1 图 1 公式 1 列表 1 代码块）
跑通 normalize → compose --no-pdf → lint --strict → report，并断言产物与计数。

用法：
    python scripts/smoke_test.py [--keep]        # --keep 保留临时项目目录便于排查

不需要 LibreOffice；需要 pandoc、python-docx。临时项目建在系统临时目录，跑完即删。
"""
from __future__ import annotations

import argparse
import json
import os
import shutil
import struct
import subprocess
import sys
import tempfile
import zlib
from pathlib import Path

REPO = Path(__file__).resolve().parents[1]
SKILLS = REPO / ".agents" / "skills"
STEPS = {
    "normalize": SKILLS / "text-normalize" / "scripts" / "normalize_md.py",
    "compose": SKILLS / "docx-compose" / "scripts" / "compose_docx.py",
    "lint": SKILLS / "layout-lint" / "scripts" / "lint_layout.py",
    "report": SKILLS / "layout-report" / "scripts" / "report_layout.py",
}

ABSTRACT = """# 摘要

本文针对"最小样例"问题建立了一个模型，结果见表 @tbl:demo 与图 @fig:demo，核心关系由式 @eq:demo 给出。

**关键词：** 冒烟测试；排版；模板
"""

BODY = """# 问题重述

## 问题背景

这是一段普通正文，包含直引号 "术语" 和一个数值 3.14，以及 25% 的比例。模型假设如下：

1. 假设一：数据无缺失；
2. 假设二：噪声独立同分布；
3. 假设三：参数在观测期内不变。

## 模型建立

目标函数为

$$
J(\\theta)=\\frac1n\\sum_{i=1}^{n}\\left(y_i-f(x_i;\\theta)\\right)^2
$$ {#eq:demo}

式中 $\\theta$ 为待估参数。表 表 @tbl:demo 给出三种方法的对比。

| 方法 | 准确率 | 宏 F1 | 备注 |
| --- | ---: | ---: | --- |
| 基线 | 0.812 | 0.790 | 规则 |
| 随机森林 | 0.943 | 0.951 | max_depth=10 |
| 本文方法 | 0.960 | 0.958 | 分层 |

: 三种方法在留出集上的表现 {#tbl:demo}

![模型流程示意](figures/demo.png){#fig:demo}

### 算法

```python
def fit(x, y):
    return sum(x) / len(x), "ok"
```
"""

REFS = """# 参考文献

[1] 张三, 李四. 一种最小样例方法[J]. 测试学报, 2024, 1(1): 1-2.

[2] Doe J. Smoke Testing for Layout Pipelines[M]. Nowhere: Nopress, 2023.
"""

PAPER_YAML = """title: "最小样例：排版冒烟测试"
lang: zh
keywords: []
sections:
  - 00_abstract.md
  - 01_body.md
  - 10_references.md
"""


def tiny_png(path: Path, w: int = 64, h: int = 32) -> None:
    """不依赖 Pillow，手写一张纯色 PNG。"""
    raw = b"".join(b"\x00" + b"\x40\x80\xc0" * w for _ in range(h))

    def chunk(tag: bytes, data: bytes) -> bytes:
        return struct.pack(">I", len(data)) + tag + data + struct.pack(">I", zlib.crc32(tag + data) & 0xFFFFFFFF)

    ihdr = struct.pack(">IIBBBBB", w, h, 8, 2, 0, 0, 0)
    path.write_bytes(b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", ihdr) + chunk(b"IDAT", zlib.compress(raw, 9)) + chunk(b"IEND", b""))


def make_project(root: Path) -> None:
    (root / "paper").mkdir(parents=True)
    (root / "polish" / "sections").mkdir(parents=True)
    (root / "figures").mkdir()
    (root / "paper" / "paper.yaml").write_text(PAPER_YAML, encoding="utf-8")
    for name, text in (("00_abstract.md", ABSTRACT), ("01_body.md", BODY), ("10_references.md", REFS)):
        (root / "polish" / "sections" / name).write_text(text, encoding="utf-8")
    tiny_png(root / "figures" / "demo.png")


def run(name: str, argv: list[str]) -> None:
    cmd = [sys.executable, str(STEPS[name]), *argv]
    print(f"--- {name}: {' '.join(cmd[2:])}", flush=True)
    r = subprocess.run(cmd, check=False)
    if r.returncode != 0:
        raise SystemExit(f"[smoke] {name} 失败 rc={r.returncode}")


def expect(cond: bool, msg: str, failures: list[str]) -> None:
    print(("  ok   " if cond else "  FAIL ") + msg)
    if not cond:
        failures.append(msg)


def main() -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--keep", action="store_true", help="保留临时项目目录")
    args = ap.parse_args()
    if shutil.which(os.environ.get("PANDOC_BIN", "pandoc")) is None:
        print("[smoke] 未找到 pandoc（或设 PANDOC_BIN）", file=sys.stderr)
        return 2

    tmp = Path(tempfile.mkdtemp(prefix="mm-layout-smoke-"))
    proj = tmp / "proj"
    failures: list[str] = []
    try:
        make_project(proj)
        run("normalize", [str(proj), "--strict"])
        run("compose", [str(proj), "--no-pdf"])
        run("lint", [str(proj), "--strict"])
        run("report", [str(proj)])

        layout = proj / "layout"
        for f in ("_src/sections/01_body.md", "NORMALIZE_LOG.md", "main_layout.docx", "compose_report.json",
                  "LAYOUT_LINT.md", "lint_report.json", "LAYOUT_REPORT.md"):
            expect((layout / f).is_file(), f"产物存在 layout/{f}", failures)

        src = (layout / "_src" / "sections" / "01_body.md").read_text(encoding="utf-8")
        expect("表 表 @tbl" not in src and "表 @tbl:demo" not in src, "normalize 去掉了“表 @tbl:”双前缀", failures)
        expect('"术语"' not in src and "“术语”" in src, "normalize 把直引号改成弯引号", failures)
        expect("0.943" in src and "3.14" in src and "25%" in src, "normalize 不改数字", failures)
        expect('"ok"' in src, "normalize 不动代码块里的直引号", failures)

        rep = json.loads((layout / "compose_report.json").read_text(encoding="utf-8"))
        c = rep["counts"]
        expect(c.get("tables") == 1 and c.get("tables_defined") == 1, f"compose 1 张表（{c.get('tables')}/{c.get('tables_defined')}）", failures)
        expect(c.get("images") == 1, f"compose 1 张图（{c.get('images')}）", failures)
        expect(c.get("display_equations") == 1, f"compose 1 个编号公式（{c.get('display_equations')}）", failures)
        expect(c.get("code_blocks") == 1, f"compose 1 个代码块（{c.get('code_blocks')}）", failures)
        expect(c.get("list_items", 0) >= 3, f"compose 列表项 ≥ 3（{c.get('list_items')}）", failures)
        expect(c.get("reference_entries") == 2, f"compose 2 条参考文献（{c.get('reference_entries')}）", failures)
        expect(c.get("h1", 0) >= 1 and c.get("h2", 0) >= 2 and c.get("h3", 0) >= 1, "compose 标题三级齐全", failures)
        expect(rep["tables"] and rep["tables"][0]["status"] == "fit" and rep["tables"][0]["font_pt"] == 12.0,
               "窄表 fit 且保持 12 pt", failures)
        expect(rep["files"].get("pdf") is None, "--no-pdf 时不产出 PDF", failures)

        lint = json.loads((layout / "lint_report.json").read_text(encoding="utf-8"))
        expect(not lint["fails"], f"lint 无 FAIL：{lint['fails'][:3]}", failures)
        expect(lint["counts"].get("tables") == 1, "lint 只统计 hwb-table 表（封面表不算）", failures)
        expect(lint["counts"].get("code_tables") == 1 and lint["counts"].get("equation_tables") == 1,
               "lint 识别公式表/代码表", failures)
        noisy = [w for w in lint["warns"] if "直引号" in w or "双前缀" in w or "圆点" in w or "表注" in w]
        expect(not noisy, f"lint 无引号/双前缀/圆点/表注误报：{noisy[:2]}", failures)

        report_md = (layout / "LAYOUT_REPORT.md").read_text(encoding="utf-8")
        expect(report_md.startswith("# LAYOUT_REPORT — ") and "## 人工待办" in report_md, "report 结构完整", failures)

        touched = [p for p in (proj / "polish").rglob("*") if p.is_file() and p.stat().st_mtime > (layout / "main_layout.docx").stat().st_mtime]
        expect(not touched, "polish/ 未被修改", failures)
    finally:
        if args.keep:
            print(f"[smoke] 保留临时项目：{proj}")
        else:
            shutil.rmtree(tmp, ignore_errors=True)

    if failures:
        print(f"\n[smoke] FAIL ×{len(failures)}")
        return 1
    print("\n[smoke] PASS")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
