"""run_layout.py — 一键跑完排版管线：normalize → compose → lint → report。任一步失败立即停止。

用法：
    python run_layout.py <proj> [--no-pdf] [--fig-max-cm 15] [--toc-depth 3] [--strict]

只写 <proj>/layout/；不改 paper/、polish/、数据、代码、图。
不依赖 bash：用 subprocess 直接以当前解释器调用各步脚本，Windows / Linux 均可。
"""
from __future__ import annotations

import argparse
import subprocess
import sys
import time
from pathlib import Path

SKILLS = Path(__file__).resolve().parents[2]
STEPS = {
    "normalize": SKILLS / "text-normalize" / "scripts" / "normalize_md.py",
    "compose": SKILLS / "docx-compose" / "scripts" / "compose_docx.py",
    "lint": SKILLS / "layout-lint" / "scripts" / "lint_layout.py",
    "report": SKILLS / "layout-report" / "scripts" / "report_layout.py",
}


def run_step(name: str, argv: list[str]) -> int:
    print(f"\n=== [{name}] {' '.join(argv)}", flush=True)
    t0 = time.monotonic()
    rc = subprocess.run([sys.executable, str(STEPS[name]), *argv], check=False).returncode
    print(f"=== [{name}] rc={rc}  {time.monotonic() - t0:.1f}s", flush=True)
    return rc


def main() -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("proj", help="项目根目录（含 paper/paper.yaml 与 polish/sections 或 paper/sections）")
    ap.add_argument("--src", default=None, help="章节目录（相对项目根），默认 polish/sections → paper/sections")
    ap.add_argument("--no-pdf", action="store_true", help="不调 LibreOffice 渲 PDF（目录页码留空，Word 中 F9 更新）")
    ap.add_argument("--fig-max-cm", type=float, default=15.0)
    ap.add_argument("--toc-depth", type=int, default=3, choices=(1, 2, 3))
    ap.add_argument("--strict", action="store_true", help="normalize 漂移 / lint FAIL 时以非零退出")
    args = ap.parse_args()
    proj = Path(args.proj).resolve()
    if not proj.is_dir():
        print(f"[error] 项目目录不存在：{proj}", file=sys.stderr)
        return 2
    missing = [p for p in STEPS.values() if not p.is_file()]
    if missing:
        print(f"[error] 缺少脚本：{missing}", file=sys.stderr)
        return 2

    strict = ["--strict"] if args.strict else []
    src = ["--src", args.src] if args.src else []
    plan = [
        ("normalize", [str(proj), *src, *strict]),
        ("compose", [str(proj), "--fig-max-cm", str(args.fig_max_cm), "--toc-depth", str(args.toc_depth),
                     *(["--no-pdf"] if args.no_pdf else [])]),
        ("lint", [str(proj), *strict]),
        ("report", [str(proj)]),
    ]
    for name, argv in plan:
        rc = run_step(name, argv)
        if rc != 0:
            print(f"\n[run_layout] 在 {name} 步失败（rc={rc}），已停止；产物见 {proj / 'layout'}", file=sys.stderr)
            return rc
    print(f"\n[run_layout] 完成 → {proj / 'layout' / 'LAYOUT_REPORT.md'}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
