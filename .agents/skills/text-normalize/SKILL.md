---
name: text-normalize
description: 排版前的 Markdown 机械规范化。把 polish/sections（或 paper/sections）复制到 layout/_src/sections，去掉“表 @tbl:”双前缀、直引号转弯引号、统一分页标记，并做事实零漂移校验（数字/公式/引用/图路径不变）。docx-compose 之前必跑。
---

# text-normalize

只改**标点与引用写法**，不改任何事实。原文件不动，全部在 `layout/_src/sections/` 副本上操作。

## 用法

```bash
python .agents/skills/text-normalize/scripts/normalize_md.py <proj> [--src polish/sections] [--strict]
```

- 源目录默认 `polish/sections`，没有则回退 `paper/sections`；`--src` 可显式指定（相对项目根）。
- 输出：`layout/_src/sections/*.md` + `layout/NORMALIZE_LOG.md`。
- `--strict`：事实漂移或未定义交叉引用时返回非零（run_layout.py 会据此停下）。

## 做什么

1. 交叉引用双前缀：`表 @tbl:x` / `图 @fig:x` / `式 @eq:x` → `@tbl:x` …（compose 会把 `@ref` 展开成“表N/图N/式(N)”，
   不去重 Word 里就是“表 表30”）；已写死的 `表 表3`、`图 图2`、`式 式(4)` 同样去重。
2. 直引号成对 → 中文弯引号：`"术语"` → `“术语”`；`”术语”` 这类错配也修成 `“术语”`。
   代码块、行内代码、公式内部不动。
3. 段尾空白、连续 3 个以上空行压成 2 个；`\newpage` / `\clearpage` / `<div style="page-break…">` 统一成 `::: {.page-break}`。

## 零漂移校验

处理前后各抽一次“数字（含小数/百分号/单位）、行内公式、行间公式、@ref 引用、图片路径”多重集，
任何差异都记 FAIL 写进 `NORMALIZE_LOG.md`，并且**不写出**该文件的副本。看到 FAIL 先怀疑正则误伤，
修脚本，不要手改论文。

## 不要做

- 不改句子、术语、数字、公式、图路径；不合并/拆分段落；不改标题层级。
- 不碰 `polish/`、`paper/` 原文件。
