---
name: table-fit
description: 把 DOCX 普通表格整理成模板三线表并按内容自适应列宽、12→10.5→9 pt 自动降字号，超宽时压列宽并报 WARN 建议转置/拆表。由 docx-compose 自动调用；也可单独修一份已有 Word。改表格宽度/换行/居中/字号策略时看这里。
---

# table-fit

## 用法

```bash
# 通常不用单独跑，compose_docx.py 会调用 fit_all_tables()
python .agents/skills/table-fit/scripts/table_fit.py in.docx [--out out.docx]
```

## 规则（用户已确认，改之前先问）

1. 列宽按“不换行所需宽度”估算（`text_width_twips`：全角 1 em、大写 0.68、数字 0.5、窄字符 0.3，×1.06 安全系数，
   单元格边距 `CELL_PAD = 216`）；窄表按内容宽度**居中**，不拉满版心 `TEXT_W = 9070`。
2. 超版心时依次尝试 12 → 10.5 → 9 pt（`FONT_LADDER`）。每档：
   - 全部列不换行装得下 → `fit`；
   - 让长文本列在 ≥ `COMFY_EM = 8` em 前提下换行 → `wrap`；
3. 9 pt 仍超：文本列压到最长不可断单元 → `tight`（WARN）；再超 → 边距收到 `TIGHT_PAD = 80` 并等比压缩 → `overflow`（WARN）。
   **不自动横排**，WARN 文案固定给“建议转置或拆表”。
4. 三线表：顶/底线 1.5 pt、表头下线 0.5 pt，表头加粗居中、`tblHeader` 跨页重复；单元格垂直居中；行 `cantSplit`；
   ≤ `SHORT_TABLE_ROWS = 8` 行的短表整体 keepNext，长表允许跨页。
5. 含合并单元格的表只套样式不算宽（`merged`）；公式表 `hwb-equation`、代码表 `hwb-code`、嵌套表不处理。
6. 处理完在 `tblPr/tblCaption` 打 `hwb-table` 标记，layout-lint 只检查带该标记的表。

## 输出

`fit_all_tables` 返回每表的 `{index, caption, rows, cols, status, font_pt, note, width}`，
compose 写进 `compose_report.json["tables"]`，lint / report 据此列 WARN。

## 不要做

- 不改单元格文字；不删列/合并列；不横排；不给长表每个单元格加 keepNext（会把整表推到下一页留白）。
