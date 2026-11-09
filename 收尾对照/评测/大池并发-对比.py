# -*- coding: utf-8 -*-
"""大池：串行基线 vs 并发结果，逐格比对。

为什么不用「五表标准答案」评大池：
  那套答案对应的是日常小池那批源文档（702 块）。大池是另外 16 份大文档
  （50 万块），记录完全不同，用旧答案打分只会得到 10% 这种无意义的数。
  大池的正确性口径是"串行跑出来的结果"——同一池、同一模板、同一模型，
  并发如果没串，两次结果应当逐格一致。

比对规则：
  · 数值按 round(2) 比，空与空算一致；
  · 差异分三类：只有并发有值（多填）、只有串行有值（漏填）、两边都有但不同（错值）。
"""
import json
from pathlib import Path

from openpyxl import load_workbook

BASE = Path(r"C:\Users\覃敬涛\Desktop\研一项目")
SERIAL = BASE / "压测-大文档" / "结果"                 # 模板1..5/模板N-*.xlsx
CONC = BASE / "测试数据" / "并发压测-1103"             # 结果-模板N/模板N-*.xlsx
OUT = Path(__file__).resolve().parent / "_大池并发对比.txt"


def norm(v):
    if v is None:
        return ""
    if isinstance(v, bool):
        return str(v)
    if isinstance(v, (int, float)):
        return "%.2f" % round(float(v), 2)
    s = str(v).strip().replace(" ", "")
    try:
        return "%.2f" % round(float(s), 2)
    except Exception:
        return s


def find_xlsx(folder):
    hits = sorted(Path(folder).glob("*已填写.xlsx"))
    return hits[0] if hits else None


def main():
    lines = []
    total_cells = diff_cells = 0
    for n in (1, 4, 5):
        a = find_xlsx(SERIAL / ("模板%d" % n))
        b = find_xlsx(CONC / ("结果-模板%d" % n))
        lines.append("=" * 62)
        lines.append("模板%d" % n)
        lines.append("  串行：%s" % (a.name if a else "缺失"))
        lines.append("  并发：%s" % (b.name if b else "缺失"))
        if not (a and b):
            lines.append("  ✗ 有一边缺文件，跳过")
            continue

        wa = load_workbook(a, data_only=True).worksheets[0]
        wb = load_workbook(b, data_only=True).worksheets[0]
        rows = max(wa.max_row, wb.max_row)
        cols = max(wa.max_column, wb.max_column)
        extra_only_b, miss_in_b, wrong = [], [], []
        n_diff = 0
        for r in range(1, rows + 1):
            for c in range(1, cols + 1):
                va = norm(wa.cell(row=r, column=c).value)
                vb = norm(wb.cell(row=r, column=c).value)
                total_cells += 1
                if va == vb:
                    continue
                n_diff += 1
                diff_cells += 1
                coord = "%s%d" % (chr(64 + c) if c <= 26 else "?%d" % c, r)
                if not va and vb:
                    extra_only_b.append((coord, vb))
                elif va and not vb:
                    miss_in_b.append((coord, va))
                else:
                    wrong.append((coord, va, vb))

        lines.append("  对比格数：%d ｜ 差异：%d" % (rows * cols, n_diff))
        lines.append("    并发多填：%d ｜ 并发漏填：%d ｜ 值不同：%d"
                     % (len(extra_only_b), len(miss_in_b), len(wrong)))
        for tag, items in (("多填", extra_only_b), ("漏填", miss_in_b)):
            for coord, v in items[:5]:
                lines.append("      %s %s = %s" % (tag, coord, v))
        for coord, va, vb in wrong[:5]:
            lines.append("      值不同 %s：串行=%s ｜ 并发=%s" % (coord, va, vb))
        lines.append("  结论：%s" % ("逐格一致，并发未串数据" if n_diff == 0
                                  else "有差异，需逐项确认"))
    lines.append("")
    lines.append("全部：%d 格中 %d 格有差异" % (total_cells, diff_cells))
    text = "\n".join(lines)
    OUT.write_text(text, encoding="utf-8")
    print(text)


if __name__ == "__main__":
    main()
