# -*- coding: utf-8 -*-
"""看模板1 / 模板4 合计行附近的单元格结构（A 列标签 + 各行值）。"""
from pathlib import Path

from openpyxl import load_workbook

BASE = Path(r"C:\Users\覃敬涛\Desktop\研一项目\测试数据\模板表格")
OUT = Path(__file__).resolve().parent / "_公式上下文.txt"

TARGETS = [("模板1-设备采购汇总表.xlsx", 16), ("模板4-客诉处理台账.xlsx", 14)]


def dump(p, n):
    lines = ["=" * 60, p.name]
    ws = load_workbook(p).worksheets[0]
    lines.append("工作表：%s ｜ %d 行 × %d 列" % (ws.title, ws.max_row, ws.max_column))
    for r in range(1, min(n, ws.max_row) + 1):
        cells = []
        for c in range(1, ws.max_column + 1):
            cc = ws.cell(row=r, column=c)
            if cc.value is None:
                continue
            cells.append("%s=%s" % (cc.column_letter, cc.value))
        lines.append("  第%2d行  %s" % (r, " ｜ ".join(cells)))
    return lines


def main():
    out = []
    for name, n in TARGETS:
        out += dump(BASE / name, n)
        out.append("")
    text = "\n".join(out)
    OUT.write_text(text, encoding="utf-8")
    print(text)


if __name__ == "__main__":
    main()
