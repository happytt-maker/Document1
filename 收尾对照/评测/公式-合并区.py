# -*- coding: utf-8 -*-
"""看模板1/4 的合并单元格，确认合计行的值能写到哪一格。"""
from pathlib import Path

from openpyxl import load_workbook

BASE = Path(r"C:\Users\覃敬涛\Desktop\研一项目\测试数据\模板表格")
OUT = Path(__file__).resolve().parent / "_合并区.txt"


def main():
    lines = []
    for name in ("模板1-设备采购汇总表.xlsx", "模板4-客诉处理台账.xlsx"):
        p = BASE / name
        ws = load_workbook(p).worksheets[0]
        lines.append("=" * 56)
        lines.append(name)
        lines.append("合并区：")
        for rng in ws.merged_cells.ranges:
            lines.append("  %s" % rng)
        lines.append("")
    text = "\n".join(lines)
    OUT.write_text(text, encoding="utf-8")
    print(text)


if __name__ == "__main__":
    main()
