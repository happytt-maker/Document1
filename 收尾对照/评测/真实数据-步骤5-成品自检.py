# -*- coding: utf-8 -*-
"""真实数据试跑 · 步骤5：读回成品的 xlsx 自检（不是读中间 json）。

要证明的是"交出去的那个文件本身没问题"，所以这里只认 writer/输出/ 下的成品：
  ① 行数、列数
  ② 表头是否原样保留
  ③ 每一格是不是数值类型（文本型的数字会让评委的公式算不出来）
  ④ 有无空行、有无漏格
  ⑤ 前 10 行 + 末 3 行原样打印
"""
import json
import os
from pathlib import Path

from openpyxl import load_workbook

XLSX = Path(os.environ.get("A23_OUT_DIR")
            or r"C:\Users\覃敬涛\Desktop\研一项目\测试数据\真实数据\04-试跑产物") \
    / "2025年中国城市经济百强全景报告-模板-已填写.xlsx"
OUT = Path(__file__).with_name("_成品自检.txt")


def main():
    L = []
    wb = load_workbook(XLSX)
    ws = wb[wb.sheetnames[0]]
    L.append("文件：%s" % XLSX.name)
    L.append("工作表：%s ｜ 尺寸 %s ｜ 最大行 %d 最大列 %d"
             % (ws.title, ws.dimensions, ws.max_row, ws.max_column))
    L.append("")

    head = [ws.cell(row=1, column=c).value for c in range(1, ws.max_column + 1)]
    L.append("【表头原样】%s" % head)

    n_data = ws.max_row - 1
    L.append("【数据行】%d 行（表头 1 行 + 数据 %d 行）" % (ws.max_row, n_data))

    empty_rows, nonnum, total = [], [], 0
    for r in range(2, ws.max_row + 1):
        vals = [ws.cell(row=r, column=c).value for c in range(1, ws.max_column + 1)]
        if all(v is None for v in vals):
            empty_rows.append(r)
            continue
        for c in range(1, ws.max_column + 1):
            v = ws.cell(row=r, column=c).value
            total += 1
            if v is None:
                nonnum.append((r, c, "空"))
            elif c > 1 and not isinstance(v, (int, float)):
                nonnum.append((r, c, "文本型：%r" % v))

    L.append("【空格/文本型数字】共 %d 格检查 → 异常 %d" % (total, len(nonnum)))
    for r, c, why in nonnum[:20]:
        L.append("    R%dC%d %s" % (r, c, why))
    L.append("【空行】%s" % (empty_rows if empty_rows else "无"))

    L.append("")
    L.append("【前 10 行】")
    L.append("  " + " ｜ ".join(str(h) for h in head))
    for r in range(2, 12):
        L.append("  " + " ｜ ".join(str(ws.cell(row=r, column=c).value)
                                    for c in range(1, ws.max_column + 1)))
    L.append("")
    L.append("【末 3 行】")
    for r in range(ws.max_row - 2, ws.max_row + 1):
        L.append("  " + " ｜ ".join(str(ws.cell(row=r, column=c).value)
                                    for c in range(1, ws.max_column + 1)))

    # 数字格式：看几格的 number_format，确认没被改成文本格式
    L.append("")
    L.append("【单元格格式抽样】")
    for coord in ("B2", "C2", "D2", "E2"):
        cell = ws[coord]
        L.append("  %s 值=%r 类型=%s number_format=%s"
                 % (coord, cell.value, type(cell.value).__name__, cell.number_format))

    txt = "\n".join(L)
    OUT.write_text(txt, encoding="utf-8")
    print(txt)


if __name__ == "__main__":
    main()
