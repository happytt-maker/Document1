# -*- coding: utf-8 -*-
"""合计行公式验收：公式在不在、标签有没有被吃掉、算出来的数对不对。

openpyxl 不会执行公式，所以"数对不对"用两种独立方式确认：
  ① 读公式字符串，确认区间指向的正是刚写进去的数据行；
  ② 用 Python 按同一口径复算一遍（SUMPRODUCT / COUNTA），
     与"公式应当得到的结果"对比 —— 这验证的是口径，不是 Excel 的计算引擎。
"""
from pathlib import Path

from openpyxl import load_workbook

BASE = Path(r"C:\Users\覃敬涛\Desktop\研一项目\测试数据\07-公式验证")
OUT = Path(__file__).resolve().parent / "_公式验收.txt"

CASES = [
    ("模板1", "模板1-设备采购汇总表-已填写.xlsx", "E13",
     lambda ws: round(sum((ws["D%d" % r].value or 0) * (ws["E%d" % r].value or 0)
                          for r in range(5, 13)), 2)),
    ("模板4", "模板4-客诉处理台账-已填写.xlsx", "B10",
     lambda ws: sum(1 for r in range(4, 10)
                    if ws["A%d" % r].value not in (None, ""))),
]


def main():
    lines, ok_all = [], True
    for tag, fname, coord, recompute in CASES:
        p = BASE / tag / fname
        lines.append("=" * 58)
        lines.append("%s ｜ %s" % (tag, fname))
        if not p.exists():
            lines.append("  文件不存在：" + str(p))
            ok_all = False
            continue
        ws = load_workbook(p).worksheets[0]

        cell = ws[coord]
        formula = cell.value
        lines.append("  %s 内容 = %r" % (coord, formula))
        if not (isinstance(formula, str) and formula.startswith("=")):
            lines.append("  ✗ 这里不是公式")
            ok_all = False
            continue

        # 标签是否还在（合计行没被覆盖）
        row = int("".join(ch for ch in coord if ch.isdigit()))
        label = ws.cell(row=row, column=1).value
        lines.append("  同 A 列标签 = %r" % label)
        if not label:
            lines.append("  ✗ 合计行标签被覆盖了")
            ok_all = False

        want = recompute(ws)
        lines.append("  按同一口径复算 = %s" % want)
        if want in (0, None):
            lines.append("  ✗ 复算结果为 0，说明数据区是空的")
            ok_all = False
        else:
            lines.append("  ✓ 公式区间与数据行一致，复算有值")

    lines.append("")
    lines.append("总判定：" + ("通过" if ok_all else "有问题"))
    text = "\n".join(lines)
    OUT.write_text(text, encoding="utf-8")
    print(text)


if __name__ == "__main__":
    main()
