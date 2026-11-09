# -*- coding: utf-8 -*-
"""两次串行之间差多少 —— 用来判断大池并发对比里的差异是不是模型随机性。"""
from pathlib import Path

from openpyxl import load_workbook

A = Path(r"C:\Users\覃敬涛\Desktop\研一项目\压测-大文档\结果\模板1\模板1-设备采购汇总表-已填写.xlsx")
B = Path(r"C:\Users\覃敬涛\Desktop\研一项目\测试数据\08-复测-大池串行\模板1\模板1-设备采购汇总表-已填写.xlsx")
OUT = Path(r"C:\Users\覃敬涛\Desktop\研一项目\收尾对照\评测\_大池串行两次对比.txt")


def norm(v):
    if v is None:
        return ""
    if isinstance(v, (int, float)):
        return "%.2f" % round(float(v), 2)
    s = str(v).strip().replace(" ", "")
    try:
        return "%.2f" % round(float(s), 2)
    except Exception:
        return s


def main():
    wa = load_workbook(A, data_only=True).worksheets[0]
    wb = load_workbook(B, data_only=True).worksheets[0]
    rows = max(wa.max_row, wb.max_row)
    cols = max(wa.max_column, wb.max_column)
    diff = []
    for r in range(1, rows + 1):
        for c in range(1, cols + 1):
            va, vb = norm(wa.cell(row=r, column=c).value), norm(wb.cell(row=r, column=c).value)
            if va != vb:
                diff.append(("第%d行第%d列" % (r, c), va, vb))
    lines = ["两次串行（同池同模板）对比", "第一次：%s" % A.parent.name,
             "第二次：%s" % B.parent.name,
             "对比 %d 格 ｜ 差异 %d 格" % (rows * cols, len(diff))]
    for coord, va, vb in diff[:12]:
        lines.append("  %s：第一次=%s ｜ 第二次=%s" % (coord, va, vb))
    text = "\n".join(lines)
    OUT.write_text(text, encoding="utf-8")
    print(text)


if __name__ == "__main__":
    main()
