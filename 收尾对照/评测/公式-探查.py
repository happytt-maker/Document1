# -*- coding: utf-8 -*-
"""扫模板里的"合计/总计/小计"行，看它们长什么样、现在有没有公式。

为什么要这一步：fill.py 目前只写值，完全不懂公式。模板里的合计行
如果被当成"待填数据行"，要么抽不到值空着，要么被填进一条无关记录。
先把现状摸清楚，再决定在哪一层补 SUM。
"""
from pathlib import Path

from openpyxl import load_workbook

BASE = Path(r"C:\Users\覃敬涛\Desktop\研一项目")
DIRS = [BASE / "测试数据" / "模板表格",
        BASE / "测试数据" / "模板-复杂表头",
        BASE / "测试数据" / "真实数据" / "01-模板"]
OUT = Path(__file__).resolve().parent / "_公式探查.txt"

KEYS = ("合计", "总计", "小计", "总额", "合计：", "总 计")


def scan_xlsx(p):
    lines = []
    for ws in load_workbook(p).worksheets:
        for row in ws.iter_rows():
            for cell in row:
                v = cell.value
                if not isinstance(v, str):
                    continue
                s = v.strip()
                if not any(k in s for k in KEYS) or len(s) > 12:
                    continue
                r = cell.row
                # 这一行有值/有公式的格子，逐列列出来
                cells = []
                for c in range(1, ws.max_column + 1):
                    cc = ws.cell(row=r, column=c)
                    if cc.value is None:
                        continue
                    val = cc.value
                    if isinstance(val, str) and len(val) > 24:
                        val = val[:24] + "…"
                    kind = "公式" if (isinstance(val, str) and val.startswith("=")) else "值"
                    cells.append("%s%d[%s]=%s" % (cc.column_letter, r, kind, val))
                # 下一行是不是数据行（判断合计在数据上方还是下方）
                lines.append("  %s「%s」第 %d 行：%s"
                             % (ws.title, s, r, " ｜ ".join(cells)))
    return lines


def scan_docx(p):
    from docx import Document
    out = []
    doc = Document(str(p))
    for ti, t in enumerate(doc.tables, start=1):
        for ri, row in enumerate(t.rows, start=1):
            texts = [c.text.strip() for c in row.cells]
            if any(any(k in x for k in KEYS) and len(x) <= 12 for x in texts):
                out.append("  表格%d 第 %d 行：%s" % (ti, ri, " ｜ ".join(texts)))
    return out


def main():
    lines = []
    for d in DIRS:
        if not d.exists():
            continue
        for p in sorted(d.glob("*.xlsx")):
            hits = scan_xlsx(p)
            lines.append("【%s】 %s" % (p.name, "%d 处" % len(hits) if hits else "无"))
            lines += hits
            lines.append("")
        for p in sorted(d.glob("*.docx")):
            hits = scan_docx(p)
            lines.append("【%s】 %s" % (p.name, "%d 处" % len(hits) if hits else "无"))
            lines += hits
            lines.append("")
    text = "\n".join(lines)
    OUT.write_text(text, encoding="utf-8")
    print(text)


if __name__ == "__main__":
    main()
