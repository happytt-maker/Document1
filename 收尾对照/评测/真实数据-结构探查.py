# -*- coding: utf-8 -*-
"""摸清真实数据的结构：模板长什么样、源文档有多大、人工样例是不是空的。"""
import sys, os, json
from pathlib import Path

BASE = Path(r"C:\Users\覃敬涛\Desktop\研一项目\测试数据\真实数据")
OUT = []


def w(s=""):
    OUT.append(str(s))


def dump_xlsx(p, tag):
    import openpyxl
    wb = openpyxl.load_workbook(p, data_only=False)
    w("【%s】%s" % (tag, p.name))
    w("  工作表：%s" % wb.sheetnames)
    for ws in wb.worksheets:
        w("  -- %s ｜ 尺寸 %s ｜ 最大行 %d 最大列 %d" % (ws.title, ws.dimensions, ws.max_row, ws.max_column))
        w("     合并单元格 %d 处：%s" % (len(ws.merged_cells.ranges),
                                    ", ".join(str(r) for r in list(ws.merged_cells.ranges)[:12])))
        # 逐行打印前 15 行
        for r in range(1, min(ws.max_row, 15) + 1):
            cells = []
            for c in range(1, min(ws.max_column, 14) + 1):
                v = ws.cell(r, c).value
                if v is None:
                    continue
                cells.append("%s=%r" % (ws.cell(r, c).coordinate, v))
            if cells:
                w("     R%-3d %s" % (r, " | ".join(cells)))
        # 公式
        fml = [(c.coordinate, c.value) for row in ws.iter_rows() for c in row
               if isinstance(c.value, str) and c.value.startswith("=")]
        w("     公式格 %d 个%s" % (len(fml), ("：" + ", ".join("%s%s" % (a, b) for a, b in fml[:8])) if fml else ""))
        # 数据校验/下拉
        try:
            dvs = ws.data_validations.dataValidation
            if dvs:
                w("     数据校验 %d 条：%s" % (len(dvs), [d.sqref for d in dvs][:6]))
        except Exception:
            pass
        # 列宽/换行（判断是否是刻意排版过的）
    w()


def dump_docx(p):
    from docx import Document
    doc = Document(str(p))
    w("【源文档】%s ｜ %.1f KB" % (p.name, p.stat().st_size / 1024))
    w("  段落 %d 个 ｜ 表格 %d 个 ｜ 章节 %d 个" % (len(doc.paragraphs), len(doc.tables), len(doc.sections)))
    # 标题层级
    heads = [(pg.style.name, pg.text.strip()) for pg in doc.paragraphs
             if pg.style.name.startswith("Heading") and pg.text.strip()]
    w("  标题 %d 个，前 20 个：" % len(heads))
    for s, t in heads[:20]:
        w("     [%s] %s" % (s, t[:50]))
    # 表格概况
    for i, t in enumerate(doc.tables[:6], 1):
        w("  表%d：%d 行 × %d 列" % (i, len(t.rows), len(t.columns)))
        for r in t.rows[:3]:
            w("     | " + " | ".join((c.text or "").strip().replace("\n", "/")[:12] for c in r.cells[:8]))
    # 正文取样
    body = [pg.text.strip() for pg in doc.paragraphs if pg.text.strip()]
    w("  非空段落 %d 个，字符总数 %d" % (len(body), sum(len(x) for x in body)))
    w("  开头 3 段：")
    for x in body[:3]:
        w("     " + x[:90])
    # 值探测：数字密度
    import re
    nums = re.findall(r"\d{4,}", " ".join(body))
    w("  4 位以上数字出现 %d 次（其中前 12 个：%s）" % (len(nums), nums[:12]))
    # 图片
    try:
        w("  内嵌图片 %d 张" % len(doc.inline_shapes))
    except Exception:
        pass
    w()


tpl = BASE / "01-模板" / "2025年中国城市经济百强全景报告-模板.xlsx"
hum = BASE / "03-人工填好的样例" / "2025年中国城市经济百强全景报告-人工.xlsx"
src = BASE / "02-源文档" / "2025年中国城市经济百强全景报告.docx"

dump_xlsx(tpl, "模板")
w("=" * 70)
dump_xlsx(hum, "人工样例（用户说还没填）")
w("=" * 70)
dump_docx(src)

(BASE / "_结构探查.txt").write_text("\n".join(OUT), encoding="utf-8")
print("ok")
