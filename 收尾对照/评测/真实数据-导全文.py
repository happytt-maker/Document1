# -*- coding: utf-8 -*-
"""把源文档全文导出来（含文本框/页眉），并看看 449KB 里到底装了什么。"""
import zipfile, re
from pathlib import Path
from docx import Document

BASE = Path(r"C:\Users\覃敬涛\Desktop\研一项目\测试数据\真实数据")
src = BASE / "02-源文档" / "2025年中国城市经济百强全景报告.docx"
OUT = []


def w(s=""):
    OUT.append(str(s))

# 1) zip 里有什么，谁占体积
w("【压缩包内容】")
with zipfile.ZipFile(src) as z:
    items = sorted(z.infolist(), key=lambda i: -i.file_size)
    for i in items[:12]:
        w("  %-45s %8.1f KB" % (i.filename, i.file_size / 1024))
    w("  条目总数 %d ｜ 解压后总大小 %.1f KB" % (len(items), sum(i.file_size for i in items) / 1024))
w()

# 2) 正文全文
doc = Document(str(src))
w("【正文全文】共 %d 段" % len(doc.paragraphs))
for i, pg in enumerate(doc.paragraphs):
    t = pg.text.strip()
    if t:
        w("[%03d] %s" % (i, t))
w()

# 3) 文本框 / 图形里的文字（python-docx 的 paragraphs 读不到这些）
w("【文本框等非正文文字】")
with zipfile.ZipFile(src) as z:
    xml = z.read("word/document.xml").decode("utf-8", "replace")
w("  document.xml 长度 %d 字符" % len(xml))
w("  <w:txbxContent> 出现 %d 次" % xml.count("txbxContent"))
w("  <w:tbl> 出现 %d 次" % xml.count("<w:tbl>"))
w("  <w:drawing> 出现 %d 次" % xml.count("<w:drawing>"))
w("  <w:pict> 出现 %d 次" % xml.count("<w:pict>"))

(BASE / "_全文.txt").write_text("\n".join(OUT), encoding="utf-8")
print("ok")
