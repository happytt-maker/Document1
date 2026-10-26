# -*- coding: utf-8 -*-
"""
check_docx.py
-------------
docx 读取器的专项检查。

不依赖 main.py 的终端输出，自己直接读文件、检查关键指标，
结果写进「docx检查报告.txt」。
"""

import json
import sys
from pathlib import Path

BASE = Path(__file__).parent
sys.path.insert(0, str(BASE))

from readers import read_one, SUPPORTED

SRC = Path(r"C:\Users\覃敬涛\Desktop\研一项目\测试数据\源文档\docx")

out = []
def w(s=""):
    out.append(str(s))


w("=" * 70)
w("docx 读取器检查报告")
w("=" * 70)
w()
w("已注册格式: " + ", ".join(SUPPORTED))
w()

docx_files = sorted(SRC.glob("*.docx"))
w("发现 {0} 个 docx 文件".format(len(docx_files)))
w()

for path in docx_files:
    doc = read_one(path)
    w("-" * 70)
    w("文件: {0}".format(doc["file"]))
    w("  段落数(原始): {0}".format(doc["paragraph_count"]))
    w("  表格数: {0}".format(doc["table_count"]))
    w("  内容块数: {0}".format(doc["block_count"]))

    para_blocks = [b for b in doc["blocks"] if b["locator"]["kind"] == "paragraph"]
    cell_blocks = [b for b in doc["blocks"] if b["locator"]["kind"] == "table_cell"]
    w("    - 段落块: {0}".format(len(para_blocks)))
    w("    - 表格单元格块: {0}".format(len(cell_blocks)))

    # 各表格的唯一单元格数
    for t_idx in range(1, doc["table_count"] + 1):
        cells = [b for b in cell_blocks if b["locator"]["table_index"] == t_idx]
        coords = set((b["locator"]["row"], b["locator"]["col"]) for b in cells)
        w("    表{0}: 去重后单元格 {1} 个（坐标对 {2} 个）".format(
            t_idx, len(cells), len(coords)))

    # 标题检测
    headings = [b for b in para_blocks if b.get("is_heading")]
    if headings:
        w("  识别到的标题（前 6 个）:")
        for h in headings[:6]:
            w("    第{0:>3}段  {1}".format(h["locator"]["para_index"], h["text"][:44]))

    # 段落序号是否连续
    idxs = [b["locator"]["para_index"] for b in para_blocks]
    if idxs:
        w("  段落序号范围: {0} — {1}（共 {2} 个非空段落）".format(
            min(idxs), max(idxs), len(idxs)))

    # 合并单元格检查
    merged = [b for b in cell_blocks if b.get("grid_span") or b.get("v_merge")]
    if merged:
        w("  合并单元格标记 {0} 个:".format(len(merged)))
        for m in merged[:4]:
            loc = m["locator"]
            w("    表{0} 行{1} 列{2}: {3!r}  grid_span={4} v_merge={5}".format(
                loc["table_index"], loc["row"], loc["col"],
                m["text"][:24], m.get("grid_span"), m.get("v_merge")))

    # 叠加内容检查
    overlaid = [b for b in cell_blocks if b.get("raw_text")]
    if overlaid:
        w("  有内容叠加的单元格 {0} 个:".format(len(overlaid)))
        for o in overlaid[:4]:
            w("    {0!r}  ->  清理为  {1!r}".format(o["raw_text"][:26], o["text"][:26]))
    w()

# 汇总
w("=" * 70)
total_blocks = 0
w("  {0:<34}{1:>8}{2:>8}{3:>8}".format("文件", "段落", "表格", "块数"))
w("  " + "-" * 60)
for path in docx_files:
    doc = read_one(path)
    total_blocks += doc["block_count"]
    w("  {0:<34}{1:>8}{2:>8}{3:>8}".format(
        doc["file"], doc["paragraph_count"], doc["table_count"], doc["block_count"]))
w("  " + "-" * 60)
w("  合计 {0} 个文件，{1} 个内容块".format(len(docx_files), total_blocks))
w()

(BASE / "docx检查报告.txt").write_text("\n".join(out), encoding="utf-8")
print("OK")
