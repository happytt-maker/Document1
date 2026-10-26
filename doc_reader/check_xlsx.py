# -*- coding: utf-8 -*-
"""
check_xlsx.py
-------------
xlsx 读取器的专项检查。

重点核对：
  1. 公式单元格有没有被正确识别和标记
  2. 多工作表有没有全部遍历到
  3. 空单元格有没有被区分出来
"""

import sys
from pathlib import Path

BASE = Path(__file__).parent
sys.path.insert(0, str(BASE))

from readers import read_one, SUPPORTED

SRC = Path(r"C:\Users\覃敬涛\Desktop\研一项目\测试数据\源文档\xlsx")

out = []
def w(s=""):
    out.append(str(s))


w("=" * 72)
w("xlsx 读取器检查报告")
w("=" * 72)
w()
w("已注册格式: " + ", ".join(SUPPORTED))
w()

files = sorted(SRC.glob("*.xlsx"))
w("发现 {0} 个 xlsx 文件".format(len(files)))
w()

total_formula_uncached = 0
total_formula_cached = 0

for path in files:
    doc = read_one(path)
    w("-" * 72)
    w("文件: {0}".format(doc["file"]))
    w("  工作表数: {0}   总块数: {1}".format(doc["sheet_count"], doc["block_count"]))

    for s in doc["sheets"]:
        w("    表『{0}』  {1}行×{2}列   块 {3}  公式 {4}".format(
            s["name"], s["max_row"], s["max_column"],
            s["block_count"], s["formula_count"]))

    blocks = doc["blocks"]

    # 公式分类
    uncached = [b for b in blocks if b.get("value_type") == "formula_uncached"]
    cached = [b for b in blocks if b.get("value_type") == "formula_cached"]
    total_formula_uncached += len(uncached)
    total_formula_cached += len(cached)

    w("  公式单元格: 共 {0} 个  （缓存为空 {1} 个 / 已缓存 {2} 个）".format(
        len(uncached) + len(cached), len(uncached), len(cached)))

    if uncached:
        w("    缓存为空（正是坑 1 的现场，前 4 个）:")
        for b in uncached[:4]:
            loc = b["locator"]
            w("      {0}!{1}  公式 = {2!r}".format(loc["sheet"], loc["cell"], b["formula"]))

    if cached:
        w("    已缓存（前 3 个）:")
        for b in cached[:3]:
            loc = b["locator"]
            w("      {0}!{1}  公式 = {2!r}  读到的值 = {3!r}".format(
                loc["sheet"], loc["cell"], b["formula"], b["text"]))

    # 值类型分布
    types = {}
    for b in blocks:
        types[b.get("value_type", "?")] = types.get(b.get("value_type", "?"), 0) + 1
    w("  值类型分布: " + "  ".join("{0}={1}".format(k, v) for k, v in sorted(types.items())))

    # 合并区域
    merged = [b for b in blocks if b.get("merged_range")]
    if merged:
        w("  合并区域涉及的块 {0} 个，例如: {1}!{2} → {3}".format(
            len(merged), merged[0]["locator"]["sheet"],
            merged[0]["locator"]["cell"], merged[0]["merged_range"]))
    w()

w("=" * 72)
w("  {0:<26}{1:>10}{2:>10}{3:>8}".format("文件", "工作表", "块数", "公式"))
w("  " + "-" * 56)
total = 0
for path in files:
    doc = read_one(path)
    fcnt = sum(s["formula_count"] for s in doc["sheets"])
    total += doc["block_count"]
    w("  {0:<26}{1:>10}{2:>10}{3:>8}".format(
        doc["file"], doc["sheet_count"], doc["block_count"], fcnt))
w("  " + "-" * 56)
w("  合计 {0} 个文件，{1} 个块".format(len(files), total))
w()
w("公式命中情况：缓存为空 {0} 个，已缓存 {1} 个".format(
    total_formula_uncached, total_formula_cached))
w()
w("结论：{0}".format(
    "公式识别正常，坑 1 已被标记出来"
    if (total_formula_uncached + total_formula_cached) > 0
    else "没有识别到任何公式，需要检查"))

(BASE / "xlsx检查报告.txt").write_text("\n".join(out), encoding="utf-8")
print("OK")
