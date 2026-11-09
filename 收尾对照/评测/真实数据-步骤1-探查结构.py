# -*- coding: utf-8 -*-
"""真实数据试跑 · 步骤1：确认模板槽位情况 / 抽取结果现状"""
import json, sys
from pathlib import Path

BASE = Path(__file__).resolve().parents[2]
TPL = "2025年中国城市经济百强全景报告-模板.xlsx"
out = []

sj = BASE / "template_reader" / "模板表结构.json"
data = json.loads(sj.read_text(encoding="utf-8"))
out.append("【模板表结构.json】共 %d 个模板" % len(data))
hit = None
for k, v in (data.items() if isinstance(data, dict) else []):
    if TPL in k or k == TPL:
        hit = (k, v)
if hit:
    k, info = hit
    out.append("  命中：%s" % k)
    out.append("  工作表：%s" % info.get("name"))
    out.append("  表头行：%s" % info.get("header_rows") or info.get("header_row"))
    out.append("  列：%s" % [(c.get("col"), c.get("title")) for c in info.get("columns", [])])
    out.append("  segments：%s" % json.dumps(info.get("segments"), ensure_ascii=False))
    out.append("  fill_region：%s" % json.dumps(info.get("fill_region"), ensure_ascii=False))
    segs = info.get("segments") or []
    n = 0
    for s in segs:
        n += s["end_row"] - s["start_row"] + 1
    out.append("  >>> 槽位总数 = %d" % n)
else:
    out.append("  没找到该模板，已有键：")
    for k in (data.keys() if isinstance(data, dict) else range(len(data))):
        out.append("    %s" % k)

out.append("")
out.append("【已产出的抽取结果】")
for f in sorted((BASE / "matcher").glob("抽取结果-*.json")):
    try:
        d = json.loads(f.read_text(encoding="utf-8"))
        out.append("  %s ｜ items=%d" % (f.name, len(d.get("items", []))))
    except Exception as e:
        out.append("  %s ｜ 读取失败 %s" % (f.name, e))

out.append("")
out.append("【池子】")
for f in sorted((BASE / "doc_reader").glob("*.json")):
    if "pool" in f.name:
        out.append("  %s ｜ %d KB" % (f.name, f.stat().st_size // 1024))

Path(__file__).with_name("_步骤1.txt").write_text("\n".join(out), encoding="utf-8")
print("\n".join(out))
