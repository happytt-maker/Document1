# -*- coding: utf-8 -*-
"""真实数据诊断：叙述体文档 → 明细表。

不调模型，只看三件事：
  1. 池子结构（113 块都是什么形态）
  2. 用正则能机械地数出多少个城市（作为"应抽条数"的独立参照）
  3. 模板要的 5 列，在这份文档里的"字面/结构线索"有多弱（判断现有语义召回能不能用）
"""
import json, re, sys, os
from pathlib import Path

BASE = Path(r"C:\Users\覃敬涛\Desktop\研一项目")
sys.path.insert(0, str(BASE / "matcher"))
os.environ["A23_OUTPUT_JSON"] = str(BASE / "doc_reader" / "真实数据-pool.json")

OUT = []


def w(s=""):
    OUT.append(str(s))


import extract
import retrieve

pool = extract.load_pool()
w("【池子】%d 块" % len(pool))
from collections import Counter
kind = Counter(((b.get("locator") or {}).get("kind") or "?") for b in pool)
w("  形态分布：%s" % dict(kind))
lens = [len((b.get("text") or "")) for b in pool]
w("  块长度：最小 %d 最大 %d 中位 %d" % (min(lens), max(lens), sorted(lens)[len(lens) // 2]))
w("  带列头(_head)的块：%d 个（0 说明这是纯叙述文档，列头机制失效）"
  % sum(1 for b in pool if b.get("_head")))
w()

# ---------- 2. 正则独立计数 ----------
w("【正则独立计数】")
CITY_LINE = re.compile(r"^([\u4e00-\u9fa5]{2,6})\s*(?:以|GDP)")
gdp = re.compile(r"GDP\s*总量\s*(?:达|达到)?\s*([\d,]+(?:\.\d+)?)\s*亿元")
peo = re.compile(r"常住人口\s*(?:达|突破)?\s*([\d,]+(?:\.\d+)?)\s*万")
pcg = re.compile(r"人均\s*GDP\s*(?:高达|为|达)?\s*([\d,]+(?:\.\d+)?)\s*元")
rev = re.compile(r"一般公共预算收入\s*(?:达|达到|突破|跃升至|大幅增长至)?\s*([\d,]+(?:\.\d+)?)\s*亿元")
exp = re.compile(r"支出\s*(?:高达|达|高达达)?\s*([\d,]+(?:\.\d+)?)\s*亿元")

n_city = n4 = 0
miss = []
for b in pool:
    t = b.get("text") or ""
    m = CITY_LINE.match(t.strip())
    if not m:
        continue
    n_city += 1
    got = [bool(x.search(t)) for x in (gdp, peo, pcg, rev)]
    if all(got):
        n4 += 1
    else:
        miss.append((m.group(1), got, t[:60]))

w("  以「城市名…GDP」开头的段落：%d 个" % n_city)
w("  5 列里能正则抽出 4 个数的：%d 个" % n4)
w("  缺项段落：%d 个" % len(miss))
for nm, g, t in miss[:12]:
    w("     %-6s 命中=%s ｜ %s" % (nm, "".join("✓" if x else "✗" for x in g), t))
w()

# ---------- 3. 现有召回对模板 5 列的表现 ----------
w("【现有字面/语义召回的分数】（模板列 → 池内最高分 / 来源块是否有列头）")
cols = ["城市名", "GDP总量（亿元）", "常住人口（万）", "人均GDP（元）", "一般公共预算收入（亿元）"]
for c in cols:
    scored = retrieve.score_blocks(pool, retrieve.strip_unit(c))
    top = scored[:3]
    w("  %-22s 最高分 %3d" % (c, top[0][0] if top else 0))
    for s, b in top[:2]:
        w("       %.0f 分 → %s" % (s, (b.get("text") or "")[:56].replace("\n", " ")))
w()

# ---------- 4. 关键干扰项统计 ----------
w("【干扰项】")
w("  同段落里还有「支出 X 亿元」的：%d 段 ← 和「一般公共预算收入」单位相同，最容易串"
  % sum(1 for b in pool if exp.search(b.get("text") or "")))
w("  带「同比增长 X%%」的：%d 段 ← 百分比不能当 GDP"
  % sum(1 for b in pool if "同比" in (b.get("text") or "")))
w("  数字带千分位逗号的：%d 段"
  % sum(1 for b in pool if re.search(r"\d,\d{3}", b.get("text") or "")))
w("  带括号修正值的：%s"
  % [ (b.get("text") or "")[:20] for b in pool if "修正值" in (b.get("text") or "")])

Path(BASE / "测试数据" / "真实数据" / "_诊断.txt").write_text("\n".join(OUT), encoding="utf-8")
print("ok")
