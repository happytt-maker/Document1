# -*- coding: utf-8 -*-
"""真实数据试跑 · 步骤2：不依赖标准答案的自动交叉校验。

思路：**用数字之间的关系自证**，不需要人工先填答案。
   人均GDP（元） ≈ GDP总量（亿元） ÷ 常住人口（万） × 10000
这一条能把三类错一次照出来：
   ① 整行错位（某行的 GDP 配了邻行的人口）→ 比值必然离谱
   ② 串列（收入填成支出、增长率填进 GDP）→ 量级差 3~10 倍
   ③ 抄错小数位 → 比值差 10 倍

另加三类结构检查：
   ④ 同一列出现重复值（叙述体里每个城市只出现一次，重复=抄串）
   ⑤ GDP 是否按原文排名单调递减（原文是"第一/第二/第三…"排下来的）
   ⑥ 每列条数是否一致（不一致说明有段落漏抽）
"""
import json
import os
import re
from pathlib import Path

# 产物目录：真实数据试跑的抽取结果不落在 matcher/（那里是 5 张正式模板的产物，
# 且五表连测的"还原"会按固定清单把它清掉），统一放 04-试跑产物。
MATCHER = Path(os.environ.get("A23_MATCHER_DIR")
               or r"C:\Users\覃敬涛\Desktop\研一项目\测试数据\真实数据\04-试跑产物")
COLS = ["城市名", "GDP总量（亿元）", "常住人口（万）", "人均GDP（元）",
        "一般公共预算收入（亿元）"]
OUT = Path(__file__).with_name("_校验.txt")


def num(v):
    try:
        return float(re.sub(r"[,，\s]", "", str(v)))
    except Exception:
        return None


def load(col):
    p = MATCHER / ("抽取结果-%s.json" % col)
    if not p.exists():
        return None
    return json.loads(p.read_text(encoding="utf-8"))


def main():
    L = []
    data = {}
    for c in COLS:
        d = load(c)
        if d is None:
            L.append("缺少结果：%s" % c)
            return
        data[c] = d["items"]

    L.append("【一、条数】")
    for c in COLS:
        L.append("  %-24s %d 条" % (c, len(data[c])))
    n = min(len(data[c]) for c in COLS)
    L.append("  各行最小条数 = %d" % n)
    if len({len(data[c]) for c in COLS}) > 1:
        L.append("  ⚠ 各列条数不一致 —— 有段落被漏抽")

    # 按出现顺序对齐成行（整表抽取本来就是一段一行，顺序即对齐）
    rows = []
    for i in range(n):
        r = {}
        for c in COLS:
            it = data[c][i]
            r[c] = it.get("value")
            r["_loc"] = it.get("loc_text")
            r["_ev"] = it.get("evidence")
        rows.append(r)

    L.append("")
    L.append("【二、人均GDP 反算校验】人均 = GDP(亿) / 人口(万) × 10000 元")
    bad = []
    for i, r in enumerate(rows, 1):
        g, p, h = num(r["GDP总量（亿元）"]), num(r["常住人口（万）"]), num(r["人均GDP（元）"])
        if not (g and p and h):
            bad.append((i, r, None, "有值不是数字"))
            continue
        calc = g / p * 10000
        err = abs(calc - h) / h * 100
        if err > 1.0:
            bad.append((i, r, calc, "偏差 %.1f%%" % err))
    L.append("  通过 %d / %d 行" % (n - len(bad), n))
    for i, r, calc, why in bad:
        L.append("  ✗ 第%d行 %s ｜GDP=%s 人口=%s 人均=%s ｜反算=%s ｜%s"
                 % (i, r["城市名"], r["GDP总量（亿元）"], r["常住人口（万）"],
                    r["人均GDP（元）"], ("%.0f" % calc) if calc else "-", why))
        L.append("      原文：%s" % (r["_ev"] or ""))

    L.append("")
    L.append("【三、重复值检查】叙述体里每个城市/每个数只该出现一次")
    for c in COLS:
        vals = [str(it.get("value")) for it in data[c]]
        dup = {v: vals.count(v) for v in set(vals) if vals.count(v) > 1}
        L.append("  %-24s %s" % (c, ("重复：" + str(dup)) if dup else "无重复"))

    L.append("")
    L.append("【四、GDP 排名单调性】原文按名次叙述，GDP 应递减")
    drops = 0
    for i in range(1, n):
        a, b = num(rows[i - 1]["GDP总量（亿元）"]), num(rows[i]["GDP总量（亿元）"])
        if a and b and b > a:
            drops += 1
            L.append("  ↑ 第%d行 %s(%s) > 上一行 %s(%s)"
                     % (i + 1, rows[i]["城市名"], b, rows[i - 1]["城市名"], a))
    if not drops:
        L.append("  全表严格递减 ✓")

    L.append("")
    L.append("【五、前 15 行对照表（人工核验用）】")
    L.append("  %-3s %-8s %12s %10s %10s %12s" % ("#", "城市", "GDP(亿)", "人口(万)", "人均(元)", "预算收入(亿)"))
    for i, r in enumerate(rows[:15], 1):
        L.append("  %-3d %-8s %12s %10s %10s %12s" % (
            i, r["城市名"], r["GDP总量（亿元）"], r["常住人口（万）"],
            r["人均GDP（元）"], r["一般公共预算收入（亿元）"]))

    L.append("")
    L.append("【六、末 5 行】")
    for i, r in enumerate(rows[-5:], n - 4):
        L.append("  %-3d %-8s %12s %10s %10s %12s" % (
            i, r["城市名"], r["GDP总量（亿元）"], r["常住人口（万）"],
            r["人均GDP（元）"], r["一般公共预算收入（亿元）"]))

    L.append("")
    L.append("【七、耗时】")
    for c in COLS:
        d = load(c)
        L.append("  %-24s %s 秒 ｜ 候选 %s 块%s" % (
            c, d.get("elapsed_sec"), d.get("candidate_count"),
            "（整表一次，耗时共用）" if d.get("shared_cost") else ""))

    txt = "\n".join(L)
    OUT.write_text(txt, encoding="utf-8")
    print(txt)


if __name__ == "__main__":
    main()
