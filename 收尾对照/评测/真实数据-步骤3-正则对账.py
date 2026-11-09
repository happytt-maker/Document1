# -*- coding: utf-8 -*-
"""真实数据试跑 · 步骤3：正则 vs 模型 全量对账（不依赖人工标准答案）。

为什么要这一层：
  步骤2 的人均反算只能证明"GDP/人口/人均三者自洽"，如果 GDP 和人口同时
  错到另一行去，反算照样通过。所以再开一条**完全独立**的通道 ——
  用正则按"原文句式"重新抽一遍，两边逐格比对：
     · 一致 → 强证据（两条互不相干的路径得出同一个答案）
     · 不一致 → 逐条打出来，人工只看差异项（100 行里通常只有几行）

正则通道的局限（必须写清楚，否则会误判模型）：
  开头几个头部城市（上海/北京/深圳/重庆/广州）的句式是变体
  （"上海以 56,708.71 亿元的 GDP 总量稳居第一"、"重庆凭借 3,191.43 万的
  庞大人口基数，拿下 33,757.93 亿元 GDP"），正则会漏或抽错。
  这类行标成「正则未命中」，不算模型错。
"""
import json
import os
import re
from pathlib import Path

MATCHER = Path(os.environ.get("A23_MATCHER_DIR")
               or r"C:\Users\覃敬涛\Desktop\研一项目\测试数据\真实数据\04-试跑产物")
POOL = Path(r"C:\Users\覃敬涛\Desktop\研一项目\doc_reader\真实数据-pool.json")
COLS = ["城市名", "GDP总量（亿元）", "常住人口（万）", "人均GDP（元）",
        "一般公共预算收入（亿元）"]
OUT = Path(__file__).with_name("_对账.txt")


def n(v):
    if v is None:
        return None
    try:
        return round(float(re.sub(r"[,，\s]", "", str(v))), 4)
    except Exception:
        return None


# ---- 正则通道：按原文句式抽（与模型完全无关联） ----
R_GDP = [r"GDP\s*总量\s*([\d,]+\.?\d*)\s*亿元",
         r"([\d,]+\.?\d*)\s*亿元的\s*GDP\s*总量",
         r"拿下\s*([\d,]+\.?\d*)\s*亿元\s*GDP"]
R_POP = [r"常住人口\s*(?:达|达到|为)?\s*([\d,]+\.?\d*)\s*万",
         r"([\d,]+\.?\d*)\s*万的(?:庞大)?人口",
         r"人口严控至\s*([\d,]+\.?\d*)\s*万",
         r"([\d,]+\.?\d*)\s*万常住人口"]
R_PC = [r"人均\s*GDP\s*(?:为|高达|达)?\s*([\d,]+\.?\d*)\s*元",
        r"([\d,]+\.?\d*)\s*元的(?:惊人)?人均\s*GDP"]
R_REV = [r"一般公共预算收入\s*(?:突破|跃升至|大幅增长至|为|达)?\s*([\d,]+\.?\d*)\s*亿元"]


def grab(pats, text):
    for p in pats:
        m = re.search(p, text)
        if m:
            return m.group(1)
    return None


def para_of(block):
    return ((block.get("locator") or {}).get("para_index"))


def main():
    pool = json.loads(POOL.read_text(encoding="utf-8"))
    # 池子结构：{"documents": [{"file":..., "blocks":[...]}, ...]}
    blocks = [b for d in pool.get("documents", []) for b in d.get("blocks", [])]

    model = {}
    for c in COLS:
        d = json.loads((MATCHER / ("抽取结果-%s.json" % c)).read_text(encoding="utf-8"))
        model[c] = {it.get("loc_text"): it.get("value") for it in d["items"]}

    # 原文段落表：段号 → 文本
    raw = {}
    for b in blocks:
        pi = para_of(b)
        if pi:
            raw["第 %d 段" % pi] = b.get("text") or ""

    L = []
    L.append("【对账范围】模型 %d 行 × %d 列 ｜ 原文 %d 段" % (len(model[COLS[0]]), len(COLS), len(raw)))
    L.append("")

    locks = list(model[COLS[0]].keys())       # 模型给出的段号，按输出顺序
    diff, rex_miss, ok = [], [], 0

    for loc in locks:
        text = raw.get(loc, "")
        city_model = model["城市名"].get(loc)
        # 城市名正则：段首到「GDP/以/凭/在」之前的中文专名
        m = re.match(r"^([\u4e00-\u9fa5]{2,6}?)(?:以|凭借|在|GDP|作为|紧随)", text)
        city_rex = m.group(1) if m else None
        if city_rex and city_model and city_rex != city_model:
            diff.append((loc, "城市名", city_model, city_rex))
        elif city_model and not city_rex:
            rex_miss.append((loc, "城市名"))
        elif city_model == city_rex:
            ok += 1

        for col, pats in ((("GDP总量（亿元）"), R_GDP),
                          (("常住人口（万）"), R_POP),
                          (("人均GDP（元）"), R_PC),
                          (("一般公共预算收入（亿元）"), R_REV)):
            mv = n(model[col].get(loc))
            rv = n(grab(pats, text))
            if rv is None:
                rex_miss.append((loc, col))
            elif mv is None:
                diff.append((loc, col, None, rv))
            elif abs(mv - rv) < 0.01:
                ok += 1
            else:
                diff.append((loc, col, mv, rv))

    total = len(locks) * len(COLS)
    L.append("【结果】%d 格对比：一致 %d ｜ 不一致 %d ｜ 正则未命中 %d"
             % (total, ok, len(diff), len(rex_miss)))
    L.append("  一致率（命中部分）= %.1f%%"
             % (ok / max(1, ok + len(diff)) * 100))
    L.append("")

    if diff:
        L.append("【不一致明细】（每行：段号 ｜ 列 ｜ 模型 ｜ 正则）")
        for loc, col, a, b in diff:
            L.append("  %-10s %-22s 模型=%s  正则=%s" % (loc, col, a, b))
            L.append("      原文：%s" % raw.get(loc, "")[:120])
    else:
        L.append("【不一致明细】无 —— 两条独立通道结果完全相同")
    L.append("")

    if rex_miss:
        L.append("【正则未命中】（不算错：原文句式是变体，正则表达不出来）")
        from collections import Counter
        c = Counter(col for _l, col in rex_miss)
        for k, v in c.most_common():
            L.append("  %-22s %d 格" % (k, v))
        for loc, col in rex_miss[:12]:
            L.append("     %s ｜ %s" % (loc, col))

    txt = "\n".join(L)
    OUT.write_text(txt, encoding="utf-8")
    print(txt)


if __name__ == "__main__":
    main()
