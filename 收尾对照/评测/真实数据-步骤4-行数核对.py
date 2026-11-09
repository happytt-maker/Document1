# -*- coding: utf-8 -*-
"""真实数据试跑 · 步骤4：城市名与行数核对。

两问：
  ① 原文到底描述了多少个城市？模型抽了 100 行，多/少的那几行是什么？
  ② 抽出来的 100 个名字里，有没有混进"段落标题/综述句"这类不是城市的东西？

判据（不靠人工）：
  · 真城市段 = 段落里同时出现 GDP 数值和"亿元"
  · 非记录段 = 章节标题（一、二、三…）、总述（"2025 年，中国城市经济…"）、
    过渡句（"第二梯队的竞争…"）
"""
import json
import os
import re
from pathlib import Path

POOL = Path(r"C:\Users\覃敬涛\Desktop\研一项目\doc_reader\真实数据-pool.json")
MATCHER = Path(os.environ.get("A23_MATCHER_DIR")
               or r"C:\Users\覃敬涛\Desktop\研一项目\测试数据\真实数据\04-试跑产物")
CITY = MATCHER / "抽取结果-城市名.json"
OUT = Path(__file__).with_name("_行数核对.txt")


def main():
    pool = json.loads(POOL.read_text(encoding="utf-8"))
    blocks = [b for d in pool.get("documents", []) for b in d.get("blocks", [])]
    city_items = json.loads(CITY.read_text(encoding="utf-8"))["items"]

    L = []
    # ---- 原文侧：哪些段是真记录 ----
    real, other = [], []
    for b in blocks:
        t = b.get("text") or ""
        pi = (b.get("locator") or {}).get("para_index")
        has_money = bool(re.search(r"[\d,]+\.\d+|[\d,]{3,}", t)) and "亿元" in t
        is_head = bool(re.match(r"^[一二三四五六七八九十]+、", t))
        if has_money and not is_head:
            real.append((pi, t))
        else:
            other.append((pi, t))

    L.append("【原文侧】共 %d 段：像记录的 %d 段 ｜ 不像记录的 %d 段"
             % (len(blocks), len(real), len(other)))
    L.append("")
    L.append("  不像记录的段落（应当**不**出现在表里）：")
    for pi, t in other:
        L.append("    第 %-4s 段 ｜ %s" % (pi, t[:60]))

    # ---- 模型侧 ----
    L.append("")
    L.append("【模型侧】%d 行" % len(city_items))
    names = [it.get("value") for it in city_items]
    bad = []
    for it in city_items:
        v = str(it.get("value") or "")
        if len(v) > 8 or re.search(r"[0-9，。、；：]", v):
            bad.append(it)
    L.append("  名字里带数字/标点或超过 8 字的：%s" % (bad if bad else "无"))
    L.append("")
    L.append("  全部 %d 个名字：" % len(names))
    for i in range(0, len(names), 10):
        L.append("    %s" % "、".join(str(x) for x in names[i:i + 10]))

    # ---- 交叉：模型行 vs 原文记录段 ----
    L.append("")
    L.append("【交叉核对】")
    real_idx = {pi for pi, _t in real}
    used = {(it.get("loc_text") or "").replace("第 ", "").replace(" 段", "") for it in city_items}
    used = {int(u) for u in used if u.isdigit()}
    L.append("  模型用到的段落号：%d 个 ｜ 原文像记录的段落：%d 个" % (len(used), len(real_idx)))
    miss = sorted(real_idx - used)
    extra = sorted(used - real_idx)
    L.append("  原文是记录但模型没取：%s" % (miss if miss else "无"))
    L.append("  模型取了但原文不像记录：%s" % (extra if extra else "无"))
    for pi in extra:
        t = next((t for p, t in other if p == pi), "")
        L.append("      第 %s 段原文：%s" % (pi, t[:80]))

    # 每行是否只有一个值命中（一条记录里出现两次同值会让"是否唯一"存疑）
    L.append("")
    L.append("【每段的城市名是否真在段首】")
    bad2 = 0
    by_loc = {(it.get("loc_text") or ""): it.get("value") for it in city_items}
    for pi, t in real:
        v = by_loc.get("第 %d 段" % pi)
        if not v:
            continue
        if not t.startswith(v):
            bad2 += 1
            L.append("  第 %d 段：名字=%s ｜段首=%s" % (pi, v, t[:30]))
    L.append("  段首不匹配：%d 行（原文句式变体导致，需要人看一眼）" % bad2)

    txt = "\n".join(L)
    OUT.write_text(txt, encoding="utf-8")
    print(txt)


if __name__ == "__main__":
    main()
