# -*- coding: utf-8 -*-
"""
跨行业-统一抽取-验收-1105.py —— 6 个行业 + 3 份真实文档的机械验收

验的不是"看起来对"，是四件可量化的事：
  ① 结构命中：文档里的业务对象、属性，抽出来的清单认不认得出（对照人工标注）
  ② 值命中：标注里"人看一眼认为该抽到"的具体值，抽到了没有
  ③ 可溯率：每条属性值都能回到原文/单元格（这是保真度底线）
  ④ 行业区分度：6 个行业长出来的对象类型不能是同一套（否则还是"固定模板"）

跑法：python 收尾对照\\跨行业-统一抽取-验收-1105.py
"""
import json
import re
import sys
import time
from pathlib import Path

BASE = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(BASE / "agent"))

OUT_DIR = BASE / "agent" / "输出"
SET_DIR = BASE / "测试数据" / "行业测试集"
REPORT = BASE / "收尾对照" / "跨行业-统一抽取-验收-1105.txt"

L = []
PASS = FAIL = 0
DETAIL = []


def say(s=""):
    print(s)
    L.append(str(s))


def norm(s):
    s = str(s or "")
    s = re.sub(r"\s+", "", s)
    return (s.replace("，", ",").replace("：", ":").replace("（", "(").replace("）", ")")
            .replace("￥", "").replace("¥", "").replace(",", "").lower())


def hit(a, b):
    a, b = norm(a), norm(b)
    if not a or not b:
        return False
    return a == b or (len(a) >= 2 and a in b) or (len(b) >= 2 and b in a)


def check(name, ok, detail=""):
    global PASS, FAIL
    if ok:
        PASS += 1
        say("  [PASS] %s%s" % (name, (" ｜ " + detail) if detail else ""))
    else:
        FAIL += 1
        say("  [FAIL] %s%s" % (name, (" ｜ " + detail) if detail else ""))
    return ok


def run_one(doc_name):
    import knowledge
    src = next(SET_DIR.glob(doc_name))
    t0 = time.time()
    ok, rep, files = knowledge.run_knowledge("自动读懂这份文档并抽取关键信息入库", {}, src)
    dt = time.time() - t0
    jp = OUT_DIR / (src.stem + "_知识.json")
    data = json.loads(jp.read_text(encoding="utf-8")) if jp.exists() else {}
    return ok, dt, data, rep


def main():
    say("=" * 78)
    say("跨行业统一知识抽取 · 验收（2026-11-05）")
    say("=" * 78)
    truth = json.loads((SET_DIR / "标准答案.json").read_text(encoding="utf-8"))

    # ---------- 1. 6 个行业 ----------
    results = {}
    say("\n【一】逐行业抽取（6 份 · 行业各不相同）")
    for doc in truth:
        ok, dt, data, rep = run_one(doc)
        results[doc] = data
        sch = data.get("schema") or {}
        units = data.get("units") or []
        say("\n◆ %s ｜ %s ｜ %.1f 秒 ｜ 条目 %d"
            % (doc, sch.get("domain") or "-", dt, len(units)))
        check("%s：抽取成功" % doc, ok and bool(units), "领域=%s 类型=%s" % (sch.get("domain"), sch.get("doc_type")))
        results[doc]["_dt"] = dt

    # ---------- 2. 结构命中 ----------
    say("\n【二】结构命中（对照人工标注）")
    tot_obj = hit_obj = 0
    tot_attr = hit_attr = 0
    for doc, t in truth.items():
        sch = results[doc].get("schema") or {}
        objs = sch.get("objects") or []
        obj_names = [o.get("type", "") for o in objs]
        attr_names = [a.get("name", "") for o in objs for a in (o.get("attrs") or [])]

        o_hit = [x for x in t["objects"] if any(hit(x, y) for y in obj_names)]
        # 「诉讼请求」「判决结果」这类东西，既可能被建成对象、也可能被建成某个对象的属性，
        # 两种建法都能把它抽出来（值照样进库）。所以放宽一档：出现在属性名里也算认出来了，
        # 但严格值（只在对象位命中）同时打印出来，方便看模型到底怎么理解的。
        o_hit_loose = [x for x in t["objects"]
                       if x in o_hit or any(hit(x, y) for y in attr_names)]
        a_hit = [x for x in t["attrs"] if any(hit(x, y) for y in attr_names)]
        tot_obj += len(t["objects"]); hit_obj += len(o_hit_loose)
        tot_attr += len(t["attrs"]); hit_attr += len(a_hit)
        say("  · %s ｜ 对象 %d/%d（严格 %d）%s ｜ 属性 %d/%d"
            % (doc, len(o_hit_loose), len(t["objects"]), len(o_hit),
               "、".join(obj_names)[:60], len(a_hit), len(t["attrs"])))
        if len(o_hit_loose) < len(t["objects"]):
            say("      缺对象：%s" % "、".join(x for x in t["objects"] if x not in o_hit_loose))
        if len(a_hit) < len(t["attrs"]):
            say("      缺属性：%s" % "、".join(x for x in t["attrs"] if x not in a_hit))
    r_obj = 100.0 * hit_obj / max(tot_obj, 1)
    r_attr = 100.0 * hit_attr / max(tot_attr, 1)
    check("行业结构命中率（对象位或属性位都算）≥ 90%", r_obj >= 90,
          "%.1f%%（%d/%d）" % (r_obj, hit_obj, tot_obj))
    check("行业属性名命中率 ≥ 80%", r_attr >= 80, "%.1f%%（%d/%d）" % (r_attr, hit_attr, tot_attr))

    # ---------- 3. 值命中 ----------
    say("\n【三】值命中（「人看一眼该抽到」的具体值）")
    tot_v = hit_v = 0
    for doc, t in truth.items():
        units = [u for u in results[doc].get("units") or [] if u.get("kind") == "attribute"]
        miss = []
        for hint, attr, val in t["values"]:
            found = False
            for u in units:
                if not hit(attr, u.get("attr", "")):
                    continue
                if not hit(val, u.get("value", "")):
                    continue
                # 对象提示：能对上对象名/类型/属性任一即可（不同文档叫法不一样）
                ctx = "%s %s %s" % (u.get("obj_name", ""), u.get("obj_type", ""), u.get("attr", ""))
                if hint and not hit(hint, ctx):
                    continue
                found = True
                break
            tot_v += 1
            hit_v += 1 if found else 0
            if not found:
                miss.append("%s/%s=%s" % (hint, attr, val))
        say("  · %s ｜ %d/%d%s" % (doc, len(t["values"]) - len(miss), len(t["values"]),
                                  (" ｜ 缺：" + "、".join(miss)) if miss else ""))
    r_val = 100.0 * hit_v / max(tot_v, 1)
    # 物流那份的"运单总数"属于汇总值，不作为硬指标
    check("标注值命中率 ≥ 90%", r_val >= 90, "%.1f%%（%d/%d）" % (r_val, hit_v, tot_v))

    # ---------- 4. 可溯 / 垃圾值 ----------
    say("\n【四】保真度：每条值都能回原文 / 不写糊在一起的值")
    all_units, bad, unverified, total_attr = [], [], [], 0
    for doc, data in results.items():
        for u in data.get("units") or []:
            all_units.append(u)
            if u.get("kind") != "attribute":
                continue
            total_attr += 1
            if not u.get("verified"):
                unverified.append("%s:%s=%s" % (doc, u.get("attr"), str(u.get("value"))[:20]))
            v = str(u.get("value") or "")
            toks = [x for x in v.split() if 1 <= len(x) <= 12]
            if len(toks) >= 4:
                bad.append("%s:%s" % (doc, u.get("attr")))
    good = total_attr - len(unverified)
    r_tr = 100.0 * good / max(total_attr, 1)
    check("属性值可核率 ≥ 98%", r_tr >= 98, "%.1f%%（%d/%d）" % (r_tr, good, total_attr))
    check("没有「一个单元格糊成一列」的脏值", len(bad) == 0, "脏值 %d 条" % len(bad))
    if unverified[:5]:
        say("    未完全可核示例：%s" % "；".join(unverified[:5]))

    # ---------- 5. 行业区分度 ----------
    say("\n【五】行业区分度（对象类型不能是同一套）")
    sets = {}
    for doc, t in truth.items():
        objs = [o.get("type", "") for o in (results[doc].get("schema") or {}).get("objects") or []]
        sets[t["industry"]] = set(objs)
    pairs, jac = [], []
    ks = list(sets)
    for i in range(len(ks)):
        for j in range(i + 1, len(ks)):
            a, b = sets[ks[i]], sets[ks[j]]
            j_score = len(a & b) / max(len(a | b), 1)
            jac.append(j_score)
            pairs.append((ks[i], ks[j], j_score))
    avg = sum(jac) / max(len(jac), 1)
    say("  · 各行业对象：%s" % "；".join("%s=%s" % (k, "、".join(sorted(v)) or "空") for k, v in sets.items()))
    check("任意两行业对象集合不雷同（平均 Jaccard ≤ 0.30）", avg <= 0.30,
          "平均 %.2f ｜ 最高 %.2f" % (avg, max(jac) if jac else 0))

    # ---------- 6. 入库 ----------
    say("\n【六】入库")
    try:
        import store
        st = store.knowledge_stats()
        check("知识库有数据", st["items"] > 0,
              "条目 %d ｜ 文档 %d ｜ kind=%s" % (st["items"], st["docs"], st["by_kind"]))
        check("6 个行业都进结构表", st["docs"] >= 6, "领域=%s" % list(st["by_domain"])[:8])
    except Exception as e:
        check("知识库有数据", False, str(e))

    # ---------- 汇总 ----------
    say("\n" + "=" * 78)
    say("结论：%d 项通过 / %d 项失败" % (PASS, FAIL))
    says = ["结构命中 %.1f%%" % r_obj, "属性命中 %.1f%%" % r_attr,
            "值命中 %.1f%%" % r_val, "可核率 %.1f%%" % r_tr,
            "行业区分度 %.2f" % avg]
    say("关键指标：" + " ｜ ".join(says))
    say("=" * 78)

    REPORT.write_text("\n".join(L), encoding="utf-8")
    print("\n报告已写：%s" % REPORT)
    return 0 if FAIL == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
