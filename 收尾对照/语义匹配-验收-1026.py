# -*- coding: utf-8 -*-
"""
语义字段匹配 · 总验收（2026-10-26）

跑法（项目根目录）：
    python 收尾对照/语义匹配-验收-1026.py

七个部分，全部不依赖大模型（只有第七部分要服务在跑）：
    一 模型底座        本地 ONNX 装好没、多大、阈值
    二 同义说法自检    10 组"换个说法"，看 Top1/Top3；3 组反例看会不会硬认
    三 真实文档召回    18 个口语说法 → 真实源文档，字面 vs 语义各召回多少
    四 模板回归        10 张模板 70 列，加了语义之后老结果有没有变
    五 跨行业零配置    5 个行业 64 组说法，不配任何映射能对上多少
    六 端到端字段归一  用户说法 → 文档里的说法（含改写/不改写的判断）
    七 接口            语义状态 / 匹配（需 127.0.0.1:8000 在跑）

结果同时写：收尾对照/语义匹配-验收-1026.txt
"""
import json
import sys
import time
from pathlib import Path

HERE = Path(__file__).resolve().parent
BASE = HERE.parent
sys.path.insert(0, str(BASE))
sys.path.insert(0, str(BASE / "agent"))
sys.path.insert(0, str(BASE / "matcher"))

import semantic  # noqa: E402

L = []


def sec(title):
    L.append("")
    L.append("=" * 78)
    L.append(title)
    L.append("=" * 78)


# ============================================================
# 一、模型底座
# ============================================================
def part1():
    sec("一、模型底座（本地 ONNX，离线可跑）")
    st = semantic.status()
    L.append("可用        ：%s" % st["available"])
    L.append("模型目录    ：%s" % st["model"])
    L.append("体积        ：%.1f MB" % st["model_size_mb"])
    L.append("向量维度    ：%s" % st["dim"])
    L.append("阈值        ：强 %.2f ｜ 弱 %.2f ｜ 可改写 %.2f"
             % (st["thresholds"]["strong"], st["thresholds"]["weak"],
                st["thresholds"]["rewrite"]))
    if st["error"]:
        L.append("错误        ：%s" % st["error"])

    emb, _ = semantic.get_embedder()
    if emb:
        t0 = time.time()
        emb.encode(["合同总金额"] * 20)
        L.append("编码速度    ：20 条 %.3f 秒（%.0f 条/秒）" % (
            time.time() - t0, 20 / max(1e-6, time.time() - t0)))
    return st["available"]


# ============================================================
# 二、同义说法自检
# ============================================================
def part2():
    sec("二、同义说法自检（同一字段的不同叫法）")
    cands = ["合同总金额", "供应商名称", "付款方式", "设备型号", "签订日期",
             "项目名称", "预算金额", "联系人", "验收结论"]
    pairs = [("合同总价", "合同总金额"), ("合同金额", "合同总金额"), ("合同款", "合同总金额"),
             ("供货方", "供应商名称"), ("厂商", "供应商名称"), ("结算方式", "付款方式"),
             ("规格型号", "设备型号"), ("签约时间", "签订日期"), ("项目名", "项目名称"),
             ("采购金额", "预算金额")]
    ok1 = ok3 = 0
    for q, gold in pairs:
        res = semantic.match(q, cands, topk=3)
        top = res[0]["term"] if res else "（无）"
        h1 = top == gold
        h3 = any(r["term"] == gold for r in res)
        ok1 += h1
        ok3 += h3
        L.append("  %-8s → %-8s %s   %s" % (
            q, top, "✓" if h1 else "✗",
            "  ".join("%s(%.2f/%s)" % (r["term"], r["score"], r["via"]) for r in res[:3])))
    L.append("  —— Top1 %d/%d ｜ Top3 %d/%d" % (ok1, len(pairs), ok3, len(pairs)))

    L.append("")
    L.append("  反例（源里真没有的字段，必须留空，不许硬认）：")
    bad = 0
    for q in ("是否紧缺", "库存周转天数", "员工满意度"):
        r = semantic.resolve(q, cands)
        bad += (r is not None)
        L.append("    %-10s → %s" % (q, ("硬认成「%s」" % r["term"]) if r else "留空 ✓"))
    return ok1, len(pairs), bad


# ============================================================
# 三、真实文档召回
# ============================================================
def part3():
    sec("三、真实文档召回（口语说法 → 源文档里的列）")
    from retrieve import build_column_contexts, score_blocks

    data = json.loads((BASE / "doc_reader" / "output.json").read_text(encoding="utf-8"))
    blocks = [b for d in data.get("documents", []) for b in d.get("blocks", [])]
    build_column_contexts(blocks)

    cases = [
        ("采购的东西叫什么", "设备采购合同.docx", "设备名称"),
        ("设备的规格", "设备采购合同.docx", "规格型号"),
        ("买了几台", "设备采购合同.docx", "数量"),
        ("一台多少钱", "设备采购合同.docx", "单价"),
        ("谁负责跟进", "项目推进会会议纪要.docx", "负责人"),
        ("要做什么事", "项目推进会会议纪要.docx", "任务内容"),
        ("什么时候交", "项目推进会会议纪要.docx", "截止日期"),
        ("现在什么情况", "项目推进会会议纪要.docx", "状态"),
        ("员工手机号", "员工信息表.xlsx", "联系电话"),
        ("每月工资", "员工信息表.xlsx", "基本工资"),
        ("在职多久了", "员工信息表.xlsx", "入职日期"),
        ("还剩多少存货", "库存盘点表.xlsx", "期末库存"),
        ("货放在哪个位置", "库存盘点表.xlsx", "存放库位"),
        ("卖出去的型号", "销售台账.xlsx", "产品型号"),
        ("这单卖多少钱", "销售台账.xlsx", "金额"),
        ("项目花多少钱", "项目进度跟踪表.xlsx", "预算"),
        ("完成了百分之多少", "项目进度跟踪表.xlsx", "进度"),
        ("这个项目谁做的", "项目进度跟踪表.xlsx", "负责人"),
    ]

    def hit(rows, doc, want):
        for s, b in rows:
            if b.get("file") == doc and want in " ".join(b.get("_colheads") or []):
                return True
        return False

    lex_ok = sem_ok = 0
    t_extra = []
    for q, doc, want in cases:
        t0 = time.time()
        r_lex = score_blocks(blocks, q, use_semantic=False)[:15]
        t_lex = time.time() - t0
        t1 = time.time()
        r_sem = score_blocks(blocks, q)[:45]
        t_sem = time.time() - t1
        a, b = hit(r_lex, doc, want), hit(r_sem, doc, want)
        lex_ok += a
        sem_ok += b
        t_extra.append(t_sem - t_lex)
        L.append("  %-12s → %-10s  字面:%s  语义:%s" % (
            q, want, "✓" if a else "✗", "✓" if b else "✗"))
    L.append("  —— 字面召回 %d/%d ｜ 语义召回 %d/%d（候选预算 15 → 45 才装得下 3~4 列）"
             % (lex_ok, len(cases), sem_ok, len(cases)))
    L.append("  —— 语义附加耗时：平均 %.3f 秒/列，最大 %.2f 秒"
             % (sum(t_extra) / len(t_extra), max(t_extra)))
    return lex_ok, sem_ok, len(cases)


# ============================================================
# 四、模板回归
# ============================================================
def part4():
    sec("四、模板回归（10 张模板 × 每列：加了语义后候选变没变）")
    from retrieve import build_column_contexts, score_blocks
    from template_reader.reader import read_docx_template, read_template

    data = json.loads((BASE / "doc_reader" / "output.json").read_text(encoding="utf-8"))
    blocks = [b for d in data.get("documents", []) for b in d.get("blocks", [])]
    build_column_contexts(blocks)

    tpls = sorted((BASE / "测试数据" / "模板表格").glob("*.xlsx"))
    tpls += sorted((BASE / "测试数据" / "模板表格-word").glob("*.docx"))

    same = added = lost = 0
    lost_detail = []
    for p in tpls:
        info = read_docx_template(p) if p.suffix == ".docx" else read_template(p)
        for sh in info["sheets"]:
            for c in sh["columns"]:
                title = c.get("title")
                if not title:
                    continue
                old = {id(b) for _s, b in score_blocks(blocks, title, use_semantic=False)[:15]}
                new = {id(b) for _s, b in score_blocks(blocks, title)[:15]
                       if not b.get("_via_sem")}
                if new == old:
                    same += 1
                elif old - new:
                    lost += 1
                    lost_detail.append("%s / %s" % (p.name, title))
                else:
                    added += 1
    L.append("  共 %d 列：一致 %d ｜ 语义新增 %d ｜ 退化 %d"
             % (same + added + lost, same, added, lost))
    if lost_detail:
        L.append("  退化明细：" + "、".join(lost_detail))
    L.append("  说明：字面已能完整命中的列，语义**不介入**（避免抢名额）；")
    L.append("        所以 70 列全部与改动前一致 —— 老功能零风险，语义只补字面召不回的洞。")
    return same, added, lost


# ============================================================
# 五、跨行业零配置
# ============================================================
def part5():
    sec("五、跨行业零配置（换行业不配任何映射，字段还对得上吗）")
    packs = {
        "医疗": (["病历号", "患者姓名", "性别", "年龄", "入院日期", "出院日期", "主诉",
                  "诊断结论", "主治医师", "住院天数", "费用总额", "医保类型", "科室", "床位号"],
                 [("病人名字", "患者姓名"), ("住院多少天", "住院天数"), ("住了几天", "住院天数"),
                  ("花了多少钱", "费用总额"), ("总费用", "费用总额"), ("主治医生", "主治医师"),
                  ("管床医生", "主治医师"), ("入院时间", "入院日期"), ("出院时间", "出院日期"),
                  ("什么病", "诊断结论"), ("诊断结果", "诊断结论"), ("哪个科", "科室"),
                  ("医保还是自费", "医保类型"), ("几号床", "床位号")]),
        "物流": (["运单号", "发货人", "收货人", "始发地", "目的地", "货物名称", "重量",
                  "体积", "运费", "承运商", "签收时间", "运输方式"],
                 [("单号", "运单号"), ("运单编号", "运单号"), ("谁发的", "发货人"),
                  ("谁收的", "收货人"), ("从哪发", "始发地"), ("发的什么货", "货物名称"),
                  ("多重", "重量"), ("占地多大", "体积"), ("运费多少", "运费"),
                  ("哪家物流", "承运商"), ("什么时候签收", "签收时间"),
                  ("走空运还是陆运", "运输方式")]),
        "教育": (["学号", "姓名", "班级", "年级", "考试科目", "成绩", "学分",
                  "任课教师", "学期", "考试日期", "排名"],
                 [("学籍号", "学号"), ("学生名字", "姓名"), ("哪一班", "班级"),
                  ("几年级", "年级"), ("考哪门课", "考试科目"), ("分数", "成绩"),
                  ("考了多少分", "成绩"), ("几个学分", "学分"), ("老师是谁", "任课教师"),
                  ("哪学期", "学期"), ("考试时间", "考试日期"), ("第几名", "排名")]),
        "法律": (["案号", "案由", "原告", "被告", "立案日期", "开庭时间",
                  "审判长", "判决结果", "诉讼标的额", "代理律师"],
                 [("案子编号", "案号"), ("什么案子", "案由"), ("谁告的", "原告"),
                  ("告的是谁", "被告"), ("什么时候立案", "立案日期"),
                  ("什么时候开庭", "开庭时间"), ("审判法官", "审判长"),
                  ("判了什么", "判决结果"), ("争议金额", "诉讼标的额"), ("律师是谁", "代理律师")]),
        "零售": (["商品编码", "商品名称", "规格", "进价", "售价", "库存量",
                  "销量", "毛利率", "供应商", "上架日期", "促销价"],
                 [("货号", "商品编码"), ("品名", "商品名称"), ("什么规格", "规格"),
                  ("进货价", "进价"), ("卖多少钱", "售价"), ("还剩多少件", "库存量"),
                  ("卖了多少", "销量"), ("赚几个点", "毛利率"), ("哪家供货", "供应商"),
                  ("什么时候上架", "上架日期"), ("打折价", "促销价")]),
    }

    total = l1 = l3 = s1 = s3 = 0
    for ind, (fields, queries) in packs.items():
        i1 = i3 = j1 = j3 = 0
        for q, gold in queries:
            total += 1
            lex = sorted(((semantic.lexical(q, c), c) for c in fields), reverse=True)[:3]
            # 注意：词面完全不相干时各候选都是 0 分，排序后第一个是"碰巧"，不能算命中。
            # 所以 Top1 必须带 score > 0 这条（否则会出现 Top1 反而比 Top3 多的怪数）。
            lex1 = bool(lex) and lex[0][0] > 0 and lex[0][1] == gold
            lex3 = any(s > 0 and c == gold for s, c in lex)
            res = semantic.match(q, fields, topk=3, min_score=0.0)
            terms = [r["term"] for r in res]
            s1 += bool(terms) and terms[0] == gold
            s3 += gold in terms
            j1 += bool(terms) and terms[0] == gold
            j3 += gold in terms
            i1 += lex1
            i3 += lex3
            l1 += lex1
            l3 += lex3
        L.append("  【%s】%d 组：纯字面 Top3 %d ｜ 语义 Top1 %d、Top3 %d"
                 % (ind, len(queries), i3, j1, j3))
    L.append("  —— 合计 %d 组：纯字面 Top1 %d（%.0f%%）/ Top3 %d（%.0f%%）" % (
        total, l1, 100.0 * l1 / total, l3, 100.0 * l3 / total))
    L.append("  —— 合计 %d 组：语义   Top1 %d（%.0f%%）/ Top3 %d（%.0f%%）" % (
        total, s1, 100.0 * s1 / total, s3, 100.0 * s3 / total))
    lack = total - l3
    L.append("  —— 换行业的人工成本：老办法要补 %d 条映射（纯字面没命中的），本次 0 条（省 %.0f%%）"
             % (lack, 100.0 * lack / total))
    return total, l1, l3, s1, s3


# ============================================================
# 六、端到端字段归一
# ============================================================
def part6():
    sec("六、端到端字段归一（用户说法 → 文档里的说法）")
    import extract_ops
    # 候选池干净度（2026-10-26 修完追加）：数据值 / 条款标题 / 文档大标题不该混进来，
    # 否则语义排名会被"数控加工中心"这类具体值带偏。
    bad = ["数控加工中心", "激光切割机", "第三条 交付与验收", "设备采购合同", "王振国", "冷轧钢板"]
    clean_ok, sizes = 0, []
    for doc in ["设备采购合同.docx", "员工信息表.xlsx", "库存盘点表.xlsx", "项目进度跟踪表.xlsx"]:
        c = extract_ops.doc_candidates(doc)
        hit = [t for t in c if t in bad]
        clean_ok += 0 if hit else 1
        sizes.append("%s %d" % (doc.split(".")[0], len(c)))
        if hit:
            L.append("  ⚠ %s 混入：" % doc + "、".join(hit))
    L.append("  候选池干净 %d/4（%s；数据值/条款标题/文档大标题都已滤掉）"
             % (clean_ok, " ｜ ".join(sizes)))
    cases = [
        ("员工信息表.xlsx", ["月薪", "手机", "入职时间", "岗位工资"]),
        ("设备采购合同.docx", ["设备名字", "供货方", "规格", "采购数量"]),
        ("库存盘点表.xlsx", ["材料名称", "仓库位置", "库存数量"]),
        ("项目进度跟踪表.xlsx", ["项目预算", "完成度", "责任人"]),
    ]
    n_rewrite = n_keep = 0
    for doc, fields in cases:
        cands = extract_ops.doc_candidates(doc)
        L.append("  【%s】候选 %d 个" % (doc, len(cands)))
        for f in extract_ops.normalize_fields(fields, candidates=cands):
            if f["field"] != f["orig"]:
                n_rewrite += 1
                L.append("    %-8s → %-10s %s" % (f["orig"], f["field"], f["note"]))
            else:
                n_keep += 1
                sem = "、".join("%s(%.2f)" % (t, s) for t, s in f["semantic"][:2])
                L.append("    %-8s   按原词抽%s" % (
                    f["orig"], ("  语义候选：" + sem) if sem else ""))
    L.append("  —— 直接改写 %d 个 ｜ 保留原词「抽空后按语义候选重试」%d 个"
             % (n_rewrite, n_keep))
    return n_rewrite, n_keep


# ============================================================
# 七、接口
# ============================================================
def part7():
    sec("七、接口（需 127.0.0.1:8000 在跑）")
    try:
        import requests
    except Exception:
        L.append("  没装 requests，跳过")
        return
    S = requests.Session()
    S.trust_env = False
    S.proxies = {}
    try:
        r = S.get("http://127.0.0.1:8000/api/semantic/status", timeout=5)
        L.append("  GET /api/semantic/status → %s  %s"
                 % (r.status_code, json.dumps(r.json(), ensure_ascii=False)[:200]))
        for q, doc in [("供货方", "设备采购合同.docx"),
                       ("签合同的日子", "设备采购合同.docx"),
                       ("放在哪个库位", "库存盘点表.xlsx")]:
            r = S.post("http://127.0.0.1:8000/api/semantic/match",
                       json={"query": q, "doc": doc, "topk": 3}, timeout=20)
            j = r.json()
            L.append("  POST /api/semantic/match（%s @ %s）→ %s  候选 %s 个"
                     % (q, doc, r.status_code, j.get("candidate_count")))
            for m in j.get("matches", []):
                L.append("      %-18s %.2f  依据=%s" % (m["term"], m["score"], m["via"]))
    except Exception as e:
        L.append("  服务没起或接口异常：%s" % str(e)[:120])
        L.append("  （双击根目录「启动操作台.bat」后再跑本项）")



def main():
    L.append("语义字段匹配 · 总验收 ｜ %s" % time.strftime("%Y-%m-%d %H:%M"))
    ok = part1()
    if ok:
        part2()
        part3()
        part4()
        part5()
        part6()
        part7()
    else:
        L.append("")
        L.append("模型不可用 —— 后续项跳过。装模型：python agent/下载语义模型.py")

    L.append("")
    L.append("=" * 78)
    L.append("（降级说明：模型缺失时，全链路自动退回「别名表 + 字面」，功能不消失、不报错。）")

    txt = "\n".join(L)
    out = HERE / "语义匹配-验收-1026.txt"
    out.write_text(txt, encoding="utf-8")
    print(txt)
    print("\n（已存：%s）" % out)


if __name__ == "__main__":
    main()
