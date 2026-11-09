# -*- coding: utf-8 -*-
"""
知识-接口冒烟-1105.py —— 统一知识抽取的接口验收

默认打 8000（正式实例）；要打别的端口就设环境变量 A23_TEST_PORT。
（早先固定打 8012，是因为 8000 上跑着旧进程、不能动；现在 8000 也是新代码了。）

验 7 件事：
  1 GET  /api/knowledge/stats      有数据、字段齐
  2 GET  /api/knowledge/items      能按 kind / 关键词筛
  3 GET  /api/knowledge/schema     结构能解析成对象数组（不是一坨字符串）
  4 GET  /api/knowledge/schemas    文档清单带 attr_count
  5 POST /api/knowledge/run        异步任务跑通（真抽一份医疗报告）
  6 GET  /api/job/{id}             报告里有"领域 / 对象 / 入库条目"
  7 关注点生效                     requirement 传进去后 focus 命中 ≥1 条
"""
import json
import os
import sys
import time
import urllib.parse
import urllib.request

BASE = "http://127.0.0.1:%s" % (os.environ.get("A23_TEST_PORT") or "8000")
OP = urllib.request.build_opener(urllib.request.ProxyHandler({}))   # 系统代理会拦 127.0.0.1
PASS = FAIL = 0
L = []


def say(s=""):
    print(s)
    L.append(str(s))


def check(name, ok, detail=""):
    global PASS, FAIL
    if ok:
        PASS += 1
        say("  [PASS] %s%s" % (name, (" ｜ " + detail) if detail else ""))
    else:
        FAIL += 1
        say("  [FAIL] %s%s" % (name, (" ｜ " + detail) if detail else ""))


def get(path):
    r = OP.open(BASE + path, timeout=30)
    return json.loads(r.read().decode("utf-8"))


def post(path, body):
    req = urllib.request.Request(BASE + path, method="POST",
                                 data=json.dumps(body, ensure_ascii=False).encode("utf-8"),
                                 headers={"Content-Type": "application/json"})
    return json.loads(OP.open(req, timeout=30).read().decode("utf-8"))


def main():
    say("=" * 70)
    say("统一知识抽取 · 接口冒烟（%s）" % BASE)
    say("=" * 70)

    # 0 等服务起来（注意：服务没有 /api/health，健康检查走 /api/status）
    for _ in range(40):
        try:
            get("/api/status")
            break
        except Exception:
            time.sleep(1)
    else:
        say("服务没起来")
        return 1

    say("\n【1】统计")
    d = get("/api/knowledge/stats")
    check("stats.ok", d.get("ok") is True)
    s = d.get("stats") or {}
    check("有知识条目", (s.get("items") or 0) > 0, "items=%s docs=%s" % (s.get("items"), s.get("docs")))
    check("按 kind 分组齐全", set(s.get("by_kind") or {}) >= {"attribute"}, str(s.get("by_kind")))
    check("有领域分布", bool(s.get("by_domain")), str(list(s.get("by_domain") or {}))[:80])

    say("\n【2】条目查询")
    d = get("/api/knowledge/items?kind=attribute&limit=5")
    rows = d.get("rows") or []
    check("按 kind=attribute 取到", d.get("ok") and len(rows) > 0, "%d 条" % len(rows))
    if rows:
        r = rows[0]
        need = {"doc_name", "kind", "obj_type", "attr", "value", "loc", "verified"}
        check("字段齐（文档/对象/属性/值/出处/核验）", need <= set(r), str(sorted(need - set(r)) or "齐"))
    d2 = get("/api/knowledge/items?q=%s&limit=5" % urllib.parse.quote("血红蛋白"))
    check("关键词能搜到", d2.get("ok") and len(d2.get("rows") or []) > 0,
          "命中 %d 条" % len(d2.get("rows") or []))

    say("\n【3】结构（单份）")
    d = get("/api/knowledge/schema?doc=%s" % urllib.parse.quote("判决书"))
    check("找到结构", d.get("ok") and d.get("found"), "")
    sch = d.get("schema") or {}
    objs = sch.get("objects") or []
    check("objects 是数组（不是字符串）", isinstance(objs, list) and len(objs) > 0, "%d 个对象" % len(objs))
    if objs:
        check("对象带 attrs", isinstance(objs[0].get("attrs"), list),
              "首个=%s" % objs[0].get("type"))
    check("领域不是空", bool(sch.get("domain")), sch.get("domain"))

    say("\n【4】结构清单")
    d = get("/api/knowledge/schemas?limit=20")
    rows = d.get("rows") or []
    check("清单非空", d.get("ok") and len(rows) >= 6, "%d 份" % len(rows))
    check("带属性数", all("attr_count" in r for r in rows), str(rows[0].get("attr_count") if rows else ""))

    say("\n【5】异步抽取（民政公报 · 带关注点）")
    # 注意：/api/knowledge/run 会去「测试数据\源文档 + 导入文档」找文件，
    # 用行业测试集里的名字会 404（那批文件只给离线验收脚本用）。
    j = post("/api/knowledge/run", {"docs": ["2021年民政事业发展统计公报.docx"],
                                    "requirement": "重点关注床位和农村的情况"})
    jid = j.get("job_id")
    check("提交成功", bool(jid), "job=%s" % jid)
    st = {}
    # 叙述体会分批（每批 8 块、最多 4 批并发）再加结构归纳，长文档 90 秒以上是正常的；
    # 早先这里只等 90 秒，民政公报这种就判成"没完成"。放宽到 4 分钟。
    for _ in range(160):
        time.sleep(1.5)
        st = get("/api/job/" + jid)
        if st.get("status") in ("done", "failed"):
            break
    check("任务完成", st.get("status") == "done", "status=%s ok=%s" % (st.get("status"), st.get("ok")))
    # 识别类任务的报告不在顶层 report，而是挂在 result.meta.docs[].lines 下
    rep = st.get("report") or ""
    if not rep:
        for _d in (((st.get("result") or {}).get("meta") or {}).get("docs") or []):
            for _ln in (_d.get("lines") or {}).values():
                rep += (_ln.get("report") or "")
    for kw in ("领域", "对象", "入库条目"):
        check("报告含「%s」" % kw, kw in rep)
    say("    报告片段：")
    for line in rep.splitlines():
        if line.strip().startswith(("◆ 领域", "◆ 对象", "◆ 入库条目", "◆ 关注点")):
            say("      " + line.strip())

    say("\n【6】关注点是否落库")
    d = get("/api/knowledge/items?doc=%s&kind=attribute&limit=200" % urllib.parse.quote("民政"))
    rows = d.get("rows") or []
    n_focus = sum(1 for r in rows if r.get("focus"))
    check("focus 标记至少 1 条", n_focus >= 1, "focus=%d / %d" % (n_focus, len(rows)))
    d2 = get("/api/knowledge/schema?doc=%s" % urllib.parse.quote("民政"))
    say("    requirement 原话：%s" % ((d2.get("schema") or {}).get("requirement") or "（空）"))

    say("\n" + "=" * 70)
    say("结论：%d 通过 / %d 失败" % (PASS, FAIL))
    say("=" * 70)
    open(r"C:\Users\覃敬涛\Desktop\研一项目\收尾对照\知识-接口冒烟-1105.txt",
         "w", encoding="utf-8").write("\n".join(L))
    return 0 if FAIL == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
