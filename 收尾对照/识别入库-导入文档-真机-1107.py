# -*- coding: utf-8 -*-
"""识别入库 · 上传的文档（导入文档目录）能否抽取 —— 端到端真机验收（2026-11-07）

背景（用户反馈）：在「识别入库」勾选"我上传的"Excel → 报
  【失败】文档《…xlsx》不在解析池里——如果是刚导入的，可跑 agent/pool.py --rebuild
根因：解析池 doc_reader/output.json 只扫描「测试数据/源文档」，
     用户上传的文件落在「导入文档」，永远不在池里。

本脚本走 HTTP 接口（不直接调函数），验证：
  A. 文档库接口能看到"我上传的"那份
  B. 提交识别（knowledge）→ 任务成功、报告里没有"不在解析池"
  C. 入库条数 > 0，且文档视图能查到这个文档
  D. 明细表对象名带时间（不再把同站点多条记录合并成一个对象）

跑法：python 收尾对照/识别入库-导入文档-真机-1107.py
"""
import json
import os
import sys
import time
import urllib.request

HOST = "http://127.0.0.1:%s" % os.environ.get("A23_TEST_PORT", "8000")
DOC = "山东省环境空气质量监测数据信息202512171921_0.xlsx"

# 本机测试：系统代理会把 127.0.0.1 的 GET 劫持成假 404
_OPENER = urllib.request.build_opener(urllib.request.ProxyHandler({}))

PASS, FAIL = [], []


def check(name, cond, detail=""):
    (PASS if cond else FAIL).append(name)
    detail = "" if not detail else str(detail)
    print("  %s %s%s" % ("✓" if cond else "✗", name, ("  —— " + detail) if detail else ""))


def api(path, body=None, method=None, timeout=30):
    url = HOST + path
    data = None
    headers = {}
    if body is not None:
        data = json.dumps(body, ensure_ascii=False).encode("utf-8")
        headers["Content-Type"] = "application/json"
    req = urllib.request.Request(url, data=data, headers=headers,
                                 method=method or ("POST" if data else "GET"))
    with _OPENER.open(req, timeout=timeout) as r:
        return json.loads(r.read().decode("utf-8"))


print("=" * 60)
print("识别入库 · 导入文档端到端验收 · %s" % HOST)
print("=" * 60)

# ---------- A. 文档库 ----------
print("\n[A] 文档库清单")
docs = api("/api/docs")
names = docs.get("docs") or []
detail = docs.get("detail") or []
mine = next((d for d in detail if d.get("name") == DOC), None)
check("文档库里能列出目标文件", DOC in names, DOC if DOC in names else "没找到")
if mine:
    check("来源标为「我上传的」", mine.get("src") == "import", str(mine.get("src")))
check("文档库非空", len(names) > 0, "%d 份" % len(names))

# ---------- B. 提交识别 ----------
print("\n[B] 提交识别（knowledge）")
reuse = (os.environ.get("A23_REUSE_JOB") or "").strip()
if reuse:
    print("  · 复用已有任务 %s（不重新抽取，用于复查结果）" % reuse)
    jid = reuse
    t0 = time.time()
else:
    t0 = time.time()
    sub = api("/api/recognize", {"docs": [DOC], "tasks": ["knowledge"], "requirement": ""})
    jid = sub.get("job_id")
    check("接口受理并返回 job_id", bool(jid), str(sub)[:160])

job = {}
while time.time() - t0 < 300:
    job = api("/api/job/%s" % jid)
    if job.get("status") in ("done", "failed"):
        break
    time.sleep(1.0)
elapsed = time.time() - t0
res = job.get("result") or {}
rep = res.get("report") or ""
check("任务跑完（status=done）", job.get("status") == "done", "%s %s" % (job.get("status"), (res.get("error") or "")[:80]))
check("没有「不在解析池」这类失败", "不在解析池" not in rep, rep[:200].replace("\n", " / "))
check("报告里有知识线 ✓", "知识 ✓" in rep, [l.strip() for l in rep.splitlines() if "知识" in l][:1])
check("耗时在 90 秒内", elapsed <= 90, "%.1fs" % elapsed)

# ---------- C. 入库 ----------
print("\n[C] 入库结果")
meta = res.get("meta") or {}
k = meta.get("knowledge") or {}
n_items = k.get("items") or 0
check("知识库有条目", n_items > 0, "%s 条" % n_items)
check("对象数 > 300（明细表逐行建对象）", (k.get("by_kind", {}) or {}).get("object", 0) > 300,
      "对象 %s" % (k.get("by_kind", {}) or {}).get("object"))
dd = api("/api/knowledge/doc?doc=%s" % urllib.request.quote(DOC))
check("文档视图能查到这份文档", bool(dd.get("ok")) and bool(dd.get("found")), str(dd)[:120])
d_doc = dd
check("文档视图有分组条目", (dd.get("total") or 0) > 0, "total=%s" % dd.get("total"))

# ---------- D. 明细表对象命名 ----------
print("\n[D] 明细表对象名（应带时间，不能把同站点多条合并）")
oi = api("/api/knowledge/items?doc=%s&kind=object&limit=6" % urllib.request.quote(DOC))
orows = oi.get("rows") or []
sample_names = [r.get("obj_name") for r in orows]
check("能查到对象条目", len(orows) > 0, "前 6 条：%s" % " / ".join(str(s) for s in sample_names))
check("对象名带时间（明细不再折叠）",
      any("2025-11-25" in (s or "") for s in sample_names),
      " / ".join(str(s) for s in sample_names)[:180])

print("\n" + "=" * 60)
print("通过 %d 项 · 失败 %d 项 · 用时 %.1fs" % (len(PASS), len(FAIL), elapsed))
if FAIL:
    print("失败项：" + "、".join(FAIL))
print("=" * 60)
sys.exit(0 if not FAIL else 1)
