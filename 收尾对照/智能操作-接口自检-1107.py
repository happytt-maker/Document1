# -*- coding: utf-8 -*-
"""智能操作 · 接口自检（打 8000 真服务）"""
import json, sys, time, urllib.request

sys.stdout.reconfigure(encoding="utf-8")
BASE = "http://127.0.0.1:8000"
OPENER = urllib.request.build_opener(urllib.request.ProxyHandler({}))
PASS = []


def req(path, data=None, method=None):
    body = json.dumps(data).encode() if data is not None else None
    r = urllib.request.Request(BASE + path, data=body, method=method or ("POST" if body else "GET"),
                               headers={"Content-Type": "application/json"})
    with OPENER.open(r, timeout=30) as f:
        return json.loads(f.read().decode())


def wait(jid, limit=90):
    t0 = time.time()
    while time.time() - t0 < limit:
        j = req("/api/job/" + jid)
        if j.get("status") in ("done", "failed"):
            return j
        time.sleep(0.8)
    raise RuntimeError("超时")


def check(name, cond, extra=""):
    PASS.append((name, bool(cond)))
    print(("  ✓ " if cond else "  ✗ ") + name + ("  " + extra if extra else ""))


print("① 快捷按钮（不走模型）")
t0 = time.time()
j0 = req("/api/op", {"op": "set_bold", "doc": "设备采购合同.docx", "hint": "标题", "by": "keyword"})
j = wait(j0["job_id"])
dt = time.time() - t0
res = j.get("result") or {}
check("POST /api/op 快捷 mode=quick", j0.get("mode") == "quick")
check("执行成功 + 有产物", res.get("ok") and (res.get("meta") or {}).get("out_path"),
      "耗时 %.1fs" % dt)
check("走的是 preset（报告里有「快捷操作」）", "快捷操作" in (res.get("report") or ""))
out_file = (res.get("meta") or {}).get("out_path", "").split("\\")[-1]

print("② 选中文件 + 自然语言「标题加粗」")
j0 = req("/api/op", {"text": "标题加粗", "doc": "设备采购合同.docx"})
j = wait(j0["job_id"])
res = j.get("result") or {}
check("执行成功", res.get("ok"), "耗时 %ss" % j.get("elapsed"))
check("按选中文件处理", "已选文件" in (res.get("report") or ""))

print("③ 不选文件 + 「设备采购合同的标题加粗」")
j0 = req("/api/op", {"text": "设备采购合同的标题加粗"})
j = wait(j0["job_id"])
res = j.get("result") or {}
check("执行成功", res.get("ok"), "耗时 %ss" % j.get("elapsed"))
check("文档=设备采购合同.docx", (res.get("meta") or {}).get("doc") == "设备采购合同.docx")

print("④ 打开产物 / 打开目录（服务端算路径）")
r = req("/api/op/open", {"file": out_file, "mode": "folder"})
check("打开目录返回路径", r.get("ok") and r.get("path", "").endswith(out_file), r.get("path", ""))
try:
    req("/api/op/open", {"file": "../../secret.txt", "mode": "file"})
    check("越权路径被拒", False)
except urllib.error.HTTPError as e:
    check("越权路径被拒", e.code == 404, "HTTP %d" % e.code)

print("⑤ 首页对话框没被破坏（/api/command 原路照旧）")
j0 = req("/api/command", {"text": "把设备采购合同的标题居中"})
j = wait(j0["job_id"])
res = j.get("result") or {}
check("/api/command 仍可用", res.get("ok"), "耗时 %ss" % j.get("elapsed"))
check("kind=command", j.get("kind") == "command")

print("\n通过 %d/%d" % (sum(1 for _n, o in PASS if o), len(PASS)))
sys.exit(0 if all(o for _n, o in PASS) else 1)
