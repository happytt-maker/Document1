# -*- coding: utf-8 -*-
"""8000 服务健康检查：业务端点真实请求（ProxyHandler({}) 绕过本机代理劫持）"""
import sys, json, urllib.request

sys.stdout.reconfigure(errors="replace")
B = "http://127.0.0.1:8000"

_opener = urllib.request.build_opener(urllib.request.ProxyHandler({}))

# ---- 0. 等待服务就绪（最多 25s）----
import time
ready = False
for i in range(25):
    try:
        _opener.open(B + "/openapi.json", timeout=3)
        print("[0] 服务就绪（等了 %.0f 秒）" % (i * 1.0))
        ready = True
        break
    except Exception:
        time.sleep(1)
if not ready:
    print("[0] 25 秒内未就绪 —— 后面各项预期全失败")
    sys.exit(1)

def get(path):
    r = _opener.open(B + path, timeout=8)
    return json.loads(r.read().decode("utf-8", "replace"))

try:
    s = get("/api/status")
    print("[1] /api/status →", "ok" if s.get("ok") else s, "｜", str(s)[:150])
except Exception as e:
    print("[1] /api/status 失败:", e)

try:
    t = get("/api/fill/templates")
    names = t.get("templates") or t.get("data") or []
    print("[2] /api/fill/templates →", len(names), "张 ｜", names[:3])
except Exception as e:
    print("[2] /api/fill/templates 失败:", e)

try:
    d = get("/api/docs")
    n = len(d.get("docs") or d.get("data") or [])
    print("[3] /api/docs →", n, "份")
except Exception as e:
    print("[3] /api/docs 失败:", e)

try:
    r = _opener.open(B + "/", timeout=8)
    html = r.read().decode("utf-8", "replace")
    print("[4] 网页操作台 → HTTP", r.status, "｜", len(html), "字节 ｜ 含识别按钮:", "rec-toggle" in html, "｜ 含填写按钮:", "fill-toggle" in html)
except Exception as e:
    print("[4] 网页失败:", e)
