# -*- coding: utf-8 -*-
"""确认重启后的 8000 认识新的「意见反馈」功能项（13 项），并清掉测试记录。"""
import json
import os
import sys
import urllib.request

BASE = "http://127.0.0.1:8000"
AGENT = r"C:\Users\覃敬涛\Desktop\研一项目\agent"
sys.path.insert(0, AGENT)
for k in ("HTTP_PROXY", "HTTPS_PROXY", "ALL_PROXY", "http_proxy", "https_proxy", "all_proxy"):
    os.environ.pop(k, None)
_op = urllib.request.build_opener(urllib.request.ProxyHandler({}))


def call(path, data=None, cookie=None):
    body = json.dumps(data).encode() if data is not None else None
    req = urllib.request.Request(BASE + path, data=body)
    if body:
        req.add_header("Content-Type", "application/json")
    if cookie:
        req.add_header("Cookie", cookie)
    r = _op.open(req, timeout=20)
    return json.loads(r.read())


tok = None
r = _op.open(urllib.request.Request(BASE + "/api/auth/login",
             data=json.dumps({"username": "2482991994", "password": "123456"}).encode(),
             headers={"Content-Type": "application/json"}), timeout=20)
for c in (r.headers.get_all("Set-Cookie") or []):
    for p in c.split(";"):
        if p.strip().startswith("a23_token="):
            tok = p.strip()
print("login:", bool(tok), flush=True)

FEATS = ["doc_edit", "warehouse", "extract", "fill", "template", "ocr",
         "voice", "image", "chat", "room", "account", "ui", "other"]
bad = 0
ids = []
for f in FEATS:
    try:
        d = call("/api/feedback", {"feature": f, "rating": 5, "content": "自检-可删", "tags": ""}, tok)
        ids.append(d["item"]["id"])
        print("  OK  ", f, "->", d["item"]["feature_text"], flush=True)
    except Exception as e:
        bad += 1
        print("  FAIL", f, "->", e, flush=True)

print("通过 %d / 失败 %d" % (len(ids), bad), flush=True)

# 清掉自检记录
if ids:
    try:
        from db import get_conn
        conn = get_conn()
        cur = conn.cursor()
        q = "DELETE FROM feedback WHERE id IN (%s)" % ",".join(["%s"] * len(ids))
        cur.execute(q, ids)
        conn.commit()
        print("已清理测试记录:", cur.rowcount, "条", flush=True)
        cur.close()
        conn.close()
    except Exception as e:
        print("清理失败（不影响功能）:", e, flush=True)
