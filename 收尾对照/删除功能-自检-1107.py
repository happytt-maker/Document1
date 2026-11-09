# -*- coding: utf-8 -*-
"""删除功能端到端自检：上传一份 txt → 从文档库删除；上传一张模板 → 删掉（归档）"""
import json
import os
import urllib.request
import urllib.parse
import urllib.error

BASE = "http://127.0.0.1:8000"
for k in ("HTTP_PROXY", "HTTPS_PROXY", "ALL_PROXY", "http_proxy", "https_proxy", "all_proxy"):
    os.environ.pop(k, None)
op = urllib.request.build_opener(urllib.request.ProxyHandler({}))
out = []

# 关掉系统代理对 127.0.0.1 的劫持
r = op.open(BASE + "/api/status", timeout=10)
out.append("status=%s" % r.status)

# 1) 上传 txt
boundary = "----a23test"
body = (
    "--%s\r\nContent-Disposition: form-data; name=\"file\"; filename=\"selftest-delete.txt\"\r\n"
    "Content-Type: text/plain\r\n\r\n这是一份删除功能自检用的临时文档。\r\n--%s--\r\n"
    % (boundary, boundary)).encode("utf-8")
req = urllib.request.Request(BASE + "/api/upload", data=body,
                             headers={"Content-Type": "multipart/form-data; boundary=" + boundary})
j = json.loads(op.open(req, timeout=60).read())
out.append("upload ok=%s stored=%s" % (j.get("ok"), j.get("stored_name")))

# 2) 删除它
name = j.get("stored_name")
try:
    rq = urllib.request.Request(BASE + "/api/docs?name=" + urllib.parse.quote(name), method="DELETE")
    d = json.loads(op.open(rq, timeout=30).read())
    out.append("doc delete ok=%s archived=%s" % (d.get("ok"), bool(d.get("archived"))))
    out.append("archived_exists=%s" % os.path.exists(d.get("archived") or ""))
except urllib.error.HTTPError as e:
    out.append("doc delete HTTP %s %s" % (e.code, e.read()[:200]))

# 3) 上传一张模板（真 xlsx 的最小内容由 openpyxl 生成）
try:
    from openpyxl import Workbook
    tmp = os.path.join(os.environ.get("TEMP", "."), "selftest-tpl.xlsx")
    wb = Workbook()
    wb.active["A1"] = "自检"
    wb.save(tmp)
    data = open(tmp, "rb").read()
    b2 = "----a23tpl"
    body2 = (
        "--%s\r\nContent-Disposition: form-data; name=\"file\"; filename=\"selftest-tpl.xlsx\"\r\n"
        "Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet\r\n\r\n"
        % b2).encode("utf-8") + data + ("\r\n--%s--\r\n" % b2).encode()
    rq = urllib.request.Request(BASE + "/api/fill/templates/upload", data=body2,
                                headers={"Content-Type": "multipart/form-data; boundary=" + b2})
    t = json.loads(op.open(rq, timeout=30).read())
    out.append("tpl upload ok=%s name=%s" % (t.get("ok"), t.get("name")))
    rq2 = urllib.request.Request(BASE + "/api/fill/templates?name=" +
                                 urllib.parse.quote(t.get("name")), method="DELETE")
    td = json.loads(op.open(rq2, timeout=30).read())
    out.append("tpl delete ok=%s archived_exists=%s" % (td.get("ok"), os.path.exists(td.get("archived") or "")))
except urllib.error.HTTPError as e:
    out.append("tpl HTTP %s %s" % (e.code, e.read()[:300]))
except Exception as e:
    out.append("tpl err %s: %s" % (type(e).__name__, e))

# 4) 系统模板不允许走"不存在"分支误判
try:
    rq3 = urllib.request.Request(BASE + "/api/fill/templates?name=" + urllib.parse.quote("根本没有这张表.xlsx"),
                                 method="DELETE")
    op.open(rq3, timeout=20)
    out.append("missing tpl NOT rejected !!")
except urllib.error.HTTPError as e:
    out.append("missing tpl rejected=%s" % e.code)

print("\n".join(out))
open(r"C:\Users\覃敬涛\Desktop\研一项目\收尾对照\删除功能-自检-1107.txt", "w", encoding="utf-8").write("\n".join(out))
