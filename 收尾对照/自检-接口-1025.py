# -*- coding: utf-8 -*-
"""运行时自检 v2：按真实接口签名（HttpOnly cookie 登录）"""
import http.cookiejar
import json
import os
import time
import urllib.error
import urllib.parse
import urllib.request

BASE = "http://127.0.0.1:8000"
OUT = r"C:\Users\覃敬涛\WorkBuddy\2026-09-24-10-48-03\_runtime2.txt"

ok = 0
bad = 0
lines = []


def chk(name, cond, detail=""):
    global ok, bad
    if cond:
        ok += 1
        lines.append("  [OK]   %-40s %s" % (name, detail))
    else:
        bad += 1
        lines.append("  [FAIL] %-40s %s" % (name, detail))


class Client(object):
    def __init__(self):
        self.jar = http.cookiejar.CookieJar()
        self.op = urllib.request.build_opener(
            urllib.request.ProxyHandler({}),
            urllib.request.HTTPCookieProcessor(self.jar),
        )

    def call(self, path, method="GET", data=None, raw=None, ctype="application/json",
             timeout=120):
        headers = {}
        body = None
        if raw is not None:
            body = raw
            headers["Content-Type"] = ctype
        elif data is not None:
            body = json.dumps(data, ensure_ascii=False).encode("utf-8")
            headers["Content-Type"] = ctype
        req = urllib.request.Request(BASE + path, data=body, headers=headers, method=method)
        try:
            r = self.op.open(req, timeout=timeout)
            return r.status, r.read()
        except urllib.error.HTTPError as e:
            return e.code, e.read()
        except Exception as e:
            return -1, ("%s: %s" % (type(e).__name__, e)).encode("utf-8")

    def login(self, username, password, portal="user"):
        code, b = self.call("/api/auth/login", "POST",
                            {"username": username, "password": password, "portal": portal})
        try:
            j = json.loads(b.decode("utf-8"))
        except Exception:
            j = {}
        return code, j


def getj(c, path, **kw):
    code, b = c.call(path, **kw)
    try:
        return code, json.loads(b.decode("utf-8"))
    except Exception:
        return code, {}


# ---------- 0 ----------
lines.append("=== 0. 服务 ===")
anon = Client()
up = False
for _ in range(20):
    code, _b = anon.call("/login", timeout=3)
    if code == 200:
        up = True
        break
    time.sleep(1)
chk("服务在 8000 响应", up, "GET /login")
if not up:
    open(OUT, "w", encoding="utf-8").write("\n".join(lines))
    raise SystemExit(1)

# ---------- 1 页面 ----------
lines.append("")
lines.append("=== 1. 页面与静态资源 ===")
for path, name, minsize in [("/login", "登录页", 10000),
                            ("/docs", "接口文档", 800),
                            ("/", "未登录跳登录", 0)]:
    code, b = anon.call(path)
    chk(name, code in (200, 302, 307) and len(b) >= minsize, "HTTP %s, %d 字节" % (code, len(b)))

# ---------- 2 登录（MySQL） ----------
lines.append("")
lines.append("=== 2. 登录与账号（MySQL 通路） ===")
u = Client()
code, j = u.login("demo_ai", "123456", "user")
usr = j.get("user") or {}
chk("demo_ai 登录", code == 200 and bool(usr), "HTTP %s, uid=%s, roles=%s"
    % (code, usr.get("id"), usr.get("roles")))
chk("登录后拿到 HttpOnly cookie", len(u.jar) > 0, "cookie 数=%d" % len(u.jar))

code, j = getj(u, "/api/auth/me")
chk("/api/auth/me 可用", code == 200 and j.get("user"), "HTTP %s" % code)

code, b = u.call("/")
chk("登录后首页可打开（操作台）", code == 200 and len(b) > 50000, "HTTP %s, %d 字节" % (code, len(b)))

adm = Client()
code, j = adm.login("2482991994", "123456", "admin")
chk("全角色账号 2482991994 登录", code == 200 and (j.get("user") or {}).get("roles"),
    "HTTP %s, roles=%s" % (code, (j.get("user") or {}).get("roles")))
for p, name in [("/staff", "客服端 /staff"), ("/admin", "管理端 /admin")]:
    code, b = adm.call(p)
    chk(name + " 可打开", code == 200 and len(b) > 5000, "HTTP %s, %d 字节" % (code, len(b)))

# ---------- 3 数据库 ----------
lines.append("")
lines.append("=== 3. 数据库（模块2 五表） ===")
code, j = getj(u, "/api/db/stats")
st = j.get("stats") or {}
chk("/api/db/stats", code == 200 and bool(st), "HTTP %s" % code)
if st:
    lines.append("       文档 %s / 实体 %s / 画像 %s / 画像条目 %s / 字段 %s"
                 % (st.get("documents"), st.get("entities"), st.get("profiles"),
                    st.get("profile_items"), st.get("extractions")))

code, j = getj(u, "/api/db/entities?limit=5")
chk("实体查询（含来源定位）", code == 200 and len(j.get("rows", [])) > 0,
    "HTTP %s, %d 条" % (code, len(j.get("rows", []))))
if j.get("rows"):
    r0 = j["rows"][0]
    keys = ",".join(list(r0.keys())[:8])
    lines.append("       首条字段：%s" % keys)

code, j = getj(u, "/api/db/fields?limit=5")
chk("字段查询（用户指定内容）", code == 200 and len(j.get("rows", [])) > 0,
    "HTTP %s, %d 条" % (code, len(j.get("rows", []))))

# 画像：用真实文档名
code, j = getj(u, "/api/db/entities?limit=1")
docname = (j.get("rows") or [{}])[0].get("doc") or ""
if docname:
    code, j2 = getj(u, "/api/db/profile?doc=" + urllib.parse.quote(docname))
    chk("画像查询", code == 200 and j2.get("found"), "HTTP %s, doc=%s" % (code, docname[:24]))

code, j = getj(u, "/api/db/search?q=" + urllib.parse.quote("合同"))
chk("全库检索", code == 200 and j.get("ok") is not False, "HTTP %s, %d 条"
    % (code, len(j.get("rows", []))))

# ---------- 4 文档库与模块 ----------
lines.append("")
lines.append("=== 4. 文档库与三模块接口 ===")
code, j = getj(u, "/api/status")
chk("/api/status", code == 200 and j.get("service") == "ok",
    "服务=%s 数据库=%s 文档=%s 份" % (j.get("service"), j.get("database"), j.get("docs_count")))

code, j = getj(u, "/api/docs")
docs = j if isinstance(j, list) else j.get("docs", [])
chk("/api/docs 列文档", code == 200 and len(docs) > 0, "HTTP %s, %d 份" % (code, len(docs)))

code, j = getj(u, "/api/fill/templates")
tpls = j if isinstance(j, list) else j.get("templates", [])
chk("模块3 模板列表", code == 200 and len(tpls) >= 5, "HTTP %s, %d 张" % (code, len(tpls)))

code, j = getj(u, "/api/imports")
chk("导入记录", code == 200, "HTTP %s" % code)

code, j = getj(u, "/api/history")
hist = j if isinstance(j, list) else j.get("history", [])
chk("操作历史", code == 200, "HTTP %s, %d 条" % (code, len(hist)))

code, j = getj(u, "/api/db/audit")
chk("模块2 质量体检（零模型审计）", code == 200 and j.get("ok") is not False,
    "HTTP %s, 耗时=%ss" % (code, j.get("elapsed")))

# ---------- 5 语音 ----------
lines.append("")
lines.append("=== 5. 语音（小橘嗓子 + 本机识别） ===")
code, j = getj(u, "/api/digital/status")
chk("数字人状态", code == 200, "HTTP %s" % code)

code, j = getj(u, "/api/asr/status")
chk("识别服务状态", code == 200, "HTTP %s %s" % (code, json.dumps(j, ensure_ascii=False)[:100]))

t0 = time.time()
code, b = u.call("/api/tts?text=" + urllib.parse.quote("你好呀，我是小橘。"), timeout=60)
dt = time.time() - t0
chk("TTS 合成", code == 200 and len(b) > 5000, "HTTP %s, %d 字节, %.2fs" % (code, len(b), dt))

mp3 = None
for p in [os.path.join(r"C:\Users\覃敬涛\Desktop\研一项目\agent\输出\声音候选\A-当前.mp3")]:
    if os.path.exists(p) and os.path.getsize(p) > 3000:
        mp3 = p
if not mp3:
    cdir = r"C:\Users\覃敬涛\Desktop\研一项目\agent\digital\cache"
    if os.path.isdir(cdir):
        fs = sorted((os.path.getsize(os.path.join(cdir, f)), f) for f in os.listdir(cdir)
                    if f.endswith(".mp3"))
        if fs:
            mp3 = os.path.join(cdir, fs[-1][1])
if mp3:
    audio = open(mp3, "rb").read()
    bd = "----a23check"
    body = (("--%s\r\nContent-Disposition: form-data; name=\"file\"; filename=\"t.mp3\"\r\n"
             "Content-Type: audio/mpeg\r\n\r\n" % bd).encode() + audio
            + ("\r\n--%s--\r\n" % bd).encode())
    t0 = time.time()
    code, b = u.call("/api/asr", "POST", raw=body,
                     ctype="multipart/form-data; boundary=" + bd, timeout=180)
    dt = time.time() - t0
    txt = ""
    try:
        txt = json.loads(b.decode("utf-8")).get("text") or ""
    except Exception:
        txt = b[:100].decode("utf-8", "ignore")
    chk("本机识别（离线 whisper）", code == 200 and len(txt) > 0,
        "HTTP %s, %.1fs, 识别：%s" % (code, dt, txt[:50]))
else:
    chk("本机识别（离线 whisper）", False, "无测试音频")

# ---------- 6 客服 + 评价 ----------
lines.append("")
lines.append("=== 6. 客服与评价 ===")
code, j = getj(u, "/api/support/faq")
faq = j if isinstance(j, list) else j.get("faq", [])
chk("FAQ 快捷指令", code == 200 and len(faq) >= 6, "HTTP %s, %d 条" % (code, len(faq)))

code, j = getj(u, "/api/support/start", method="POST", data={"mode": "ai"})
sess = j.get("session") or {}
sid = sess.get("id") or 0
chk("开启 AI 客服会话", code == 200 and sid > 0, "HTTP %s, id=%s, 状态=%s"
    % (code, sid, sess.get("status")))

if sid:
    code, j = getj(u, "/api/support/send", method="POST",
                   data={"session_id": sid, "content": "你好，自检一下"})
    chk("发消息（AI 当场回复）", code == 200, "HTTP %s, 回复=%s"
        % (code, (j.get("message") or {}).get("content", "")[:36]))

    code, j = getj(u, "/api/support/poll?session_id=%d&after=0" % sid)
    msgs = j.get("messages", [])
    chk("轮询消息", code == 200 and len(msgs) >= 1, "HTTP %s, %d 条" % (code, len(msgs)))

code, j = getj(u, "/api/feedback", method="POST",
               data={"feature": "other", "rating": 5, "content": "自检留痕，可忽略", "tags": ""})
chk("提交评价", code == 200 and j.get("ok"), "HTTP %s" % code)

code, j = getj(u, "/api/feedback/mine?limit=5")
chk("我的评价可读", code == 200, "HTTP %s, %d 条" % (code, len(j.get("items", j.get("rows", [])))))

# ---------- 7 管理端 ----------
lines.append("")
lines.append("=== 7. 管理端 ===")
for p, name in [("/api/admin/overview", "总览统计"),
                ("/api/admin/users", "用户列表"),
                ("/api/admin/feedback", "评价反馈"),
                ("/api/admin/sessions", "会话列表")]:
    code, j = getj(adm, p)
    chk(name, code == 200, "HTTP %s" % code)

lines.append("")
lines.append("总计：%d 项通过，%d 项失败" % (ok, bad))
open(OUT, "w", encoding="utf-8").write("\n".join(lines))
print("PASS=%d FAIL=%d" % (ok, bad))
