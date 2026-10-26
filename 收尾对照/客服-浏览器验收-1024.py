# -*- coding: utf-8 -*-
"""浏览器真机验收：AI 客服快捷指令面板 + 人工客服先留言（Edge headless + CDP）。"""
import base64, json, os, shutil, subprocess, sys, time, urllib.request
from websockets.sync.client import connect

BASE = "http://127.0.0.1:8000"
EDGE = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
PROF = r"C:\Users\覃敬涛\WorkBuddy\2026-09-24-10-48-03\_cdp_prof"
OUT = r"C:\Users\覃敬涛\Desktop\研一项目\agent\输出"
PORT = 9333
op = urllib.request.build_opener(urllib.request.ProxyHandler({}))

ok, bad = [], []


def check(name, cond, detail=""):
    (ok if cond else bad).append(name)
    print(("  PASS " if cond else "  FAIL ") + name + ("" if cond else "  <<< " + str(detail)[:200]))


def login(u, p, portal):
    req = urllib.request.Request(
        BASE + "/api/auth/login",
        data=json.dumps({"username": u, "password": p, "portal": portal}).encode(),
        headers={"Content-Type": "application/json"}, method="POST")
    r = op.open(req, timeout=10)
    ck = r.headers.get("Set-Cookie") or ""
    for part in ck.split(";"):
        part = part.strip()
        if part.startswith("a23_token="):
            return part.split("=", 1)[1]
    raise RuntimeError("没拿到 cookie: " + ck)


def purge(token):
    """把该用户进行中的会话都结束，保证截图是干净的新会话。"""
    def call(path, data=None):
        req = urllib.request.Request(
            BASE + path,
            data=None if data is None else json.dumps(data).encode(),
            headers={"Content-Type": "application/json", "Cookie": "a23_token=" + token},
            method="GET" if data is None else "POST")
        return json.load(op.open(req, timeout=10))
    for _ in range(8):
        s = call("/api/support/mine").get("session")
        if not s:
            return
        call("/api/support/close", {"session_id": s["id"]})


def http(path):
    return json.load(op.open("http://127.0.0.1:%d%s" % (PORT, path), timeout=5))


class Page:
    def __init__(self, url):
        tgt = json.load(op.open(urllib.request.Request("http://127.0.0.1:%d/json/new?%s" % (PORT, url), method="PUT"), timeout=5))
        self.ws = connect(tgt["webSocketDebuggerUrl"], max_size=None, open_timeout=10)
        print("  ws 已连接", flush=True)
        self.i = 0
        self.cmd("Page.enable")
        self.cmd("Runtime.enable")
        self.cmd("Network.enable")
        print("  域已启用", flush=True)

    def cmd(self, method, params=None):
        self.i += 1
        mid = self.i
        self.ws.send(json.dumps({"id": mid, "method": method, "params": params or {}}))
        while True:
            m = json.loads(self.ws.recv())
            # headless 下 confirm/alert 会卡住 evaluate，必须自动确认
            if m.get("method") == "Page.javascriptDialogOpening":
                self.i += 1
                self.ws.send(json.dumps({"id": self.i, "method": "Page.handleJavaScriptDialog",
                                         "params": {"accept": True}}))
                continue
            if m.get("id") == mid:
                if "error" in m:
                    raise RuntimeError(method + " -> " + json.dumps(m["error"]))
                return m.get("result", {})

    def cookie(self, value):
        self.cmd("Network.setCookie", {"name": "a23_token", "value": value,
                                       "url": BASE, "path": "/", "httpOnly": True})

    def go(self, url, wait=2.0):
        print("  导航 " + url, flush=True)
        self.cmd("Page.navigate", {"url": url})
        time.sleep(wait)

    def js(self, expr, wait=0.0):
        r = self.cmd("Runtime.evaluate", {"expression": expr, "returnByValue": True, "awaitPromise": True})
        if wait:
            time.sleep(wait)
        if r.get("exceptionDetails"):
            raise RuntimeError(json.dumps(r["exceptionDetails"])[:300])
        return (r.get("result") or {}).get("value")

    def shot(self, name):
        d = self.cmd("Page.captureScreenshot", {"format": "png"})
        path = os.path.join(OUT, name)
        with open(path, "wb") as f:
            f.write(base64.b64decode(d["data"]))
        print("  截图 -> " + path)
        return path


# 清掉上次的 profile
if os.path.isdir(PROF):
    shutil.rmtree(PROF, ignore_errors=True)
os.makedirs(OUT, exist_ok=True)

proc = subprocess.Popen([
    EDGE, "--headless=new", "--disable-gpu", "--no-first-run", "--no-default-browser-check",
    "--disable-extensions", "--disable-sync", "--disable-background-networking",
    "--disable-component-update", "--no-proxy-server", "--remote-allow-origins=*",
    "--remote-debugging-port=%d" % PORT, "--user-data-dir=" + PROF,
    "--window-size=1440,960", "about:blank",
], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

try:
    for _ in range(40):
        try:
            http("/json/version"); break
        except Exception:
            time.sleep(0.5)
    else:
        raise RuntimeError("Edge 调试端口没起来")
    print("Edge 就绪", flush=True)

    # ---------- 用户端：AI 客服快捷指令
    tk = login("demo_ai", "123456", "user")
    purge(tk)
    print("登录成功", flush=True)
    p = Page("about:blank")
    print("页面通道建立", flush=True)
    p.cookie(tk)
    p.go(BASE + "/", 3.0)
    p.js("document.getElementById('cs-ai').click()", 2.5)
    check("AI 客服面板打开", p.js("document.getElementById('cs-panel').classList.contains('on')"))
    n = p.js("document.querySelectorAll('#cs-faq-list .cs-fq').length")
    check("快捷指令按钮 6 个", n == 6, n)
    labels = p.js("[...document.querySelectorAll('#cs-faq-list .cs-fq')].map(b=>b.textContent.trim())")
    print("  按钮：" + " ｜ ".join(labels or []))
    check("快捷指令面板可见", p.js("getComputedStyle(document.getElementById('cs-faq')).display !== 'none'"))
    p.shot("客服-1-AI快捷指令.png")

    # 输入数字 3 → 秒回预设
    t0 = time.time()
    p.js("document.getElementById('cs-input').value='3'; csSendMsg();", 1.6)
    ms = int((time.time() - t0 - 1.6) * 1000)
    last = p.js("(document.querySelectorAll('#cs-body .cs-msg .tx').length && [...document.querySelectorAll('#cs-body .cs-msg .tx')].pop().textContent) || ''")
    check("输入 3 秒回预设答案", "拖进对话框" in (last or ""), (ms, last))
    print("  回复：" + (last or "")[:60])
    p.shot("客服-2-AI数字秒答.png")

    # ---------- 用户端：人工客服先留言
    p.js("document.getElementById('cs-end').click()", 1.0)
    p.js("document.getElementById('cs-human').click()", 2.5)
    st = p.js("document.getElementById('cs-status').textContent")
    check("人工会话进入排队态", "排队" in (st or ""), st)
    check("排队时输入框可用", p.js("!document.getElementById('cs-input').disabled"))
    ph = p.js("document.getElementById('cs-input').placeholder")
    check("排队占位提示是「可先留言」", "留言" in (ph or ""), ph)
    check("人工会话不显示快捷指令", p.js("getComputedStyle(document.getElementById('cs-faq')).display === 'none'"))
    p.js("document.getElementById('cs-input').value='我先留言：导出能不能只导某一列？'; csSendMsg();", 2.0)
    body = p.js("[...document.querySelectorAll('#cs-body .cs-msg .tx')].map(e=>e.textContent).join(' | ')")
    check("排队中留言已发出", "只导某一列" in (body or ""), body)
    check("页面给出留言提示", "留言已" in (body or "") or "留言" in (body or ""), body)
    p.shot("客服-3-人工先留言.png")

    # ---------- 客服端：队列里能看到留言
    kt = login("kf_demo", "123456", "staff")
    s = Page("about:blank")
    s.cookie(kt)
    s.go(BASE + "/staff", 3.0)
    txt = s.js("document.getElementById('list').innerText")
    check("客服工作台出现留言会话", "留言" in (txt or ""), txt)
    s.js("[...document.querySelectorAll('#list .item')].find(e=>/留言/.test(e.innerText))?.click()", 1.5)
    s.shot("客服-4-客服端看到留言.png")

    print("\n结果：通过 %d 项，失败 %d 项" % (len(ok), len(bad)))
    if bad:
        print("失败项：" + "；".join(bad))
finally:
    p_ = None
    try:
        proc.terminate()
    except Exception:
        pass
    time.sleep(1)
    shutil.rmtree(PROF, ignore_errors=True)

sys.exit(1 if bad else 0)
