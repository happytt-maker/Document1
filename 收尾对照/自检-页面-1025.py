# -*- coding: utf-8 -*-
"""页面健康检查：能否正常渲染 / 有无 JS 报错 / 关键控件在不在"""
import base64
import json
import os
import shutil
import subprocess
import time
import urllib.request

from websockets.sync.client import connect

BASE = "http://127.0.0.1:8000"
EDGE = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
PROF = r"C:\Users\覃敬涛\WorkBuddy\2026-09-24-10-48-03\_cdp_health_%d" % int(time.time())
OUT = r"C:\Users\覃敬涛\Desktop\研一项目\agent\输出"
PORT = 9351
op = urllib.request.build_opener(urllib.request.ProxyHandler({}))
lines = []
ok = 0
bad = 0


def chk(name, cond, detail=""):
    global ok, bad
    if cond:
        ok += 1
        lines.append("  [OK]   %-36s %s" % (name, detail))
    else:
        bad += 1
        lines.append("  [FAIL] %-36s %s" % (name, str(detail)[:220]))


def login(u, p, portal):
    req = urllib.request.Request(
        BASE + "/api/auth/login",
        data=json.dumps({"username": u, "password": p, "portal": portal}).encode(),
        headers={"Content-Type": "application/json"}, method="POST")
    r = op.open(req, timeout=10)
    for part in (r.headers.get("Set-Cookie") or "").split(";"):
        part = part.strip()
        if part.startswith("a23_token="):
            return part.split("=", 1)[1]
    raise RuntimeError("没拿到 cookie")


def http(path):
    return json.load(op.open("http://127.0.0.1:%d%s" % (PORT, path), timeout=5))


class Page:
    def __init__(self, url):
        tgt = json.load(op.open(urllib.request.Request(
            "http://127.0.0.1:%d/json/new?%s" % (PORT, url), method="PUT"), timeout=5))
        self.ws = connect(tgt["webSocketDebuggerUrl"], max_size=None, open_timeout=10,
                          legacy=True)
        self.i = 0
        self.errors = []
        self.cmd("Page.enable")
        self.cmd("Runtime.enable")
        self.cmd("Network.enable")

    def cmd(self, method, params=None):
        self.i += 1
        mid = self.i
        self.ws.send(json.dumps({"id": mid, "method": method, "params": params or {}}))
        while True:
            m = json.loads(self.ws.recv())
            meth = m.get("method")
            if meth == "Runtime.exceptionThrown":
                d = m["params"].get("exceptionDetails", {})
                self.errors.append("EXC: " + str(d.get("text")) + " "
                                   + str((d.get("exception") or {}).get("description", ""))[:160])
            elif meth == "Runtime.consoleAPICalled":
                if m["params"].get("type") == "error":
                    args = m["params"].get("args", [])
                    txt = " ".join(str(a.get("value", a.get("description", ""))) for a in args)
                    self.errors.append("CONSOLE: " + txt[:180])
            elif meth == "Page.javascriptDialogOpening":
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
        """导航并真正等到新页面加载完（否则读到的是上一页的 DOM）。"""
        try:
            before = self.js("document.location.href")
        except Exception:
            before = None
        self.cmd("Page.navigate", {"url": url})
        t0 = time.time()
        while time.time() - t0 < 15:
            try:
                href = self.js("document.location.href") or ""
                ready = self.js("document.readyState")
                if href and href != before and ready == "complete":
                    break
            except Exception:
                pass
            time.sleep(0.25)
        time.sleep(wait)

    def js(self, expr):
        r = self.cmd("Runtime.evaluate", {"expression": expr, "returnByValue": True,
                                          "awaitPromise": True})
        if r.get("exceptionDetails"):
            raise RuntimeError(json.dumps(r["exceptionDetails"])[:260])
        return (r.get("result") or {}).get("value")

    def wait_for(self, expr, timeout=15.0):
        t0 = time.time()
        while time.time() - t0 < timeout:
            try:
                if self.js(expr):
                    return True
            except Exception:
                pass
            time.sleep(0.3)
        return False

    def shot(self, name):
        d = self.cmd("Page.captureScreenshot", {"format": "png"})
        with open(os.path.join(OUT, name), "wb") as f:
            f.write(base64.b64decode(d["data"]))


if os.path.isdir(PROF):
    pass          # 每次用新的唯一目录，跑完由外部清理（避免批量删除被安全策略拦）
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
            http("/json/version")
            break
        except Exception:
            time.sleep(0.5)
    else:
        raise RuntimeError("Edge 调试端口没起来")

    tk = login("demo_ai", "123456", "user")
    p = Page("about:blank")
    p.cookie(tk)

    lines.append("=== 1. 主页面加载 ===")
    p.go(BASE + "/", 4.0)
    chk("页面标题正常", "A23" in str(p.js("document.title")) or len(str(p.js("document.title"))) > 0,
        p.js("document.title"))
    chk("body 有内容（不是白屏）", (p.js("document.body.innerText.length") or 0) > 500,
        "可见文字 %s 字" % p.js("document.body.innerText.length"))

    lines.append("")
    lines.append("=== 2. 关键控件 ===")
    for expr, name in [
        ("!!document.getElementById('dh')", "数字人小橘容器"),
        ("document.querySelectorAll('#dh svg *').length > 60", "小橘 SVG 已渲染"),
        ("!!document.querySelector('textarea')", "指令输入框"),
        ("document.querySelectorAll('button').length > 8", "操作按钮"),
        ("!!document.getElementById('voice-btn')", "语音输入按钮"),
    ]:
        chk(name, bool(p.js(expr)), p.js(expr))

    lines.append("")
    lines.append("=== 3. 小橘是否真在动 ===")
    p.js("document.getElementById('cs-ai') && document.getElementById('cs-ai').click()")
    time.sleep(2.5)
    chk("客服面板可打开", bool(p.js(
        "(()=>{const e=document.querySelector('#cs-panel,#csPanel,.cs-panel');"
        "if(!e) return false;const r=e.getBoundingClientRect();return r.width>100;})()")))
    chk("小橘 SVG 有几何尺寸", bool(p.js(
        "(()=>{const s=document.querySelector('#dh svg');if(!s)return false;"
        "const r=s.getBoundingClientRect();return r.width>50&&r.height>50;})()")),
        p.js("(()=>{const s=document.querySelector('#dh svg');if(!s)return null;"
             "const r=s.getBoundingClientRect();return Math.round(r.width)+'x'+Math.round(r.height);})()"))

    lines.append("")
    lines.append("=== 4. 运行时错误 ===")
    real_errs = [e for e in p.errors
                 if "favicon" not in e.lower()
                 and "Failed to load resource" not in e
                 and "ERR_" not in e]
    chk("没有 JS 未捕获异常 / console 错误", len(real_errs) == 0,
        "%d 条" % len(real_errs))
    for e in real_errs[:6]:
        lines.append("       " + e[:200])

    p.shot("运行检查-主页.png")
    lines.append("       截图 -> agent\\输出\\运行检查-主页.png")

    lines.append("")
    lines.append("=== 5. 三个端口的页面（用全角色账号） ===")
    tk_all = login("2482991994", "123456", "user")
    p.cookie(tk_all)
    for path, name, minsize in [("/staff", "客服端", 50),
                                ("/admin", "管理端", 200)]:
        p.errors = []
        p.go(BASE + path, 2.5)
        url = str(p.js("document.location.href"))
        txt = str(p.js("document.body.innerText") or "")
        errs = [e for e in p.errors if "favicon" not in e.lower()
                and "Failed to load resource" not in e and "ERR_" not in e]
        right = path in url
        chk(name + " 打开的是正确页面", right, "URL=%s" % url.replace(BASE, ""))
        chk(name + " 渲染正常", len(txt) >= minsize and len(errs) == 0,
            "%d 字, 首行：%s, 错误 %d 条" % (len(txt), txt.strip().split("\n")[0][:24], len(errs)))
        if path == "/admin":
            p.shot("运行检查-管理端.png")
        if path == "/staff":
            p.shot("运行检查-客服端.png")

    # 登录页必须用「没登录」的干净会话测（已登录会被 302 掉）
    p2 = Page("about:blank")
    p2.cmd("Network.clearBrowserCookies")      # 清掉共享的登录态
    p2.go(BASE + "/login", 2.5)
    url2 = str(p2.js("document.location.href"))
    txt2 = str(p2.js("document.body.innerText") or "")
    errs2 = [e for e in p2.errors if "favicon" not in e.lower()
             and "Failed to load resource" not in e and "ERR_" not in e]
    chk("登录页 打开的是正确页面（未登录）", "/login" in url2, "URL=%s" % url2.replace(BASE, ""))
    chk("登录页 渲染正常", len(txt2) > 20 and len(errs2) == 0,
        "%d 字, 含端口选择=%s, 错误 %d 条"
        % (len(txt2), bool(p2.js("document.body.innerText.indexOf('用户')>=0")), len(errs2)))
    p2.shot("运行检查-登录页.png")

finally:
    try:
        proc.terminate()
    except Exception:
        pass
    time.sleep(1)

lines.append("")
lines.append("总计：%d 项通过，%d 项失败" % (ok, bad))
open(r"C:\Users\覃敬涛\WorkBuddy\2026-09-24-10-48-03\_page.txt", "w", encoding="utf-8").write("\n".join(lines))
print("PASS=%d FAIL=%d" % (ok, bad))
