# -*- coding: utf-8 -*-
"""真机验收：数字人「小橘」（Edge headless + CDP）。

验的是四件事：
1. 形象在客服浮层里正常显示（纯 SVG，无外部素材）
2. AI 回答时小橘进入 speak、嘴在动、音频真的在播
3. 点小橘自己会打招呼；点静音就闭嘴
4. 转人工后小橘退场（不插话）
"""
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
PROF = r"C:\Users\覃敬涛\WorkBuddy\2026-09-24-10-48-03\_cdp_prof_dh"
OUT = r"C:\Users\覃敬涛\Desktop\研一项目\agent\输出"
PORT = 9336
op = urllib.request.build_opener(urllib.request.ProxyHandler({}))

ok, bad = [], []


def check(name, cond, detail=""):
    (ok if cond else bad).append(name)
    print(("  PASS " if cond else "  FAIL ") + name + ("" if cond else "  <<< " + str(detail)[:220]), flush=True)


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
    raise RuntimeError("没拿到 cookie")


def http(path):
    return json.load(op.open("http://127.0.0.1:%d%s" % (PORT, path), timeout=5))


def purge(tok):
    """把该用户进行中的会话结束掉，保证这次开的是全新的 AI 会话。"""
    def call(path, data=None):
        req = urllib.request.Request(
            BASE + path, None if data is None else json.dumps(data).encode(),
            {"Content-Type": "application/json", "Cookie": "a23_token=" + tok},
            method="GET" if data is None else "POST")
        return json.load(op.open(req, timeout=10))
    for _ in range(8):
        s = call("/api/support/mine").get("session")
        if not s:
            return
        call("/api/support/close", {"session_id": s["id"]})


class Page:
    def __init__(self, url):
        tgt = json.load(op.open(urllib.request.Request(
            "http://127.0.0.1:%d/json/new?%s" % (PORT, url), method="PUT"), timeout=5))
        self.ws = connect(tgt["webSocketDebuggerUrl"], max_size=None, open_timeout=10)
        self.i = 0
        self.cmd("Page.enable")
        self.cmd("Runtime.enable")
        self.cmd("Network.enable")

    def cmd(self, method, params=None):
        self.i += 1
        mid = self.i
        self.ws.send(json.dumps({"id": mid, "method": method, "params": params or {}}))
        while True:
            m = json.loads(self.ws.recv())
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
        self.cmd("Page.navigate", {"url": url})
        time.sleep(wait)

    def js(self, expr, wait=0.0):
        r = self.cmd("Runtime.evaluate", {"expression": expr, "returnByValue": True, "awaitPromise": True})
        if wait:
            time.sleep(wait)
        if r.get("exceptionDetails"):
            raise RuntimeError(json.dumps(r["exceptionDetails"])[:300])
        return (r.get("result") or {}).get("value")

    def wait_for(self, expr, timeout=15.0):
        t0 = time.time()
        while time.time() - t0 < timeout:
            try:
                if self.js(expr):
                    return True
            except Exception:
                pass
            time.sleep(0.35)
        return False

    def shot(self, name, sel=None, pad=10):
        if sel:
            r = self.js("(()=>{const b=document.querySelector('%s').getBoundingClientRect();"
                        "return {x:b.x,y:b.y,w:b.width,h:b.height};})()" % sel)
            if not r or r["w"] < 2:
                raise RuntimeError("元素没尺寸：" + sel)
            params = {"format": "png", "clip": {"x": max(r["x"] - pad, 0), "y": max(r["y"] - pad, 0),
                                                "width": r["w"] + pad * 2, "height": r["h"] + pad * 2,
                                                "scale": 1.6}}
        else:
            params = {"format": "png"}
        d = self.cmd("Page.captureScreenshot", params)
        path = os.path.join(OUT, name)
        with open(path, "wb") as f:
            f.write(base64.b64decode(d["data"]))
        print("  截图 -> " + name, flush=True)
        return path


if os.path.isdir(PROF):
    shutil.rmtree(PROF, ignore_errors=True)
os.makedirs(OUT, exist_ok=True)

proc = subprocess.Popen([
    EDGE, "--headless=new", "--disable-gpu", "--no-first-run", "--no-default-browser-check",
    "--disable-extensions", "--disable-sync", "--disable-background-networking",
    "--disable-component-update", "--no-proxy-server", "--remote-allow-origins=*",
    # 无头环境没有真实点击手势，放开自动播放才能验到「出声」；真实浏览器里点按钮本身就是手势
    "--autoplay-policy=no-user-gesture-required",
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
    purge(tk)
    p = Page("about:blank")
    p.cookie(tk)
    p.go(BASE + "/", 3.0)
    p.js("localStorage.setItem('a23_dh_hi','1')")     # 先不放欢迎语，拍一张干净的待机
    p.js("document.getElementById('cs-ai').click()", 2.5)

    check("AI 客服面板打开", p.js("document.getElementById('cs-panel').classList.contains('on')"))
    check("数字人舞台可见", p.js("getComputedStyle(document.getElementById('cs-stage')).display !== 'none'"))
    check("小橘是 SVG 画出来的（无外部图片）",
          p.js("document.querySelectorAll('#dh svg .dh-eye').length") == 2
          and p.js("document.querySelectorAll('#dh svg .dh-blink').length") == 2,
          p.js("document.querySelectorAll('#dh svg *').length"))
    check("初始状态是待机",
          p.js("document.getElementById('dh').dataset.state") == "idle",
          p.js("document.getElementById('dh').dataset.state"))
    check("状态文案写着「随便问」", "随便问" in (p.js("document.getElementById('dh-state-tx').textContent") or ""))
    check("面板变成左右分栏",
          p.js("Math.round(document.getElementById('cs-panel').getBoundingClientRect().width)") >= 600,
          p.js("document.getElementById('cs-panel').getBoundingClientRect().width"))
    p.shot("数字人-1-小橘待机.png", "#cs-panel")

    # ---------- 点快捷指令 → 小橘念答案 ----------
    p.js("document.querySelectorAll('#cs-faq-list .cs-fq')[2].click()")
    spoke = p.wait_for("document.getElementById('dh').dataset.state === 'speak'", 25)
    check("AI 回复后小橘进入「说话」状态", spoke,
          p.js("document.getElementById('dh').dataset.state"))
    if spoke:
        p.shot("数字人-2-小橘念答案.png", "#cs-panel")
        played = p.wait_for("(()=>{const a=window.__dh.audio;"
                            "return !!(a && !a.paused && a.currentTime > 0.1);})()", 5)
        audio = p.js("(()=>{const a=window.__dh.audio;return a?{paused:a.paused,t:+a.currentTime.toFixed(2),"
                     "dur:+(a.duration||0).toFixed(2),src:a.src.slice(0,40)}:null;})()")
        check("音频真的在播（不是只动嘴）", played, audio)
        print("  音频：" + str(audio), flush=True)
        m1 = p.js("getComputedStyle(document.getElementById('dh')).getPropertyValue('--mouth')")
        time.sleep(0.4)
        m2 = p.js("getComputedStyle(document.getElementById('dh')).getPropertyValue('--mouth')")
        check("嘴型幅度在变（口型跟着节奏动）", m1 != m2, (m1, m2))
        check("说话时嘴里是张开的（不是一条线）",
              p.js("getComputedStyle(document.querySelector('.dh-mouth')).opacity") == "1")

    # ---------- 静音 → 立刻闭嘴 ----------
    p.js("document.getElementById('dh-mute').click()", 0.6)
    check("静音按钮变成「已静音」", "已静音" in (p.js("document.getElementById('dh-mute').textContent") or ""))
    check("静音后立即停止播放", p.js("!window.__dh.audio || window.__dh.audio.paused") is True)
    check("静音后回到待机", p.js("document.getElementById('dh').dataset.state") == "idle")
    p.js("document.getElementById('dh-mute').click()", 0.4)     # 恢复出声

    # ---------- 点小橘 → 主动打招呼 ----------
    p.js("document.getElementById('dh').click()", 0.3)
    check("被点后先竖耳朵听（listen）",
          p.js("document.getElementById('dh').dataset.state") == "listen"
          or p.js("document.getElementById('dh').dataset.state") == "think",
          p.js("document.getElementById('dh').dataset.state"))
    greet = p.wait_for("document.getElementById('dh').dataset.state === 'speak'", 20)
    check("点一下会自己开口打招呼", greet, p.js("document.getElementById('dh').dataset.state"))
    if greet:
        p.shot("数字人-3-点一下打招呼.png", "#cs-panel")
    p.wait_for("document.getElementById('dh').dataset.state === 'idle'", 30)

    # ---------- 转人工 → 小橘退场 ----------
    p.js("document.getElementById('cs-x').click()", 0.8)
    p.js("document.getElementById('cs-human').click()", 3.0)
    check("转人工后小橘改成「不插嘴」",
          "不插嘴" in (p.js("document.getElementById('dh-state-tx').textContent") or ""),
          p.js("document.getElementById('dh-state-tx').textContent"))
    check("转人工后不说话（state=idle）",
          p.js("document.getElementById('dh').dataset.state") == "idle")
    p.shot("数字人-4-转人工后小橘退场.png", "#cs-panel")

    # ---------- 窄屏：自动收起数字人，不挡对话 ----------
    p.cmd("Emulation.setDeviceMetricsOverride", {"width": 700, "height": 900, "deviceScaleFactor": 1, "mobile": False})
    time.sleep(0.8)
    check("窄屏自动隐藏数字人",
          p.js("getComputedStyle(document.getElementById('cs-stage')).display === 'none'"))
    check("窄屏面板贴满视口（不再固定 648 宽）",
          p.js("document.getElementById('cs-panel').getBoundingClientRect().width") <= 700,
          p.js("document.getElementById('cs-panel').getBoundingClientRect().width"))
    p.shot("数字人-5-窄屏收起.png", "#cs-panel")
    p.cmd("Emulation.clearDeviceMetricsOverride")

finally:
    proc.terminate()

print("\n结果：PASS %d / FAIL %d" % (len(ok), len(bad)))
if bad:
    print("失败项：" + "；".join(bad))
