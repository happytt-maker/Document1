# -*- coding: utf-8 -*-
"""验收：新小橘形象 + 新声音。"""
import base64
import hashlib
import json
import os
import shutil
import subprocess
import time
import urllib.request

from websockets.sync.client import connect

BASE = "http://127.0.0.1:8000"
EDGE = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
PROF = r"C:\Users\覃敬涛\WorkBuddy\2026-09-24-10-48-03\_cdp_cat_verify"
OUT = r"C:\Users\覃敬涛\Desktop\研一项目\agent\输出"
VOICE_REF = r"C:\Users\覃敬涛\Desktop\研一项目\agent\输出\声音候选\B-奶音.mp3"
PORT = 9341
op = urllib.request.build_opener(urllib.request.ProxyHandler({}))
ok, bad = [], []


def check(name, cond, detail=""):
    (ok if cond else bad).append(name)
    print(("  PASS " if cond else "  FAIL ") + name + ("" if cond else "  <<< " + str(detail)[:240]), flush=True)


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
            time.sleep(0.3)
        return False

    def shot(self, name, sel=None, pad=10, scale=1.8):
        if sel:
            r = self.js("(()=>{const b=document.querySelector('%s').getBoundingClientRect();"
                        "return {x:b.x,y:b.y,w:b.width,h:b.height};})()" % sel)
            if not r or r["w"] < 2:
                raise RuntimeError("元素没尺寸：" + sel)
            params = {"format": "png", "clip": {"x": max(r["x"] - pad, 0), "y": max(r["y"] - pad, 0),
                                               "width": r["w"] + pad * 2, "height": r["h"] + pad * 2,
                                               "scale": scale}}
        else:
            params = {"format": "png"}
        d = self.cmd("Page.captureScreenshot", params)
        with open(os.path.join(OUT, name), "wb") as f:
            f.write(base64.b64decode(d["data"]))
        print("  截图 -> " + name, flush=True)


if os.path.isdir(PROF):
    shutil.rmtree(PROF, ignore_errors=True)
os.makedirs(OUT, exist_ok=True)

proc = subprocess.Popen([
    EDGE, "--headless=new", "--disable-gpu", "--no-first-run", "--no-default-browser-check",
    "--disable-extensions", "--disable-sync", "--disable-background-networking",
    "--disable-component-update", "--no-proxy-server", "--remote-allow-origins=*",
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
    p = Page("about:blank")
    p.cookie(tk)
    p.go(BASE + "/", 3.0)
    p.js("localStorage.setItem('a23_dh_hi','1')")
    p.js("document.getElementById('cs-ai').click()", 2.5)

    print("\n--- 1. 新形象是否完整渲染 ---", flush=True)
    check("两只眼睛", p.js("document.querySelectorAll('#dh svg .dh-eye').length") == 2)
    check("两只耳朵", p.js("document.querySelectorAll('#dh svg .dh-ear').length") == 2)
    check("虹膜用渐变（cgIris）",
          p.js("document.querySelectorAll('#dh svg [fill=\"url(#cgIris)\"]').length") == 2,
          p.js("document.querySelectorAll('#dh svg [fill^=\"url(\"]').length"))
    check("白口鼻区存在", p.js("document.querySelectorAll('#dh svg [fill=\"url(#cgMuz)\"]').length") >= 2)
    check("四根胡须", p.js("document.querySelectorAll('#dh svg .dh-whis path').length") == 4,
          p.js("document.querySelectorAll('#dh svg .dh-whis path').length"))
    check("尾巴环纹", bool(p.js("(()=>{const n=document.querySelectorAll('#dh svg .dh-tail path')[1];"
                                "return n?n.getAttribute('stroke-dasharray'):'';})()")))
    check("嘴默认关闭（不露黑洞）",
          float(p.js("getComputedStyle(document.querySelector('#dh svg .dh-mouth')).opacity")) == 0)
    check("微笑唇线可见",
          float(p.js("getComputedStyle(document.querySelector('#dh svg .dh-mouthline')).opacity")) > 0.5)

    geo = p.js("""(()=>{
      const head=document.querySelector('#dh svg .dh-head')||document.querySelector('#dh svg path');
      const hb=head.getBoundingClientRect();
      const eyes=[...document.querySelectorAll('#dh .dh-eye')].map(e=>e.getBoundingClientRect());
      return {head:{x:+hb.x.toFixed(1),y:+hb.y.toFixed(1),w:+hb.width.toFixed(1),h:+hb.height.toFixed(1)},
              eyes:eyes.map(b=>({x:+b.x.toFixed(1),y:+b.y.toFixed(1),w:+b.width.toFixed(1),h:+b.height.toFixed(1)}))};
    })()""")
    print("  几何：" + json.dumps(geo, ensure_ascii=False), flush=True)
    hb = geo["head"]
    inside = all(e["x"] >= hb["x"] - 2 and e["x"] + e["w"] <= hb["x"] + hb["w"] + 2 and
                 e["y"] >= hb["y"] - 2 and e["y"] + e["h"] <= hb["y"] + hb["h"] + 2 for e in geo["eyes"])
    check("两只眼睛都在脸内（没跑出去）", inside, geo)
    check("两只眼睛大小一致",
          abs(geo["eyes"][0]["w"] - geo["eyes"][1]["w"]) < 1.5 and
          abs(geo["eyes"][0]["h"] - geo["eyes"][1]["h"]) < 1.5, geo["eyes"])
    check("猫头够大（头宽占舞台宽 > 60%）",
          hb["w"] / p.js("document.querySelector('#dh').getBoundingClientRect().width") > 0.6,
          hb["w"])

    p.shot("小橘新版-1-待机.png", "#cs-panel")
    p.shot("小橘新版-2-特写.png", "#dh", pad=2, scale=3.0)

    print("\n--- 2. 说话时张口（且嘴在上颚处张开） ---", flush=True)
    p.js("dhSay('你好呀，我是小橘。今天有什么想问的直接说就行。')")
    p.wait_for("document.getElementById('dh').dataset.state==='speak'", 12)
    time.sleep(0.55)
    check("进入说话状态", p.js("document.getElementById('dh').dataset.state") == "speak",
          p.js("document.getElementById('dh').dataset.state"))
    m1 = float(p.js("parseFloat(getComputedStyle(document.getElementById('dh')).getPropertyValue('--mouth'))||0"))
    p.shot("小橘新版-3-说话特写.png", "#dh", pad=2, scale=3.0)
    # 在页面里按真实帧率采样（RPC 取样间隔太大，会把平滑曲线看成抖动）
    vals = p.js("""new Promise(res=>{
      const el=document.getElementById('dh'), out=[];
      const mo=new MutationObserver(()=>{ out.push(parseFloat(el.style.getPropertyValue('--mouth'))||0); });
      mo.observe(el,{attributes:true,attributeFilter:['style']});
      setTimeout(()=>{ mo.disconnect(); res(out); }, 900);
    })""")
    jumps = [abs(vals[i] - vals[i - 1]) for i in range(1, len(vals))] or [0]
    print("  采样 %d 帧，mouth 范围 %.2f~%.2f" % (len(vals), min(vals), max(vals)), flush=True)
    check("口型有起伏（不是死的）", max(vals) - min(vals) > 0.08, (min(vals), max(vals)))
    check("相邻帧不硬跳（<0.08/帧）", max(jumps) < 0.08, round(max(jumps), 3))
    p.wait_for("document.getElementById('dh').dataset.state!=='speak'", 20)
    time.sleep(0.6)
    check("说完回到待机", p.js("document.getElementById('dh').dataset.state") == "idle",
          p.js("document.getElementById('dh').dataset.state"))
    p.shot("小橘新版-4-说完.png", "#dh", pad=2, scale=3.0)

    print("\n--- 3. 声音换了吗 ---", flush=True)
    text = "你好呀，我是小橘。今天有什么可以帮你的吗？"
    req = urllib.request.Request(BASE + "/api/tts?text=" + urllib.request.quote(text),
                                 headers={"Cookie": "a23_token=" + tk})
    r = op.open(req, timeout=40)
    data = r.read()
    check("/api/tts 返回音频", r.status == 200 and len(data) > 8000, (r.status, len(data)))
    ref = open(VOICE_REF, "rb").read() if os.path.exists(VOICE_REF) else b""
    same = hashlib.md5(data).hexdigest() == hashlib.md5(ref).hexdigest() if ref else False
    check("线上音频 = 参数 B（奶音）那一版", same,
          "线上 %d B / 参照 %d B" % (len(data), len(ref)))

    print("\n===== 通过 %d 项，失败 %d 项 =====" % (len(ok), len(bad)), flush=True)
    if bad:
        print("失败：" + " ｜ ".join(bad), flush=True)
finally:
    try:
        proc.terminate()
    except Exception:
        pass
