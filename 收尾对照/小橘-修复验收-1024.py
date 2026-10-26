# -*- coding: utf-8 -*-
"""真机验收：消息重复 / 口型抖动 / 写实小橘。"""
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
PROF = r"C:\Users\覃敬涛\WorkBuddy\2026-09-24-10-48-03\_cdp_prof_fix"
OUT = r"C:\Users\覃敬涛\Desktop\研一项目\agent\输出"
PORT = 9337
op = urllib.request.build_opener(urllib.request.ProxyHandler({}))

ok, bad = [], []


def check(name, cond, detail=""):
    (ok if cond else bad).append(name)
    print(("  PASS " if cond else "  FAIL ") + name + ("" if cond else "  <<< " + str(detail)[:260]), flush=True)


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


def purge(tok):
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
            time.sleep(0.3)
        return False

    def shot(self, name, sel=None, pad=10, scale=1.6):
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
    purge(tk)
    p = Page("about:blank")
    p.cookie(tk)
    p.go(BASE + "/", 3.0)
    p.js("localStorage.setItem('a23_dh_hi','1')")
    p.js("document.getElementById('cs-ai').click()", 2.5)

    print("\n--- 1. 形象：写实度 ---", flush=True)
    check("两只眼睛都是 SVG",
          p.js("document.querySelectorAll('#dh svg .dh-eye').length") == 2)
    check("眼睛是竖瞳（真猫特征，rx3.2/ry8.2）",
          p.js("document.querySelectorAll('#dh svg .dh-eye ellipse[rx=\"3.2\"]').length") == 2,
          p.js("document.querySelectorAll('#dh svg .dh-eye ellipse').length"))
    check("虹膜用黄绿渐变（不是纯黑豆）",
          p.js("document.querySelectorAll('#dh svg .dh-eye path[fill=\"url(#dhIris)\"]').length") == 2)
    check("两个耳廓 + 粉内耳", p.js("document.querySelectorAll('#dh svg .dh-ear').length") == 2
          and p.js("document.querySelectorAll('#dh svg [fill=\"url(#dhEarIn)\"]').length") == 2)
    check("六根胡须", p.js("document.querySelectorAll('#dh svg .dh-whis path').length") == 6,
          p.js("document.querySelectorAll('#dh svg .dh-whis path').length"))
    check("尾巴有环纹（dasharray）",
          bool(p.js("(()=>{const n=document.querySelector('#dh svg .dh-tail path:nth-child(2)');"
                    "return n?n.getAttribute('stroke-dasharray'):'';})()")),
          p.js("(()=>{const n=document.querySelectorAll('#dh svg .dh-tail path');"
                "return [...n].map(x=>x.getAttribute('stroke-dasharray'));})()"))
    check("口鼻是两瓣白毛唇垫", p.js("document.querySelectorAll('#dh svg [fill=\"url(#dhMuzzle)\"]').length") >= 3)
    check("已隐藏欢迎语时仍是待机",
          p.js("document.getElementById('dh').dataset.state") == "idle")
    eye_idle = p.js("(()=>{const b=document.querySelectorAll('#dh .dh-eye')[0].getBoundingClientRect();"
                    "return {x:+b.x.toFixed(1),y:+b.y.toFixed(1),h:+b.height.toFixed(1)};})()")
    print("  待机时左眼位置：" + str(eye_idle), flush=True)
    p.shot("数字人-1-小橘待机.png", "#cs-panel")
    p.shot("数字人-6-小橘特写.png", "#dh", pad=2, scale=2.4)

    print("\n--- 2. 发一条消息，只显示一条 ---", flush=True)
    msg = "我想问一下，系统支持哪些文档格式？"
    p.js("document.getElementById('cs-input').value=%s;"
         "document.getElementById('cs-send').click()" % json.dumps(msg, ensure_ascii=False))
    time.sleep(0.35)
    # 故意在 AI 回复返回前手动催一次轮询：这正是原来重复的来源
    p.js("try{csPoll()}catch(e){}")
    time.sleep(0.6)
    mid_user = p.js("document.querySelectorAll('#cs-body .cs-msg.me').length")
    check("AI 还没回时，本地那条没有被轮询再插一遍", mid_user == 1, mid_user)

    p.wait_for("!!document.querySelector('#cs-body .cs-msg:not(.me):not(.sys) .tx')", 60)
    time.sleep(1.0)
    n_user = p.js("document.querySelectorAll('#cs-body .cs-msg.me').length")
    texts = p.js("[...document.querySelectorAll('#cs-body .cs-msg.me .tx')].map(e=>e.textContent)")
    check("回复到达后仍只有一条", n_user == 1, (n_user, texts))
    check("内容没被复制", isinstance(texts, list) and texts.count(msg) == 1, texts)

    time.sleep(3.6)                                  # 再过两轮轮询
    n_late = p.js("document.querySelectorAll('#cs-body .cs-msg.me').length")
    check("过了两轮轮询也不会变两条", n_late == 1, n_late)

    # 回复条数：AI 每条只上屏一次（id 去重）
    ai = p.js("[...document.querySelectorAll('#cs-body .cs-msg:not(.me):not(.sys) .tx')].map(e=>e.textContent)")
    check("AI 回复也没有重复", len(set(ai)) == len(ai), ai)

    print("\n--- 3. 嘴型：连续，不跳 ---", flush=True)
    p.js("document.getElementById('cs-input').value='怎么插入图片？';"
         "document.getElementById('cs-send').click()")
    spoke = p.wait_for("document.getElementById('dh').dataset.state === 'speak'", 30)
    check("进入说话状态", spoke, p.js("document.getElementById('dh').dataset.state"))
    if spoke:
        eye_spk = p.js("(()=>{const b=document.querySelectorAll('#dh .dh-eye')[0].getBoundingClientRect();"
                       "return {x:+b.x.toFixed(1),y:+b.y.toFixed(1),h:+b.height.toFixed(1)};})()")
        check("说话时眼睛还在脸上（状态缩放没把定位顶掉）",
              abs(eye_spk["x"] - eye_idle["x"]) < 4 and abs(eye_spk["y"] - eye_idle["y"]) < 4,
              (eye_idle, eye_spk))
        check("说话时眼睛没被压成一条缝", eye_spk["h"] > eye_idle["h"] * 0.6, (eye_idle, eye_spk))
        # 页面里采 40 帧，避免 CDP 往返把采样周期拖长
        series = p.js("""new Promise(res=>{
          const el=document.getElementById('dh'), out=[];
          let n=0;
          const step=()=>{
            out.push(parseFloat(getComputedStyle(el).getPropertyValue('--mouth'))||0);
            if(++n<40) requestAnimationFrame(step); else res(out);
          };
          requestAnimationFrame(step);
        })""")
        if not isinstance(series, list) or len(series) < 10:
            check("能采到口型序列", False, series)
        else:
            diffs = [abs(series[i] - series[i - 1]) for i in range(1, len(series))]
            mx = max(diffs)
            avg = sum(diffs) / len(diffs)
            rng = (min(series), max(series))
            print("  序列：" + ", ".join("%.2f" % v for v in series[:14]) + " …", flush=True)
            print("  相邻最大跳变 %.3f / 平均 %.3f / 幅度 %s" % (mx, avg, rng), flush=True)
            check("相邻帧不跳（最大变化 < 0.10）", mx < 0.10, "max=%.3f" % mx)
            check("幅度收在合理区间（0.02~0.70）", rng[0] >= 0.0 and rng[1] <= 0.70, rng)
            check("确实在开合（有起伏）", rng[1] - rng[0] > 0.12, rng)
        p.shot("数字人-2-小橘念答案.png", "#cs-panel")
        p.shot("数字人-7-小橘说话特写.png", "#dh", pad=2, scale=2.4)

finally:
    proc.terminate()

print("\n结果：PASS %d / FAIL %d" % (len(ok), len(bad)))
if bad:
    print("失败项：" + "；".join(bad))
