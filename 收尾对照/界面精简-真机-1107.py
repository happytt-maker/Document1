# -*- coding: utf-8 -*-
"""界面精简 + 浮层拖动缩放：真机验收（Edge headless + CDP）

验的东西：
  1) 客服面板 / 聊天室面板：拖动标题栏、右下角拉伸、最大化还原
  2) 意见反馈：多余文案已删、功能项变多、布局正常
  3) 四个核心弹窗：标题副文案变短、分步结构在、主按钮是实心蓝
  4) 对话框：placeholder 在、无 out-hint
  5) 全程 console 无报错
"""
import base64
import json
import os
import shutil
import subprocess
import sys
import time
import urllib.request
from urllib.parse import urlparse

EDGE = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
BASE = "http://127.0.0.1:8000"
PROFILE = os.path.join(os.environ.get("TEMP", r"C:\Windows\Temp"), "_cdp_ui1107")
OUT = r"C:\Users\覃敬涛\Desktop\研一项目\agent\输出"
LOG = r"C:\Users\覃敬涛\Desktop\研一项目\收尾对照\界面精简-真机-1107.txt"

# 关掉代理：系统代理会劫持 127.0.0.1
for k in ("HTTP_PROXY", "HTTPS_PROXY", "ALL_PROXY", "http_proxy", "https_proxy", "all_proxy"):
    os.environ.pop(k, None)

_opener = urllib.request.build_opener(urllib.request.ProxyHandler({}))

lines = []


def say(*a):
    s = " ".join(str(x) for x in a)
    print(s, flush=True)
    lines.append(s)


def http(path, data=None, headers=None):
    url = BASE + path
    body = json.dumps(data).encode() if data is not None else None
    req = urllib.request.Request(url, data=body, headers=headers or {})
    if body:
        req.add_header("Content-Type", "application/json")
    return _opener.open(req, timeout=20)


class CDP:
    def __init__(self, ws):
        self.ws = ws
        self.n = 0
        self.events = []

    def send(self, method, params=None, timeout=30):
        self.n += 1
        mid = self.n
        self.ws.send(json.dumps({"id": mid, "method": method, "params": params or {}}))
        t0 = time.time()
        while time.time() - t0 < timeout:
            m = json.loads(self.ws.recv())
            if m.get("id") == mid:
                if "error" in m:
                    raise RuntimeError("%s -> %s" % (method, m["error"]))
                return m.get("result", {})
            if m.get("method"):
                self.events.append(m)
        raise TimeoutError(method)

    def js(self, expr, timeout=30):
        r = self.send("Runtime.evaluate",
                      {"expression": expr, "returnByValue": True, "awaitPromise": True},
                      timeout=timeout)
        res = r.get("result", {})
        if r.get("exceptionDetails"):
            raise RuntimeError(json.dumps(r["exceptionDetails"], ensure_ascii=False)[:400])
        return res.get("value")

    def shot(self, name):
        d = self.send("Page.captureScreenshot", {"format": "png"}, timeout=60)
        p = os.path.join(OUT, name)
        with open(p, "wb") as f:
            f.write(base64.b64decode(d["data"]))
        say("  [图] " + name)
        return p


def login_cookie():
    r = http("/api/auth/login", {"username": "2482991994", "password": "123456"})
    raw = r.headers.get_all("Set-Cookie") or []
    for c in raw:
        for part in c.split(";"):
            part = part.strip()
            if part.startswith("a23_token="):
                return part.split("=", 1)[1]
    raise RuntimeError("没拿到 a23_token")


def main():
    # 清掉上次残留
    for pid in os.popen('wmic process where "name=\'msedge.exe\'" get processid,commandline').read().split("\n"):
        if "_cdp_ui1107" in pid:
            try:
                subprocess.run(["taskkill", "/F", "/PID", pid.strip().split()[-1]],
                               stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            except Exception:
                pass
    shutil.rmtree(PROFILE, ignore_errors=True)

    try:
        http("/api/status")
    except Exception as e:
        say("!! 8000 没起：", e)
        return 1

    proc = subprocess.Popen([
        EDGE, "--headless=new", "--disable-gpu", "--no-first-run", "--no-default-browser-check",
        "--disable-extensions", "--disable-sync", "--disable-background-networking",
        "--disable-component-update", "--no-proxy-server", "--remote-allow-origins=*",
        "--remote-debugging-port=9333", "--user-data-dir=" + PROFILE,
        "--window-size=1440,960", "about:blank",
    ], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

    ws_url = None
    try:
        for _ in range(50):
            try:
                j = json.loads(_opener.open("http://127.0.0.1:9333/json/version", timeout=2).read())
                if j.get("webSocketDebuggerUrl"):
                    break
            except Exception:
                time.sleep(0.5)
        req = urllib.request.Request("http://127.0.0.1:9333/json/new?about:blank", method="PUT")
        tab = json.loads(_opener.open(req, timeout=10).read())
        ws_url = tab["webSocketDebuggerUrl"]

        import websockets.sync.client as wsc
        ws = wsc.connect(ws_url, max_size=None, open_timeout=20)
        c = CDP(ws)
        c.send("Page.enable")
        c.send("Runtime.enable")
        c.send("Network.enable")
        c.send("Console.enable")

        tok = login_cookie()
        c.send("Network.setCookie", {"name": "a23_token", "value": tok,
                                     "url": BASE, "path": "/", "httpOnly": True})
        c.send("Page.navigate", {"url": BASE + "/"})
        time.sleep(3.5)

        ok = 0
        bad = 0

        def chk(name, cond, extra=""):
            nonlocal ok, bad
            if cond:
                ok += 1
                say("  ✅", name, extra)
            else:
                bad += 1
                say("  ❌", name, extra)

        say("== 0. 页面 ==")
        title = c.js("document.title")
        chk("页面打开", bool(title), title)

        # ---------- 1. 对话框 ----------
        say("\n== 1. 对话框 ==")
        chk("输入框有 placeholder", bool(c.js("(document.getElementById('input')||{}).placeholder || ''")))
        chk("out-hint 已移除", c.js("document.querySelectorAll('.out-hint').length") == 0)
        chk("输入框所在组 .cp-field", c.js("!!document.querySelector('.cp-field #input')"))
        c.shot("界面精简-1-对话区.png")

        # ---------- 2. 意见反馈 ----------
        say("\n== 2. 意见反馈 ==")
        c.js("fbOpen('other', {title:'意见反馈'})")
        time.sleep(0.8)
        body = c.js("document.querySelector('.fb-card').innerText")
        chk("删掉『打个分，顺手写两句』", "打个分，顺手写两句" not in body)
        chk("删掉『他们真的会看』", "他们真的会看" not in body)
        chk("删掉『也可只写意见』", "也可只写意见" not in body)
        nfeat = c.js("document.querySelectorAll('#fb-feats .fb-feat').length")
        chk("功能项 ≥ 12", nfeat >= 12, "实有 %s" % nfeat)
        chk("有数据仓库 / 文字识别 等新项",
            c.js("!!document.querySelector('.fb-feat[data-f=\"warehouse\"]') && !!document.querySelector('.fb-feat[data-f=\"ocr\"]')"))
        chk("星星 5 颗", c.js("document.querySelectorAll('#fb-stars .fb-star').length") == 5)
        c.js("document.querySelector('.fb-feat[data-f=\"fill\"]').click()")
        time.sleep(0.4)
        chk("点功能后标签跟着换", "填得对" in c.js("document.getElementById('fb-tags').innerText"))
        c.shot("界面精简-2-意见反馈.png")
        c.js("fbClose()")
        time.sleep(0.3)

        # ---------- 3. 四个核心弹窗 ----------
        say("\n== 3. 核心弹窗 ==")
        def panel_sub(oid):
            return c.js("(document.querySelector('#%s .dw-hd .dw-sub')||{}).innerText||''" % oid)

        cases = [("dw-overlay", "数据仓库"), ("rec-overlay", "识别入库"),
                 ("ocr-overlay", "文字识别"), ("fill-overlay", "表格填写")]
        for oid, nm in cases:
            c.js("document.getElementById('%s').classList.add('open')" % oid)
            time.sleep(0.6)
            sub = panel_sub(oid)
            chk(nm + " 副标题简短", 0 < len(sub) <= 26, repr(sub))
            if oid == "dw-overlay":
                chk(nm + " 保留 文档库 + 查询栏",
                    c.js("!!document.querySelector('#dw-overlay .dw-docs') && "
                         "!!document.querySelector('#dw-overlay .dw-ctl')"))
            else:
                chk(nm + " 用分步结构", c.js("document.querySelectorAll('#%s .pstep').length" % oid) >= 1)
            c.shot("界面精简-3-%s.png" % nm)
            c.js("document.getElementById('%s').classList.remove('open')" % oid)
            time.sleep(0.2)

        chk("数据仓库 不再写 MySQL/对象属性",
            "MySQL" not in c.js("document.querySelector('#dw-overlay .dw-hd').innerText"))
        chk("识别入库 主按钮实心蓝",
            c.js("getComputedStyle(document.getElementById('rec-go')).color") == "rgb(255, 255, 255)")
        chk("表格填写 主按钮在 .fill-run",
            c.js("!!document.querySelector('.fill-run #fill-go')"))
        chk("文字识别 精度卡可点",
            c.js("document.querySelectorAll('.ocr-opt').length") == 2)

        # ---------- 4. 客服浮层：拖 / 缩 / 最大化 ----------
        say("\n== 4. 客服浮层 ==")
        c.js("document.getElementById('cs-ai').click()")
        time.sleep(1.2)
        chk("客服面板已打开", c.js("document.getElementById('cs-panel').classList.contains('on')"))
        r0 = c.js("JSON.stringify((()=>{const r=document.getElementById('cs-panel').getBoundingClientRect();"
                  "return {x:Math.round(r.x),y:Math.round(r.y),w:Math.round(r.width),h:Math.round(r.height)}})())")
        r0 = json.loads(r0)
        say("  初始：", r0)

        # 拖标题栏（客服面板的标题栏在右侧 cs-main 里，不在左栏，按真实矩形取坐标）
        hrect = json.loads(c.js("JSON.stringify((()=>{const r=document.querySelector('#cs-panel .cs-hd')"
                                ".getBoundingClientRect();return {x:Math.round(r.x),y:Math.round(r.y),"
                                "w:Math.round(r.width),h:Math.round(r.height)}})())"))
        hx = hrect["x"] + 70
        hy = hrect["y"] + hrect["h"] // 2
        c.send("Input.dispatchMouseEvent", {"type": "mousePressed", "x": hx, "y": hy,
                                            "button": "left", "clickCount": 1})
        c.send("Input.dispatchMouseEvent", {"type": "mouseMoved", "x": hx - 120, "y": hy - 90,
                                            "button": "left", "buttons": 1})
        c.send("Input.dispatchMouseEvent", {"type": "mouseReleased", "x": hx - 120, "y": hy - 90,
                                            "button": "left", "clickCount": 1})
        time.sleep(0.5)
        r1 = json.loads(c.js("JSON.stringify((()=>{const r=document.getElementById('cs-panel').getBoundingClientRect();"
                             "return {x:Math.round(r.x),y:Math.round(r.y),w:Math.round(r.width),h:Math.round(r.height)}})())"))
        chk("拖动后位置变了", abs(r1["x"] - r0["x"]) > 40 or abs(r1["y"] - r0["y"]) > 40,
            "→ %s,%s" % (r1["x"], r1["y"]))

        # 右下角拉伸
        gx = r1["x"] + r1["w"] - 8
        gy = r1["y"] + r1["h"] - 8
        c.send("Input.dispatchMouseEvent", {"type": "mousePressed", "x": gx, "y": gy,
                                            "button": "left", "clickCount": 1})
        c.send("Input.dispatchMouseEvent", {"type": "mouseMoved", "x": gx - 150, "y": gy - 110,
                                            "button": "left", "buttons": 1})
        c.send("Input.dispatchMouseEvent", {"type": "mouseReleased", "x": gx - 150, "y": gy - 110,
                                            "button": "left", "clickCount": 1})
        time.sleep(0.5)
        r2 = json.loads(c.js("JSON.stringify((()=>{const r=document.getElementById('cs-panel').getBoundingClientRect();"
                             "return {x:Math.round(r.x),y:Math.round(r.y),w:Math.round(r.width),h:Math.round(r.height)}})())"))
        chk("拉伸后变小", r2["w"] < r1["w"] - 60 and r2["h"] < r1["h"] - 50,
            "%sx%s → %sx%s" % (r1["w"], r1["h"], r2["w"], r2["h"]))
        c.shot("界面精简-4-客服-拖缩后.png")

        # 最大化 / 还原
        c.js("document.getElementById('cs-max').click()")
        time.sleep(0.4)
        r3 = json.loads(c.js("JSON.stringify((()=>{const r=document.getElementById('cs-panel').getBoundingClientRect();"
                             "return {x:Math.round(r.x),y:Math.round(r.y),w:Math.round(r.width),h:Math.round(r.height)}})())"))
        chk("最大化铺满窗口", r3["w"] >= 1300 and r3["x"] <= 20,
            "%sx%s @%s" % (r3["w"], r3["h"], r3["x"]))
        c.shot("界面精简-5-客服-最大化.png")
        c.js("document.getElementById('cs-max').click()")
        time.sleep(0.4)
        r4 = json.loads(c.js("JSON.stringify((()=>{const r=document.getElementById('cs-panel').getBoundingClientRect();"
                             "return {x:Math.round(r.x),y:Math.round(r.y),w:Math.round(r.width),h:Math.round(r.height)}})())"))
        chk("还原回拉伸后的尺寸", abs(r4["w"] - r2["w"]) <= 2 and abs(r4["h"] - r2["h"]) <= 2,
            "%sx%s" % (r4["w"], r4["h"]))

        # 关掉重开：尺寸要记住
        c.js("document.getElementById('cs-x').click()")
        time.sleep(0.4)
        c.js("document.getElementById('cs-ai').click()")
        time.sleep(1.0)
        r5 = json.loads(c.js("JSON.stringify((()=>{const r=document.getElementById('cs-panel').getBoundingClientRect();"
                             "return {x:Math.round(r.x),y:Math.round(r.y),w:Math.round(r.width),h:Math.round(r.height)}})())"))
        chk("重开后记住尺寸", abs(r5["w"] - r2["w"]) <= 2 and abs(r5["h"] - r2["h"]) <= 2,
            "%sx%s" % (r5["w"], r5["h"]))

        # ---------- 5. 聊天室 ----------
        say("\n== 5. 聊天室 ==")
        c.js("crOpen()")
        time.sleep(1.6)
        chk("聊天室已打开", c.js("document.getElementById('cr-panel').classList.contains('on')"))
        q0 = json.loads(c.js("JSON.stringify((()=>{const r=document.getElementById('cr-panel').getBoundingClientRect();"
                             "return {x:Math.round(r.x),y:Math.round(r.y),w:Math.round(r.width),h:Math.round(r.height)}})())"))
        hx = q0["x"] + 60
        hy = q0["y"] + 20
        c.send("Input.dispatchMouseEvent", {"type": "mousePressed", "x": hx, "y": hy,
                                            "button": "left", "clickCount": 1})
        c.send("Input.dispatchMouseEvent", {"type": "mouseMoved", "x": hx - 100, "y": hy - 70,
                                            "button": "left", "buttons": 1})
        c.send("Input.dispatchMouseEvent", {"type": "mouseReleased", "x": hx - 100, "y": hy - 70,
                                            "button": "left", "clickCount": 1})
        time.sleep(0.5)
        q1 = json.loads(c.js("JSON.stringify((()=>{const r=document.getElementById('cr-panel').getBoundingClientRect();"
                             "return {x:Math.round(r.x),y:Math.round(r.y),w:Math.round(r.width),h:Math.round(r.height)}})())"))
        chk("聊天室可拖动", abs(q1["x"] - q0["x"]) > 30 or abs(q1["y"] - q0["y"]) > 30,
            "→ %s,%s" % (q1["x"], q1["y"]))
        c.js("document.getElementById('cr-max').click()")
        time.sleep(0.4)
        q2 = json.loads(c.js("JSON.stringify((()=>{const r=document.getElementById('cr-panel').getBoundingClientRect();"
                             "return {x:Math.round(r.x),y:Math.round(r.y),w:Math.round(r.width),h:Math.round(r.height)}})())"))
        chk("聊天室可最大化", q2["w"] >= 1300, "%sx%s" % (q2["w"], q2["h"]))
        c.shot("界面精简-6-聊天室-最大化.png")
        c.js("document.getElementById('cr-max').click()")
        time.sleep(0.3)
        c.js("crClose()")

        # ---------- 6. console 报错 ----------
        say("\n== 6. console ==")
        errs = []
        for ev in c.events:
            if ev.get("method") == "Runtime.exceptionThrown":
                errs.append(json.dumps(ev["params"], ensure_ascii=False)[:200])
            if ev.get("method") == "Runtime.consoleAPICalled" and ev["params"].get("type") == "error":
                errs.append("console.error")
        chk("无 JS 异常", len(errs) == 0, "; ".join(errs[:3]))

        say("\n===== 结果：%d 通过 / %d 失败 =====" % (ok, bad))
        return 0 if bad == 0 else 2
    finally:
        try:
            proc.terminate()
            proc.wait(timeout=8)
        except Exception:
            pass
        time.sleep(0.6)
        shutil.rmtree(PROFILE, ignore_errors=True)
        with open(LOG, "w", encoding="utf-8") as f:
            f.write("\n".join(lines))
        print("记录已写：", LOG)


if __name__ == "__main__":
    sys.exit(main())
