# -*- coding: utf-8 -*-
"""收尾复核：
  1) 结果卡片：文档名不再重复（「X.docx · 已处理「X.docx」」→ 只留一句）
  2) 聊天室面板：拖动 + 右下角拉伸 + 最大化 + 还原 + 记忆尺寸
"""
import base64
import json
import os
import shutil
import subprocess
import time
import urllib.request

EDGE = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
BASE = "http://127.0.0.1:8000"
PROFILE = os.path.join(os.environ.get("TEMP", r"C:\Windows\Temp"), "_cdp_ui1107d")
OUT = r"C:\Users\覃敬涛\Desktop\研一项目\agent\输出"
LOG = r"C:\Users\覃敬涛\Desktop\研一项目\收尾对照\界面精简-收尾复核-1107.txt"

for k in ("HTTP_PROXY", "HTTPS_PROXY", "ALL_PROXY", "http_proxy", "https_proxy", "all_proxy"):
    os.environ.pop(k, None)
_opener = urllib.request.build_opener(urllib.request.ProxyHandler({}))
lines = []
ok = bad = 0


def say(*a):
    s = " ".join(str(x) for x in a)
    print(s, flush=True)
    lines.append(s)


def chk(name, cond, extra=""):
    global ok, bad
    if cond:
        ok += 1
        say("  ✅", name, extra)
    else:
        bad += 1
        say("  ❌", name, extra)


class CDP:
    def __init__(self, ws):
        self.ws, self.n = ws, 0

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
        raise TimeoutError(method)

    def js(self, expr, timeout=40):
        r = self.send("Runtime.evaluate",
                      {"expression": expr, "returnByValue": True, "awaitPromise": True}, timeout=timeout)
        if r.get("exceptionDetails"):
            raise RuntimeError(json.dumps(r["exceptionDetails"], ensure_ascii=False)[:400])
        return r.get("result", {}).get("value")

    def rect(self, sel):
        return self.js("(function(){const e=document.querySelector('%s');if(!e)return null;"
                       "const r=e.getBoundingClientRect();return {x:r.x,y:r.y,w:r.width,h:r.height};})()" % sel)

    def drag(self, x0, y0, x1, y1, steps=8):
        self.send("Input.dispatchMouseEvent", {"type": "mousePressed", "x": x0, "y": y0,
                                               "button": "left", "clickCount": 1})
        for i in range(1, steps + 1):
            self.send("Input.dispatchMouseEvent", {"type": "mouseMoved",
                                                   "x": x0 + (x1 - x0) * i / steps,
                                                   "y": y0 + (y1 - y0) * i / steps,
                                                   "button": "left", "buttons": 1})
            time.sleep(0.02)
        self.send("Input.dispatchMouseEvent", {"type": "mouseReleased", "x": x1, "y": y1,
                                               "button": "left", "clickCount": 1})
        time.sleep(0.35)

    def shot(self, name):
        d = self.send("Page.captureScreenshot", {"format": "png"}, timeout=60)
        with open(os.path.join(OUT, name), "wb") as f:
            f.write(base64.b64decode(d["data"]))
        say("  [图] " + name)


def login_cookie():
    data = json.dumps({"username": "2482991994", "password": "123456"}).encode()
    req = urllib.request.Request(BASE + "/api/auth/login", data=data,
                                 headers={"Content-Type": "application/json"})
    r = _opener.open(req, timeout=20)
    for c in (r.headers.get_all("Set-Cookie") or []):
        for part in c.split(";"):
            part = part.strip()
            if part.startswith("a23_token="):
                return part.split("=", 1)[1]
    raise RuntimeError("no token")


FAKE = {
    "status": "done", "elapsed": 1.8,
    "result": {
        "ok": True,
        "report": "✓ 已处理「设备采购合同.docx」\n- 标题段落：3 处加粗\n- 输出：agent\\输出\\设备采购合同-已处理.docx",
        "meta": {"action": "格式调整", "op": "set_bold", "doc": "设备采购合同.docx",
                 "out_path": "agent\\输出\\设备采购合同-已处理.docx"},
    },
}


def main():
    shutil.rmtree(PROFILE, ignore_errors=True)
    proc = subprocess.Popen([
        EDGE, "--headless=new", "--disable-gpu", "--no-first-run", "--no-default-browser-check",
        "--disable-extensions", "--disable-sync", "--disable-background-networking",
        "--disable-component-update", "--no-proxy-server", "--remote-allow-origins=*",
        "--remote-debugging-port=9333", "--user-data-dir=" + PROFILE,
        "--window-size=1440,960", "about:blank",
    ], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
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
        import websockets.sync.client as wsc
        ws = wsc.connect(tab["webSocketDebuggerUrl"], max_size=None, open_timeout=20)
        c = CDP(ws)
        for m in ("Page.enable", "Runtime.enable", "Network.enable"):
            c.send(m)
        tok = login_cookie()
        c.send("Network.setCookie", {"name": "a23_token", "value": tok,
                                     "url": BASE, "path": "/", "httpOnly": True})
        c.send("Page.navigate", {"url": BASE + "/"})
        time.sleep(4)

        say("== 1. 结果卡片一句话 ==")
        c.js("addUserMsg('把设备采购合同的标题加粗')")
        c.js("(function(){const card=addTaskCard();renderResult(card,%s);})()" % json.dumps(FAKE, ensure_ascii=False))
        time.sleep(0.6)
        rline = c.js("(document.querySelector('.msg-sys .card .rline')||{}).innerText || ''")
        say("  一句话结果:", repr(rline))
        chk("文档名只出现一次", rline.count("设备采购合同.docx") == 1, rline)
        c.shot("界面精简-12-结果卡片-去重.png")

        say("\n== 2. 聊天室拖动 / 缩放 / 最大化 ==")
        c.js("crOpen()")
        time.sleep(1.6)
        r0 = c.rect("#cr-panel")
        say("  初始:", r0)
        chk("聊天室已打开", bool(r0 and r0["w"] > 100), str(r0))

        # 拖标题栏
        hx, hy = r0["x"] + r0["w"] / 2 - 40, r0["y"] + 18
        c.drag(hx, hy, hx - 130, hy - 70)
        r1 = c.rect("#cr-panel")
        chk("可拖动", abs(r1["x"] - r0["x"]) > 40 or abs(r1["y"] - r0["y"]) > 30,
            "→ %.0f,%.0f" % (r1["x"], r1["y"]))

        # 右下角拉伸（抓手法 20x20 在圆角内侧，取它的中心，别取面板矩形角点）
        gr = c.rect("#cr-panel .flt-grip")
        gx, gy = gr["x"] + gr["w"] / 2, gr["y"] + gr["h"] / 2
        say("  抓手法:", gr)
        c.drag(gx, gy, gx - 90, gy - 80)
        r2 = c.rect("#cr-panel")
        chk("可缩放", r2["w"] < r1["w"] - 30 and r2["h"] < r1["h"] - 30,
            "%.0fx%.0f → %.0fx%.0f" % (r1["w"], r1["h"], r2["w"], r2["h"]))
        c.shot("界面精简-13-聊天室-拖动缩放.png")

        # 最大化
        c.js("document.getElementById('cr-max').click()")
        time.sleep(0.5)
        r3 = c.rect("#cr-panel")
        chk("可最大化", r3["w"] > 1200, "%.0fx%.0f" % (r3["w"], r3["h"]))
        c.js("document.getElementById('cr-max').click()")
        time.sleep(0.5)
        r4 = c.rect("#cr-panel")
        chk("可还原", abs(r4["w"] - r2["w"]) < 6 and abs(r4["h"] - r2["h"]) < 6,
            "%.0fx%.0f" % (r4["w"], r4["h"]))
        c.shot("界面精简-14-聊天室-还原.png")

        ws.close()
    finally:
        try:
            proc.terminate()
        except Exception:
            pass
        shutil.rmtree(PROFILE, ignore_errors=True)

    say("\n===== 结果：%d 通过 / %d 失败 =====" % (ok, bad))
    with open(LOG, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
    return 0 if bad == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())
