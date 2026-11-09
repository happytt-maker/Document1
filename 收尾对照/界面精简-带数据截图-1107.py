# -*- coding: utf-8 -*-
"""界面精简：带真实数据的四弹窗截图（上一版只加 open class，数据没加载）
   目的：确认「数据仓库」等弹窗在真实数据下的排版是否简洁好读。
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
PROFILE = os.path.join(os.environ.get("TEMP", r"C:\Windows\Temp"), "_cdp_ui1107b")
OUT = r"C:\Users\覃敬涛\Desktop\研一项目\agent\输出"
LOG = r"C:\Users\覃敬涛\Desktop\研一项目\收尾对照\界面精简-带数据截图-1107.txt"

for k in ("HTTP_PROXY", "HTTPS_PROXY", "ALL_PROXY", "http_proxy", "https_proxy", "all_proxy"):
    os.environ.pop(k, None)
_opener = urllib.request.build_opener(urllib.request.ProxyHandler({}))

lines = []


def say(*a):
    s = " ".join(str(x) for x in a)
    print(s, flush=True)
    lines.append(s)


class CDP:
    def __init__(self, ws):
        self.ws = ws
        self.n = 0

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
                      {"expression": expr, "returnByValue": True, "awaitPromise": True},
                      timeout=timeout)
        if r.get("exceptionDetails"):
            raise RuntimeError(json.dumps(r["exceptionDetails"], ensure_ascii=False)[:400])
        return r.get("result", {}).get("value")

    def shot(self, name):
        d = self.send("Page.captureScreenshot", {"format": "png"}, timeout=60)
        p = os.path.join(OUT, name)
        with open(p, "wb") as f:
            f.write(base64.b64decode(d["data"]))
        say("  [图] " + name)
        return p


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
    raise RuntimeError("没拿到 a23_token")


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
        c.send("Page.enable")
        c.send("Runtime.enable")
        c.send("Network.enable")

        tok = login_cookie()
        c.send("Network.setCookie", {"name": "a23_token", "value": tok,
                                     "url": BASE, "path": "/", "httpOnly": True})
        c.send("Page.navigate", {"url": BASE + "/"})
        time.sleep(4)

        # 数据仓库：真开 + 真加载
        c.js("document.getElementById('dw-overlay').classList.add('open')")
        time.sleep(0.4)
        c.js("(async()=>{dwLoadStats();dwLoadDocs();dwLoadEntities();})()")
        time.sleep(4.5)
        st = c.js("document.getElementById('dw-stats').innerText")
        rows = c.js("document.querySelectorAll('#dw-body .kdoc, #dw-body .dw-item, #dw-body .card').length")
        say("数据仓库统计行:", repr(st))
        say("仓库条目数:", rows)
        c.shot("界面精简-7-数据仓库-有数据.png")

        # 识别入库：真加载文档清单
        c.js("document.getElementById('dw-overlay').classList.remove('open')")
        c.js("document.getElementById('rec-overlay').classList.add('open')")
        time.sleep(0.3)
        c.js("recLoadDocs(true)")
        time.sleep(3)
        say("识别入库文档数:", c.js("document.querySelectorAll('#rec-docs input[type=checkbox]').length"))
        c.shot("界面精简-8-识别入库-有数据.png")

        # 表格填写：真加载模板
        c.js("document.getElementById('rec-overlay').classList.remove('open')")
        c.js("document.getElementById('fill-overlay').classList.add('open')")
        time.sleep(0.3)
        c.js("fillLoadTemplates(true)")
        time.sleep(3)
        say("模板数:", c.js("document.getElementById('fill-tpl').options.length"))
        c.shot("界面精简-9-表格填写-有数据.png")

        # 意见反馈
        c.js("document.getElementById('fill-overlay').classList.remove('open')")
        c.js("fbOpen('other', {title:'意见反馈'})")
        time.sleep(0.8)
        c.shot("界面精简-10-意见反馈-终版.png")

        ws.close()
    finally:
        try:
            proc.terminate()
        except Exception:
            pass
        shutil.rmtree(PROFILE, ignore_errors=True)

    with open(LOG, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
