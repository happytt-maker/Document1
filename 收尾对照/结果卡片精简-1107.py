# -*- coding: utf-8 -*-
"""结果卡片精简：注入一份真实结构的返回，看卡片是否只留一句人话 + 详情折叠。"""
import base64
import json
import os
import shutil
import subprocess
import time
import urllib.request

EDGE = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
BASE = "http://127.0.0.1:8000"
PROFILE = os.path.join(os.environ.get("TEMP", r"C:\Windows\Temp"), "_cdp_ui1107c")
OUT = r"C:\Users\覃敬涛\Desktop\研一项目\agent\输出"
LOG = r"C:\Users\覃敬涛\Desktop\研一项目\收尾对照\结果卡片精简-1107.txt"

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
        "report": "✓ 已处理「设备采购合同.docx」\n- 标题段落：3 处加粗\n- 正文段落：12 处未改动\n- 输出：agent\\输出\\设备采购合同-已处理.docx\n- 耗时 1.8s（解析 0.4s / 执行 0.9s / 写盘 0.5s）",
        "meta": {
            "action": "格式调整", "op": "set_bold",
            "doc": "设备采购合同.docx",
            "out_path": "agent\\输出\\设备采购合同-已处理.docx",
        },
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
        c.send("Page.enable")
        c.send("Runtime.enable")
        c.send("Network.enable")
        tok = login_cookie()
        c.send("Network.setCookie", {"name": "a23_token", "value": tok,
                                     "url": BASE, "path": "/", "httpOnly": True})
        c.send("Page.navigate", {"url": BASE + "/"})
        time.sleep(4)

        c.js("addUserMsg('把设备采购合同的标题加粗')")
        c.js("(function(){const card=addTaskCard();renderResult(card,%s);})()" % json.dumps(FAKE, ensure_ascii=False))
        time.sleep(0.8)

        head = c.js("(document.querySelector('.msg-sys .card .head')||{}).innerText || ''")
        rline = c.js("(document.querySelector('.msg-sys .card .rline')||{}).innerText || ''")
        detOpen = c.js("!!document.querySelector('.msg-sys .card details.rdet')")
        detVisible = c.js("(function(){const d=document.querySelector('details.rdet'); if(!d) return null; const r=d.getBoundingClientRect(); return r.height;})().valueOf()")
        metaVisible = c.js("document.querySelectorAll('.msg-sys .card .meta-line').length")
        say("卡片标题行:", repr(head))
        say("一句话结果:", repr(rline))
        say("有详情折叠:", detOpen, "｜折叠高度:", detVisible)
        say("残留 meta-line 数:", metaVisible)
        c.shot("界面精简-11-结果卡片.png")
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
