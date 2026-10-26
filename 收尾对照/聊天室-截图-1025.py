# -*- coding: utf-8 -*-
"""拍 README 用的聊天室截图：用内置演示账号演一段真实对话。

跑法：managed 3.13（带 websockets）
     python 收尾对照\\聊天室-截图-1025.py --port 8010
产出：docs\\screenshots\\08-聊天室.png
"""
import argparse
import base64
import json
import os
import shutil
import subprocess
import sys
import tempfile
import time
import urllib.request
import uuid
from http.cookiejar import CookieJar
from pathlib import Path

for k in ("HTTP_PROXY", "HTTPS_PROXY", "ALL_PROXY", "http_proxy", "https_proxy", "all_proxy"):
    os.environ.pop(k, None)

EDGE = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
ROOT = Path(__file__).resolve().parent.parent
SHOT = ROOT / "docs" / "screenshots" / "08-聊天室.png"


class Client:
    def __init__(self, base):
        self.base = base
        self.cj = CookieJar()
        self.op = urllib.request.build_opener(urllib.request.ProxyHandler({}),
                                              urllib.request.HTTPCookieProcessor(self.cj))

    def call(self, path, data=None):
        body = json.dumps(data).encode() if data is not None else None
        hdr = {"Content-Type": "application/json"} if body else {}
        req = urllib.request.Request(self.base + path, data=body, headers=hdr,
                                     method="POST" if body else "GET")
        try:
            with self.op.open(req, timeout=20) as r:
                return r.status, json.loads(r.read().decode() or "{}")
        except urllib.error.HTTPError as e:
            return e.code, json.loads(e.read().decode() or "{}")

    def token(self):
        for c in self.cj:
            if c.name == "a23_token":
                return c.value
        return ""


from websockets.sync.client import connect


class CDP:
    def __init__(self, url):
        self.ws = connect(url, max_size=None)
        self.id = 0
        self.errors = []

    def cmd(self, method, params=None, timeout=25):
        self.id += 1
        mid = self.id
        self.ws.send(json.dumps({"id": mid, "method": method, "params": params or {}}))
        end = time.time() + timeout
        while time.time() < end:
            m = json.loads(self.ws.recv())
            if m.get("id") == mid:
                return m.get("result", {})
            if m.get("method") == "Runtime.exceptionThrown":
                self.errors.append(json.dumps(m, ensure_ascii=False)[:200])
        raise TimeoutError(method)

    def js(self, expr):
        r = self.cmd("Runtime.evaluate", {"expression": expr, "returnByValue": True,
                                         "awaitPromise": True, "userGesture": True})
        return r.get("result", {}).get("value")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--port", type=int, default=8010)
    args = ap.parse_args()
    base = "http://127.0.0.1:%d" % args.port

    lin = Client(base)
    zhou = Client(base)
    zhang = Client(base)
    for cli, user in ((lin, "demo_ai"), (zhou, "demo_human"), (zhang, "kf_demo")):
        code, j = cli.call("/api/auth/login", {"username": user, "password": "123456",
                                               "portal": "user"})
        print("login", user, code, (j.get("user") or {}).get("display_name"))
        if code != 200:
            raise SystemExit("演示账号登录失败，先确认服务在跑：%s" % user)
        cli.call("/api/room/poll?after=0&alive=1")

    msgs = [
        (lin, "模板三里的「库存数量」是按期末数填，还是按公式算？我在原文里没找到依据"),
        (zhou, "@林同学 我这边是按期末数填的，报告里会标出原文出处，不确定的格子会留空"),
        (zhang, "两位好，这个口径我再跟开发确认一遍，结论同步到群里"),
    ]
    for cli, text in msgs:
        code, j = cli.call("/api/room/send", {"content": text})
        print("send", code, str(j.get("detail") or "")[:60])
        time.sleep(1.2)

    prof = Path(tempfile.gettempdir()) / ("_a23_shot_" + uuid.uuid4().hex[:6])
    proc = None
    cdp = None
    try:
        proc = subprocess.Popen([EDGE, "--headless=new", "--disable-gpu", "--no-first-run",
                                 "--no-default-browser-check", "--disable-extensions",
                                 "--disable-sync", "--disable-background-networking",
                                 "--disable-component-update", "--no-proxy-server",
                                 "--remote-allow-origins=*", "--remote-debugging-port=9336",
                                 "--user-data-dir=" + str(prof), "--window-size=1440,960",
                                 "about:blank"],
                                stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        op = urllib.request.build_opener(urllib.request.ProxyHandler({}))
        for _ in range(42):
            try:
                with op.open("http://127.0.0.1:9336/json/version", timeout=3) as r:
                    if json.loads(r.read().decode()).get("Browser"):
                        break
            except Exception:
                pass
            time.sleep(0.5)
        req = urllib.request.Request("http://127.0.0.1:9336/json/new?about:blank", method="PUT")
        with op.open(req) as r:
            ws = json.loads(r.read().decode())["webSocketDebuggerUrl"]
        cdp = CDP(ws)
        cdp.cmd("Page.enable")
        cdp.cmd("Runtime.enable")
        cdp.cmd("Network.setCookie", {"name": "a23_token", "value": lin.token(),
                                      "url": base, "path": "/", "httpOnly": True})
        cdp.cmd("Page.navigate", {"url": base + "/"})
        time.sleep(4)
        cdp.js("document.getElementById('room-open').click()")
        time.sleep(3)
        # 收起右侧操作历史，让截图只留聊天室
        cdp.js("(()=>{const b=document.getElementById('side-toggle');if(b)b.click();})()")
        time.sleep(1.2)
        n = cdp.js("document.querySelectorAll('.cr-msg').length")
        print("面板消息数：", n, "在线：", cdp.js("document.getElementById('cr-online-tx').textContent"))
        r = cdp.cmd("Page.captureScreenshot", {"format": "png"})
        SHOT.parent.mkdir(parents=True, exist_ok=True)
        SHOT.write_bytes(base64.b64decode(r["data"]))
        print("截图 -> %s (%.1f KB)" % (SHOT, SHOT.stat().st_size / 1024))
        print("JS 报错：", cdp.errors or "无")
    finally:
        if cdp:
            try:
                cdp.ws.close()
            except Exception:
                pass
        if proc:
            proc.terminate()
            try:
                proc.wait(timeout=8)
            except Exception:
                pass
        shutil.rmtree(prof, ignore_errors=True)


if __name__ == "__main__":
    main()
