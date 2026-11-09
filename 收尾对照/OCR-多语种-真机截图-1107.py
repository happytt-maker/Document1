# -*- coding: utf-8 -*-
"""真机验收：文字识别 · 繁体 / 竖排 / 英文（2026-11-07）

Edge headless + CDP，走真实前端：打开「文字识别」→ 选语言 → 上传图片 → 识别 → 截图。

跑法：
  C:\\Users\\覃敬涛\\.workbuddy\\binaries\\python\\versions\\3.13.12\\python.exe -u 收尾对照/OCR-多语种-真机截图-1107.py
"""
import asyncio
import base64
import json
import os
import shutil
import subprocess
import sys
import time
import urllib.request

BASE = "http://127.0.0.1:%s" % os.environ.get("A23_TEST_PORT", "8000")
EDGE = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PROFILE = os.path.join(ROOT, "收尾对照", "_cdp_prof_ocr")
OUT = os.path.join(ROOT, "agent", "输出")
IMGS = [
    ("繁英混排.png", "auto", "OCR多语种-1-繁英混排.png"),
    ("竖排繁体.png", "cht", "OCR多语种-2-竖排繁体.png"),
    ("英文小字.png", "en", "OCR多语种-3-英文.png"),
]

_OPENER = urllib.request.build_opener(urllib.request.ProxyHandler({}))
import websockets  # noqa: E402


async def main():
    tok = None
    data = json.dumps({"username": "2482991994", "password": "123456"}).encode()
    req = urllib.request.Request(BASE + "/api/auth/login", data=data,
                                 headers={"Content-Type": "application/json"})
    with _OPENER.open(req, timeout=20) as r:
        for c in r.headers.get_all("Set-Cookie") or []:
            if c.startswith("a23_token="):
                tok = c.split(";")[0].split("=", 1)[1]

    if os.path.isdir(PROFILE):
        shutil.rmtree(PROFILE, ignore_errors=True)
    proc = subprocess.Popen([
        EDGE, "--headless=new", "--disable-gpu", "--no-first-run", "--no-default-browser-check",
        "--disable-extensions", "--disable-sync", "--disable-background-networking",
        "--disable-component-update", "--no-proxy-server", "--remote-allow-origins=*",
        "--remote-debugging-port=9334", "--user-data-dir=" + PROFILE,
        "--window-size=1440,960", "about:blank",
    ], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

    ws = None
    try:
        for _ in range(42):
            try:
                with _OPENER.open("http://127.0.0.1:9334/json/version", timeout=2) as r:
                    if json.loads(r.read().decode()).get("Browser"):
                        break
            except Exception:
                pass
            time.sleep(0.5)

        req = urllib.request.Request("http://127.0.0.1:9334/json/new?about:blank", method="PUT")
        with _OPENER.open(req, timeout=5) as r:
            ws_url = json.loads(r.read().decode())["webSocketDebuggerUrl"]

        ws = await websockets.connect(ws_url, max_size=None)
        seq = [0]

        async def cmd(method, params=None):
            seq[0] += 1
            await ws.send(json.dumps({"id": seq[0], "method": method, "params": params or {}}))
            while True:
                m = json.loads(await ws.recv())
                if m.get("method") == "Page.javascriptDialogOpening":
                    await ws.send(json.dumps({"id": seq[0] + 90000,
                                              "method": "Page.handleJavaScriptDialog",
                                              "params": {"accept": True}}))
                    continue
                if m.get("id") == seq[0]:
                    if "error" in m:
                        raise RuntimeError("%s → %s" % (method, m["error"]))
                    return m.get("result") or {}

        async def js(expr):
            r = await cmd("Runtime.evaluate", {"expression": expr, "returnByValue": True,
                                               "awaitPromise": True})
            if r.get("exceptionDetails"):
                return "JS-ERR: " + str(r["exceptionDetails"])[:200]
            return (r.get("result") or {}).get("value")

        async def shot(name):
            res = await cmd("Page.captureScreenshot", {"format": "png"})
            p = os.path.join(OUT, name)
            with open(p, "wb") as f:
                f.write(base64.b64decode(res["data"]))
            print("  截图 →", p)

        await cmd("Page.enable")
        await cmd("Runtime.enable")
        await cmd("Network.enable")
        await cmd("DOM.enable")
        if tok:
            await cmd("Network.setCookie", {"name": "a23_token", "value": tok, "url": BASE,
                                            "path": "/", "httpOnly": True})
        await cmd("Page.navigate", {"url": BASE + "/"})
        await asyncio.sleep(3.5)

        print("打开文字识别:", await js("""(() => {
          const b = [...document.querySelectorAll('button, .card, [onclick]')]
            .find((e) => (e.innerText || '').includes('文字识别'));
          if (!b) return 'no-card';
          b.click();
          return document.getElementById('ocr-overlay').classList.contains('open') ? 'ok' : 'clicked';
        })()"""))
        await asyncio.sleep(1.2)

        for fname, lang, out in IMGS:
            path = os.path.join(ROOT, "测试数据", "OCR-多语种", fname)
            # 选语言
            print("  语言 →", await js("""(() => {
              const b = document.querySelector('#ocr-langs .kd[data-l="%s"]');
              if (!b) return 'no-chip';
              b.click();
              return b.classList.contains('on') ? 'on' : 'off';
            })()""" % lang))
            # 上传文件（CDP 直接塞进 input，会触发 change）
            doc = await cmd("DOM.getDocument", {"depth": -1})
            node = await cmd("DOM.querySelector", {"nodeId": doc["root"]["nodeId"],
                                                   "selector": "#ocr-input"})
            await cmd("DOM.setFileInputFiles", {"files": [path], "nodeId": node["nodeId"]})
            await asyncio.sleep(1.0)

            print("  点开始:", await js("""(() => {
              const b = document.getElementById('ocr-go');
              if (!b || b.disabled) return 'disabled';
              b.click(); return 'clicked';
            })()"""))
            t0 = time.time()
            status = ""
            while time.time() - t0 < 120:
                status = await js("(document.getElementById('ocr-status')||{}).innerText || ''") or ""
                if "识别完成" in status or "失败" in status:
                    break
                await asyncio.sleep(0.8)
            print("  状态:", status.replace("\n", " ")[:90])
            txt = await js("(document.querySelector('#ocr-result .ocr-text')||{}).innerText || ''") or ""
            print("  文本:", txt.replace("\n", " ⏎ ")[:150])
            await asyncio.sleep(0.8)
            await shot(out)

            # 清空，准备下一张
            await js("""(() => {
              const b = document.getElementById('ocr-reset');
              if (b) { b.click(); return; }
            })()""")
            await asyncio.sleep(0.6)

        return 0
    finally:
        if ws:
            await ws.close()
        proc.terminate()
        time.sleep(1)
        shutil.rmtree(PROFILE, ignore_errors=True)


if __name__ == "__main__":
    sys.exit(asyncio.run(main()))
