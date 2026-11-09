# -*- coding: utf-8 -*-
"""真机验收：识别入库 · 我上传的 Excel 能否抽出来（2026-11-07）

用 Edge headless + CDP：登录 → 打开「识别入库」→ 选 Excel 分类 → 勾那份上传的空气质量数据
→ 开始识别 → 等任务结束 → 截图。

跑法（managed 3.13.12，带 websockets）：
  C:\\Users\\覃敬涛\\.workbuddy\\binaries\\python\\versions\\3.13.12\\python.exe -u 收尾对照/识别入库-导入文档-真机截图-1107.py
"""
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
PROFILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "_cdp_prof_rec")
OUT = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "agent", "输出")
DOC = "山东省环境空气质量监测数据信息202512171921_0.xlsx"

_OPENER = urllib.request.build_opener(urllib.request.ProxyHandler({}))

import asyncio
import websockets  # noqa: E402


def http_json(path, body=None, method=None):
    data = json.dumps(body, ensure_ascii=False).encode() if body is not None else None
    req = urllib.request.Request(BASE + path, data=data,
                                 headers={"Content-Type": "application/json"} if data else {},
                                 method=method or ("POST" if data else "GET"))
    return _OPENER.open(req, timeout=30)


def shot_env():
    return None


async def _shot(cmd, name):
    res = await cmd("Page.captureScreenshot", {"format": "png"})
    path = os.path.join(OUT, name)
    with open(path, "wb") as f:
        f.write(base64.b64decode(res["data"]))
    print("  截图 →", path)


async def main():
    # 0. 登录拿 cookie
    tok = None
    with http_json("/api/auth/login", {"username": "2482991994", "password": "123456"}) as r:
        for c in r.headers.get_all("Set-Cookie") or []:
            if c.startswith("a23_token="):
                tok = c.split(";")[0].split("=", 1)[1]
    print("登录 cookie:", "有" if tok else "无")

    if os.path.isdir(PROFILE):
        shutil.rmtree(PROFILE, ignore_errors=True)
    proc = subprocess.Popen([
        EDGE, "--headless=new", "--disable-gpu", "--no-first-run", "--no-default-browser-check",
        "--disable-extensions", "--disable-sync", "--disable-background-networking",
        "--disable-component-update", "--no-proxy-server", "--remote-allow-origins=*",
        "--remote-debugging-port=9333", "--user-data-dir=" + PROFILE,
        "--window-size=1440,960", "about:blank",
    ], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

    ws = None
    try:
        for _ in range(42):
            try:
                with _OPENER.open("http://127.0.0.1:9333/json/version", timeout=2) as r:
                    if json.loads(r.read().decode()).get("Browser"):
                        break
            except Exception:
                pass
            time.sleep(0.5)

        req = urllib.request.Request("http://127.0.0.1:9333/json/new?about:blank", method="PUT")
        with _OPENER.open(req, timeout=5) as r:
            ws_url = json.loads(r.read().decode())["webSocketDebuggerUrl"]

        ws = await websockets.connect(ws_url, max_size=None)
        _id = [0]

        async def cmd(method, params=None):
            _id[0] += 1
            await ws.send(json.dumps({"id": _id[0], "method": method, "params": params or {}}))
            while True:
                m = json.loads(await ws.recv())
                if m.get("method") == "Page.javascriptDialogOpening":
                    await ws.send(json.dumps({"id": _id[0] + 100000, "method": "Page.handleJavaScriptDialog",
                                              "params": {"accept": True}}))
                    continue
                if m.get("id") == _id[0]:
                    if "error" in m:
                        raise RuntimeError("%s → %s" % (method, m["error"]))
                    return m.get("result") or {}

        async def js(expr):
            r = await cmd("Runtime.evaluate", {"expression": expr, "returnByValue": True,
                                               "awaitPromise": True})
            return (r.get("result") or {}).get("value")

        await cmd("Page.enable")
        await cmd("Runtime.enable")
        await cmd("Network.enable")
        if tok:
            await cmd("Network.setCookie", {"name": "a23_token", "value": tok, "url": BASE,
                                            "path": "/", "httpOnly": True})
        await cmd("Page.navigate", {"url": BASE + "/"})
        await asyncio.sleep(3.5)

        # 1. 打开「识别入库」
        opened = await js("""(() => {
          const b = [...document.querySelectorAll('button, .card, [onclick]')]
            .find((e) => (e.innerText || '').includes('识别入库'));
          if (!b) return 'no-card';
          b.click();
          return document.getElementById('rec-overlay').classList.contains('open') ? 'ok' : 'clicked';
        })()""")
        print("打开识别入库:", opened)
        await asyncio.sleep(2.0)

        # 2. 切到 Excel 分类
        await js("""(() => {
          const c = [...document.querySelectorAll('#rec-kinds *')]
            .find((e) => (e.innerText || '').trim().startsWith('Excel'));
          if (c) c.click();
        })()""")
        await asyncio.sleep(1.2)
        await _shot(cmd, "识别入库-1-Excel分类.png")

        # 3. 勾选目标文档
        picked = await js("""(() => {
          const rows = [...document.querySelectorAll('#rec-docs *')];
          const row = rows.find((e) => (e.innerText || '').includes(%s));
          if (!row) return 'no-row';
          let host = row;
          while (host && !host.querySelector('input[type=checkbox]')) host = host.parentElement;
          const box = host && host.querySelector('input[type=checkbox]');
          if (!box) return 'no-box';
          if (!box.checked) box.click();
          return 'checked=' + box.checked;
        })()""" % json.dumps(DOC))
        print("勾选文档:", picked)
        await asyncio.sleep(0.6)
        await _shot(cmd, "识别入库-2-已勾选.png")

        # 4. 开始识别
        await js("document.getElementById('rec-go').click()")
        t0 = time.time()
        last = ""
        while time.time() - t0 < 210:
            txt = await js("""(() => {
              const s = document.getElementById('rec-status');
              const r = document.getElementById('rec-result');
              return (s ? s.innerText : '') + '\\n' + (r ? r.innerText : '');
            })()""") or ""
            if txt != last:
                last = txt
                print("  …%s" % txt.replace("\n", " | ")[:120])
            if ("完成" in txt or "失败" in txt) and "识别中" not in txt:
                break
            await asyncio.sleep(2.0)
        elapsed = time.time() - t0
        print("任务耗时 %.1fs" % elapsed)
        await asyncio.sleep(1.5)
        await _shot(cmd, "识别入库-3-结果.png")

        final = await js("""(() => {
          const s = document.getElementById('rec-status');
          const r = document.getElementById('rec-result');
          return (s ? s.innerText : '') + '\\n' + (r ? r.innerText : '');
        })()""")
        print("\n=== 结果区文案 ===\n" + (final or "(空)"))

        ok = ("识别完成" in (final or "")) and ("不在解析池" not in (final or ""))
        print("\n断言：%s" % ("✓ 这次成功了，没有『不在解析池』" if ok else "✗ 仍然失败"))
        return 0 if ok else 1
    finally:
        if ws:
            await ws.close()
        proc.terminate()
        time.sleep(1)
        shutil.rmtree(PROFILE, ignore_errors=True)


if __name__ == "__main__":
    sys.exit(asyncio.run(main()))
