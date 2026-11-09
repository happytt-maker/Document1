# -*- coding: utf-8 -*-
"""真机验收：欢迎卡片改版（简介 + 操作指南，点击直达）+ 输入框占位文字清空（2026-11-07）

Edge headless + CDP：登录 → 打开首页 → 断言欢迎卡内容 / 占位符 / 点击直达 → 截图。

跑法：
  C:\\Users\\覃敬涛\\.workbuddy\\binaries\\python\\versions\\3.13.12\\python.exe -u 收尾对照/欢迎页-改版-真机-1107.py
"""
import asyncio
import base64
import json
import os
import shutil
import subprocess
import time
import urllib.request

BASE = "http://127.0.0.1:%s" % os.environ.get("A23_TEST_PORT", "8000")
EDGE = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PROFILE = os.path.join(ROOT, "收尾对照", "_cdp_prof_welcome")
OUT = os.path.join(ROOT, "agent", "输出")

_OPENER = urllib.request.build_opener(urllib.request.ProxyHandler({}))
import websockets  # noqa: E402

PASS, FAIL = [], []


def check(name, ok, detail=""):
    (PASS if ok else FAIL).append(name)
    print(("  ✓ " if ok else "  ✗ ") + name + (("  —— " + str(detail)) if detail and not ok else ""))


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
        "--remote-debugging-port=9335", "--user-data-dir=" + PROFILE,
        "--window-size=1440,960", "about:blank",
    ], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

    ws = None
    try:
        for _ in range(42):
            try:
                with _OPENER.open("http://127.0.0.1:9335/json/version", timeout=2) as r:
                    if json.loads(r.read().decode()).get("Browser"):
                        break
            except Exception:
                pass
            time.sleep(0.5)

        req = urllib.request.Request("http://127.0.0.1:9335/json/new?about:blank", method="PUT")
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
        if tok:
            await cmd("Network.setCookie", {"name": "a23_token", "value": tok, "url": BASE,
                                            "path": "/", "httpOnly": True})
        await cmd("Page.navigate", {"url": BASE + "/"})
        await asyncio.sleep(3.5)

        print("\n【1】输入框占位文字已清空")
        ph = await js("document.getElementById('input').placeholder")
        check("占位符为空", ph == "", repr(ph))

        print("\n【2】欢迎卡：系统简介 + 操作指南")
        lead = await js("(document.querySelector('.welcome .w-lead')||{}).textContent||''")
        check("有系统简介一句话", "智能文档助手" in str(lead), repr(lead))
        n = await js("document.querySelectorAll('.welcome .guide p').length")
        check("指南 8 条", n == 8, n)
        txt = await js("document.querySelector('.welcome').textContent")
        for kw in ("智能操作", "识别入库", "数据仓库", "表格填写", "文字识别",
                   "繁体", "竖排", "客服", "聊天室", "意见反馈", "语音"):
            check("指南含「%s」" % kw, kw in str(txt))
        chips = await js("document.querySelectorAll('.welcome .chip').length")
        check("示例胶囊按钮已移除", chips == 0, chips)

        print("\n【3】指南条目点击直达对应面板")
        r1 = await js("(()=>{document.querySelector('.welcome .guide p[data-go=\\'ocr-toggle\\']').click();"
                      "return document.getElementById('ocr-overlay').classList.contains('open')})()")
        check("点「文字识别」打开 OCR 弹窗", r1 is True, r1)
        await js("document.getElementById('ocr-x')&&document.getElementById('ocr-x').click();1")
        await js("(()=>{const m=document.getElementById('ocr-overlay');if(m)m.classList.remove('open');return 1})()")
        r2 = await js("(()=>{document.querySelector('.welcome .guide p[data-go=\\'dw-toggle\\']').click();"
                      "return document.getElementById('dw-overlay').classList.contains('open')})()")
        check("点「数据仓库」打开数据仓库弹窗", r2 is True, r2)
        await js("(()=>{const m=document.getElementById('dw-overlay');if(m)m.classList.remove('open');return 1})()")
        r3 = await js("(()=>{document.querySelector('.welcome .guide p[data-go=\\'input\\']').click();"
                      "return document.activeElement&&document.activeElement.id==='input'})()")
        check("点「直接说」聚焦输入框", r3 is True, r3)
        await js("document.activeElement&&document.activeElement.blur();1")

        print("\n【4】卡片整体更大（内边距 / 双列）")
        pad = await js("getComputedStyle(document.querySelector('.welcome')).padding")
        cols = await js("getComputedStyle(document.querySelector('.welcome .guide')).gridTemplateColumns")
        check("内边距加大（≥24px）", "24px" in str(pad), pad)
        check("指南双列布局", str(cols).count(" ") >= 1 and "0px" in str(cols) or len(str(cols).split()) >= 2, cols)

        await shot("欢迎页-改版-1107.png")

    finally:
        try:
            if ws:
                await ws.close()
        except Exception:
            pass
        proc.terminate()
        try:
            proc.wait(timeout=6)
        except Exception:
            proc.kill()

    print("\n" + "=" * 46)
    print("通过 %d 项 · 失败 %d 项" % (len(PASS), len(FAIL)))
    if FAIL:
        print("失败项：", "、".join(FAIL))
    print("=" * 46)
    return 0 if not FAIL else 1


if __name__ == "__main__":
    raise SystemExit(asyncio.run(main()))
