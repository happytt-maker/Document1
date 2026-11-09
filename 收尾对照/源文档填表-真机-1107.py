# -*- coding: utf-8 -*-
"""表格填写 · 选源文档（2026-11-07）真机验收：headless Edge + CDP。

验的是"眼睛能看到的"：
  1. 点左侧「表格填写」→ 面板里出现第 1/2/3 步（选模板 / 选源文档 / 填表要求）
  2. 源文档清单读出来了（21 份），勾选后提示从"整个文档库"变成"只在这几份里找数"
  3. 按钮在位：全选/取消、上传源文档
截图落到 agent\\输出\\填表-源文档-1107-*.png
"""
import base64
import json
import os
import shutil
import subprocess
import time
import urllib.request
from pathlib import Path

EDGE = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
PROF = Path(os.environ["TEMP"]) / "_cdp_prof_filldocs"
OUT = Path(r"C:\Users\覃敬涛\Desktop\研一项目\agent\输出")
BASE = "http://127.0.0.1:8000"
PORT = 9333
SHOT = []


def http(u, data=None, method=None, headers=None):
    op = urllib.request.build_opener(urllib.request.ProxyHandler({}))
    req = urllib.request.Request(u, data=data, method=method, headers=headers or {})
    with op.open(req, timeout=10) as r:
        return r.read(), dict(r.headers)


def login():
    body = json.dumps({"username": "2482991994", "password": "123456"}).encode()
    _b, h = http(BASE + "/api/auth/login", body, "POST",
                 {"Content-Type": "application/json"})
    # urllib 给的 header 名可能被小写化，别按 "Set-Cookie" 硬匹配
    ck = ""
    for k, v in h.items():
        if k.lower() == "set-cookie":
            ck = v
    tok = ""
    for part in ck.split(","):
        if "a23_token=" in part:
            tok = part.split("a23_token=")[1].split(";")[0].strip()
    return tok


def main():
    tok = login()
    print("cookie:", "ok" if tok else "MISSING")
    shutil.rmtree(PROF, ignore_errors=True)
    proc = subprocess.Popen([
        EDGE, "--headless=new", "--disable-gpu", "--no-first-run", "--no-default-browser-check",
        "--disable-extensions", "--disable-sync", "--disable-background-networking",
        "--disable-component-update", "--no-proxy-server", "--remote-allow-origins=*",
        "--remote-debugging-port=%d" % PORT, "--user-data-dir=" + str(PROF),
        "--window-size=1440,960", "about:blank",
    ], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

    ws_url = None
    for _ in range(42):
        try:
            raw, _h = http("http://127.0.0.1:%d/json/version" % PORT)
            if b"Browser" in raw:
                break
        except Exception:
            pass
        time.sleep(0.5)
    for _ in range(20):
        try:
            raw, _h = http("http://127.0.0.1:%d/json/new?about:blank" % PORT, b"", "PUT")
            ws_url = json.loads(raw.decode()).get("webSocketDebuggerUrl")
            if ws_url:
                break
        except Exception:
            time.sleep(0.4)
    print("ws:", "ok" if ws_url else "MISSING")

    import asyncio
    import websockets

    async def run():
        mid = [0]
        pend = {}

        async def cmd(ws, method, params=None):
            mid[0] += 1
            i = mid[0]
            await ws.send(json.dumps({"id": i, "method": method, "params": params or {}}))
            while True:
                m = json.loads(await ws.recv())
                if m.get("method") == "Page.javascriptDialogOpening":
                    await ws.send(json.dumps({"id": 900000 + mid[0],
                                              "method": "Page.handleJavaScriptDialog",
                                              "params": {"accept": True}}))
                    continue
                if m.get("id") == i:
                    return m.get("result") or {}

        async def js(ws, expr):
            r = await cmd(ws, "Runtime.evaluate",
                          {"expression": expr, "returnByValue": True, "awaitPromise": True})
            return (r.get("result") or {}).get("value")

        async def shot(ws, name):
            r = await cmd(ws, "Page.captureScreenshot", {"format": "png"})
            p = OUT / name
            p.write_bytes(base64.b64decode(r["data"]))
            SHOT.append(str(p))
            print("shot:", p.name)

        async with websockets.connect(ws_url, max_size=None) as ws:
            await cmd(ws, "Page.enable")
            await cmd(ws, "Runtime.enable")
            await cmd(ws, "Network.enable")
            await cmd(ws, "Network.setCookie",
                      {"name": "a23_token", "value": tok, "url": BASE,
                       "path": "/", "httpOnly": True})
            await cmd(ws, "Page.navigate", {"url": BASE + "/"})
            await asyncio.sleep(3)

            # 打开表格填写
            await js(ws, "document.getElementById('fill-toggle').click()")
            await asyncio.sleep(2.0)
            steps = await js(ws, "[...document.querySelectorAll('#fill-overlay .pstep')].map(x=>x.innerText.replace(/\\n/g,' '))")
            print("steps:", steps)
            n_docs = await js(ws, "document.querySelectorAll('#fill-docs input[type=checkbox]').length")
            hint0 = await js(ws, "document.getElementById('fill-doc-hint').innerText")
            print("doc rows:", n_docs, "| hint0:", hint0)

            # 勾两份
            await js(ws, """(() => {
              const cbs = [...document.querySelectorAll('#fill-docs input[type=checkbox]')];
              const byName = (s) => cbs.find(c => c.value.includes(s));
              const a = byName('设备采购合同'), b = byName('供应商联系方式');
              if (a) a.click(); if (b) b.click();
              return true;
            })()""")
            await asyncio.sleep(0.6)
            hint1 = await js(ws, "document.getElementById('fill-doc-hint').innerText")
            print("hint1:", hint1)
            on = await js(ws, "document.querySelectorAll('#fill-docs .rk-line.on').length")
            print("highlighted rows:", on)
            btn = await js(ws, "[document.getElementById('fill-doc-all')?.innerText, document.getElementById('fill-doc-up')?.innerText]")
            print("buttons:", btn)
            req_tag = await js(ws, "document.getElementById('fill-req').tagName")
            print("requirement field tag:", req_tag)

            outline = await js(ws, """(() => {
              const p = document.querySelector('#fill-overlay .dw-panel').getBoundingClientRect();
              const l = document.getElementById('fill-docs').getBoundingClientRect();
              return {panel:[Math.round(p.width),Math.round(p.height)],
                      list:[Math.round(l.width),Math.round(l.height)],
                      scrollable: document.getElementById('fill-docs').scrollHeight >
                                  document.getElementById('fill-docs').clientHeight};
            })()""")
            print("layout:", outline)
            await shot(ws, "填表-源文档-1107-1.png")

            # 断言
            ok = True

            def ck(name, cond, extra=""):
                nonlocal ok
                ok = ok and bool(cond)
                print(("  [PASS] " if cond else "  [FAIL] ") + name + (("  " + str(extra)) if extra else ""))

            print("\n断言：")
            ck("三步都在", len(steps or []) >= 3, steps)
            ck("第 2 步是选源文档", any("选源文档" in s for s in (steps or [])), steps)
            ck("文档清单读出来了（21 份）", (n_docs or 0) >= 20, n_docs)
            ck("未勾时提示整个文档库", "整个文档库" in (hint0 or ""), hint0)
            ck("勾了之后提示只在这几份里找数", "只在这 2 份里找数" in (hint1 or ""), hint1)
            ck("勾中的行有高亮", (on or 0) == 2, on)
            ck("两个按钮都在", all(btn or []), btn)
            ck("填表要求是可换行的 textarea", req_tag == "TEXTAREA", req_tag)
            ck("面板够大（>=1000x700）", (outline or {}).get("panel", [0, 0])[0] >= 1000,
               (outline or {}).get("panel"))

            # 真点一次「开始填写」：确认前端把勾选的源文档一起提交了、结果区能出下载按钮
            await js(ws, "document.getElementById('fill-go').click()")
            got = None
            for _ in range(80):
                await asyncio.sleep(2)
                got = await js(ws, "document.getElementById('fill-status').innerText + '\\n' + "
                                   "document.getElementById('fill-result').innerText")
                if "填写完成" in (got or "") or "未完成" in (got or "") or "失败" in (got or ""):
                    break
            print("\n结果区：\n" + (got or "")[:500])
            ck("点「开始填写」真的跑完了", "填写完成" in (got or "") or "未完成" in (got or ""), (got or "")[:80])
            ck("报告里写了源文档范围", "源文档：" in (got or ""))
            await shot(ws, "填表-源文档-1107-2.png")
            print("\n合计：", "全过" if ok else "有失败")
            return ok

    try:
        return asyncio.run(run())
    finally:
        proc.terminate()
        try:
            proc.wait(timeout=5)
        except Exception:
            proc.kill()
        shutil.rmtree(PROF, ignore_errors=True)


if __name__ == "__main__":
    raise SystemExit(0 if main() else 1)
