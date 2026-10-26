# -*- coding: utf-8 -*-
"""「我上传的模板 · 删除」真机浏览器验收（Edge headless + CDP）

验的是人真的点上去：
  填表面板 → 🗂 我上传的模板 → 空清单提示 → 传一张 → 清单出现这一条
  → 点「删除」→ 这一条消失 → 下拉框同步少一条 → 0 条 JS 报错。
"""
import argparse
import asyncio
import base64
import json
import os
import shutil
import subprocess
import sys
import time
import urllib.request
import uuid

EDGE = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
PORT_CDP = 9334
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SHOT_DIR = os.path.join(ROOT, "agent", "输出")
TMP_SRC = os.path.join(os.environ["TEMP"], "删除验收模板-%s.xlsx" % uuid.uuid4().hex[:6])

PASS, FAIL, LOG, JSErr, NETFAIL = [], [], [], [], []


def say(s):
    LOG.append(s)
    print(s)


def ok(name, cond, extra=""):
    (PASS if cond else FAIL).append(name)
    say("%s %s%s" % ("[通过]" if cond else "[失败]", name,
                     ("  —— " + str(extra)) if extra else ""))


def _opener():
    return urllib.request.build_opener(urllib.request.ProxyHandler({}))


def http_json(url, data=None, method=None, cookie=None):
    op = _opener()
    body = json.dumps(data).encode() if data is not None else None
    req = urllib.request.Request(url, data=body,
                                 method=method or ("POST" if body else "GET"))
    req.add_header("Content-Type", "application/json")
    if cookie:
        req.add_header("Cookie", cookie)
    with op.open(req, timeout=60) as r:
        return r.status, r.read().decode("utf-8", "replace"), r.headers


class CDP:
    def __init__(self, ws):
        self.ws = ws
        self.n = 0

    async def cmd(self, method, params=None, timeout=40):
        self.n += 1
        mid = self.n
        await self.ws.send(json.dumps({"id": mid, "method": method,
                                       "params": params or {}}))
        t0 = time.time()
        while True:
            if time.time() - t0 > timeout:
                raise TimeoutError(method)
            raw = await asyncio.wait_for(self.ws.recv(), timeout=timeout)
            m = json.loads(raw)
            if m.get("method") == "Page.javascriptDialogOpening":
                await self.ws.send(json.dumps({
                    "id": 900000 + self.n, "method": "Page.handleJavaScriptDialog",
                    "params": {"accept": True}}))
                say("  [dialog] 自动点确定：%s"
                    % m.get("params", {}).get("message", "")[:50].replace("\n", " "))
                continue
            if m.get("method") == "Runtime.exceptionThrown":
                d = m.get("params", {}).get("exceptionDetails", {})
                JSErr.append(str(d.get("text")) + " "
                             + str((d.get("exception") or {}).get("description", ""))[:160])
                continue
            if m.get("method") == "Runtime.consoleAPICalled":
                if m.get("params", {}).get("type") == "error":
                    JSErr.append("console.error: "
                                 + json.dumps(m["params"].get("args"),
                                              ensure_ascii=False)[:160])
                continue
            if m.get("method") == "Log.entryAdded":
                e = m.get("params", {}).get("entry", {})
                if e.get("level") == "error":
                    u = str(e.get("url") or "")
                    if "favicon" in u:
                        continue
                    JSErr.append("log: " + str(e.get("text"))[:160] + (" @ " + u if u else ""))
                continue
            if m.get("method") == "Network.responseReceived":
                resp = m.get("params", {}).get("response", {})
                if int(resp.get("status") or 0) >= 400:
                    NETFAIL.append("%s %s" % (resp.get("status"), resp.get("url")))
                continue
            if m.get("id") == mid:
                if "error" in m:
                    raise RuntimeError("%s -> %s" % (method, m["error"]))
                return m.get("result", {})

    async def js(self, expr, timeout=40):
        r = await self.cmd("Runtime.evaluate", {"expression": expr,
                                                "returnByValue": True,
                                                "awaitPromise": True}, timeout)
        if r.get("exceptionDetails"):
            return "__ERR__ " + str(r["exceptionDetails"].get("text"))
        return (r.get("result") or {}).get("value")

    async def wait(self, expr, sec=10):
        t0 = time.time()
        while time.time() - t0 < sec:
            if await self.js(expr) is True:
                return True
            await asyncio.sleep(0.35)
        return False

    async def shot(self, name):
        r = await self.cmd("Page.captureScreenshot", {"format": "png"})
        p = os.path.join(SHOT_DIR, name)
        open(p, "wb").write(base64.b64decode(r["data"]))
        say("  截图 → %s" % p)
        return p


async def main(port):
    base = "http://127.0.0.1:%d" % port
    say("=" * 62)
    say("模板删除（我上传的模板）· 真机浏览器验收（%s）" % base)
    say("=" * 62)

    try:
        code, body, _ = http_json(base + "/api/fill/templates")
        base_cnt = len(json.loads(body)["templates"])
        say("服务在：模板 %d 张" % base_cnt)
    except Exception as e:
        say("服务没起来：%s" % e)
        return 1

    src = os.path.join(ROOT, "测试数据", "模板表格", "模板1-设备采购汇总表.xlsx")
    shutil.copyfile(src, TMP_SRC)
    up_name = os.path.basename(TMP_SRC)

    cookie = None
    for u, p in (("2482991994", "123456"), ("demo_ai", "123456")):
        try:
            code, body, headers = http_json(base + "/api/auth/login",
                                            {"username": u, "password": p})
            sc = headers.get("Set-Cookie") or ""
            m = [x for x in sc.split(";") if x.strip().startswith("a23_token=")]
            if m:
                cookie = m[0].strip()
                say("登录用账号：%s" % u)
                break
        except Exception as e:
            say("  账号 %s 登录失败：%s" % (u, str(e)[:80]))
    if not cookie:
        say("拿不到登录态，浏览器会停在登录页")
        return 1

    prof = os.path.join(os.environ["TEMP"], "_cdp_prof_del_%s" % uuid.uuid4().hex[:6])
    proc = subprocess.Popen([
        EDGE, "--headless=new", "--disable-gpu", "--no-first-run",
        "--no-default-browser-check", "--disable-extensions", "--disable-sync",
        "--disable-background-networking", "--disable-component-update",
        "--no-proxy-server", "--remote-allow-origins=*",
        "--remote-debugging-port=%d" % PORT_CDP, "--user-data-dir=" + prof,
        "--window-size=1440,980", "about:blank",
    ], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

    try:
        ws_url = None
        for _ in range(50):
            time.sleep(0.4)
            try:
                with _opener().open("http://127.0.0.1:%d/json/version" % PORT_CDP,
                                    timeout=3) as r:
                    j = json.loads(r.read().decode())
                ws_url = j.get("webSocketDebuggerUrl")
                if ws_url:
                    break
            except Exception:
                continue
        if not ws_url:
            say("Edge 没起来（CDP 端口连不上）")
            return 1

        req = urllib.request.Request("http://127.0.0.1:%d/json/new?about:blank" % PORT_CDP,
                                     method="PUT")
        with _opener().open(req, timeout=10) as r:
            page = json.loads(r.read().decode())

        from websockets.asyncio.client import connect
        async with connect(page["webSocketDebuggerUrl"], max_size=None) as ws:
            cdp = CDP(ws)
            for m in ("Page.enable", "Runtime.enable", "Network.enable",
                      "DOM.enable", "Log.enable"):
                await cdp.cmd(m)

            await cdp.cmd("Network.setCookie", {"name": "a23_token",
                                                "value": cookie.split("=", 1)[1],
                                                "url": base, "path": "/",
                                                "httpOnly": True})
            await cdp.cmd("Page.navigate", {"url": base + "/"})
            await asyncio.sleep(3.2)
            ok("页面加载完（顶栏在）", await cdp.wait("!!document.getElementById('fill-toggle')", 12))

            # ---------- 1 打开填表面板 + 管理入口
            say("\n【1】填表面板与「我上传的模板」入口")
            await cdp.js("document.getElementById('fill-toggle').click()")
            ok("填表面板打开",
               await cdp.wait("document.getElementById('fill-overlay').classList.contains('open')", 8))
            await asyncio.sleep(1.0)
            ok("面板顶上有「🗂 我上传的模板」按钮", await cdp.js("!!document.getElementById('fill-mg-btn')"))
            ok("按钮文字对得上",
               "我上传的模板" in str(await cdp.js("document.getElementById('fill-mg-btn').textContent")))

            await cdp.js("document.getElementById('fill-mg-btn').click()")
            ok("清单弹窗打开",
               await cdp.wait("document.getElementById('tplmg-overlay').classList.contains('open')", 8))
            ok("没传过时给的是「还没有上传过模板」提示",
               await cdp.wait("document.getElementById('tplmg-list').innerText.includes('还没有上传过模板')", 8),
               await cdp.js("document.getElementById('tplmg-list').innerText"))
            ok("没传过时没有删除行",
               await cdp.js("document.querySelectorAll('#tplmg-list .tplmg-row').length") == 0)
            await cdp.shot("模板删除-1-空清单.png")
            await cdp.js("document.getElementById('tplmg-x').click()")
            await asyncio.sleep(0.4)
            ok("弹窗能关掉",
               await cdp.js("document.getElementById('tplmg-overlay').classList.contains('open')") is False)

            # ---------- 2 传一张
            say("\n【2】先传一张模板")
            doc = await cdp.cmd("DOM.getDocument", {"depth": -1})
            node = await cdp.cmd("DOM.querySelector",
                                 {"nodeId": doc["root"]["nodeId"], "selector": "#fill-file"})
            ok("找得到文件输入框", node.get("nodeId", 0) > 0, node.get("nodeId"))
            await cdp.cmd("DOM.setFileInputFiles", {"files": [TMP_SRC],
                                                    "nodeId": node["nodeId"]})
            ok("上传完成（出现「已上传」提示）",
               await cdp.wait("document.getElementById('fill-up-tip').innerText.includes('已上传')", 25),
               await cdp.js("document.getElementById('fill-up-tip').innerText"))

            # ---------- 3 清单里出现这一条
            say("\n【3】清单里能看到刚上传的那张")
            await cdp.js("document.getElementById('fill-mg-btn').click()")
            ok("弹窗打开",
               await cdp.wait("document.getElementById('tplmg-overlay').classList.contains('open')", 8))
            ok("清单里出现 1 条", await cdp.wait(
                "document.querySelectorAll('#tplmg-list .tplmg-row').length === 1", 8),
               await cdp.js("document.querySelectorAll('#tplmg-list .tplmg-row').length"))
            row_txt = await cdp.js("document.querySelector('#tplmg-list .tplmg-row').innerText")
            ok("这一条就是刚传的那张", up_name in str(row_txt), row_txt)
            ok("这一条带大小 / 时间", ("KB" in str(row_txt) or "MB" in str(row_txt)) and ":" in str(row_txt), row_txt)
            ok("计数显示「共 1 张」", "1 张" in str(await cdp.js("document.getElementById('tplmg-cnt').innerText")),
               await cdp.js("document.getElementById('tplmg-cnt').innerText"))
            ok("系统自带模板没混进来",
               await cdp.js("document.getElementById('tplmg-list').innerText.includes('模板1-设备采购')") is False)
            await cdp.shot("模板删除-2-清单有它.png")

            # ---------- 4 点删除
            say("\n【4】点这一条的「删除」")
            await cdp.js("document.querySelector('#tplmg-list .tplmg-row button.del').click()")
            ok("删除完成（出现「已删除」）", await cdp.wait(
                "document.getElementById('tplmg-msg').innerText.includes('已删除')", 20),
               await cdp.js("document.getElementById('tplmg-msg').innerText"))
            ok("清单回到空状态", await cdp.wait(
                "document.getElementById('tplmg-list').innerText.includes('还没有上传过模板')", 8),
               (await cdp.js("document.getElementById('tplmg-list').innerText"))[:80])
            ok("删除行也没了",
               await cdp.js("document.querySelectorAll('#tplmg-list .tplmg-row').length") == 0)
            await asyncio.sleep(0.5)
            ok("下拉框数量回到原来的 %d 张" % base_cnt,
               await cdp.js("document.getElementById('fill-tpl').options.length") == base_cnt,
               await cdp.js("document.getElementById('fill-tpl').options.length"))
            ok("上传目录 / 接口都没它了",
               up_name not in str(await cdp.js(
                   "fetch('/api/fill/templates').then(r=>r.json()).then(j=>j.templates.map(t=>t.name).join('|'))")),
               "接口已刷新")
            await cdp.shot("模板删除-3-删完.png")

            # ---------- 5 面板上的「删除」按钮仍然能删（选中态那条路）
            say("\n【5】关掉弹窗后，面板里的「删除」按钮也不受影响")
            await cdp.js("document.getElementById('tplmg-x').click()")
            await asyncio.sleep(0.4)
            await cdp.cmd("DOM.setFileInputFiles", {"files": [TMP_SRC], "nodeId": node["nodeId"]})
            ok("再传一张", await cdp.wait(
                "document.getElementById('fill-up-tip').innerText.includes('已上传')", 25))
            await asyncio.sleep(0.5)
            ok("传完自动选中、删除按钮可用",
               await cdp.js("document.getElementById('fill-del-btn').disabled") is False)
            await cdp.js("document.getElementById('fill-del-btn').click()")
            ok("面板删除也能删（出现已删除）", await cdp.wait(
                "document.getElementById('fill-up-tip').innerText.includes('已删除')", 20),
               await cdp.js("document.getElementById('fill-up-tip').innerText"))
            ok("下拉框又回到 %d 张" % base_cnt,
               await cdp.js("document.getElementById('fill-tpl').options.length") == base_cnt,
               await cdp.js("document.getElementById('fill-tpl').options.length"))

            ok("全程 0 条 JS 报错", len(JSErr) == 0, JSErr[:3])
            bad_net = [u for u in NETFAIL if "favicon" not in u]
            ok("没有失败的网络请求", len(bad_net) == 0, bad_net[:3])

    finally:
        try:
            proc.terminate()
            proc.wait(timeout=8)
        except Exception:
            pass
        shutil.rmtree(prof, ignore_errors=True)
        try:
            os.remove(TMP_SRC)
        except OSError:
            pass

    say("\n" + "=" * 62)
    say("通过 %d 项 ｜ 失败 %d 项" % (len(PASS), len(FAIL)))
    if FAIL:
        say("失败项：" + "；".join(FAIL))
    say("=" * 62)

    out = os.path.join(ROOT, "收尾对照", "模板删除-浏览器验收-1026.txt")
    open(out, "w", encoding="utf-8").write("\n".join(LOG))
    print("写出：" + out)
    return 0 if not FAIL else 1


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--port", type=int, default=8012)
    a = ap.parse_args()
    raise SystemExit(asyncio.run(main(a.port)))
