# -*- coding: utf-8 -*-
"""模板上传 · 真机浏览器验收（Edge headless + CDP）

验的是「人真的点上去能不能用」：
  面板上有上传按钮 → 选文件 → 下拉框自动选中新模板 → 删除按钮可用
  → 删掉 → 清单恢复。全程记录 JS 报错。
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

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

EDGE = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
PORT = 9333
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SHOT_DIR = os.path.join(ROOT, "agent", "输出")
TMP_SRC = os.path.join(os.environ["TEMP"], "验收模板-%s.xlsx" % uuid.uuid4().hex[:6])

PASS, FAIL, LOG, JSErr = [], [], [], []
NETFAIL = []


def say(s):
    LOG.append(s)
    print(s)


def ok(name, cond, extra=""):
    (PASS if cond else FAIL).append(name)
    say("%s %s%s" % ("[通过]" if cond else "[失败]", name, ("  —— " + str(extra)) if extra else ""))


def _opener():
    return urllib.request.build_opener(urllib.request.ProxyHandler({}))


def http_json(url, data=None, method=None, cookie=None):
    op = _opener()
    body = json.dumps(data).encode() if data is not None else None
    req = urllib.request.Request(url, data=body, method=method or ("POST" if body else "GET"))
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
        await self.ws.send(json.dumps({"id": mid, "method": method, "params": params or {}}))
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
                say("  [dialog] 自动确认：%s" % (m.get("params", {}).get("message", "")[:60]))
                continue
            if m.get("method") == "Runtime.exceptionThrown":
                d = m.get("params", {}).get("exceptionDetails", {})
                JSErr.append(str(d.get("text")) + " " + str((d.get("exception") or {}).get("description", ""))[:160])
                continue
            if m.get("method") == "Runtime.consoleAPICalled":
                if m.get("params", {}).get("type") == "error":
                    JSErr.append("console.error: " + json.dumps(m["params"].get("args"), ensure_ascii=False)[:160])
                continue
            if m.get("method") == "Log.entryAdded":
                e = m.get("params", {}).get("entry", {})
                if e.get("level") == "error":
                    u = str(e.get("url") or "")
                    if "favicon" in u:          # 页面图标，与业务无关，不算进来
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
        r = await self.cmd("Runtime.evaluate", {
            "expression": expr, "returnByValue": True, "awaitPromise": True}, timeout=timeout)
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
    say("模板上传 · 真机浏览器验收（%s）" % base)
    say("=" * 62)

    # ---------- 服务探测
    try:
        code, body, _ = http_json(base + "/api/fill/templates")
        say("服务在：/api/fill/templates -> %s，模板 %d 张" % (code, len(json.loads(body)["templates"])))
    except Exception as e:
        say("服务没起来：%s" % e)
        return 1

    # ---------- 准备一份待上传的模板副本
    src = os.path.join(ROOT, "测试数据", "模板表格", "模板1-设备采购汇总表.xlsx")
    shutil.copyfile(src, TMP_SRC)
    up_name = os.path.basename(TMP_SRC)
    say("待上传文件：%s（%d 字节）" % (up_name, os.path.getsize(TMP_SRC)))

    # ---------- 登录拿 cookie
    cookie = None
    for u, p in (("2482991994", "123456"), ("demo_ai", "123456")):
        try:
            code, body, headers = http_json(base + "/api/auth/login", {"username": u, "password": p})
            sc = headers.get("Set-Cookie") or ""
            m = [x for x in sc.split(";") if x.strip().startswith("a23_token=")]
            if m:
                cookie = m[0].strip()
                say("登录用账号：%s" % u)
                break
        except Exception as e:
            say("  账号 %s 登录失败：%s" % (u, str(e)[:80]))
    if not cookie:
        say("拿不到登录态，浏览器会停在登录页 —— 验收没法继续")
        return 1

    # ---------- 拉 browser
    prof = os.path.join(os.environ["TEMP"], "_cdp_prof_up_%s" % uuid.uuid4().hex[:6])
    proc = subprocess.Popen([
        EDGE, "--headless=new", "--disable-gpu", "--no-first-run", "--no-default-browser-check",
        "--disable-extensions", "--disable-sync", "--disable-background-networking",
        "--disable-component-update", "--no-proxy-server", "--remote-allow-origins=*",
        "--remote-debugging-port=%d" % PORT, "--user-data-dir=" + prof,
        "--window-size=1440,980", "about:blank",
    ], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

    try:
        ws_url = None
        for _ in range(50):
            time.sleep(0.4)
            try:
                op = _opener()
                with op.open("http://127.0.0.1:%d/json/version" % PORT, timeout=3) as r:
                    j = json.loads(r.read().decode())
                ws_url = j.get("webSocketDebuggerUrl")
                if ws_url:
                    break
            except Exception:
                continue
        if not ws_url:
            say("Edge 没起来（CDP 端口连不上）")
            return 1

        req = urllib.request.Request("http://127.0.0.1:%d/json/new?about:blank" % PORT, method="PUT")
        with _opener().open(req, timeout=10) as r:
            page = json.loads(r.read().decode())

        from websockets.asyncio.client import connect
        async with connect(page["webSocketDebuggerUrl"], max_size=None) as ws:
            cdp = CDP(ws)
            for m in ("Page.enable", "Runtime.enable", "Network.enable", "DOM.enable", "Log.enable"):
                await cdp.cmd(m)

            await cdp.cmd("Network.setCookie", {"name": "a23_token",
                                                "value": cookie.split("=", 1)[1],
                                                "url": base, "path": "/", "httpOnly": True})
            await cdp.cmd("Page.navigate", {"url": base + "/"})
            await asyncio.sleep(3.2)

            ok("页面加载完（有顶栏）", await cdp.wait("!!document.getElementById('fill-toggle')", 12))

            # ---------- 1 打开填表面板
            say("\n【1】打开填表面板")
            await cdp.js("document.getElementById('fill-toggle').click()")
            ok("面板打开", await cdp.wait("document.getElementById('fill-overlay').classList.contains('open')", 8))
            await asyncio.sleep(1.2)
            ok("看到「⬆ 上传模板」按钮", await cdp.js("!!document.getElementById('fill-up-btn')"))
            ok("看到「删除」按钮", await cdp.js("!!document.getElementById('fill-del-btn')"))
            ok("按钮文字是「⬆ 上传模板」",
               "上传模板" in str(await cdp.js("document.getElementById('fill-up-btn').textContent")))
            ok("删除按钮初始是禁用的", await cdp.js("document.getElementById('fill-del-btn').disabled") is True)
            base_cnt = await cdp.js("document.getElementById('fill-tpl').options.length")
            say("  下拉框里现有 %s 张模板" % base_cnt)
            await cdp.shot("模板上传-1-面板.png")

            # ---------- 2 模拟选文件上传
            say("\n【2】选文件上传（走真实的 change 事件）")
            doc = await cdp.cmd("DOM.getDocument", {"depth": -1})
            node = await cdp.cmd("DOM.querySelector", {"nodeId": doc["root"]["nodeId"],
                                                       "selector": "#fill-file"})
            ok("找得到隐藏的 file input", node.get("nodeId", 0) > 0, node.get("nodeId"))
            await cdp.cmd("DOM.setFileInputFiles", {"files": [TMP_SRC],
                                                    "nodeId": node["nodeId"]})
            got_tip = await cdp.wait(
                "document.getElementById('fill-up-tip').innerText.includes('已上传')", 25)
            ok("出现「✓ 已上传」提示", got_tip,
               await cdp.js("document.getElementById('fill-up-tip').innerText"))
            await asyncio.sleep(0.6)
            sel_val = await cdp.js("document.getElementById('fill-tpl').value")
            ok("下拉框自动选中刚传的那张", sel_val == up_name, sel_val)
            opt_txt = await cdp.js(
                "(function(){var s=document.getElementById('fill-tpl');"
                "return s.options[s.selectedIndex]?s.options[s.selectedIndex].textContent:''})()")
            ok("下拉项标着「我上传的」", "我上传的" in str(opt_txt), opt_txt)
            ok("下拉框数量 +1", await cdp.js("document.getElementById('fill-tpl').options.length") == base_cnt + 1,
               "%s -> %s" % (base_cnt, await cdp.js("document.getElementById('fill-tpl').options.length")))
            ok("删除按钮变成可用", await cdp.js("document.getElementById('fill-del-btn').disabled") is False)
            await cdp.shot("模板上传-2-已上传.png")

            # ---------- 3 接口侧核对确实落盘了
            say("\n【3】落盘核对")
            code, body, _ = http_json(base + "/api/fill/templates")
            tps = json.loads(body)["templates"]
            hit = next((t for t in tps if t["name"] == up_name), None)
            ok("接口清单里有它", hit is not None, hit)
            ok("文件真的写进了 模板上传 目录",
               os.path.isfile(os.path.join(ROOT, "测试数据", "模板上传", up_name)))

            # ---------- 4 删除
            say("\n【4】删除自己上传的模板")
            await cdp.js("document.getElementById('fill-del-btn').click()")
            got_del = await cdp.wait(
                "document.getElementById('fill-up-tip').innerText.includes('已删除')", 20)
            ok("出现「✓ 已删除」提示", got_del,
               await cdp.js("document.getElementById('fill-up-tip').innerText"))
            await asyncio.sleep(0.5)
            ok("下拉框数量回到原来的", await cdp.js("document.getElementById('fill-tpl').options.length") == base_cnt,
               await cdp.js("document.getElementById('fill-tpl').options.length"))
            ok("下拉框里已经没有它",
               up_name not in str(await cdp.js("Array.from(document.getElementById('fill-tpl').options).map(o=>o.value).join('|')")))
            code, body, _ = http_json(base + "/api/fill/templates")
            ok("接口清单也回到 %s 张" % base_cnt, len(json.loads(body)["templates"]) == base_cnt,
               len(json.loads(body)["templates"]))
            ok("上传目录里文件已删",
               not os.path.isfile(os.path.join(ROOT, "测试数据", "模板上传", up_name)))
            await cdp.shot("模板上传-3-已删除.png")

            # ---------- 5 系统模板不可删
            say("\n【5】切到系统自带模板：删除按钮应变灰")
            await cdp.js("(function(){var s=document.getElementById('fill-tpl');"
                         "for(var i=0;i<s.options.length;i++){if(s.options[i].dataset.upload==='0'){"
                         "s.selectedIndex=i;break;}}s.dispatchEvent(new Event('change'));})()")
            await asyncio.sleep(0.5)
            ok("选中系统模板时删除按钮变禁用",
               await cdp.js("document.getElementById('fill-del-btn').disabled") is True)
            ok("按钮 title 解释了为什么不能删",
               "不能删" in str(await cdp.js("document.getElementById('fill-del-btn').title")))

            ok("全程 0 条 JS 报错", len(JSErr) == 0, JSErr[:3])
            bad_net = [u for u in NETFAIL if "favicon" not in u]
            ok("没有失败的网络请求", len(bad_net) == 0, bad_net[:3])
            say("  网络 4xx/5xx：%s" % (NETFAIL if NETFAIL else "无"))

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

    out = os.path.join(ROOT, "收尾对照", "模板上传-浏览器验收-1026.txt")
    open(out, "w", encoding="utf-8").write("\n".join(LOG))
    print("写出：" + out)
    return 0 if not FAIL else 1


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--port", type=int, default=8010)
    a = ap.parse_args()
    raise SystemExit(asyncio.run(main(a.port)))
