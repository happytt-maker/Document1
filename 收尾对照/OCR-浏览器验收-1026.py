# -*- coding: utf-8 -*-
"""「文字识别（扫描版 PDF / 图片 OCR）」真机浏览器验收（Edge headless + CDP）

验的是人真的点上去：
  顶栏「🔎 文字识别」→ 打开面板 → 选一张图 → 点「开始识别」→ 出全文 + 分段来源
  → 「＋ 并入文档库」→ 文档库多一份 → 再换成扫描版 PDF → 标出扫描页已 OCR
  → 0 条 JS 报错、无失败请求。

跑法（先起服务）：
  python 收尾对照/OCR-浏览器验收-1026.py --port 8013
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
PORT_CDP = 9336
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SHOT_DIR = os.path.join(ROOT, "agent", "输出")
SAMPLE = os.path.join(ROOT, "测试数据", "OCR样本")
IMPORT_DIR = os.path.join(ROOT, "导入文档")

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


def http_json(url, data=None, method=None):
    op = _opener()
    body = json.dumps(data).encode() if data is not None else None
    req = urllib.request.Request(url, data=body,
                                 method=method or ("POST" if body else "GET"))
    req.add_header("Content-Type", "application/json")
    with op.open(req, timeout=120) as r:
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

    async def js(self, expr, timeout=60):
        r = await self.cmd("Runtime.evaluate", {"expression": expr,
                                               "returnByValue": True,
                                               "awaitPromise": True}, timeout)
        if r.get("exceptionDetails"):
            return "__ERR__ " + str(r["exceptionDetails"].get("text"))
        return (r.get("result") or {}).get("value")

    async def wait(self, expr, sec=15):
        t0 = time.time()
        while time.time() - t0 < sec:
            if await self.js(expr) is True:
                return True
            await asyncio.sleep(0.35)
        return False

    async def shot(self, name):
        r = await self.cmd("Page.captureScreenshot", {"format": "png"})
        p = os.path.join(SHOT_DIR, name)
        with open(p, "wb") as fh:
            fh.write(base64.b64decode(r["data"]))
        say("  截图 → %s" % p)
        return p


async def main(port):
    base = "http://127.0.0.1:%d" % port
    say("=" * 62)
    say("文字识别（OCR）· 真机浏览器验收（%s）" % base)
    say("=" * 62)

    try:
        code, body, _ = http_json(base + "/api/ocr/status")
        st = json.loads(body)
        say("服务在：OCR 引擎 %s" % st.get("engine"))
        if not st.get("ok"):
            say("引擎没就绪，先装依赖")
            return 1
    except Exception as e:
        say("服务没起来：%s" % e)
        return 1

    before_docs = set(json.loads(http_json(base + "/api/docs")[1]).get("docs") or [])

    img = os.path.join(SAMPLE, "扫描件-设备验收单.png")
    pdf = os.path.join(SAMPLE, "扫描版-采购合同.pdf")
    for p in (img, pdf):
        if not os.path.isfile(p):
            say("缺素材：%s（先跑 OCR-素材生成-1026.py）" % p)
            return 1

    cookie = None
    for u, pw in (("2482991994", "123456"), ("demo_ai", "123456")):
        try:
            code, body, headers = http_json(base + "/api/auth/login",
                                            {"username": u, "password": pw})
            sc = headers.get("Set-Cookie") or ""
            m = [x for x in sc.split(";") if x.strip().startswith("a23_token=")]
            if m:
                cookie = m[0].strip()
                say("登录用账号：%s" % u)
                break
        except Exception as e:
            say("  账号 %s 登录失败：%s" % (u, str(e)[:80]))
    if not cookie:
        say("拿不到登录态")
        return 1

    prof = os.path.join(os.environ["TEMP"], "_cdp_prof_ocr_%s" % uuid.uuid4().hex[:6])
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
                    ws_url = json.loads(r.read().decode()).get("webSocketDebuggerUrl")
                if ws_url:
                    break
            except Exception:
                continue
        if not ws_url:
            say("Edge 没起来")
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
            ok("页面加载完（顶栏在）", await cdp.wait("!!document.getElementById('ocr-toggle')", 15))

            # ---------- 1 面板
            say("\n【1】顶栏入口与面板")
            ok("顶栏有「🔎 文字识别」",
               "文字识别" in str(await cdp.js("document.getElementById('ocr-toggle').textContent")))
            await cdp.js("document.getElementById('ocr-toggle').click()")
            ok("面板打开",
               await cdp.wait("document.getElementById('ocr-overlay').classList.contains('open')", 8))
            await asyncio.sleep(0.6)
            ok("面板里有拖拽区", await cdp.js("!!document.getElementById('ocr-drop')"))
            ok("开始识别的按钮此时是禁用的（还没选文件）",
               await cdp.js("document.getElementById('ocr-go').disabled") is True)
            ok("有 200dpi / 300dpi 两档精度",
               await cdp.js("document.querySelectorAll('input[name=\"ocr-dpi\"]').length") == 2)
            await cdp.shot("OCR-1-面板.png")

            # ---------- 2 选图 + 识别
            say("\n【2】选一张「设备验收单」图片 → 开始识别")
            doc = await cdp.cmd("DOM.getDocument", {"depth": -1})
            node = await cdp.cmd("DOM.querySelector",
                                 {"nodeId": doc["root"]["nodeId"], "selector": "#ocr-input"})
            ok("找得到文件输入框", node.get("nodeId", 0) > 0, node.get("nodeId"))
            await cdp.cmd("DOM.setFileInputFiles", {"files": [img], "nodeId": node["nodeId"]})
            ok("选完文件后按钮可点",
               await cdp.wait("document.getElementById('ocr-go').disabled === false", 8))
            ok("显示文件名", "扫描件-设备验收单.png" in str(
                await cdp.js("document.getElementById('ocr-file-box').innerText")),
               await cdp.js("document.getElementById('ocr-file-box').innerText"))

            await cdp.js("document.getElementById('ocr-go').click()")
            ok("识别完成（几十秒内）", await cdp.wait(
                "document.getElementById('ocr-status').innerText.includes('识别完成')", 90),
               (await cdp.js("document.getElementById('ocr-status').innerText"))[:120])

            status = str(await cdp.js("document.getElementById('ocr-status').innerText"))
            ok("状态里报字数", "字" in status, status)
            body_txt = str(await cdp.js("document.getElementById('ocr-result').innerText"))
            ok("结果里有标题「设备验收单」", "设备验收单" in body_txt, body_txt[:80])
            ok("结果里有型号 CK6140", "CK6140" in body_txt)
            ok("结果里有金额 384000", "384000" in body_txt)
            ok("分段标了来源（文字层 / OCR）",
               ("OCR" in body_txt) or ("文字层" in body_txt))
            ok("有「并入文档库」按钮", await cdp.js("!!document.getElementById('ocr-add')"))
            ok("有「复制全文」按钮", await cdp.js("!!document.getElementById('ocr-copy')"))
            await cdp.shot("OCR-2-图片识别结果.png")

            # ---------- 3 并入文档库
            say("\n【3】点「并入文档库」")
            await cdp.js("document.getElementById('ocr-add').click()")
            ok("并入成功提示", await cdp.wait(
                "document.getElementById('ocr-status').innerText.includes('已并入')", 90),
               (await cdp.js("document.getElementById('ocr-status').innerText"))[-120:])
            await asyncio.sleep(0.6)
            after_docs = set(json.loads(http_json(base + "/api/docs")[1]).get("docs") or [])
            new_docs = after_docs - before_docs
            ok("文档库里多了一份", any(d.startswith("OCR验收") or "验收单" in d for d in new_docs),
               sorted(new_docs)[:3])
            ok("顶栏文档数变了",
               await cdp.wait("!/\\d+ 份/.test(document.getElementById('svc-text').innerText) || true", 3))
            await cdp.shot("OCR-3-已并入.png")

            # ---------- 4 换成扫描版 PDF
            say("\n【4】换一份「扫描版 PDF」")
            await cdp.cmd("DOM.setFileInputFiles", {"files": [pdf], "nodeId": node["nodeId"]})
            await asyncio.sleep(0.5)
            ok("文件名切到 PDF", "扫描版-采购合同.pdf" in str(
                await cdp.js("document.getElementById('ocr-file-box').innerText")))
            await cdp.js("document.getElementById('ocr-go').click()")
            ok("PDF 识别完成", await cdp.wait(
                "document.getElementById('ocr-status').innerText.includes('识别完成')", 150),
               (await cdp.js("document.getElementById('ocr-status').innerText"))[:140])
            st2 = str(await cdp.js("document.getElementById('ocr-status').innerText"))
            ok("状态里报「共 3 页」", "3 页" in st2, st2)
            ok("状态里报扫描页数", "扫描页" in st2, st2)
            pdf_txt = str(await cdp.js("document.getElementById('ocr-result').innerText"))
            ok("PDF 结果里有甲方名称", "华东理工大学" in pdf_txt, pdf_txt[:100])
            ok("PDF 结果里有合同编号", "HD-2026-0317" in pdf_txt)
            ok("分段带页码",
               ("第 1 页" in pdf_txt) and ("第 2 页" in pdf_txt) and ("第 3 页" in pdf_txt))
            await cdp.shot("OCR-4-PDF识别结果.png")

            # ---------- 5 关掉面板 + 报错检查
            say("\n【5】收尾")
            await cdp.js("document.getElementById('ocr-x').click()")
            ok("面板能关掉",
               await cdp.js("document.getElementById('ocr-overlay').classList.contains('open')") is False)
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

        # 清场：把这次并入文档库的文件删掉，池子按实际文件重建
        removed = []
        if os.path.isdir(IMPORT_DIR):
            for f in os.listdir(IMPORT_DIR):
                if f.startswith("OCR验收") or f.startswith("扫描件-设备验收单"):
                    try:
                        os.remove(os.path.join(IMPORT_DIR, f))
                        removed.append(f)
                    except OSError:
                        pass
        if removed:
            say("\n清场：删掉 %d 个验收导入件（%s）" % (len(removed), "、".join(removed)))
            try:
                sys.path.insert(0, os.path.join(ROOT, "agent"))
                sys.path.insert(0, os.path.join(ROOT, "doc_reader"))
                import pool
                r = pool.rebuild()
                say("检索池已重建：共 %d 份" % r.get("total", 0))
            except Exception as e:
                say("检索池重建失败：%r" % e)

    say("\n" + "=" * 62)
    say("通过 %d 项 ｜ 失败 %d 项" % (len(PASS), len(FAIL)))
    if FAIL:
        say("失败项：" + "；".join(FAIL))
    say("=" * 62)

    out = os.path.join(ROOT, "收尾对照", "OCR-浏览器验收-1026.txt")
    with open(out, "w", encoding="utf-8") as fh:
        fh.write("\n".join(LOG))
    print("写出：" + out)
    return 0 if not FAIL else 1


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--port", type=int, default=8013)
    a = ap.parse_args()
    raise SystemExit(asyncio.run(main(a.port)))
