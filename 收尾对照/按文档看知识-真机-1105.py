# -*- coding: utf-8 -*-
"""真机验收：数据仓库「按文档看知识」

验三件事：
  1 文档库里每份抽过的文档带「已入库 N」徽章
  2 点文档名 → 进入该文档的知识视图（不是全部混在一起）
  3 视图里能按属性列出全部值（"这份文档里有哪些城市/省份"）

跑法（managed 3.13.12，带 websockets）：
  python -u 按文档看知识-真机-1105.py > 输出.txt 2>&1
"""
import base64
import json
import os
import shutil
import subprocess
import sys
import time
import urllib.parse
import urllib.request

import asyncio
import websockets

BASE = "http://127.0.0.1:8000"
EDGE = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
PROFILE = os.path.join(os.environ["TEMP"], "_cdp_prof_dw")
OUT = r"C:\Users\覃敬涛\Desktop\研一项目\agent\输出"

op = urllib.request.build_opener(urllib.request.ProxyHandler({}))
log = []
PASS = FAIL = 0


def say(s=""):
    log.append(s)
    print(s, flush=True)


def check(name, ok, extra=""):
    global PASS, FAIL
    if ok:
        PASS += 1
        say("  ✓ %s%s" % (name, (" ｜ " + extra) if extra else ""))
    else:
        FAIL += 1
        say("  ✗ %s%s" % (name, (" ｜ " + extra) if extra else ""))


def http_json(path, payload=None):
    data = None
    headers = {}
    if payload is not None:
        data = json.dumps(payload).encode("utf-8")
        headers["Content-Type"] = "application/json"
    req = urllib.request.Request(BASE + path, data=data, headers=headers)
    return op.open(req, timeout=30)


def wait_port(times=40):
    for _ in range(times):
        try:
            http_json("/api/status")
            return True
        except Exception:
            time.sleep(0.5)
    return False


async def shot(cmd, name):
    r = await cmd("Page.captureScreenshot", {"format": "png"})
    b64 = (r.get("result") or {}).get("data") or ""
    if not b64:
        say("    截图失败：%s" % name)
        return
    p = os.path.join(OUT, name)
    with open(p, "wb") as f:
        f.write(base64.b64decode(b64))
    say("    截图 → %s" % p)


async def main():
    global PASS, FAIL
    say("=" * 66)
    say("数据仓库 · 按文档看知识 · 真机验收")
    say("=" * 66)

    if not wait_port():
        say("服务没起来（%s）" % BASE)
        return

    # 登录拿 cookie
    try:
        r = http_json("/api/auth/login",
                      {"username": "2482991994", "password": "123456", "portal": "user"})
        ck = r.headers.get("Set-Cookie") or ""
        token = ""
        for seg in ck.split(";"):
            if seg.strip().startswith("a23_token="):
                token = seg.strip().split("=", 1)[1]
        say("登录 ok ｜ token %s" % ("拿到" if token else "没拿到"))
    except Exception as e:
        say("登录失败：%s" % e)
        token = ""

    proc = subprocess.Popen([
        EDGE, "--headless=new", "--disable-gpu", "--no-first-run", "--no-default-browser-check",
        "--disable-extensions", "--disable-sync", "--disable-background-networking",
        "--disable-component-update", "--no-proxy-server", "--remote-allow-origins=*",
        "--remote-debugging-port=9333", "--user-data-dir=" + PROFILE,
        "--window-size=1440,1100", "about:blank",
    ], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

    try:
        for _ in range(42):
            try:
                info = json.loads(op.open("http://127.0.0.1:9333/json/version", timeout=1).read())
                if info.get("webSocketDebuggerUrl"):
                    break
            except Exception:
                time.sleep(0.5)
        req = urllib.request.Request("http://127.0.0.1:9333/json/new?about:blank", method="PUT")
        tab = json.loads(op.open(req, timeout=5).read())
        ws_url = tab["webSocketDebuggerUrl"]

        async with websockets.connect(ws_url, max_size=None) as ws:
            mid = [0]

            async def cmd(method, params=None):
                mid[0] += 1
                i = mid[0]
                await ws.send(json.dumps({"id": i, "method": method, "params": params or {}}))
                while True:
                    m = json.loads(await ws.recv())
                    if m.get("method") == "Page.javascriptDialogOpening":
                        await ws.send(json.dumps({"id": mid[0] + 1,
                                                  "method": "Page.handleJavaScriptDialog",
                                                  "params": {"accept": True}}))
                        continue
                    if m.get("id") == i:
                        return m

            async def js(expr):
                r = await cmd("Runtime.evaluate", {"expression": expr, "returnByValue": True,
                                                   "awaitPromise": True})
                return ((r.get("result") or {}).get("result") or {}).get("value")

            await cmd("Page.enable")
            await cmd("Runtime.enable")
            await cmd("Network.enable")
            if token:
                await cmd("Network.setCookie", {"name": "a23_token", "value": token,
                                                "url": BASE, "path": "/", "httpOnly": True})
            await cmd("Page.navigate", {"url": BASE + "/"})
            await asyncio.sleep(3.5)

            # 打开数据仓库
            say("\n【1】打开「🗄 数据仓库」")
            await js("document.getElementById('dw-toggle').click()")
            await asyncio.sleep(2.5)
            opened = await js("document.getElementById('dw-overlay').classList.contains('open')")
            check("数据仓库面板已打开", bool(opened))

            say("\n【2】文档库每行带「已入库 N」徽章")
            n_badge = await js(
                "document.querySelectorAll('#dw-docs-list .dw-has').length")
            badges = await js(
                "Array.from(document.querySelectorAll('#dw-docs-list .dw-has'))"
                ".slice(0,4).map(e=>e.closest('.dw-doc-row').querySelector('.nm').dataset.open+'='+e.innerText)")
            check("有已入库徽章", (n_badge or 0) >= 5, "共 %s 个 ｜ %s" % (n_badge, badges))
            await shot(cmd, "按文档看知识-1-文档库.png")

            say("\n【3】点「中国COVID-19新冠疫情情况.docx」进详情")
            clicked = await js("""(() => {
                const el = [...document.querySelectorAll('#dw-docs-list .nm[data-open]')]
                    .find(e => e.dataset.open.includes('COVID'));
                if (!el) return 'not-found';
                el.click();
                return el.dataset.open;
            })()""")
            check("找到并点击了新冠文档", clicked not in (None, "not-found", False), str(clicked))
            await asyncio.sleep(3.0)

            in_detail = await js("!!document.querySelector('.kdoc')")
            title = await js("(document.querySelector('.kdoc-title')||{}).innerText || ''")
            meta = await js("(document.querySelector('.kdoc-meta')||{}).innerText || ''")
            check("已进入该文档的知识视图", bool(in_detail), title)
            check("头部显示领域和条目数", ("公共卫生" in (meta or "")) or ("条" in (meta or "")),
                  (meta or "").replace("\n", " ")[:120])

            say("\n【4】按属性列出全部值（城市/省份清单）")
            attrs = await js("""Array.from(document.querySelectorAll('.kdoc-attr-hd b')).map(e=>e.innerText)""")
            check("列出的属性里有「省份名称」", "省份名称" in (attrs or []),
                  "属性：" + "、".join((attrs or [])[:8]))
            vals = await js("""(() => {
                const blk = [...document.querySelectorAll('.kdoc-attr')].find(
                    b => (b.querySelector('.kdoc-attr-hd b')||{}).innerText === '省份名称');
                if (!blk) return [];
                return [...blk.querySelectorAll('.kdoc-v')].map(e=>e.innerText);
            })()""")
            vals = vals or []
            hit = [v for v in vals if any(x in v for x in ("北京市", "上海市", "天津市", "重庆市"))]
            check("省份名称下铺开了所有值", len(vals) >= 20, "共 %d 个：%s" % (len(vals), "、".join(vals[:8])))
            check("里面有直辖市（北京/上海/天津/重庆）", len(hit) >= 2, "命中：" + "、".join(hit))
            await shot(cmd, "按文档看知识-2-城市清单.png")

            say("\n【5】全部明细可展开")
            opened_det = await js("""(() => {
                const d = document.querySelector('.kdoc-all');
                if (!d) return 'no-details';
                d.open = true;
                return d.querySelectorAll('tbody tr, tr').length;
            })()""")
            check("明细表可展开且有行", isinstance(opened_det, int) and opened_det > 100,
                  "行数 %s" % opened_det)
            await asyncio.sleep(0.6)
            await shot(cmd, "按文档看知识-3-全部明细.png")

            say("\n【6】返回全部知识")
            back = await js("""(() => {
                const b = document.getElementById('kdoc-back');
                if (!b) return 'no-btn';
                b.click();
                return 'ok';
            })()""")
            await asyncio.sleep(2.2)
            back_ok = await js("!document.querySelector('.kdoc')")
            check("「← 全部知识」能退回全局表", bool(back_ok), str(back))

    finally:
        try:
            proc.terminate()
        except Exception:
            pass
        try:
            shutil.rmtree(PROFILE, ignore_errors=True)
        except Exception:
            pass

    say("\n" + "=" * 66)
    say("结果：%d 通过 / %d 失败 ｜ 共 %d 项" % (PASS, FAIL, PASS + FAIL))
    say("=" * 66)
    with open(os.path.join(r"C:\Users\覃敬涛\Desktop\研一项目\收尾对照",
                           "按文档看知识-真机-1105.txt"), "w", encoding="utf-8") as f:
        f.write("\n".join(log))


if __name__ == "__main__":
    asyncio.run(main())
