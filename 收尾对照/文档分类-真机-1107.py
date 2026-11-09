# -*- coding: utf-8 -*-
"""四个选文档的地方 · 按类型分类 · 真机验收（Edge headless + CDP）

用户要求：「智能操作、识别入库、数据仓库、表格填写这些里面选文档的功能要分个类，
.docx 为一类，.xlsx 为一类这样子」。

验的是"人能不能用"：
  ① 智能操作：顶上一排分类 chip（全部/Word/Excel/文本…）+ 组标题；点 Excel 只剩 xlsx
  ② 识别入库：切分类再切回来，勾选不丢；全选只作用于当前看得见的那一类
  ③ 数据仓库：同一套 chip + 分组
  ④ 表格填写：源文档分组 + 模板下拉框按 Word/Excel 分成两段（optgroup）
全程截图 + 收控制台报错。
"""
import base64
import json
import os
import shutil
import subprocess
import sys
import time
import urllib.request

sys.stdout.reconfigure(encoding="utf-8")

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SHOTS = os.path.join(ROOT, "agent", "输出")
BASE = "http://127.0.0.1:8000"
EDGE = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
PROFILE = os.path.join(os.environ.get("TEMP", "."), "_cdp_kind_prof")
OPENER = urllib.request.build_opener(urllib.request.ProxyHandler({}))

PASS = []
CONSOLE = []


def check(name, cond, extra=""):
    PASS.append((name, bool(cond)))
    tail = ("  " + str(extra)) if extra != "" else ""
    print(("  ✓ " if cond else "  ✗ ") + name + tail)


def login():
    body = json.dumps({"username": "2482991994", "password": "123456"}).encode()
    req = urllib.request.Request(BASE + "/api/auth/login", data=body,
                                 headers={"Content-Type": "application/json"})
    try:
        with OPENER.open(req, timeout=15) as f:
            f.read()
            for k, v in f.headers.items():
                if k.lower() == "set-cookie" and v.startswith("a23_token="):
                    return v.split(";")[0].split("=", 1)[1]
    except Exception as e:
        print("登录失败，可能不需要登录：", e)
    return None


def start_edge():
    if os.path.isdir(PROFILE):
        shutil.rmtree(PROFILE, ignore_errors=True)
    proc = subprocess.Popen([
        EDGE, "--headless=new", "--disable-gpu", "--no-first-run", "--no-default-browser-check",
        "--disable-extensions", "--disable-sync", "--disable-background-networking",
        "--disable-component-update", "--no-proxy-server", "--remote-allow-origins=*",
        "--remote-debugging-port=9334", "--user-data-dir=" + PROFILE,
        "--window-size=1440,960", "about:blank",
    ], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    for _ in range(42):
        try:
            with OPENER.open("http://127.0.0.1:9334/json/version", timeout=2) as f:
                if b"Browser" in f.read():
                    return proc
        except Exception:
            pass
        time.sleep(0.5)
    raise RuntimeError("Edge CDP 起不来")


def main():
    import asyncio
    import websockets

    token = login()
    proc = start_edge()
    try:
        req = urllib.request.Request("http://127.0.0.1:9334/json/new?about:blank", method="PUT")
        with OPENER.open(req, timeout=10) as f:
            target = json.loads(f.read().decode())
        ws_url = target["webSocketDebuggerUrl"]

        async def run():
            async with websockets.connect(ws_url, max_size=None) as ws:
                seq = {"n": 0}

                async def cmd(method, params=None):
                    seq["n"] += 1
                    mid = seq["n"]
                    await ws.send(json.dumps({"id": mid, "method": method, "params": params or {}}))
                    while True:
                        msg = json.loads(await ws.recv())
                        if msg.get("method") == "Page.javascriptDialogOpening":
                            await ws.send(json.dumps({"id": seq["n"] + 900,
                                                      "method": "Page.handleJavaScriptDialog",
                                                      "params": {"accept": True}}))
                            continue
                        if msg.get("method") == "Runtime.consoleAPICalled":
                            args = " ".join(str(a.get("value", a.get("description", "")))
                                            for a in msg["params"].get("args", []))
                            if msg["params"].get("type") in ("error", "warning"):
                                CONSOLE.append(args[:200])
                        if msg.get("method") == "Runtime.exceptionThrown":
                            d = msg["params"]["exceptionDetails"]
                            CONSOLE.append("EXC " + str(d.get("text"))[:200])
                        if msg.get("id") == mid:
                            return msg.get("result", {})

                async def js(expr):
                    r = await cmd("Runtime.evaluate", {
                        "expression": expr, "returnByValue": True, "awaitPromise": True})
                    res = r.get("result", {})
                    if r.get("exceptionDetails"):
                        return "__ERR__" + str(r["exceptionDetails"].get("text"))
                    return res.get("value")

                async def shot(name):
                    r = await cmd("Page.captureScreenshot", {"format": "png"})
                    p = os.path.join(SHOTS, name)
                    with open(p, "wb") as f:
                        f.write(base64.b64decode(r["data"]))
                    print("    截图 →", os.path.basename(p))

                async def chips(sel):
                    return await js(
                        "Array.from(document.querySelectorAll('%s .kd')).map("
                        "e=>({k:e.dataset.k,t:e.textContent.trim().replace(/\\s+/g,'')}))" % sel)

                async def groups(sel):
                    return await js(
                        "Array.from(document.querySelectorAll('%s .rk-grp-hd')).map("
                        "e=>e.textContent.trim().replace(/\\s+/g,''))" % sel)

                await cmd("Page.enable")
                await cmd("Runtime.enable")
                if token:
                    await cmd("Network.setCookie", {"name": "a23_token", "value": token,
                                                    "url": BASE, "path": "/", "httpOnly": True})
                await cmd("Page.navigate", {"url": BASE + "/"})
                time.sleep(3.5)

                # ---------------- ① 智能操作 ----------------
                print("\n① 智能操作：分类 chip + 分组")
                await js("document.getElementById('op-toggle').click()")
                time.sleep(1.8)
                cs = await chips("#op-kinds")
                ks = [c["k"] for c in cs] if isinstance(cs, list) else []
                check("有「全部」+ 真实存在的类别 chip",
                      ks[:1] == ["all"] and len(ks) >= 4, cs)
                check("分类里有 word / excel / text", {"word", "excel", "text"} <= set(ks), ks)
                gs = await groups("#op-docs")
                check("清单按类型分了组标题", isinstance(gs, list) and len(gs) >= 3, gs)
                n_all = await js("document.querySelectorAll('#op-docs .rk-line').length")
                # 点 Excel
                await js("(function(){const b=Array.from(document.querySelectorAll('#op-kinds .kd'))"
                         ".find(x=>x.dataset.k==='excel');b.click();})()")
                time.sleep(0.5)
                n_xl = await js("document.querySelectorAll('#op-docs .rk-line').length")
                exts = await js("Array.from(document.querySelectorAll('#op-docs .rk-line'))"
                                ".map(e=>(e.querySelector('.rk-src')||{}).textContent||'')")
                check("点「Excel」只剩 xlsx", n_xl > 0 and n_xl < n_all,
                      "%d → %d 份" % (n_all, n_xl))
                check("剩下的确实都是 xlsx",
                      isinstance(exts, list) and all("xlsx" in str(x) for x in exts), exts)
                await shot("文档分类-1-智能操作-Excel.png")
                await js("(function(){const b=Array.from(document.querySelectorAll('#op-kinds .kd'))"
                         ".find(x=>x.dataset.k==='all');b.click();})()")
                time.sleep(0.4)
                n_back = await js("document.querySelectorAll('#op-docs .rk-line').length")
                check("点回「全部」恢复完整清单", n_back == n_all, "%s 份" % n_back)
                await shot("文档分类-2-智能操作-全部.png")
                await js("document.getElementById('op-x').click()")
                time.sleep(0.4)

                # ---------------- ② 识别入库 ----------------
                print("\n② 识别入库：切分类，勾选不丢")
                await js("document.getElementById('rec-toggle').click()")
                time.sleep(1.8)
                cs = await chips("#rec-kinds")
                check("识别入库也有分类 chip", isinstance(cs, list) and len(cs) >= 4,
                      [c["k"] for c in cs] if isinstance(cs, list) else cs)
                gs = await groups("#rec-docs")
                check("识别入库清单已分组", isinstance(gs, list) and len(gs) >= 3, gs)
                # 在「全部」下勾两份 xlsx
                ticks = await js(
                    "(function(){const ls=Array.from(document.querySelectorAll('#rec-docs .rk-line'))"
                    ".filter(l=>{const c=l.querySelector('input');return c&&/\\.xlsx$/i.test(c.value);})"
                    ".slice(0,2);ls.forEach(l=>{const c=l.querySelector('input');c.checked=true;"
                    "c.dispatchEvent(new Event('change',{bubbles:true}));});"
                    "return ls.map(l=>l.querySelector('input').value);})()")
                n_recall = await js("document.querySelectorAll('#rec-docs input').length")
                check("先勾 2 份 xlsx", isinstance(ticks, list) and len(ticks) == 2,
                      "%s（全库 %s 份）" % (ticks, n_recall))
                # 切到 Word 再切回全部
                await js("(function(){const b=Array.from(document.querySelectorAll('#rec-kinds .kd'))"
                         ".find(x=>x.dataset.k==='word');b.click();})()")
                time.sleep(0.4)
                n_word = await js("document.querySelectorAll('#rec-docs .rk-line').length")
                check("切到 Word 只剩 docx", 0 < n_word < n_recall,
                      "%d 份（全库 %d 份）" % (n_word, n_recall))
                await shot("文档分类-3-识别入库-Word.png")
                await js("(function(){const b=Array.from(document.querySelectorAll('#rec-kinds .kd'))"
                         ".find(x=>x.dataset.k==='all');b.click();})()")
                time.sleep(0.4)
                still = await js("Array.from(document.querySelectorAll('#rec-docs input:checked'))"
                                 ".map(e=>e.value)")
                check("切回全部后勾选还在", isinstance(still, list) and len(still) == 2, still)
                # 全选只作用于当前类
                await js("(function(){const b=Array.from(document.querySelectorAll('#rec-kinds .kd'))"
                         ".find(x=>x.dataset.k==='word');b.click();})()")
                time.sleep(0.3)
                await js("document.getElementById('rec-all').click()")
                time.sleep(0.3)
                after_all = await js("Array.from(recPickSet)")
                check("Word 类里全选 = 勾上全部 docx（不碰 Excel 那 2 份）",
                      isinstance(after_all, list) and len(after_all) == n_word + 2,
                      "%d 个勾选（Word %d 份 + 原 2 份 Excel）" % (len(after_all or []), n_word))
                await shot("文档分类-4-识别入库-全选.png")
                await js("(function(){const b=Array.from(document.querySelectorAll('#rec-kinds .kd'))"
                         ".find(x=>x.dataset.k==='all');b.click();})()")
                time.sleep(0.3)
                n_total = await js("document.querySelectorAll('#rec-docs .rk-line').length")
                await js("document.getElementById('rec-all').click()")   # 全勾
                time.sleep(0.3)
                n_on1 = await js("document.querySelectorAll('#rec-docs input:checked').length")
                n_on1_set = await js("recPickSet.size")
                await js("document.getElementById('rec-all').click()")   # 再点 = 全不勾
                time.sleep(0.3)
                left = await js("recPickSet.size")
                check("全选 / 再点取消，一共管全库",
                      n_on1_set == n_total and n_on1 == n_total and left == 0,
                      "全勾 %s/%s（页面上 %s 个），再点剩 %s" % (n_on1_set, n_total, n_on1, left))
                await js("document.getElementById('rec-x').click()")
                time.sleep(0.4)

                # ---------------- ③ 数据仓库 ----------------
                print("\n③ 数据仓库：文档库分类")
                await js("document.getElementById('dw-toggle').click()")
                time.sleep(2.0)
                cs = await chips("#dw-kinds")
                check("数据仓库也有分类 chip", isinstance(cs, list) and len(cs) >= 4,
                      [c["k"] for c in cs] if isinstance(cs, list) else cs)
                gs = await groups("#dw-docs-list")
                check("左栏文档按类型分组", isinstance(gs, list) and len(gs) >= 3, gs)
                n_all = await js("document.querySelectorAll('#dw-docs-list .dw-doc-row').length")
                await js("(function(){const b=Array.from(document.querySelectorAll('#dw-kinds .kd'))"
                         ".find(x=>x.dataset.k==='excel');b.click();})()")
                time.sleep(0.5)
                n_xl = await js("document.querySelectorAll('#dw-docs-list .dw-doc-row').length")
                check("点 Excel 只剩 xlsx 行", 0 < n_xl < n_all, "%d → %d 行" % (n_all, n_xl))
                await shot("文档分类-5-数据仓库-Excel.png")
                await js("document.getElementById('dw-x').click()")
                time.sleep(0.4)

                # ---------------- ④ 表格填写 ----------------
                print("\n④ 表格填写：源文档分组 + 模板下拉分组")
                await js("document.getElementById('fill-toggle').click()")
                time.sleep(2.2)
                cs = await chips("#fill-kinds")
                check("源文档也有分类 chip", isinstance(cs, list) and len(cs) >= 4,
                      [c["k"] for c in cs] if isinstance(cs, list) else cs)
                gs = await groups("#fill-docs")
                check("源文档已分组", isinstance(gs, list) and len(gs) >= 3, gs)
                ogs = await js("Array.from(document.querySelectorAll('#fill-tpl optgroup'))"
                               ".map(e=>e.label)")
                check("模板下拉按 Word / Excel 分两段",
                      isinstance(ogs, list) and len(ogs) == 2, ogs)
                await js("(function(){const b=Array.from(document.querySelectorAll('#fill-kinds .kd'))"
                         ".find(x=>x.dataset.k==='word');b.click();})()")
                time.sleep(0.5)
                await shot("文档分类-6-表格填写-Word.png")
                # 勾一份 Word → 切到 Excel → 勾一份 xlsx → 看提示是 2 份
                await js("(function(){const l=document.querySelector('#fill-docs .rk-line');"
                         "const c=l.querySelector('input');c.checked=true;"
                         "c.dispatchEvent(new Event('change',{bubbles:true}));})()")
                await js("(function(){const b=Array.from(document.querySelectorAll('#fill-kinds .kd'))"
                         ".find(x=>x.dataset.k==='excel');b.click();})()")
                time.sleep(0.4)
                await js("(function(){const l=document.querySelector('#fill-docs .rk-line');"
                         "const c=l.querySelector('input');c.checked=true;"
                         "c.dispatchEvent(new Event('change',{bubbles:true}));})()")
                time.sleep(0.3)
                hint = await js("document.getElementById('fill-doc-hint').innerText")
                check("跨分类勾选都算数（提示 2 份）", "2" in str(hint), hint)
                await shot("文档分类-7-表格填写-跨类勾选.png")

                print("\n⑤ 控制台")
                bad = [c for c in CONSOLE if "favicon" not in c and "404" not in c]
                check("无 JS 报错", not bad, bad[:3])

        asyncio.run(run())
    finally:
        try:
            proc.terminate()
        except Exception:
            pass
        shutil.rmtree(PROFILE, ignore_errors=True)

    ok = sum(1 for _n, v in PASS if v)
    print("\n通过 %d/%d" % (ok, len(PASS)))
    for n, v in PASS:
        if not v:
            print("  未过：", n)
    sys.exit(0 if ok == len(PASS) else 1)


if __name__ == "__main__":
    main()
