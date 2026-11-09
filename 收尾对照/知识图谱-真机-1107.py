# -*- coding: utf-8 -*-
"""知识图谱增强 · 真机验收（Edge headless + CDP）

用户要求：「数据仓库里面的知识图谱再优化一下，根据多种属性能够确定更多的关联关系，
并且能够支持多种关联关系的转换，用户能够更加直观的看到他们之间的关系」。

验的是"人能不能用"：
  ① 图谱能出来：节点/边数量、统计口径（多少对象、多少共享值）
  ② 关系视角可切换：全部 / 显式关系 / 共享属性 / 属性 / 同类型 / 同文档
  ③ 关系可转换：把「A-属性->值<-属性-B」折叠成「A-[属性=值]->B」直连
  ④ 直观：点节点右侧列出全部关系、点一条能顺着走到对端、悬停高亮、拖拽、缩放
  ⑤ 筛选：图例里点某个关系类型 → 只留/去掉这一类
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
PROFILE = os.path.join(os.environ.get("TEMP", "."), "_cdp_kg_prof")
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
        "--remote-debugging-port=9335", "--user-data-dir=" + PROFILE,
        "--window-size=1500,980", "about:blank",
    ], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    for _ in range(42):
        try:
            with OPENER.open("http://127.0.0.1:9335/json/version", timeout=2) as f:
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
        req = urllib.request.Request("http://127.0.0.1:9335/json/new?about:blank", method="PUT")
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
                    if r.get("exceptionDetails"):
                        return "__ERR__" + str(r["exceptionDetails"].get("text"))
                    return r.get("result", {}).get("value")

                async def shot(name):
                    r = await cmd("Page.captureScreenshot", {"format": "png"})
                    p = os.path.join(SHOTS, name)
                    with open(p, "wb") as f:
                        f.write(base64.b64decode(r["data"]))
                    print("    截图 →", os.path.basename(p))

                async def wait_kg(min_nodes=20, limit=40):
                    for _ in range(limit):
                        n = await js("document.querySelectorAll('#kg-svg .kn').length")
                        if isinstance(n, (int, float)) and n >= min_nodes:
                            return n
                        time.sleep(0.5)
                    return 0

                async def counts():
                    return await js(
                        "({n:document.querySelectorAll('#kg-svg .kn').length,"
                        " e:document.querySelectorAll('#kg-svg line').length,"
                        " r:document.querySelectorAll('#kg-svg rect').length,"
                        " t:document.querySelectorAll('.kg-lg').length})")

                await cmd("Page.enable")
                await cmd("Runtime.enable")
                if token:
                    await cmd("Network.setCookie", {"name": "a23_token", "value": token,
                                                    "url": BASE, "path": "/", "httpOnly": True})
                await cmd("Page.navigate", {"url": BASE + "/"})
                time.sleep(3.5)

                print("\n① 打开数据仓库 → 知识图谱")
                await js("document.getElementById('dw-toggle').click()")
                time.sleep(1.6)
                await js("(function(){const b=document.querySelector('[data-dwview=graph]');b.click();})()")
                n0 = await wait_kg(60)
                c0 = await counts()
                check("图谱画出来了", n0 >= 60, "%s 个节点" % n0)
                check("关系边足够多（>150 条）", c0.get("e", 0) > 150, c0)
                check("统计口径有「对象」和「共享值」",
                      "个对象" in str(await js("document.getElementById('kg-stat').textContent"))
                      and "共享值" in str(await js("document.getElementById('kg-stat').textContent")),
                      await js("document.getElementById('kg-stat').textContent"))
                await shot("知识图谱-1-全部视角.png")

                print("\n② 关系视角切换")
                vs = await js("Array.from(document.querySelectorAll('.kg-v')).map(b=>b.dataset.kgview)")
                check("6 个视角 chip", isinstance(vs, list) and len(vs) == 6, vs)

                await js("Array.from(document.querySelectorAll('.kg-v')).find(b=>b.dataset.kgview==='shared').click()")
                time.sleep(1.4)
                c_sh = await counts()
                check("切到「共享属性」重画", c_sh.get("n", 0) > 20 and c_sh.get("e", 0) > 30, c_sh)
                check("共享属性视角没有紫色类型节点", c_sh.get("r", 0) == 0, c_sh)
                await shot("知识图谱-2-共享属性.png")

                await js("Array.from(document.querySelectorAll('.kg-v')).find(b=>b.dataset.kgview==='rel').click()")
                time.sleep(1.2)
                c_rel = await counts()
                check("切到「显式关系」只剩 4 条", c_rel.get("e", 0) == 4, c_rel)
                await shot("知识图谱-3-显式关系.png")

                await js("Array.from(document.querySelectorAll('.kg-v')).find(b=>b.dataset.kgview==='type').click()")
                time.sleep(1.2)
                c_ty = await counts()
                check("「同类型」出现类型中枢（方块节点）", c_ty.get("r", 0) >= 5, c_ty)
                check("同类型边够多", c_ty.get("e", 0) > 50, c_ty)
                await shot("知识图谱-4-同类型.png")

                await js("Array.from(document.querySelectorAll('.kg-v')).find(b=>b.dataset.kgview==='doc').click()")
                time.sleep(1.2)
                c_dc = await counts()
                check("「同文档」有文档中枢", c_dc.get("r", 0) >= 5, c_dc)

                print("\n③ 关系转换：值节点折叠成对象直连")
                await js("Array.from(document.querySelectorAll('.kg-v')).find(b=>b.dataset.kgview==='shared').click()")
                time.sleep(1.2)
                before = await counts()
                await js("document.getElementById('kg-collapse').click()")
                time.sleep(1.6)
                after = await counts()
                labels = await js(
                    "Array.from(document.querySelectorAll('#kg-svg text'))"
                    ".map(t=>t.textContent).filter(s=>s.indexOf('=')>=0).slice(0,5)")
                check("折叠后节点变少（值节点被收掉）", after.get("n", 0) < before.get("n", 0),
                      "%s → %s" % (before.get("n"), after.get("n")))
                check("出现「属性=值」的直接关系标签",
                      isinstance(labels, list) and len(labels) > 0, labels)
                check("折叠后按钮亮起",
                      "on" in str(await js("document.getElementById('kg-collapse').className")))
                await shot("知识图谱-5-折叠直连.png")

                print("\n④ 点节点看关系清单（直观）")
                picked = await js(
                    "(function(){const gs=Array.from(document.querySelectorAll('#kg-svg .kn'));"
                    "const g=gs.find(x=>x.querySelector('circle'))||gs[0];"
                    "const r=g.getBoundingClientRect();"
                    "const o={bubbles:true,clientX:r.x+r.width/2,clientY:r.y+r.height/2,pointerId:1,isPrimary:true};"
                    "g.dispatchEvent(new PointerEvent('pointerdown',o));"
                    "g.dispatchEvent(new PointerEvent('pointerup',o));"
                    "return g.dataset.id;})()")
                time.sleep(0.8)
                side_on = await js("document.getElementById('kg-side').classList.contains('on')")
                rows = await js("document.querySelectorAll('#kg-side .kg-rel').length")
                title1 = await js("document.querySelector('#kg-side .kg-side-hd b').textContent")
                check("点节点弹出右侧关系清单", side_on is True, picked)
                check("清单里列了多条关系", isinstance(rows, (int, float)) and rows > 0, "%s 条" % rows)
                check("清单标题就是点的那个对象", bool(title1), title1)
                await shot("知识图谱-6-关系清单.png")

                jump = await js(
                    "(function(){const el=document.querySelector('#kg-side .kg-rel');"
                    "if(!el) return '';const t=el.querySelector('b').textContent;el.click();return t;})()")
                time.sleep(0.7)
                title2 = await js("document.querySelector('#kg-side .kg-side-hd b').textContent")
                check("点关系能顺着走到对端", bool(jump) and title2 == jump,
                      "%s → %s" % (jump, title2))

                await js("document.getElementById('kg-clear').click()")
                time.sleep(0.4)
                check("取消选中关掉清单",
                      not await js("document.getElementById('kg-side').classList.contains('on')"))

                print("\n⑤ 关系筛选 / 布局 / 拖拽 / 缩放")
                lg = await js("document.querySelectorAll('.kg-lg').length")
                check("图例列出关系类型", isinstance(lg, (int, float)) and lg >= 3, "%s 类" % lg)
                e_before = (await counts()).get("e", 0)
                for i in range(3):
                    await js("(function(){const a=document.querySelectorAll('.kg-lg[data-k]');"
                             "if(a[%d]) a[%d].click();})()" % (i, i))
                    time.sleep(0.45)
                e_after = (await counts()).get("e", 0)
                check("关掉 3 类关系后边被筛掉", e_after < e_before, "%s → %s" % (e_before, e_after))
                has_reset = await js("!!document.getElementById('kg-lg-reset')")
                check("出现「恢复全部」入口", has_reset is True)
                await js("document.getElementById('kg-lg-reset').click()")
                time.sleep(0.9)
                e_back = (await counts()).get("e", 0)
                check("一键恢复后边都回来了", e_back == e_before, "%s → %s" % (e_after, e_back))

                await js("(function(){const s=document.getElementById('kg-layout');"
                         "s.value='cluster';s.dispatchEvent(new Event('change',{bubbles:true}));})()")
                time.sleep(1.6)
                c_cl = await counts()
                check("换「按类型分簇」布局还在", c_cl.get("n", 0) > 20, c_cl)
                await shot("知识图谱-7-分簇布局.png")

                pos_before = await js(
                    "(function(){const g=document.querySelector('#kg-svg .kn');"
                    "const c=g.querySelector('circle')||g.querySelector('rect');"
                    "return [c.getAttribute('cx')||c.getAttribute('x'), c.getAttribute('cy')||c.getAttribute('y')];})()")
                await js(
                    "(function(){const g=document.querySelector('#kg-svg .kn');"
                    "const r=g.getBoundingClientRect();"
                    "const x=r.x+r.width/2,y=r.y+r.height/2;"
                    "const mk=(t,dx,dy)=>new PointerEvent(t,{bubbles:true,clientX:x+dx,clientY:y+dy,pointerId:2,isPrimary:true});"
                    "g.dispatchEvent(mk('pointerdown',0,0));"
                    "g.dispatchEvent(mk('pointermove',60,40));"
                    "g.dispatchEvent(mk('pointerup',60,40));})()")
                time.sleep(0.5)
                pos_after = await js(
                    "(function(){const g=document.querySelector('#kg-svg .kn');"
                    "const c=g.querySelector('circle')||g.querySelector('rect');"
                    "return [c.getAttribute('cx')||c.getAttribute('x'), c.getAttribute('cy')||c.getAttribute('y')];})()")
                check("节点能拖动", pos_before != pos_after, "%s → %s" % (pos_before, pos_after))

                t0 = await js("document.getElementById('kg-g').getAttribute('transform')")
                await js("document.querySelector('#kg-zoom button[data-kgz=in]').click()")
                time.sleep(0.4)
                t1 = await js("document.getElementById('kg-g').getAttribute('transform')")
                check("缩放按钮生效", t0 != t1, "%s → %s" % (t0, t1))

                print("\n⑥ 换文档看（单文档图谱）")
                await js("(function(){const b=document.querySelector('[data-dwview=list]');b.click();})()")
                time.sleep(0.6)
                await js("(function(){const el=Array.from(document.querySelectorAll('#dw-docs-list .nm'))"
                         ".find(x=>x.dataset.open&&x.dataset.open.indexOf('员工')>=0);if(el) el.click();})()")
                time.sleep(1.3)
                await js("(function(){const b=document.querySelector('[data-dwview=graph]');b.click();})()")
                time.sleep(1.6)
                scope = await js("document.getElementById('kg-scope').textContent")
                c_one = await counts()
                check("点进某份文档后图谱收窄到该文档", "员工" in str(scope), scope)
                check("单文档图谱仍有内容", c_one.get("n", 0) > 3, c_one)
                await shot("知识图谱-8-单文档.png")

                errs = [c for c in CONSOLE if "EXC" in c or "error" in c.lower()]
                check("控制台无报错", len(errs) == 0, errs[:2])

        asyncio.run(run())
    finally:
        try:
            proc.terminate()
        except Exception:
            pass
        shutil.rmtree(PROFILE, ignore_errors=True)

    ok = sum(1 for _n, o in PASS if o)
    print("\n通过 %d/%d" % (ok, len(PASS)))
    for n, o in PASS:
        if not o:
            print("  未过：", n)
    sys.exit(0 if ok == len(PASS) else 1)


if __name__ == "__main__":
    main()
