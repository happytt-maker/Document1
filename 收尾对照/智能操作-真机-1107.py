# -*- coding: utf-8 -*-
"""智能操作模块 · 真机验收（Edge headless + CDP）

验的是"人能不能用"：
  ① 左边第一张卡是「智能操作」，点开弹窗正常加载文档清单
  ② 点一份文件 → 快捷按钮「加粗」点一下就生效（结果区出现产物 + 两个打开按钮）
  ③ 同一项再点一次 = 取消（报告里出现"取消加粗"）
  ④ 清除选择 → 直接说「设备采购合同的标题加粗」也能改对
  ⑤ 首页对话框照旧能下指令（没被这次改动破坏）
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
OUT = os.path.join(ROOT, "收尾对照")
SHOTS = os.path.join(ROOT, "agent", "输出")
BASE = "http://127.0.0.1:8000"
EDGE = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
PROFILE = os.path.join(os.environ.get("TEMP", "."), "_cdp_op_prof")
OPENER = urllib.request.build_opener(urllib.request.ProxyHandler({}))

PASS = []
CONSOLE = []


def check(name, cond, extra=""):
    PASS.append((name, bool(cond)))
    tail = ("  " + str(extra)) if extra != "" else ""
    print(("  ✓ " if cond else "  ✗ ") + name + tail)


def api(path, data=None):
    body = json.dumps(data).encode() if data is not None else None
    r = urllib.request.Request(BASE + path, data=body,
                               headers={"Content-Type": "application/json"})
    with OPENER.open(r, timeout=30) as f:
        return json.loads(f.read().decode())


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
        "--remote-debugging-port=9333", "--user-data-dir=" + PROFILE,
        "--window-size=1440,960", "about:blank",
    ], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    for _ in range(42):
        try:
            with OPENER.open("http://127.0.0.1:9333/json/version", timeout=2) as f:
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
        req = urllib.request.Request("http://127.0.0.1:9333/json/new?about:blank", method="PUT")
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

                await cmd("Page.enable")
                await cmd("Runtime.enable")
                await cmd("Network.enable")
                if token:
                    await cmd("Network.setCookie", {"name": "a23_token", "value": token,
                                                    "url": BASE, "path": "/", "httpOnly": True})
                await cmd("Page.navigate", {"url": BASE + "/"})
                time.sleep(3.5)

                print("\n① 左侧卡片顺序 / 弹窗打开")
                order = await js("Array.from(document.querySelectorAll('.feat-rail .fc-tx b'))"
                                 ".slice(0,3).map(e=>e.textContent.trim())")
                check("「智能操作」排第一张", isinstance(order, list) and order[0] == "智能操作", order)
                check("「识别入库」紧接其后", isinstance(order, list) and len(order) > 1 and order[1] == "识别入库", order)
                await js("document.getElementById('op-toggle').click()")
                time.sleep(1.6)
                opened = await js("document.getElementById('op-overlay').classList.contains('open')")
                check("弹窗已打开", opened)
                n_docs = await js("document.querySelectorAll('#op-docs .rk-line').length")
                check("文档清单已加载（只列可编辑的）", isinstance(n_docs, int) and n_docs >= 4, "%s 份" % n_docs)
                n_chips = await js("document.querySelectorAll('#op-tools .op-chip[data-op]').length")
                check("快捷按钮数量（6 文字 + 4 对齐 + 3 段落）", n_chips == 13, n_chips)
                names = await js("Array.from(document.querySelectorAll('#op-tools .op-chip[data-op]'))"
                                 ".map(e=>e.textContent.trim())")
                print("    按钮：", names)
                await shot("智能操作-1-面板.png")

                print("\n② 选中文件 → 点「加粗」（快捷按钮，不走模型）")
                pick = await js("(function(){const b=Array.from(document.querySelectorAll('#op-docs .rk-line'))"
                                ".find(x=>x.dataset.doc==='设备采购合同.docx');"
                                "if(!b)return 'NO_ROW';b.click();return document.getElementById('op-cur').innerText;})()")
                check("能选中《设备采购合同.docx》", pick != "NO_ROW" and "设备采购合同" in str(pick), pick)
                await js("(function(){const b=Array.from(document.querySelectorAll('#op-tools .op-chip'))"
                         ".find(x=>x.textContent.trim()==='加粗');b.click();})()")
                time.sleep(0.5)
                await shot("智能操作-2-执行中.png")
                for _ in range(24):
                    time.sleep(0.6)
                    st = await js("document.getElementById('op-status').innerText")
                    if "✓" in str(st) or "✗" in str(st) or "没做成" in str(st):
                        break
                st = await js("document.getElementById('op-status').innerText")
                check("快捷加粗成功", "✓" in str(st), st)
                has_open = await js("!!document.getElementById('op-open-file') && "
                                    "!!document.getElementById('op-open-dir')")
                check("结果区有「打开文件 / 打开所在目录」", has_open)
                rep = await js("(document.querySelector('#op-result .report')||{}).textContent||''")
                check("报告显示走了快捷通道", "快捷操作" in str(rep))
                check("报告里是设备采购合同", "设备采购合同" in str(rep))
                await shot("智能操作-3-加粗完成.png")

                print("\n③ 同一项再点一次 = 取消")
                await js("(function(){const b=Array.from(document.querySelectorAll('#op-tools .op-chip'))"
                         ".find(x=>x.textContent.trim()==='加粗');b.click();})()")
                for _ in range(24):
                    time.sleep(0.6)
                    st = await js("document.getElementById('op-status').innerText")
                    if "✓" in str(st) or "✗" in str(st):
                        break
                rep = await js("(document.querySelector('#op-result .report')||{}).textContent||''")
                check("第二次点是「取消加粗」", "取消加粗" in str(rep))
                await shot("智能操作-4-取消加粗.png")

                print("\n③b 换一份 xlsx：不支持的行距 / 缩进自动灰掉")
                pick2 = await js("(function(){const b=Array.from(document.querySelectorAll('#op-docs .rk-line'))"
                                 ".find(x=>x.dataset.doc==='员工信息表.xlsx');"
                                 "if(!b)return 'NO_ROW';b.click();return 'ok';})()")
                if pick2 == "NO_ROW":
                    check("能找到 xlsx 文档", False)
                else:
                    time.sleep(0.8)
                    dis = await js("Array.from(document.querySelectorAll('#op-tools .op-chip[data-op]'))"
                                   ".filter(b=>b.disabled).map(b=>b.textContent.trim())")
                    check("xlsx 下行距 / 缩进灰掉", "1.5 倍行距" in dis and "首行缩进 2 字" in dis, dis)
                    check("xlsx 下加粗仍可用", "加粗" not in dis)
                    check("xlsx 下字号仍可用",
                          await js("!document.getElementById('op-size').disabled"))
                    await shot("智能操作-4b-xlsx能力灰掉.png")

                print("\n④ 清除选择 → 话里点名")
                await js("document.getElementById('op-clear').click()")
                cur = await js("document.getElementById('op-cur').innerText")
                check("已清除选择", "未选择" in str(cur), cur)
                await js("document.getElementById('op-text').value='设备采购合同的标题居中'")
                await js("document.getElementById('op-go').click()")
                for _ in range(40):
                    time.sleep(0.8)
                    st = await js("document.getElementById('op-status').innerText")
                    if "✓" in str(st) or "✗" in str(st) or "没做成" in str(st):
                        break
                st = await js("document.getElementById('op-status').innerText")
                check("不选文件也能改对", "✓" in str(st), st)
                rep = await js("(document.querySelector('#op-result .report')||{}).textContent||''")
                check("命中的是设备采购合同", "设备采购合同" in str(rep))
                await shot("智能操作-5-话里点名.png")

                print("\n⑤ 首页对话框照旧")
                await js("document.getElementById('op-x').click()")
                time.sleep(0.4)
                await js("document.getElementById('input').value='把设备采购合同的标题右对齐'")
                await js("send()")
                card = ""
                for _ in range(45):
                    time.sleep(0.8)
                    card = await js("(function(){const cs=document.querySelectorAll('#msgs-inner .msg-sys');"
                                    "if(!cs.length)return '';return cs[cs.length-1].innerText;})()")
                    if "已完成" in str(card) or "没做成" in str(card) or "失败" in str(card):
                        break
                check("首页输入框仍能下指令（卡片显示已完成）", "已完成" in str(card), str(card)[:60].replace("\n", " | "))
                check("首页卡片命中的是设备采购合同", "设备采购合同" in str(card))
                await shot("智能操作-6-首页照旧.png")

        asyncio.run(run())
    finally:
        try:
            proc.terminate()
        except Exception:
            pass
        time.sleep(1)
        shutil.rmtree(PROFILE, ignore_errors=True)

    print("\n控制台报错：", CONSOLE if CONSOLE else "无")
    ok_n = sum(1 for _n, o in PASS if o)
    print("\n通过 %d/%d" % (ok_n, len(PASS)))
    for n, o in PASS:
        if not o:
            print("  未通过：", n)
    sys.exit(0 if ok_n == len(PASS) else 1)


if __name__ == "__main__":
    main()
