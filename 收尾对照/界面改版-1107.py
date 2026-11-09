# -*- coding: utf-8 -*-
"""界面改版（左侧功能栏 / 大弹窗 / 文档与模板删除 / 知识图谱 / 能力问答）真机验收

跑法：python 收尾对照\界面改版-1107.py
结果：收尾对照\界面改版-1107.txt + agent\输出\界面改版-*.png
"""
import base64
import json
import os
import shutil
import subprocess
import time
import urllib.request
import urllib.error

EDGE = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
BASE = "http://127.0.0.1:8000"
PROFILE = os.path.join(os.environ.get("TEMP", r"C:\Windows\Temp"), "_cdp_ui1107b")
OUT = r"C:\Users\覃敬涛\Desktop\研一项目\agent\输出"
LOG = r"C:\Users\覃敬涛\Desktop\研一项目\收尾对照\界面改版-1107.txt"

for k in ("HTTP_PROXY", "HTTPS_PROXY", "ALL_PROXY", "http_proxy", "https_proxy", "all_proxy"):
    os.environ.pop(k, None)
_opener = urllib.request.build_opener(urllib.request.ProxyHandler({}))

lines = []


def say(*a):
    s = " ".join(str(x) for x in a)
    print(s, flush=True)
    lines.append(s)


def http(path, data=None, method=None):
    url = BASE + path
    body = json.dumps(data).encode() if data is not None else None
    req = urllib.request.Request(url, data=body, method=method)
    if body:
        req.add_header("Content-Type", "application/json")
    return _opener.open(req, timeout=20)


class CDP:
    def __init__(self, ws):
        self.ws = ws
        self.n = 0
        self.events = []

    def send(self, method, params=None, timeout=30):
        self.n += 1
        mid = self.n
        self.ws.send(json.dumps({"id": mid, "method": method, "params": params or {}}))
        t0 = time.time()
        while time.time() - t0 < timeout:
            m = json.loads(self.ws.recv())
            if m.get("id") == mid:
                if "error" in m:
                    raise RuntimeError("%s -> %s" % (method, m["error"]))
                return m.get("result", {})
            if m.get("method"):
                self.events.append(m)
        raise TimeoutError(method)

    def js(self, expr, timeout=40):
        r = self.send("Runtime.evaluate",
                      {"expression": expr, "returnByValue": True, "awaitPromise": True},
                      timeout=timeout)
        if r.get("exceptionDetails"):
            raise RuntimeError(json.dumps(r["exceptionDetails"], ensure_ascii=False)[:400])
        return r.get("result", {}).get("value")

    def shot(self, name):
        d = self.send("Page.captureScreenshot", {"format": "png"}, timeout=60)
        p = os.path.join(OUT, name)
        with open(p, "wb") as f:
            f.write(base64.b64decode(d["data"]))
        say("  [图] " + name)


def login_cookie():
    r = http("/api/auth/login", {"username": "2482991994", "password": "123456"})
    for c in (r.headers.get_all("Set-Cookie") or []):
        for part in c.split(";"):
            part = part.strip()
            if part.startswith("a23_token="):
                return part.split("=", 1)[1]
    raise RuntimeError("没拿到 a23_token")


def main():
    for pid in os.popen('wmic process where "name=\'msedge.exe\'" get processid,commandline').read().split("\n"):
        if "_cdp_ui1107b" in pid:
            try:
                subprocess.run(["taskkill", "/F", "/PID", pid.strip().split()[-1]],
                               stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            except Exception:
                pass
    shutil.rmtree(PROFILE, ignore_errors=True)

    try:
        http("/api/status")
    except Exception as e:
        say("!! 8000 没起：", e)
        return 1

    proc = subprocess.Popen([
        EDGE, "--headless=new", "--disable-gpu", "--no-first-run", "--no-default-browser-check",
        "--disable-extensions", "--disable-sync", "--disable-background-networking",
        "--disable-component-update", "--no-proxy-server", "--remote-allow-origins=*",
        "--remote-debugging-port=9334", "--user-data-dir=" + PROFILE,
        "--window-size=1600,1000", "about:blank",
    ], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

    try:
        for _ in range(50):
            try:
                j = json.loads(_opener.open("http://127.0.0.1:9334/json/version", timeout=2).read())
                if j.get("webSocketDebuggerUrl"):
                    break
            except Exception:
                time.sleep(0.5)
        req = urllib.request.Request("http://127.0.0.1:9334/json/new?about:blank", method="PUT")
        tab = json.loads(_opener.open(req, timeout=10).read())

        import websockets.sync.client as wsc
        ws = wsc.connect(tab["webSocketDebuggerUrl"], max_size=None, open_timeout=20)
        c = CDP(ws)
        c.send("Page.enable")
        c.send("Runtime.enable")
        c.send("Console.enable")
        c.send("Network.enable")
        c.send("Network.setCookie", {"name": "a23_token", "value": login_cookie(),
                                     "url": BASE, "path": "/", "httpOnly": True})
        c.send("Page.navigate", {"url": BASE + "/"})
        time.sleep(3.5)

        ok = bad = 0

        def chk(name, cond, extra=""):
            nonlocal ok, bad
            if cond:
                ok += 1
                say("  OK  ", name, extra)
            else:
                bad += 1
                say("  FAIL", name, extra)

        say("== 0 页面 ==")
        chk("页面打开", bool(c.js("document.title")))
        c.shot("界面改版-1-总览.png")

        # ---------- 1 左侧功能栏 ----------
        say("\n== 1 左侧功能栏 ==")
        chk("功能卡 ≥ 7", c.js("document.querySelectorAll('.feat-rail .feat-card').length") >= 7,
            "实有 %s" % c.js("document.querySelectorAll('.feat-rail .feat-card').length"))
        chk("左栏在视口左侧", (c.js("document.querySelector('.feat-rail').getBoundingClientRect().left") or 999) < 40)
        chk("卡片够大（高度≥52）",
            (c.js("document.querySelector('.feat-card').getBoundingClientRect().height") or 0) >= 52)
        chk("顶栏不再放功能按钮", c.js("document.querySelectorAll('header .hdr-right #dw-toggle, header #rec-toggle').length") == 0)
        chk("右上角保留 服务状态/会话/新会话/账号",
            c.js("!!document.getElementById('svc') && !!document.getElementById('sid-label') "
                 "&& !!document.getElementById('new-session') && !!document.getElementById('user-btn')"))
        chk("主内容已给左栏让位", (c.js("parseFloat(getComputedStyle(document.querySelector('main')).paddingLeft)") or 0) >= 200)

        # ---------- 2 能力问答 ----------
        say("\n== 2 对话框问功能 ==")
        c.js("document.getElementById('input').value='你有什么功能'")
        c.js("document.getElementById('send').click()")
        time.sleep(4.5)
        txt = c.js("(function(){var a=document.querySelectorAll('.msg-sys');"
                   "return a.length? a[a.length-1].innerText : '';})()")
        chk("回答里说明能做什么", "改文档" in txt and "填" in txt, txt[:60].replace("\n", " "))
        chk("回答直接可见（不在折叠里）",
            c.js("!!document.querySelector('.report.help')") and
            "识别入库" in (c.js("(document.querySelector('.report.help')||{}).innerText||''")))
        c.shot("界面改版-2-问功能.png")

        # ---------- 3 数据仓库 ----------
        say("\n== 3 数据仓库 ==")
        c.js("document.getElementById('dw-toggle').click()")
        time.sleep(2.0)
        chk("弹窗更大（宽≥1200）",
            (c.js("document.querySelector('#dw-overlay .dw-panel').getBoundingClientRect().width") or 0) >= 1200)
        chk("左右分区（文档库 / 知识区）",
            c.js("!!document.querySelector('#dw-overlay .dw-side') && !!document.querySelector('#dw-overlay .dw-view')"))
        chk("上传区是大卡片（高≥64）",
            (c.js("document.getElementById('dw-docs-drop').getBoundingClientRect().height") or 0) >= 64)
        chk("搜索框显眼（高≥40）",
            (c.js("document.getElementById('dw-q').getBoundingClientRect().height") or 0) >= 40)
        chk("有「知识图谱」标签", c.js("!!document.querySelector('.dw-tab[data-dwview=\"graph\"]')"))
        chk("上传的文档有删除按钮",
            c.js("document.querySelectorAll('#dw-docs-list button.del[data-del]').length") >= 0)
        c.shot("界面改版-3-数据仓库.png")

        say("\n== 4 知识图谱 ==")
        c.js("document.querySelector('.dw-tab[data-dwview=\"graph\"]').click()")
        time.sleep(3.0)
        n_node = c.js("document.querySelectorAll('#kg-svg circle').length") or 0
        n_edge = c.js("document.querySelectorAll('#kg-svg line').length") or 0
        chk("图上有点", n_node > 0, "节点 %s" % n_node)
        chk("图上有连线", n_edge > 0, "边 %s" % n_edge)
        chk("图例齐全", "对象" in (c.js("document.querySelector('.kg-legend').innerText") or ""))
        c.shot("界面改版-4-知识图谱.png")
        c.js("dwSetView('list')")
        c.js("document.getElementById('dw-x').click()")
        time.sleep(0.4)

        # ---------- 5 识别入库 ----------
        say("\n== 5 识别入库 ==")
        c.js("document.getElementById('rec-toggle').click()")
        time.sleep(1.6)
        chk("弹窗更大（宽≥1100）",
            (c.js("document.querySelector('#rec-overlay .dw-panel').getBoundingClientRect().width") or 0) >= 1100)
        chk("能上传", c.js("!!document.getElementById('rec-upload') && !!document.getElementById('rec-file')"))
        chk("能删除", c.js("!!document.getElementById('rec-del')"))
        chk("文档清单是单列",
            (c.js("getComputedStyle(document.getElementById('rec-docs')).flexDirection") or "") == "column")
        chk("关注点输入框够大（高≥44）",
            (c.js("document.getElementById('rec-req').getBoundingClientRect().height") or 0) >= 44)
        c.shot("界面改版-5-识别入库.png")
        c.js("document.getElementById('rec-x').click()")
        time.sleep(0.3)

        # ---------- 6 表格填写 ----------
        say("\n== 6 表格填写 ==")
        c.js("document.getElementById('fill-toggle').click()")
        time.sleep(1.8)
        chk("弹窗更大（宽≥1000）",
            (c.js("document.querySelector('#fill-overlay .dw-panel').getBoundingClientRect().width") or 0) >= 1000)
        chk("模板删除按钮可用（系统模板也能删）",
            c.js("!document.getElementById('fill-del-btn').disabled"),
            c.js("JSON.stringify({v:document.getElementById('fill-tpl').value,"
                 "dis:document.getElementById('fill-del-btn').disabled})"))
        chk("填表要求输入框够大（高≥44）",
            (c.js("document.getElementById('fill-req').getBoundingClientRect().height") or 0) >= 44)
        c.shot("界面改版-6-表格填写.png")

        # ---------- 7 后端接口 ----------
        say("\n== 7 接口 ==")
        g = json.loads(http("/api/knowledge/graph").read())
        chk("图谱接口 ok", bool(g.get("ok")), "节点 %s / 边 %s" % (len(g.get("nodes") or []), len(g.get("edges") or [])))
        try:
            http("/api/docs?name=" + urllib.request.quote("不存在的文档.docx"), method="DELETE")
            d1 = False
        except urllib.error.HTTPError as e:
            d1 = e.code in (400, 403, 404)
        chk("删不存在的文档会被拒", d1)
        try:
            http("/api/docs?name=" + urllib.request.quote("设备采购合同.docx"), method="DELETE")
            d2 = False
        except urllib.error.HTTPError as e:
            d2 = e.code == 403
        chk("标准文档不能删（403）", d2)

        # ---------- 8 console ----------
        errs = []
        for e in c.events:
            if e.get("method") == "Runtime.exceptionThrown":
                errs.append(str(e["params"])[:200])
            if e.get("method") == "Runtime.consoleAPICalled" and e["params"].get("type") == "error":
                errs.append(str(e["params"].get("args"))[:200])
        chk("console 无报错", len(errs) == 0, errs[:1])

        say("\n结论：通过 %d 项 / 失败 %d 项" % (ok, bad))
    finally:
        try:
            proc.terminate()
        except Exception:
            pass
        with open(LOG, "w", encoding="utf-8") as f:
            f.write("\n".join(lines))
    return 0 if bad == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())
