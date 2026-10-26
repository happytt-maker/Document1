# -*- coding: utf-8 -*-
"""评价功能：浏览器真机验收（用户端入口 / 客服端弹窗 / 管理员端页面）。"""
import base64, json, os, shutil, subprocess, sys, time, urllib.request
from websockets.sync.client import connect

BASE = "http://127.0.0.1:8000"
EDGE = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
PROF = r"C:\Users\覃敬涛\WorkBuddy\2026-09-24-10-48-03\_cdp_prof"
OUT = r"C:\Users\覃敬涛\Desktop\研一项目\agent\输出"
PORT = 9336
op = urllib.request.build_opener(urllib.request.ProxyHandler({}))
ok, bad = [], []


def check(name, cond, detail=""):
    (ok if cond else bad).append(name)
    print(("  PASS " if cond else "  FAIL ") + name + ("" if cond else "  <<< " + str(detail)[:220]), flush=True)


def login(u, p, portal):
    req = urllib.request.Request(
        BASE + "/api/auth/login",
        data=json.dumps({"username": u, "password": p, "portal": portal}).encode(),
        headers={"Content-Type": "application/json"}, method="POST")
    ck = op.open(req, timeout=10).headers.get("Set-Cookie") or ""
    for part in ck.split(";"):
        if part.strip().startswith("a23_token="):
            return part.strip().split("=", 1)[1]
    raise RuntimeError("没拿到 cookie")


def http(path):
    return json.load(op.open("http://127.0.0.1:%d%s" % (PORT, path), timeout=5))


class Page:
    def __init__(self):
        tgt = json.load(op.open(urllib.request.Request(
            "http://127.0.0.1:%d/json/new?about:blank" % PORT, method="PUT"), timeout=5))
        self.ws = connect(tgt["webSocketDebuggerUrl"], max_size=None, open_timeout=10)
        self.i = 0
        self.cmd("Page.enable"); self.cmd("Runtime.enable"); self.cmd("Network.enable")

    def cmd(self, method, params=None):
        self.i += 1
        mid = self.i
        self.ws.send(json.dumps({"id": mid, "method": method, "params": params or {}}))
        while True:
            m = json.loads(self.ws.recv())
            if m.get("method") == "Page.javascriptDialogOpening":
                self.i += 1
                self.ws.send(json.dumps({"id": self.i, "method": "Page.handleJavaScriptDialog",
                                         "params": {"accept": True}}))
                continue
            if m.get("id") == mid:
                if "error" in m:
                    raise RuntimeError(method + " -> " + json.dumps(m["error"]))
                return m.get("result", {})

    def cookie(self, v):
        self.cmd("Network.setCookie", {"name": "a23_token", "value": v, "url": BASE,
                                       "path": "/", "httpOnly": True})

    def go(self, url, wait=2.5):
        self.cmd("Page.navigate", {"url": url})
        time.sleep(wait)

    def js(self, expr, wait=0.0):
        r = self.cmd("Runtime.evaluate", {"expression": expr, "returnByValue": True, "awaitPromise": True})
        if wait:
            time.sleep(wait)
        if r.get("exceptionDetails"):
            raise RuntimeError(json.dumps(r["exceptionDetails"])[:300])
        return (r.get("result") or {}).get("value")

    def shot(self, name):
        d = self.cmd("Page.captureScreenshot", {"format": "png"})
        p = os.path.join(OUT, name)
        open(p, "wb").write(base64.b64decode(d["data"]))
        print("  截图 -> " + p, flush=True)


if os.path.isdir(PROF):
    shutil.rmtree(PROF, ignore_errors=True)
os.makedirs(OUT, exist_ok=True)

proc = subprocess.Popen([
    EDGE, "--headless=new", "--disable-gpu", "--no-first-run", "--no-default-browser-check",
    "--disable-extensions", "--disable-sync", "--disable-background-networking",
    "--disable-component-update", "--no-proxy-server", "--remote-allow-origins=*",
    "--remote-debugging-port=%d" % PORT, "--user-data-dir=" + PROF,
    "--window-size=1440,960", "about:blank",
], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

try:
    for _ in range(40):
        try:
            http("/json/version"); break
        except Exception:
            time.sleep(0.5)
    else:
        raise RuntimeError("Edge 调试端口没起来")

    # ---------------- 用户端
    p = Page()
    p.cookie(login("demo_ai", "123456", "user"))
    p.go(BASE + "/", 3.2)

    check("顶栏有意见反馈入口", p.js("!!document.getElementById('fb-open')"))
    p.js("document.getElementById('fb-open').click()", 0.6)
    check("反馈弹窗打开", p.js("document.getElementById('fb-mask').classList.contains('on')"))
    check("功能下拉默认整体/其他", p.js("document.getElementById('fb-feature').value") == "other")
    tags = p.js("[...document.querySelectorAll('#fb-tags .fb-tag')].map(b=>b.textContent)")
    print("  标签：" + " / ".join(tags or []), flush=True)
    check("标签按功能变化（整体建议）", bool(tags) and len(tags) >= 3, tags)

    p.js("document.querySelector('#fb-stars .fb-star[data-v=\"4\"]').click()", 0.3)
    check("点 4 星有反馈", "满意" in (p.js("document.getElementById('fb-starhint').textContent") or ""),
          p.js("document.getElementById('fb-starhint').textContent"))
    p.js("document.querySelector('#fb-tags .fb-tag').click()", 0.2)
    p.js("document.getElementById('fb-ta').value='表格填写那次很省事，就是希望多一个「只填空白格」的选项';"
         "document.getElementById('fb-ta').dispatchEvent(new Event('input'))", 0.2)
    check("字数统计跟着走", "/1000" in (p.js("document.getElementById('fb-cnt').textContent") or ""),
          p.js("document.getElementById('fb-cnt').textContent"))
    p.shot("评价-1-用户提交弹窗.png")

    p.js("document.getElementById('fb-send').click()", 2.2)
    check("提交后弹窗关闭", not p.js("document.getElementById('fb-mask').classList.contains('on')"))
    check("提交后有提示", "收到" in (p.js("document.getElementById('toast').textContent") or ""),
          p.js("document.getElementById('toast').textContent"))

    # 操作历史每条都能评价
    p.js("loadHistory()", 1.2)
    check("操作历史每条带评价按钮", p.js("document.querySelectorAll('#his .his-fb').length") > 0,
          p.js("document.querySelectorAll('#his .his-fb').length"))
    p.js("document.querySelector('#his .his-fb').click()", 0.6)
    check("历史里点评价打开弹窗", p.js("document.getElementById('fb-mask').classList.contains('on')"))
    check("功能自动定位到文档操作", p.js("document.getElementById('fb-feature').value") == "doc_edit",
          p.js("document.getElementById('fb-feature').value"))
    p.js("document.getElementById('fb-cancel').click()", 0.3)

    # 识别入库结果区
    p.js("renderRecResult({result:{ok:true,report:'（验收用，未真跑识别）'},elapsed:1.2})", 0.5)
    check("识别结果区有评价条", p.js("!!document.querySelector('#rec-result .fb-strip')"))
    p.js("document.querySelector('#rec-result .fb-go').click()", 0.6)
    check("点评价条功能=识别入库", p.js("document.getElementById('fb-feature').value") == "extract",
          p.js("document.getElementById('fb-feature').value"))
    p.shot("评价-2-识别结果区评价.png")
    p.js("document.getElementById('fb-cancel').click()", 0.3)

    # 表格填写结果区
    p.js("renderFillResult({result:{ok:false,report:'（验收用）'},elapsed:1.0},'某模板.xlsx')", 0.5)
    check("填写结果区有评价条", p.js("!!document.querySelector('#fill-result .fb-strip')"))

    # ---------------- 客服端
    s = Page()
    s.cookie(login("kf_demo", "123456", "staff"))
    s.go(BASE + "/staff", 3.0)
    check("客服端有评价入口", s.js("!!document.getElementById('fb-btn')"))
    s.js("document.getElementById('fb-btn').click()", 1.8)
    body = s.js("document.getElementById('fbm-body').innerText") or ""
    check("客服端弹窗有内容", ("我收到的评价" in body and "平均分" in body) or "还没有人评价" in body, body[:120])
    check("客服端显示平均分", "分" in (s.js("document.getElementById('fbm-sub').textContent") or ""),
          s.js("document.getElementById('fbm-sub').textContent"))
    s.shot("评价-3-客服端看到评价.png")

    # ---------------- 管理员端
    a = Page()
    a.cookie(login("2482991994", "123456", "admin"))
    a.go(BASE + "/admin", 3.0)
    check("管理端导航有评价反馈", a.js("[...document.querySelectorAll('#nav button')].some(b=>b.dataset.view==='feedback')"))
    a.js("document.querySelector('#nav button[data-view=\"feedback\"]').click()", 2.2)
    check("切到评价页", a.js("document.getElementById('view-feedback').classList.contains('on')"))
    avg = a.js("document.getElementById('fb-avg').textContent")
    total = a.js("document.getElementById('fb-total').textContent")
    check("平均分/总数有数", avg not in ("–", "", None) and total not in ("–", "", None), (avg, total))
    print("  平均 %s 分 · 共 %s 条" % (avg, total), flush=True)
    check("列表有评价行", a.js("document.querySelectorAll('#fb-box .fb-item').length") > 0,
          a.js("document.querySelectorAll('#fb-box .fb-item').length"))
    check("按功能分区有数据", a.js("document.querySelectorAll('#fb-by .fb-row').length") > 0,
          a.js("document.querySelectorAll('#fb-by .fb-row').length"))
    a.shot("评价-4-管理端评价页.png")

    print("\n结果：通过 %d 项，失败 %d 项" % (len(ok), len(bad)), flush=True)
    if bad:
        print("失败项：" + "；".join(bad), flush=True)
finally:
    try:
        proc.terminate()
    except Exception:
        pass
    time.sleep(1)
    shutil.rmtree(PROF, ignore_errors=True)

sys.exit(1 if bad else 0)
