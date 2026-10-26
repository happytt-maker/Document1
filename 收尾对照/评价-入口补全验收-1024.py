# -*- coding: utf-8 -*-
"""评价入口补全：文档操作卡片 / 图片插入 / 语音输入 / 客服面板 —— 真机逐条验。"""
import base64, json, os, shutil, subprocess, time, urllib.request
from websockets.sync.client import connect

BASE = "http://127.0.0.1:8000"
EDGE = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
PROF = r"C:\Users\覃敬涛\WorkBuddy\2026-09-24-10-48-03\_cdp_prof_v"
OUT = r"C:\Users\覃敬涛\Desktop\研一项目\agent\输出"
PORT = 9338
op = urllib.request.build_opener(urllib.request.ProxyHandler({}))
ok, bad = [], []


def check(name, cond, detail=""):
    (ok if cond else bad).append(name)
    print(("  PASS " if cond else "  FAIL ") + name +
          ("" if cond else "  <<< " + str(detail)[:200]), flush=True)


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

    def init_js(self, src):
        self.cmd("Page.addScriptToEvaluateOnNewDocument", {"source": src})

    def go(self, url, wait=2.5):
        self.cmd("Page.navigate", {"url": url})
        time.sleep(wait)

    def js(self, expr, wait=0.0):
        r = self.cmd("Runtime.evaluate", {"expression": expr, "returnByValue": True,
                                          "awaitPromise": True})
        if wait:
            time.sleep(wait)
        if r.get("exceptionDetails"):
            raise RuntimeError(json.dumps(r["exceptionDetails"])[:300])
        return (r.get("result") or {}).get("value")

    def wait_for(self, expr, secs=45, step=0.5):
        t0 = time.time()
        while time.time() - t0 < secs:
            try:
                if self.js(expr):
                    return True
            except Exception:
                pass
            time.sleep(step)
        return False

    def shot(self, name):
        d = self.cmd("Page.captureScreenshot", {"format": "png"})
        p = os.path.join(OUT, name)
        open(p, "wb").write(base64.b64decode(d["data"]))
        print("  截图 -> " + p, flush=True)


# 1x1 透明 PNG，用来走真实的「拖图片进对话框」链路
PNG_B64 = ("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==")

# 假的语音识别：只替掉浏览器的识别引擎，页面里 voiceStart / onresult / onend 全走真代码
FAKE_SR = """
window.SpeechRecognition = function () {
  this.lang=''; this.continuous=false; this.interimResults=false; this.maxAlternatives=1;
  var self = this;
  this.start = function () {
    setTimeout(function(){ self.onstart && self.onstart(); }, 60);
    setTimeout(function(){
      self.onresult && self.onresult({
        resultIndex: 0,
        results: [{0: {transcript: '把设备采购合同的标题居中'}, isFinal: true, length: 1}]
      });
    }, 220);
    setTimeout(function(){ self.onend && self.onend(); }, 420);
  };
  this.stop = function () {};
};
"""

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

    p = Page()
    p.init_js(FAKE_SR)
    p.cookie(login("demo_ai", "123456", "user"))
    p.go(BASE + "/", 3.2)

    # ---------------- 1. 文档操作：跑一条真指令，看卡片下面是否跟评价条
    p.js("document.getElementById('input').value='把设备采购合同的标题居中';"
         "document.getElementById('send').click()")
    done = p.wait_for("!!document.querySelector('#msgs .card.ok')", 90)
    check("文档操作真跑通（卡片变绿）", done)
    if done:
        has = p.js("!!document.querySelector('#msgs .fb-strip [data-fb=\"doc_edit\"]')")
        check("文档操作卡片下有评价条", has)
        p.js("document.querySelector('#msgs .fb-strip [data-fb=\"doc_edit\"]').click()", 0.5)
        check("点它 → 弹窗定位文档操作",
              p.js("document.getElementById('fb-feature').value") == "doc_edit",
              p.js("document.getElementById('fb-feature').value"))
        p.shot("评价-5-文档操作卡片.png")
        p.js("document.getElementById('fb-cancel').click()", 0.3)

    # ---------------- 2. 语音输入：假引擎走真流程（说一句 → 再点一次收尾）
    p.js("document.getElementById('voice-btn').click()")
    got_text = p.wait_for("(document.getElementById('input').value||'').indexOf('设备采购合同') >= 0", 6)
    check("识别结果确实填进了输入框", got_text,
          p.js("document.getElementById('input').value"))
    p.js("document.getElementById('voice-btn').click()")     # 再点一次 = 停止录音
    voice_ok = p.wait_for("document.getElementById('fb-tip').classList.contains('on') "
                          "&& !!document.querySelector('#fb-tip [data-fb=\"voice\"]')", 8)
    check("语音说完弹出评价提示", voice_ok)
    p.shot("评价-6-语音输入评价提示.png")
    p.js("document.querySelector('#fb-tip [data-fb=\"voice\"]').click()", 0.5)
    check("语音评价弹窗功能=语音输入",
          p.js("document.getElementById('fb-feature').value") == "voice",
          p.js("document.getElementById('fb-feature').value"))
    p.js("document.getElementById('fb-cancel').click()", 0.3)
    check("从提示点进弹窗后提示自动收起",
          not p.js("document.getElementById('fb-tip').classList.contains('on')"))

    # ---------------- 3. 图片插入：真拖一张 PNG 进去 → 真插入 → 出现评价提示
    p.js("""
      (function(){
        var b = atob('%s'); var u = new Uint8Array(b.length);
        for (var i=0;i<b.length;i++) u[i]=b.charCodeAt(i);
        var f = new File([u], '验收图.png', {type:'image/png'});
        var dt = new DataTransfer(); dt.items.add(f);
        var ev = new DragEvent('drop', {bubbles:true, cancelable:true, dataTransfer:dt});
        document.querySelector('.chat').dispatchEvent(ev);
        return 1;
      })()
    """ % PNG_B64, 1.2)
    check("拖图后进入待插入状态",
          p.js("document.getElementById('img-note').classList.contains('on')"),
          p.js("document.getElementById('img-sub').textContent"))
    # 语音识别把字留在输入框了，先清空 —— 空回车 = 插到上一条文档的文末
    p.js("document.getElementById('input').value=''")
    p.js("document.getElementById('send').click()")
    img_ok = p.wait_for("document.getElementById('fb-tip').classList.contains('on') "
                        "&& !!document.querySelector('#fb-tip [data-fb=\"image\"]')", 90)
    check("图片插入后弹出评价提示", img_ok)
    if not img_ok:
        print("  最后一张卡片：" +
              str(p.js("(function(){var a=document.querySelectorAll('#msgs .card');"
                       "return a.length? a[a.length-1].innerText : 'none';})()"))[:300], flush=True)
    check("图片插完自动收起待插入条",
          not p.js("document.getElementById('img-note').classList.contains('on')"))
    p.shot("评价-7-图片插入评价提示.png")
    p.js("document.querySelector('#fb-tip [data-fb=\"image\"]').click()", 0.5)
    check("图片评价弹窗功能=图片插入",
          p.js("document.getElementById('fb-feature').value") == "image",
          p.js("document.getElementById('fb-feature').value"))
    p.js("document.getElementById('fb-cancel').click()", 0.3)

    # ---------------- 4. 客服面板里的「评价这次咨询」
    p.js("document.getElementById('cs-ai').click()", 2.2)
    check("客服面板打开且有评价按钮", p.js("!!document.getElementById('cs-fb')"))
    p.js("document.getElementById('cs-fb').click()", 0.6)
    check("客服评价弹窗功能=客服咨询",
          p.js("document.getElementById('fb-feature').value") == "chat",
          p.js("document.getElementById('fb-feature').value"))
    p.shot("评价-8-客服面板评价按钮.png")

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

raise SystemExit(1 if bad else 0)
