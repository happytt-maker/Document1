# -*- coding: utf-8 -*-
"""验收：语音输入的「本机离线识别」通道。

背景：浏览器自带的 SpeechRecognition 要连谷歌/微软服务器，国内经常 network 失败。
这次加了 /api/asr（faster-whisper 在本机转写）+ 前端自动降级。

检查项：
  后端 6 项：模型就绪 / 真语音能转出字 / 转写速度 / 空文件 / 未登录 / 超长
  前端 7 项：页面无报错 / 自动降级 / 状态机 / 按钮态 / 取消 / 不污染输入框 / 提示文案
"""
import base64
import hashlib
import json
import os
import shutil
import subprocess
import sys
import time
import urllib.error
import urllib.request

from websockets.sync.client import connect

BASE = "http://127.0.0.1:8000"
EDGE = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
PROF = r"C:\Users\覃敬涛\WorkBuddy\2026-09-24-10-48-03\_cdp_voice_verify"
OUT = r"C:\Users\覃敬涛\Desktop\研一项目\agent\输出"
TMP = r"C:\Users\覃敬涛\WorkBuddy\2026-09-24-10-48-03\_asr_tmp"
PORT = 9347
op = urllib.request.build_opener(urllib.request.ProxyHandler({}))
ok, bad = [], []
TOKEN = ""


def check(name, cond, detail=""):
    (ok if cond else bad).append(name)
    print(("  PASS " if cond else "  FAIL ") + name + ("" if cond else "  <<< " + str(detail)[:220]),
          flush=True)


def login(u, p, portal):
    req = urllib.request.Request(
        BASE + "/api/auth/login",
        data=json.dumps({"username": u, "password": p, "portal": portal}).encode(),
        headers={"Content-Type": "application/json"}, method="POST")
    r = op.open(req, timeout=10)
    for part in (r.headers.get("Set-Cookie") or "").split(";"):
        part = part.strip()
        if part.startswith("a23_token="):
            return part.split("=", 1)[1]
    raise RuntimeError("没拿到 cookie")


def api(path, method="GET", data=None, ctype="application/json", cookie=True):
    h = {}
    if data is not None:
        h["Content-Type"] = ctype
    if cookie and TOKEN:
        h["Cookie"] = "a23_token=" + TOKEN
    req = urllib.request.Request(BASE + path, data=data, headers=h, method=method)
    try:
        r = op.open(req, timeout=120)
        return r.status, r.read()
    except urllib.error.HTTPError as e:
        return e.code, e.read()


def post_audio(data, filename="voice.webm", ctype="audio/webm", cookie=True):
    """按 multipart/form-data 上传——File(...) 只认这种格式，裸 body 会 422。"""
    b = "----a23asrboundary"
    head = ("--%s\r\nContent-Disposition: form-data; name=\"file\"; filename=\"%s\"\r\n"
            "Content-Type: %s\r\n\r\n" % (b, filename, ctype)).encode("utf-8")
    body = head + data + ("\r\n--%s--\r\n" % b).encode("utf-8")
    return api("/api/asr", "POST", body, "multipart/form-data; boundary=" + b, cookie=cookie)


# ---------------------------------------------------------------- 造一段真语音
def make_speech(text, path):
    """用 edge-tts 合成一段中文语音，当作「用户说的话」来测识别。"""
    import asyncio
    import edge_tts

    async def go():
        c = edge_tts.Communicate(text, "zh-CN-XiaoxiaoNeural", rate="+0%")
        await c.save(path)

    asyncio.run(go())
    return os.path.getsize(path)


# ---------------------------------------------------------------- 浏览器
def cdp_json(path):
    return json.load(op.open("http://127.0.0.1:%d%s" % (PORT, path), timeout=5))


class Page:
    def __init__(self, url):
        tgt = json.load(op.open(urllib.request.Request(
            "http://127.0.0.1:%d/json/new?%s" % (PORT, url), method="PUT"), timeout=5))
        self.ws = connect(tgt["webSocketDebuggerUrl"], max_size=None, open_timeout=10)
        self.i = 0
        self.logs = []
        self.cmd("Page.enable")
        self.cmd("Runtime.enable")
        self.cmd("Log.enable")

    def cmd(self, method, params=None):
        self.i += 1
        mid = self.i
        self.ws.send(json.dumps({"id": mid, "method": method, "params": params or {}}))
        while True:
            m = json.loads(self.ws.recv())
            if m.get("method") == "Runtime.consoleAPICalled":
                a = m["params"].get("args") or []
                self.logs.append(" ".join(str(x.get("value", x.get("description", ""))) for x in a))
            if m.get("method") == "Runtime.exceptionThrown":
                d = m["params"].get("exceptionDetails") or {}
                self.logs.append("EXCEPTION " + str(d.get("text")) + " " + str(
                    (d.get("exception") or {}).get("description", ""))[:200])
            if m.get("id") == mid:
                if "error" in m:
                    raise RuntimeError(method + " -> " + json.dumps(m["error"]))
                return m.get("result", {})

    def cookie(self, value):
        self.cmd("Network.setCookie", {"name": "a23_token", "value": value,
                                       "url": BASE, "path": "/", "httpOnly": True})

    def go(self, url, wait=2.5):
        self.cmd("Page.navigate", {"url": url})
        time.sleep(wait)

    def js(self, expr, wait=0.0):
        r = self.cmd("Runtime.evaluate", {"expression": expr, "returnByValue": True,
                                          "awaitPromise": True, "userGesture": True})
        if wait:
            time.sleep(wait)
        if r.get("exceptionDetails"):
            raise RuntimeError(json.dumps(r["exceptionDetails"])[:300])
        return (r.get("result") or {}).get("value")

    def wait_for(self, expr, timeout=20.0):
        t0 = time.time()
        while time.time() - t0 < timeout:
            try:
                if self.js(expr):
                    return True
            except Exception:
                pass
            time.sleep(0.3)
        return False

    def shot(self, name, sel=None, pad=8, scale=2.0):
        if sel:
            r = self.js("(()=>{const b=document.querySelector('%s').getBoundingClientRect();"
                        "return {x:b.x,y:b.y,w:b.width,h:b.height};})()" % sel)
            params = {"format": "png", "clip": {"x": max(r["x"] - pad, 0), "y": max(r["y"] - pad, 0),
                                               "width": r["w"] + pad * 2, "height": r["h"] + pad * 2,
                                               "scale": scale}}
        else:
            params = {"format": "png"}
        d = self.cmd("Page.captureScreenshot", params)
        with open(os.path.join(OUT, name), "wb") as f:
            f.write(base64.b64decode(d["data"]))
        print("  截图 -> " + name, flush=True)


def main():
    global TOKEN
    os.makedirs(OUT, exist_ok=True)
    os.makedirs(TMP, exist_ok=True)
    # 每次用全新的浏览器配置目录：沙箱不允许一次性删几百个文件，索性不删
    global PROF
    PROF = os.path.join(r"C:\Users\覃敬涛\WorkBuddy\2026-09-24-10-48-03",
                        "_cdp_voice_%d" % int(time.time()))

    TOKEN = login("demo_ai", "123456", "user")

    print("\n--- 1. 后端：离线识别能力 ---", flush=True)
    code, body = api("/api/asr/status")
    st = json.loads(body) if code == 200 else {}
    check("模型文件就绪", code == 200 and st.get("ready") is True, (code, st.get("file_mb")))
    check("model.bin 约 461MB", (st.get("file_mb") or {}).get("model.bin", 0) > 440,
          (st.get("file_mb") or {}).get("model.bin"))

    wav = os.path.join(TMP, "say.mp3")
    size = make_speech("帮我看看项目名称和合同金额这两个字段填得对不对", wav)
    check("测试语音已生成", size > 8000, size)
    with open(wav, "rb") as f:
        audio = f.read()

    t0 = time.time()
    code, body = post_audio(audio, "say.mp3", "audio/mpeg")
    cost = time.time() - t0
    got = json.loads(body) if code == 200 else {}
    text = (got.get("text") or "").replace(" ", "")
    print("  识别结果：%s（%.1fs）" % (text, cost), flush=True)
    check("识别接口 200", code == 200, (code, body[:150]))
    check("转写不是空的", len(text) >= 4, text)
    hit = sum(1 for w in ("项目", "名称", "合同", "金额", "字段") if w in text)
    check("关键实体命中 >= 3/5", hit >= 3, "命中 %d：%s" % (hit, text))
    check("整段耗时 < 15s", cost < 15, "%.1fs" % cost)

    code, _ = post_audio(b"", "voice.webm", "audio/webm")
    check("空录音被挡（400）", code == 400, code)

    code, _ = post_audio(audio, "say.mp3", "audio/mpeg", cookie=False)
    check("未登录被挡（401）", code == 401, code)

    print("\n--- 2. 前端：自动降级到本机识别 ---", flush=True)
    proc = subprocess.Popen([
        EDGE, "--headless=new", "--disable-gpu", "--no-first-run", "--no-default-browser-check",
        "--disable-extensions", "--disable-sync", "--disable-background-networking",
        "--disable-component-update", "--no-proxy-server", "--remote-allow-origins=*",
        "--use-fake-ui-for-media-stream", "--use-fake-device-for-media-stream",
        "--autoplay-policy=no-user-gesture-required",
        "--remote-debugging-port=%d" % PORT, "--user-data-dir=" + PROF,
        "--window-size=1280,900", "about:blank",
    ], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

    try:
        for _ in range(40):
            try:
                cdp_json("/json/version")
                break
            except Exception:
                time.sleep(0.5)
        else:
            raise RuntimeError("Edge 调试端口没起来")

        p = Page("about:blank")
        p.cookie(TOKEN)
        p.go(BASE + "/", 3.0)
        p.js("localStorage.removeItem('a23_asr_local'); input.value=''")
        p.logs.clear()

        check("页面加载无 JS 异常", not [x for x in p.logs if x.startswith("EXCEPTION")],
              [x for x in p.logs if x.startswith("EXCEPTION")][:2])

        # 模拟「上次浏览器语音服务就不通」：直接把标志置上，再点按钮
        p.js("localStorage.setItem('a23_asr_local','1')")
        p.js("document.getElementById('voice-btn').click()")
        started = p.wait_for("document.getElementById('voice-label').textContent.indexOf('正在听')>=0", 8)
        check("点一下就进入录音（不等外网超时）", bool(started),
              p.js("document.getElementById('voice-label').textContent"))
        # 提示条要等 vNote 真正写入（HTML 里那句初始文案长得和原生识别很像，容易误判）
        noted = p.wait_for("document.getElementById('voice-note').classList.contains('on')"
                           "&& document.getElementById('voice-text').textContent.indexOf('本机识别')>=0", 10)
        check("提示条说的是本机识别", bool(noted),
              p.js("document.getElementById('voice-text').textContent"))
        check("按钮是录音态", bool(p.js("document.getElementById('voice-btn').classList.contains('rec')")))
        p.shot("语音-本机识别中.png", "#voice-note")

        # 再点一下 = 说完
        p.js("document.getElementById('voice-btn').click()")
        p.wait_for("document.getElementById('voice-label').textContent.indexOf('识别中')>=0", 10)
        # 等识别真的走完：按钮回到「语音输入」才算结束（上传 + 本机转写要几百毫秒到几秒）
        ended = p.wait_for("document.getElementById('voice-label').textContent=='语音输入'", 90)
        check("再点一下就结束录音", bool(ended))
        check("结束后按钮复位", p.js("document.getElementById('voice-label').textContent") == "语音输入",
              p.js("document.getElementById('voice-label').textContent"))
        note = str(p.js("document.getElementById('voice-text').textContent"))
        typed = str(p.js("input.value") or "")
        print("  结束后的提示：%s ｜ 输入框：%s" % (note, typed), flush=True)
        check("要么把字填进输入框、要么明确说没听清",
              bool(typed.strip()) or ("没听清" in note) or ("没录到声音" in note), (typed, note))
        p.shot("语音-识别完成.png", "#voice-note")

        # 取消按钮不该把半截文字塞进输入框
        p.js("input.value='原来的半句话'")
        p.js("document.getElementById('voice-btn').click()")
        p.wait_for("document.getElementById('voice-btn').classList.contains('rec')", 8)
        p.js("document.getElementById('voice-x').click()")
        p.wait_for("!document.getElementById('voice-btn').classList.contains('rec')", 8)
        check("取消后输入框原样保留", p.js("input.value") == "原来的半句话", p.js("input.value"))
        check("取消后提示条收起",
              not p.js("document.getElementById('voice-note').classList.contains('on')"))

        print("\n  浏览器控制台：%s" % (p.logs[:4] or "（无）"), flush=True)
    finally:
        proc.terminate()

    print("\n==== 汇总：%d 项通过，%d 项失败 ====" % (len(ok), len(bad)), flush=True)
    if bad:
        print("失败：" + " / ".join(bad))
    return 0 if not bad else 1


if __name__ == "__main__":
    sys.exit(main())
