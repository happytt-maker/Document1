# -*- coding: utf-8 -*-
"""端到端演示：让 Edge 的「假麦克风」播放一句真语音，走完 录音→上传→本机识别→填进输入框。

headless 浏览器没有真麦克风，Chromium 支持 --use-file-for-fake-audio-capture=<wav>：
把 wav 当成麦克风输入。这样既能在无头环境验收完整链路，也能截到「识别结果进输入框」的图。
"""
import asyncio
import base64
import json
import os
import subprocess
import time
import urllib.request
import wave

import av
import edge_tts
from websockets.sync.client import connect

BASE = "http://127.0.0.1:8000"
EDGE = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
WORK = r"C:\Users\覃敬涛\WorkBuddy\2026-09-24-10-48-03"
OUT = r"C:\Users\覃敬涛\Desktop\研一项目\agent\输出"
PORT = 9351
SAY = "把文档里的项目名称和合同金额提取出来"
op = urllib.request.build_opener(urllib.request.ProxyHandler({}))


def mp3_to_wav(src, dst, rate=48000):
    """mp3 → 单声道 16bit wav（假麦克风只吃 wav）"""
    c = av.open(src)
    st = c.streams.audio[0]
    res = av.AudioResampler(format="s16", layout="mono", rate=rate)
    frames = []
    for fr in c.decode(st):
        frames.extend(res.resample(fr))
    frames.extend(res.resample(None))
    with wave.open(dst, "wb") as w:
        w.setnchannels(1)
        w.setsampwidth(2)
        w.setframerate(rate)
        for f in frames:
            w.writeframes(bytes(f.to_ndarray().tobytes()))
    return os.path.getsize(dst)


def login(u="demo_ai", p="123456", portal="user"):
    req = urllib.request.Request(BASE + "/api/auth/login",
                                 data=json.dumps({"username": u, "password": p, "portal": portal}).encode(),
                                 headers={"Content-Type": "application/json"}, method="POST")
    r = op.open(req, timeout=10)
    for part in (r.headers.get("Set-Cookie") or "").split(";"):
        part = part.strip()
        if part.startswith("a23_token="):
            return part.split("=", 1)[1]
    raise RuntimeError("no cookie")


class Page:
    def __init__(self, url):
        tgt = json.load(op.open(urllib.request.Request(
            "http://127.0.0.1:%d/json/new?%s" % (PORT, url), method="PUT"), timeout=5))
        self.ws = connect(tgt["webSocketDebuggerUrl"], max_size=None, open_timeout=10)
        self.i = 0
        for m in ("Page.enable", "Runtime.enable"):
            self.cmd(m)

    def cmd(self, method, params=None):
        self.i += 1
        mid = self.i
        self.ws.send(json.dumps({"id": mid, "method": method, "params": params or {}}))
        while True:
            m = json.loads(self.ws.recv())
            if m.get("id") == mid:
                if "error" in m:
                    raise RuntimeError(method + " -> " + json.dumps(m["error"]))
                return m.get("result", {})

    def js(self, expr, wait=0.0):
        r = self.cmd("Runtime.evaluate", {"expression": expr, "returnByValue": True,
                                          "awaitPromise": True, "userGesture": True})
        if wait:
            time.sleep(wait)
        if r.get("exceptionDetails"):
            raise RuntimeError(json.dumps(r["exceptionDetails"])[:300])
        return (r.get("result") or {}).get("value")

    def wait_for(self, expr, timeout=30.0):
        t0 = time.time()
        while time.time() - t0 < timeout:
            try:
                if self.js(expr):
                    return True
            except Exception:
                pass
            time.sleep(0.25)
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
    tk = login()
    mp3 = os.path.join(WORK, "_asr_tmp", "demo.mp3")
    wav = os.path.join(WORK, "_asr_tmp", "demo.wav")
    os.makedirs(os.path.dirname(mp3), exist_ok=True)
    asyncio.run(edge_tts.Communicate(SAY, "zh-CN-XiaoxiaoNeural").save(mp3))
    print("  wav %.1fKB" % (mp3_to_wav(mp3, wav) / 1024), flush=True)

    prof = os.path.join(WORK, "_cdp_demo_%d" % int(time.time()))
    proc = subprocess.Popen([
        EDGE, "--headless=new", "--disable-gpu", "--no-first-run", "--no-default-browser-check",
        "--disable-extensions", "--disable-sync", "--disable-background-networking",
        "--disable-component-update", "--no-proxy-server", "--remote-allow-origins=*",
        "--use-fake-ui-for-media-stream", "--use-fake-device-for-media-stream",
        "--use-file-for-fake-audio-capture=" + wav,
        "--autoplay-policy=no-user-gesture-required",
        "--remote-debugging-port=%d" % PORT, "--user-data-dir=" + prof,
        "--window-size=1280,900", "about:blank",
    ], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

    log = []
    try:
        for _ in range(40):
            try:
                op.open("http://127.0.0.1:%d/json/version" % PORT, timeout=3)
                break
            except Exception:
                time.sleep(0.5)
        p = Page("about:blank")
        p.cmd("Network.enable")
        p.cmd("Network.setCookie", {"name": "a23_token", "value": tk, "url": BASE,
                                    "path": "/", "httpOnly": True})
        p.cmd("Page.navigate", {"url": BASE + "/"})
        time.sleep(3)
        p.js("localStorage.setItem('a23_asr_local','1'); input.value=''")
        p.js("document.getElementById('voice-btn').click()")
        p.wait_for("document.getElementById('voice-text').textContent.indexOf('本机识别')>=0", 10)
        print("  正在录……", flush=True)
        time.sleep(4.2)                       # 让「假麦克风」把这句话播完
        p.js("document.getElementById('voice-btn').click()")   # 说完了，结束
        done = p.wait_for("document.getElementById('voice-label').textContent=='语音输入'", 90)
        time.sleep(0.6)
        typed = str(p.js("input.value") or "")
        note = str(p.js("document.getElementById('voice-text').textContent") or "")
        p.shot("语音-识别结果已填入.png", "#input", pad=10, scale=1.6)
        p.shot("语音-整条输入区.png", ".composer", pad=6, scale=1.4)
        log.append("说：%s" % SAY)
        log.append("填进输入框：%s" % typed)
        log.append("提示：%s" % note)
        log.append("走完流程：%s" % done)
        ok = bool(typed.strip()) and ("项目" in typed or "金额" in typed)
        log.append("结论：%s" % ("PASS 端到端通了" if ok else "FAIL 没拿到预期文字"))
    finally:
        proc.terminate()
    open(os.path.join(WORK, "_demo_voice.txt"), "w", encoding="utf-8").write("\n".join(log))
    print("\n".join(log), flush=True)


if __name__ == "__main__":
    main()
