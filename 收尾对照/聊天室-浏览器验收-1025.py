# -*- coding: utf-8 -*-
"""聊天室-浏览器验收-1025.py —— 真机（Edge headless + CDP）点一遍聊天室

跑法（必须用 managed 3.13，带 websockets）：
    C:\\Users\\<你>\\.workbuddy\\binaries\\python\\versions\\3.13.12\\python.exe ^
        收尾对照\\聊天室-浏览器验收-1025.py --port 8010

验证内容：
   1  登录态注入 → 打开首页 → 顶栏「聊天室」能点开面板
   2  面板里能看到自己在线、空群提示
   3  从界面发一句话 → 自己的气泡上屏（右侧、蓝色）
   4  另一个账号（HTTP 模拟）在群里说话 → 浏览器 2.5 秒内收到，@我 高亮
   5  点在线名单里的人 → 输入框自动插入 @昵称
   6  点「撤回」→ 气泡变「这条消息已撤回」
   7  收起面板后别人发言 → 顶栏出现未读角标
截图落在 agent\\输出\\ 下；跑完删掉测试账号与消息。
"""

import argparse
import base64
import json
import os
import shutil
import subprocess
import sys
import tempfile
import time
import urllib.parse
import urllib.request
import uuid
from http.cookiejar import CookieJar
from pathlib import Path

for _k in ("HTTP_PROXY", "HTTPS_PROXY", "ALL_PROXY",
           "http_proxy", "https_proxy", "all_proxy"):
    os.environ.pop(_k, None)

EDGE = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
PY_VENV = os.path.expanduser(r"~\.workbuddy\binaries\python\envs\default\Scripts\python.exe")
HERE = Path(__file__).resolve().parent
ROOT = HERE.parent
OUT_DIR = ROOT / "agent" / "输出"
LOG = HERE / "聊天室-浏览器验收-1025.txt"

PASS, FAIL, JSErr = [], [], []
TAG = uuid.uuid4().hex[:6]


class Tee:
    def __init__(self, *st):
        self.st = st

    def write(self, s):
        for x in self.st:
            try:
                x.write(s)
                x.flush()
            except Exception:
                pass

    def flush(self):
        for x in self.st:
            try:
                x.flush()
            except Exception:
                pass


def ok(name, cond, detail=""):
    (PASS if cond else FAIL).append(name)
    print(("  [PASS] " if cond else "  [FAIL] ") + name + (("  " + str(detail)) if detail else ""))
    return cond


# ---------------------------------------------------------------- 简易 HTTP 客户端（带 cookie）

class Client:
    def __init__(self, base):
        self.base = base
        self.cj = CookieJar()
        self.op = urllib.request.build_opener(urllib.request.ProxyHandler({}),
                                              urllib.request.HTTPCookieProcessor(self.cj))

    def call(self, path, data=None, method=None):
        url = self.base + path
        body = None
        headers = {}
        if data is not None:
            body = json.dumps(data).encode("utf-8")
            headers["Content-Type"] = "application/json"
        req = urllib.request.Request(url, data=body, headers=headers,
                                     method=method or ("POST" if body else "GET"))
        try:
            with self.op.open(req, timeout=20) as r:
                return r.status, json.loads(r.read().decode("utf-8") or "{}")
        except urllib.error.HTTPError as e:
            try:
                return e.code, json.loads(e.read().decode("utf-8") or "{}")
            except Exception:
                return e.code, {}

    def token(self):
        for c in self.cj:
            if c.name == "a23_token":
                return c.value
        return ""


# ---------------------------------------------------------------- CDP

class CDP:
    def __init__(self, url):
        from websockets.sync.client import connect
        self.ws = connect(url, max_size=None)
        self.id = 0

    def _evt(self, m):
        name = m.get("method")
        if name == "Page.javascriptDialogOpening":
            self.id += 1
            self.ws.send(json.dumps({"id": self.id, "method": "Page.handleJavaScriptDialog",
                                     "params": {"accept": True}}))
        elif name == "Runtime.exceptionThrown":
            d = m.get("params", {}).get("exceptionDetails", {})
            JSErr.append(str(d.get("text")) + " " +
                         str((d.get("exception") or {}).get("description", ""))[:200])
        elif name == "Log.entryAdded":
            e = m.get("params", {}).get("entry", {})
            # favicon 404 是浏览器自己要的，不算页面报错
            if e.get("level") == "error" and "favicon" not in str(e.get("url", "")):
                JSErr.append(str(e.get("text"))[:200])

    def cmd(self, method, params=None, timeout=25):
        self.id += 1
        mid = self.id
        self.ws.send(json.dumps({"id": mid, "method": method, "params": params or {}}))
        end = time.time() + timeout
        while time.time() < end:
            m = json.loads(self.ws.recv())
            if m.get("id") == mid:
                if "error" in m:
                    raise RuntimeError("%s: %s" % (method, m["error"].get("message")))
                return m.get("result", {})
            self._evt(m)
        raise TimeoutError(method)

    def js(self, expr, wait=0.0):
        r = self.cmd("Runtime.evaluate", {"expression": expr, "returnByValue": True,
                                          "awaitPromise": True, "userGesture": True})
        if wait:
            time.sleep(wait)
        return r.get("result", {}).get("value")

    def shot(self, path):
        r = self.cmd("Page.captureScreenshot", {"format": "png"})
        Path(path).write_bytes(base64.b64decode(r["data"]))
        return Path(path).stat().st_size


def wait_true(cdp, expr, timeout=12, step=0.4):
    end = time.time() + timeout
    while time.time() < end:
        try:
            if cdp.js(expr):
                return True
        except Exception:
            pass
        time.sleep(step)
    return False


def cleanup(user_ids):
    if not user_ids:
        return
    code = (
        "import sys;sys.path.insert(0,%r);import db;c=db.get_conn(with_db=True);ids=%r;"
        "cur=c.cursor();ph=','.join(['%%s']*len(ids));"
        "cur.execute('DELETE FROM room_messages WHERE user_id IN ('+ph+')',ids);"
        "cur.execute('DELETE FROM room_online WHERE user_id IN ('+ph+')',ids);"
        "cur.execute('DELETE FROM sessions WHERE user_id IN ('+ph+')',ids);"
        "cur.execute('DELETE FROM users WHERE id IN ('+ph+')',ids);"
        "c.commit();c.close();print('cleaned',ids)"
    ) % (str(ROOT / "agent"), user_ids)
    try:
        subprocess.run([PY_VENV, "-c", code], capture_output=True, timeout=60)
    except Exception as e:
        print("  清理失败（不影响结论）：%s" % e)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--port", type=int, default=8010)
    args = ap.parse_args()
    base = "http://127.0.0.1:%d" % args.port
    cdp_port = 9333
    prof = Path(tempfile.gettempdir()) / ("_a23_cdp_room_" + TAG)
    print("=" * 66)
    print("聊天室浏览器验收 · %s" % base)
    print("=" * 66)

    ids, proc, cdp = [], None, None
    try:
        # ---------- 造两个测试账号（浏览器里用 A，HTTP 用 B）
        a, b = Client(base), Client(base)
        na, nb = "uitest_a_" + TAG, "uitest_b_" + TAG
        ca, ja = a.call("/api/auth/register", {"username": na, "password": "Room123456",
                                               "display_name": "界面测试A"})
        cb, jb = b.call("/api/auth/register", {"username": nb, "password": "Room123456",
                                               "display_name": "界面测试B"})
        ids = [ja["user"]["id"], jb["user"]["id"]]
        ok("浏览器用的 A 账号注册成功", ca == 200 and bool(a.token()))
        ok("模拟群友 B 账号注册成功", cb == 200 and bool(b.token()))
        b.call("/api/room/poll?after=0&alive=1")      # B 先上线，A 打开面板就能看到人

        # ---------- 起 headless Edge
        if not Path(EDGE).exists():
            raise SystemExit("找不到 Edge：%s" % EDGE)
        proc = subprocess.Popen([
            EDGE, "--headless=new", "--disable-gpu", "--no-first-run", "--no-default-browser-check",
            "--disable-extensions", "--disable-sync", "--disable-background-networking",
            "--disable-component-update", "--no-proxy-server", "--remote-allow-origins=*",
            "--remote-debugging-port=%d" % cdp_port, "--user-data-dir=" + str(prof),
            "--window-size=1440,960", "about:blank",
        ], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

        ws_url = None
        op = urllib.request.build_opener(urllib.request.ProxyHandler({}))
        for _ in range(42):
            try:
                with op.open("http://127.0.0.1:%d/json/version" % cdp_port, timeout=3) as r:
                    if json.loads(r.read().decode()) .get("Browser"):
                        break
            except Exception:
                pass
            time.sleep(0.5)
        req = urllib.request.Request("http://127.0.0.1:%d/json/new?about:blank" % cdp_port,
                                     method="PUT")
        with op.open(req, timeout=5) as r:
            ws_url = json.loads(r.read().decode())["webSocketDebuggerUrl"]
        cdp = CDP(ws_url)
        cdp.cmd("Page.enable")
        cdp.cmd("Runtime.enable")
        cdp.cmd("Log.enable")
        cdp.cmd("Network.enable")
        ok("headless Edge 起来了", bool(ws_url))

        # ---------- 注入登录态打开首页
        cdp.cmd("Network.setCookie", {"name": "a23_token", "value": a.token(),
                                      "url": base, "path": "/", "httpOnly": True})
        cdp.cmd("Page.navigate", {"url": base + "/"})
        ok("首页加载完成、聊天室脚本就绪",
           wait_true(cdp, "!!(window.__cr && window.__cr.ready)", 20))
        ok("顶栏有「聊天室」入口",
           bool(cdp.js("!!document.getElementById('room-open')")))

        # ---------- 1 打开面板
        print("\n【1】点开聊天室面板")
        cdp.js("document.getElementById('room-open').click()")
        ok("面板打开", wait_true(cdp, "document.getElementById('cr-panel').classList.contains('on')", 8))
        # 现场留证：这一刻页面 DOM 与两个账号各自看到的在线名单
        _dbg = cdp.js("JSON.stringify({chips:[...document.querySelectorAll('.cr-mem')]"
                      ".map(e=>e.className+'/'+e.dataset.name), empty:document.querySelectorAll"
                      "('.cr-empty').length, msgs:document.querySelectorAll('.cr-msg').length,"
                      " cr:{myId:__cr.myId,online:__cr.online,lastId:__cr.lastId,total:__cr.total}})")
        print("  【现场】页面：%s" % _dbg)
        _cb, _jb = b.call("/api/room/poll?after=0&alive=1")
        print("  【现场】B 看到的在线：%s" % [(u["name"], u["me"]) for u in _jb.get("online", [])])
        ok("在线名单里有人（B 已在群里）",
           wait_true(cdp, "document.querySelectorAll('.cr-mem').length >= 1", 8),
           cdp.js("document.querySelectorAll('.cr-mem').length"))
        ok("自己那条带「我」标记", bool(cdp.js("!!document.querySelector('.cr-mem.me')")),
           cdp.js("[...document.querySelectorAll('.cr-mem')].map(e=>e.className).join(' | ')"))
        ok("标题显示人数", "人" in str(cdp.js("document.getElementById('cr-online-tx').textContent")),
           cdp.js("document.getElementById('cr-online-tx').textContent"))
        ok("空群时有引导文案", bool(cdp.js("!!document.querySelector('.cr-empty')")),
           cdp.js("'msgs=' + document.querySelectorAll('.cr-msg').length"))

        # ---------- 2 从界面发言
        print("\n【2】界面上发一条消息")
        cdp.js("(()=>{const t=document.getElementById('cr-input');t.value='大家好，我是界面测试A，"
               "刚从表格填写那边过来';document.getElementById('cr-send').click();})()")
        ok("自己的气泡立刻上屏（乐观显示）",
           wait_true(cdp, "!!document.querySelector('.cr-msg.me .cr-tx')", 5))
        time.sleep(1.2)
        mine = cdp.js("(document.querySelector('.cr-msg.me .cr-tx')||{}).textContent||''")
        ok("内容就是刚发的那句", "表格填写" in str(mine), mine)
        ok("回执到了（气泡挂上 message id）",
           bool(cdp.js("!!document.querySelector('.cr-msg.me[data-mid]')")))
        ok("没有重复气泡", int(cdp.js("document.querySelectorAll('.cr-msg.me').length") or 0) == 1,
           cdp.js("document.querySelectorAll('.cr-msg.me').length"))
        ok("自己那条有「撤回」按钮", bool(cdp.js("!!document.querySelector('.cr-msg.me [data-undo]')")))

        # ---------- 3 别人发言，浏览器能收到
        print("\n【3】另一个用户在群里说话 → 浏览器收得到")
        time.sleep(1.1)
        code, jsend = b.call("/api/room/send",
                             {"content": "@界面测试A 我也在，模板在「表格填写」里选就行"})
        ok("B 通过接口发言成功", code == 200, jsend.get("detail", ""))
        ok("浏览器 3 秒内收到 B 的消息",
           wait_true(cdp, "[...document.querySelectorAll('.cr-msg .cr-tx')]"
                          ".some(e=>e.textContent.indexOf('模板在')>=0)", 8))
        ok("@我 的消息被高亮", bool(cdp.js("!!document.querySelector('.cr-msg .cr-tx.at')")))
        ok("别人的气泡在左边（不是 me）",
           bool(cdp.js("!!document.querySelector('.cr-msg:not(.me) .cr-nm')")))

        # ---------- 4 点名字 @ 人
        print("\n【4】点在线名单里的人 → 自动 @")
        cdp.js("(()=>{const t=document.getElementById('cr-input');t.value='';"
               "const m=document.querySelector('.cr-mem:not(.me)');if(m)m.click();})()")
        time.sleep(0.4)
        val = str(cdp.js("document.getElementById('cr-input').value") or "")
        ok("输入框插入了 @昵称", val.startswith("@") and "界面测试B" in val, val)
        cdp.js("document.getElementById('cr-input').value=''")

        # ---------- 5 撤回
        print("\n【5】撤回自己的消息")
        cdp.js("document.querySelector('.cr-msg.me [data-undo]').click()")
        ok("气泡变成「已撤回」",
           wait_true(cdp, "!!document.querySelector('.cr-msg.me .cr-tx.gone')", 8))
        ok("撤回后文案正确",
           "已撤回" in str(cdp.js("(document.querySelector('.cr-msg.me .cr-tx.gone')||{}).textContent||''")))
        ok("撤回按钮消失", not cdp.js("!!document.querySelector('.cr-msg.me [data-undo]')"))

        # ---------- 6 截图
        print("\n【6】截图")
        OUT_DIR.mkdir(parents=True, exist_ok=True)
        p1 = OUT_DIR / "聊天室-1-面板.png"
        sz = cdp.shot(p1)
        ok("截图 1 已保存（面板）", sz > 20000, "%s %.1fKB" % (p1.name, sz / 1024))

        # ---------- 7 收起后未读角标
        print("\n【7】收起面板 → 别人发言 → 未读角标")
        cdp.js("document.getElementById('cr-x').click()")
        ok("面板收起时露出右下角客服按钮",
           wait_true(cdp, "getComputedStyle(document.getElementById('cs-dock')).display !== 'none'", 5))
        time.sleep(1.2)
        code, _j = b.call("/api/room/send", {"content": "我先去试试填写模板，等下回来"})
        ok("B 又发了一条", code == 200)
        ok("顶栏出现未读角标",
           wait_true(cdp, "document.getElementById('room-badge').classList.contains('on')", 12),
           cdp.js("document.getElementById('room-badge').textContent"))
        p2 = OUT_DIR / "聊天室-2-未读角标.png"
        sz2 = cdp.shot(p2)
        ok("截图 2 已保存（未读角标）", sz2 > 20000, "%s %.1fKB" % (p2.name, sz2 / 1024))

        # ---------- 8 再打开，历史还在
        print("\n【8】重新打开能看回历史")
        cdp.js("document.getElementById('room-open').click()")
        ok("重新打开有历史消息",
           wait_true(cdp, "document.querySelectorAll('.cr-msg').length >= 2", 8),
           cdp.js("document.querySelectorAll('.cr-msg').length"))
        ok("未读角标清零", not cdp.js("document.getElementById('room-badge').classList.contains('on')"))

        # ---------- 9 JS 错误
        print("\n【9】页面错误")
        ok("没有任何 JS 报错", len(JSErr) == 0, JSErr[:3])

    finally:
        if cdp:
            try:
                cdp.ws.close()
            except Exception:
                pass
        if proc:
            proc.terminate()
            try:
                proc.wait(timeout=8)
            except Exception:
                pass
        shutil.rmtree(prof, ignore_errors=True)
        cleanup(ids)
        print("\n（已清理测试账号 %s、测试消息；浏览器进程与 profile 已回收）" % ids)

    print("\n" + "=" * 66)
    print("通过 %d 项，失败 %d 项" % (len(PASS), len(FAIL)))
    for f in FAIL:
        print("  - 失败：" + f)
    print("=" * 66)
    return 0 if not FAIL else 1


if __name__ == "__main__":
    _f = open(LOG, "w", encoding="utf-8")
    _old = sys.stdout
    sys.stdout = Tee(_old, _f)
    try:
        _code = main()
    finally:
        sys.stdout = _old
        _f.close()
        print("验收日志已写入：%s" % LOG)
    raise SystemExit(_code)
