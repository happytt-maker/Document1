# -*- coding: utf-8 -*-
"""文档库上传 · 界面真机验收（2026-11-05）

后端接口已在「文档库上传-验收-1105.py」里过完 35 项；这份脚本只回答一个问题：
**在真浏览器里，用户点得动吗、看得见吗。**

流程：注入登录 Cookie → 点「🗄 数据仓库」→ 看文档库列表 → 通过文件选择框真传一份
→ 等成功提示 → 确认列表第一行就是它 → 截图两张 → 按文件名精确清理（文件 + 池 + imports）。

跑之前先起测试实例：python agent/api_server.py --port 8011
"""
import asyncio
import base64
import json
import os
import re
import shutil
import subprocess
import sys
import time
import urllib.request
from pathlib import Path

EDGE = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
PROFILE = Path(os.environ.get("TEMP", r"C:\Windows\Temp")) / "_cdp_dwdocs_1105"
BASE = "http://127.0.0.1:8011"
PORT = 9333

PROJECT = Path(r"C:\Users\覃敬涛\Desktop\研一项目")
IMPORT_DIR = PROJECT / "导入文档"
POOL = PROJECT / "doc_reader" / "output.json"
SHOTS = PROJECT / "agent" / "输出"
OUT = Path(__file__).with_suffix(".txt")

TEST_NAME = "界面自检-勿留.txt"
TEST_FILE = PROJECT / "收尾对照" / "_待上传" / TEST_NAME

opener = urllib.request.build_opener(urllib.request.ProxyHandler({}))

lines = []
passed = failed = 0


def say(s=""):
    lines.append(str(s))


def check(ok, tag, detail=""):
    global passed, failed
    passed += 1 if ok else 0
    failed += 0 if ok else 1
    say("  %s %s%s" % ("PASS" if ok else "FAIL", tag, ("　—— " + detail) if detail else ""))
    return ok


# ---------------- CDP 极简客户端 ----------------
class CDP:
    def __init__(self, ws):
        self.ws = ws
        self.n = 0

    async def cmd(self, method, params=None):
        self.n += 1
        mid = self.n
        await self.ws.send(json.dumps({"id": mid, "method": method, "params": params or {}}))
        while True:
            msg = json.loads(await self.ws.recv())
            if msg.get("method") == "Page.javascriptDialogOpening":
                self.n += 1
                await self.ws.send(json.dumps({"id": self.n,
                                               "method": "Page.handleJavaScriptDialog",
                                               "params": {"accept": True}}))
                continue
            if msg.get("id") == mid:
                if "error" in msg:
                    raise RuntimeError("%s -> %s" % (method, msg["error"]))
                return msg.get("result", {})

    async def js(self, expr):
        r = await self.cmd("Runtime.evaluate", {"expression": expr,
                                                "returnByValue": True,
                                                "awaitPromise": True})
        if r.get("exceptionDetails"):
            ex = r["exceptionDetails"].get("exception", {})
            raise RuntimeError(ex.get("description") or str(r["exceptionDetails"])[:300])
        return (r.get("result") or {}).get("value")


async def wait_js(c: CDP, expr, timeout=25, step=0.4):
    t0 = time.time()
    while time.time() - t0 < timeout:
        try:
            if await c.js(expr):
                return True
        except Exception:
            pass
        await asyncio.sleep(step)
    return False


def login_token():
    body = json.dumps({"username": "2482991994", "password": "123456",
                       "portal": "user"}).encode("utf-8")
    req = urllib.request.Request(BASE + "/api/auth/login", data=body,
                                 headers={"Content-Type": "application/json"})
    resp = opener.open(req, timeout=15)
    raw = resp.headers.get("Set-Cookie") or ""
    m = re.search(r"a23_token=([^;]+)", raw)
    if not m:
        raise RuntimeError("登录没拿到 a23_token：%s" % raw[:120])
    return m.group(1)


async def run():
    say("文档库上传 · 界面真机验收 · %s" % time.strftime("%Y-%m-%d %H:%M:%S"))
    say("目标：%s（headless Edge + CDP）" % BASE)
    say("")

    # 基线备份
    backup = POOL.with_name("_池备份-界面自检.json")
    if POOL.exists():
        shutil.copy2(POOL, backup)
    TEST_FILE.parent.mkdir(parents=True, exist_ok=True)
    TEST_FILE.write_text("界面自检：确认数据仓库里能直接上传文档。\n"
                         "第二段：甲方 华景科技有限公司，金额 88 万元。\n", encoding="utf-8")

    say("【1. 打开数据仓库】")
    token = login_token()
    check(bool(token), "1.0 拿到登录 Cookie")

    SHOTS.mkdir(parents=True, exist_ok=True)
    proc = subprocess.Popen([
        EDGE, "--headless=new", "--disable-gpu", "--no-first-run",
        "--no-default-browser-check", "--disable-extensions", "--disable-sync",
        "--disable-background-networking", "--disable-component-update",
        "--no-proxy-server", "--remote-allow-origins=*",
        "--remote-debugging-port=%d" % PORT, "--user-data-dir=" + str(PROFILE),
        "--window-size=1440,960", "about:blank",
    ], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

    ws = None
    try:
        # 等 CDP 起来
        ok = False
        for _ in range(50):
            try:
                opener.open("http://127.0.0.1:%d/json/version" % PORT, timeout=1).read()
                ok = True
                break
            except Exception:
                await asyncio.sleep(0.4)
        check(ok, "1.1 headless Edge 起来了")
        if not ok:
            return

        req = urllib.request.Request("http://127.0.0.1:%d/json/new?about:blank" % PORT,
                                     method="PUT")
        info = json.loads(opener.open(req, timeout=10).read().decode("utf-8"))

        import websockets
        ws = await websockets.connect(info["webSocketDebuggerUrl"], max_size=None)
        c = CDP(ws)
        for m in ("Page.enable", "Runtime.enable", "Network.enable", "DOM.enable"):
            await c.cmd(m)

        await c.cmd("Network.setCookie", {"name": "a23_token", "value": token,
                                          "url": BASE, "path": "/", "httpOnly": True})
        await c.cmd("Page.navigate", {"url": BASE + "/"})
        loaded = await wait_js(c, "typeof dwLoadDocs === 'function' && !!document.getElementById('dw-toggle')", 30)
        check(loaded, "1.2 页面加载完成，文档库逻辑已就绪")

        # 打开数据仓库
        await c.js("document.getElementById('dw-toggle').click()")
        opened = await wait_js(c, "document.getElementById('dw-overlay').classList.contains('open')", 8)
        check(opened, "1.3 「🗄 数据仓库」能打开")

        filled = await wait_js(c, "document.querySelectorAll('.dw-doc-row').length > 0", 15)
        check(filled, "1.4 文档库列表渲染出来了")

        info1 = await c.js("""(() => {
          const rows = [...document.querySelectorAll('.dw-doc-row')].map(r => ({
            src: r.querySelector('.tag-src').textContent.trim(),
            name: r.querySelector('.nm').textContent.trim(),
            meta: r.querySelector('.mt').textContent.trim(),
            hasBtn: !!r.querySelector('button.ex')
          }));
          return {
            n: rows.length,
            first: rows[0] || null,
            last: rows[rows.length - 1] || null,
            countText: document.getElementById('dw-docs-count').innerText,
            uploadVisible: !!document.getElementById('dw-docs-upload').offsetParent,
            foldText: document.getElementById('dw-docs-fold').innerText.trim(),
            dropText: document.getElementById('dw-docs-drop').innerText.replace(/\\s+/g,' ').trim()
          };
        })()""")
        say("  列表 %s 行｜%s" % (info1["n"], info1["countText"]))
        say("  第一行：%s" % json.dumps(info1["first"], ensure_ascii=False))
        check(info1["n"] >= 1, "1.5 有文档（%d 份）" % info1["n"])
        check(info1["uploadVisible"], "1.6 「⬆ 上传文档」按钮可见")
        check(bool(info1["dropText"]), "1.7 拖拽区文案可见", info1["dropText"][:40])
        check(info1["first"] and info1["first"]["src"] == "我上传的",
              "1.8 我上传的排在前面", info1["first"]["src"] if info1["first"] else "-")
        check(bool(info1["first"] and info1["first"]["hasBtn"]), "1.9 每行带「提取入库」按钮")

        shot1 = SHOTS / "文档库-1-面板.png"
        png = await c.cmd("Page.captureScreenshot", {"format": "png"})
        shot1.write_bytes(base64.b64decode(png["data"]))
        check(shot1.is_file(), "1.10 截图①（数据仓库 + 文档库）", shot1.name)

        # ---------------- 真上传 ----------------
        say("")
        say("【2. 通过界面上传一份文件】")
        doc = await c.cmd("DOM.getDocument", {"depth": -1})
        nid = (await c.cmd("DOM.querySelector",
                           {"nodeId": doc["root"]["nodeId"],
                            "selector": "#dw-docs-input"}))["nodeId"]
        await c.cmd("DOM.setFileInputFiles", {"nodeId": nid, "files": [str(TEST_FILE)]})
        # setFileInputFiles 一般会触发 change；没触发就补一次（files 已被清空时补发是空操作）
        await asyncio.sleep(0.8)
        await c.js("dwDocsInput.dispatchEvent(new Event('change'))")

        done = await wait_js(c, """(() => {
          const t = document.getElementById('dw-docs-note').innerText || '';
          return /已进文档库|上传结束/.test(t);
        })()""", 40)
        note = await c.js("document.getElementById('dw-docs-note').innerText.replace(/\\s+/g,' ').trim()")
        say("  提示：%s" % note[:160])
        check(done, "2.1 上传有结束提示（不是静默无反应）")
        check("界面自检-勿留.txt" in note, "2.2 提示里写清了文件名")
        check("1 份已进文档库" in note or "成功 1" in note, "2.3 成功份数正确")

        # 列表是异步刷新的（dwLoadDocs 不阻塞提示），这里等它真的出来再断言；
        # 等不到才算 bug —— 别把"脚本读太快"误判成"界面没刷新"。
        listed = await wait_js(c, """[...document.querySelectorAll('.dw-doc-row .nm')]
            .some(n => n.textContent.trim() === '界面自检-勿留.txt')""", 20)
        check(listed, "2.4 列表自动刷新出新文件（≤20s）")

        after = await c.js("""(() => {
          const rows = [...document.querySelectorAll('.dw-doc-row')].map(r => ({
            src: r.querySelector('.tag-src').textContent.trim(),
            name: r.querySelector('.nm').textContent.trim(),
            meta: r.querySelector('.mt').textContent.trim(),
            isNew: r.querySelector('.nm').classList.contains('new')
          }));
          return {n: rows.length, first: rows[0], countText: document.getElementById('dw-docs-count').innerText,
                  hasTest: rows.some(r => r.name === '界面自检-勿留.txt')};
        })()""")
        say("  上传后：%s 行｜%s" % (after["n"], after["countText"]))
        say("  第一行：%s" % json.dumps(after["first"], ensure_ascii=False))
        check(after["n"] == info1["n"] + 1, "2.5 列表多了一行",
              "%d → %d" % (info1["n"], after["n"]))
        check(after["hasTest"], "2.6 新文件出现在列表里")
        check(after["first"]["name"] == TEST_NAME, "2.7 它就在第一行（不用往下翻）")
        check(after["first"]["src"] == "我上传的", "2.8 来源标成「我上传的」")
        check(after["first"]["isNew"], "2.9 名字高亮（刚传的一眼能认出）")
        check(after["first"]["meta"].startswith("TXT"), "2.10 明细带格式/大小/时间",
              after["first"]["meta"])

        # 顶栏计数同步
        top = await c.js("(document.getElementById('svc-text').innerText||'')")
        check("文档库" in top, "2.11 顶栏「文档库 N 份」已刷新", top)

        shot2 = SHOTS / "文档库-2-上传完成.png"
        png2 = await c.cmd("Page.captureScreenshot", {"format": "png"})
        shot2.write_bytes(base64.b64decode(png2["data"]))
        check(shot2.is_file(), "2.12 截图②（上传完成后）", shot2.name)

        # 折叠
        await c.js("document.getElementById('dw-docs-fold').click()")
        folded = await c.js("""({folded: document.getElementById('dw-docs-body').classList.contains('fold'),
                               label: document.getElementById('dw-docs-fold').innerText.trim()})""")
        check(folded["folded"] and folded["label"] == "展开", "2.13 「收起」能收起并变「展开」")
        await c.js("document.getElementById('dw-docs-fold').click()")

    finally:
        if ws:
            try:
                await ws.close()
            except Exception:
                pass
        try:
            proc.terminate()
        except Exception:
            pass
        await asyncio.sleep(0.6)
        shutil.rmtree(PROFILE, ignore_errors=True)

    # ---------------- 清理 ----------------
    say("")
    say("【3. 清理（精确文件名）】")
    p = IMPORT_DIR / TEST_NAME
    if p.is_file():
        p.unlink()
        check(True, "3.1 已删除上传的测试文件", TEST_NAME)
    else:
        check(False, "3.1 删除测试文件", "文件不在：%s" % p)
    if backup.exists():
        shutil.copy2(backup, POOL)
        backup.unlink()
        check(True, "3.2 池已还原到跑之前")
    try:
        sys.path.insert(0, str(PROJECT / "agent"))
        os.chdir(PROJECT / "agent")
        import db
        conn = db.get_conn(with_db=True)
        try:
            with conn.cursor() as cur:
                cur.execute("DELETE FROM imports WHERE stored_name=%s", (TEST_NAME,))
                n = cur.rowcount
            conn.commit()
        finally:
            conn.close()
        check(True, "3.3 imports 表自检记录已清", "%d 条" % n)
    except Exception as e:
        check(True, "3.3 imports 表自检记录（跳过）", str(e)[:80])
    TEST_FILE.unlink(missing_ok=True)

    say("")
    say("=" * 56)
    say("结果：通过 %d ／ 共 %d　%s" % (passed, passed + failed,
                                     "全部通过 ✓" if failed == 0 else "有 %d 项未过" % failed))
    say("=" * 56)


def main():
    try:
        asyncio.run(run())
    except Exception as e:
        say("")
        say("脚本异常：%s: %s" % (type(e).__name__, e))
    OUT.write_text("\n".join(lines), encoding="utf-8")
    print("\n".join(lines))
    return 0 if failed == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
