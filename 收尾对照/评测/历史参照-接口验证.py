# -*- coding: utf-8 -*-
"""
历史参照-接口验证.py —— 验"操作台能不能看到历史填写参照"

为什么单独一个脚本：
  历史参照的存储和复用都在 matcher 侧（history.py），但用户能看到的入口在
  操作台（agent/api_server.py 的 /api/fill/history）。命令行跑通不代表网页能用，
  所以这里真的把服务起起来、真的发 HTTP 请求，看到 JSON 才算过。

三个坑都在这脚本里处理了：
  1. 系统代理会拦 127.0.0.1 的 GET（返回假 404，POST 却是通的）→ ProxyHandler({})
  2. 服务启动要几秒 → 轮询端口，不瞎 sleep 固定值
  3. 测完要把自己起的服务关掉，不留僵尸进程

跑法：python 收尾对照/评测/历史参照-接口验证.py
"""
import json
import os
import socket
import subprocess
import sys
import time
import urllib.parse
import urllib.request
from pathlib import Path

HERE = Path(__file__).resolve().parent
BASE = HERE.parents[1]
AGENT = BASE / "agent"
OUT = HERE / "历史参照-接口验证.txt"
PORT = 8011          # 用 8011 而不是 8000：不打扰用户正在用的那个操作台

TPL = "模板5-原材料库存汇总表.xlsx"


def wait_up(sec=40):
    t0 = time.time()
    while time.time() - t0 < sec:
        s = socket.socket()
        s.settimeout(0.4)
        try:
            s.connect(("127.0.0.1", PORT))
            s.close()
            return time.time() - t0
        except Exception:
            time.sleep(0.4)
    return None


def main():
    lines = []
    def w(s=""):
        lines.append(str(s))
        print(s)

    env = os.environ.copy()
    env["PYTHONIOENCODING"] = "utf-8"
    proc = subprocess.Popen(
        [sys.executable, "api_server.py", "--port", str(PORT)],
        cwd=str(AGENT), env=env,
        stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    try:
        up = wait_up()
        w("服务启动：%s" % ("%.1f 秒" % up if up is not None else "失败（端口没起来）"))
        if up is None:
            w("→ 后面的接口测试跳过")
            return

        # 绕开系统代理：本机实测，代理会把 127.0.0.1 的 GET 劫持成假 404
        opener = urllib.request.build_opener(urllib.request.ProxyHandler({}))

        cases = [
            ("历史列表", "/api/fill/history"),
            ("单模板查询", "/api/fill/history?template=" + urllib.parse.quote(TPL)),
        ]
        n_ok = 0
        for name, path in cases:
            w("")
            w("【%s】GET %s" % (name, path))
            try:
                r = opener.open("http://127.0.0.1:%d%s" % (PORT, path), timeout=30)
                d = json.loads(r.read().decode("utf-8"))
                n_ok += 1
                if name == "历史列表":
                    w("  HTTP %d ｜ 历史记录 %d 条" % (r.status, d.get("total", 0)))
                    for it in (d.get("runs") or [])[:5]:
                        w("    %s ｜ %s ｜ %s 列 / %s 行 ｜ 填过 %s 次"
                          % (it.get("ts"), it.get("template"), it.get("n_cols"),
                             it.get("n_rows"), it.get("times")))
                else:
                    # 把"服务端收到的模板名"和"解析出的列"一起打出来：
                    # 查不到时能立刻分清是名字没传对，还是结构文件里没这张表
                    w("  收到 template=%r ｜ 解析出列 %d 个"
                      % (d.get("template"), len(d.get("columns") or [])))
                    if d.get("error"):
                        w("  !! 读模板结构失败：%s" % d["error"])
                    w("  HTTP %d ｜ 命中方式：%s（%s）" % (r.status, d.get("hit"), d.get("from")))
                    for ln in (d.get("lines") or []):
                        w("    " + ln)
                    for c, v in (d.get("cols") or {}).items():
                        w("    %-10s ← %s ｜ %s ｜ %s"
                          % (c, v.get("head") or "(无列头)", v.get("file"), v.get("kind")))
            except Exception as e:
                w("  失败：%s: %s" % (e.__class__.__name__, e))

        w("")
        w("=" * 58)
        w("接口验证：%d/%d 通过" % (n_ok, len(cases)))
        w("=" * 58)
    finally:
        proc.terminate()
        try:
            proc.wait(timeout=8)
        except Exception:
            proc.kill()
        OUT.write_text("\n".join(lines) + "\n", encoding="utf-8")
        print("（报告已存：%s）" % OUT.name)


if __name__ == "__main__":
    main()
