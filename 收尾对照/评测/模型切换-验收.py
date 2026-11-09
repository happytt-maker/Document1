# -*- coding: utf-8 -*-
"""
模型多路切换 —— 验收脚本（2026-11-02）

不需要任何真实 API key：本脚本起两个假模型服务，
一个专门模拟"免费额度用尽"，一个正常回答，
然后用项目真实的 llm_config.chat() 走一遍，验证：

  1. 403 额度用尽        → 自动换下一家，最终成功
  2. 401 key 错          → 自动换下一家
  3. 429 限流            → 自动换下一家
  4. 404 模型名不存在    → 自动换下一家
  5. 400 参数错          → 换家没用，直接报错（不该白白多打一家）
  6. 失败过的进冷却       → 第二次调用直接走能用的那家，不再撞一次
  7. 全部不可用          → 返回一句说得清的错误，不是崩掉
  8. 本机模型不走系统代理 → 开代理也能连 127.0.0.1

跑法：python 收尾对照/评测/模型切换-验收.py
"""
import json
import os
import sys
import threading
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / "agent"))

import llm_config as L  # noqa: E402

HITS = []          # 每次请求打到哪个假服务，用来验证"真的换了家"
MODE = {"bad": 403}  # 第一家的失败方式


class Handler(BaseHTTPRequestHandler):
    def do_POST(self):
        n = int(self.headers.get("Content-Length") or 0)
        raw = self.rfile.read(n) or b"{}"
        try:
            body = json.loads(raw)
        except Exception:
            body = {}
        model = body.get("model", "")

        if model == "bad-model":
            HITS.append("bad")
            code = MODE["bad"]
            msg = {
                403: "Free quota exhausted",
                401: "Incorrect API key provided",
                429: "Rate limit reached",
                404: "model not found",
                400: "invalid parameter: foo",
            }.get(code, "error")
            self._json(code, {"error": {"message": msg, "type": "mock"}})
            return

        HITS.append("good")
        self._json(200, {
            "choices": [{"message": {"content": "可用"}}],
            "usage": {"prompt_tokens": 5, "completion_tokens": 2, "total_tokens": 7},
        })

    def _json(self, code, obj):
        data = json.dumps(obj, ensure_ascii=False).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def log_message(self, *a):
        pass


def serve(port):
    srv = ThreadingHTTPServer(("127.0.0.1", port), Handler)
    threading.Thread(target=srv.serve_forever, daemon=True).start()
    return srv


def install(bad_port, good_port, order=("bad", "good")):
    """临时把两家假服务塞进配置（不动真实的 llm_key.txt）。"""
    L.PROVIDERS = {
        "bad": {
            "name": "假服务-额度用尽", "base": "http://127.0.0.1:%d/v1" % bad_port,
            "model": "bad-model", "key_env": [], "key": "fake", "thinking": None,
            "concurrency": 4,
        },
        "good": {
            "name": "假服务-正常", "base": "http://127.0.0.1:%d/v1" % good_port,
            "model": "good-model", "key_env": [], "key": "fake", "thinking": None,
            "concurrency": 4,
        },
    }
    L.DEFAULT_ORDER = list(order)
    L._LAST_OK["id"] = None
    with L._COOL_LOCK:
        L._COOL.clear()


def ask():
    return L.chat([{"role": "user", "content": "只回两个字：可用"}], timeout=8)


def main():
    L_root = []
    bad_port, good_port = 18081, 18082
    serve(bad_port)
    serve(good_port)
    msgs = [{"role": "user", "content": "只回两个字：可用"}]
    passed = 0
    total = 0

    def check(name, cond, detail=""):
        nonlocal passed, total
        total += 1
        if cond:
            passed += 1
        L_root.append("  %s %s%s" % ("✔" if cond else "✘", name,
                                     ("   → " + detail) if detail else ""))

    # 1~4：不同失败码都应该自动换家
    for code, label in [(403, "额度用尽"), (401, "key 错"),
                        (429, "限流"), (404, "模型名不存在")]:
        MODE["bad"] = code
        install(bad_port, good_port)
        HITS.clear()
        text, usage, cost, err = ask()
        check("%d %s → 自动换家并成功" % (code, label),
              err is None and (text or "").strip() == "可用",
              "err=%s" % err if err else "走 %s，用量 provider=%s"
              % ("→".join(HITS), usage.get("provider")))

    # 5：400 是请求本身的问题，换家没用，应该直接报错（不该打第二家）
    MODE["bad"] = 400
    install(bad_port, good_port)
    HITS.clear()
    text, usage, cost, err = ask()
    check("400 参数错 → 不瞎换家，直接报错", err is not None and HITS.count("good") == 0,
          "打到 %s（应只有 bad）" % (HITS or "无"))

    # 6：冷却。第一次 403 之后，第二次应该直接走 good，不再撞 bad
    MODE["bad"] = 403
    install(bad_port, good_port)
    HITS.clear()
    ask()
    first = list(HITS)
    HITS.clear()
    ask()
    second = list(HITS)
    check("失败过进冷却 → 第二次直接走能用的那家",
          first == ["bad", "good"] and second == ["good"],
          "第一次 %s，第二次 %s" % (first, second))

    # 7：两家都是"额度用尽"（fatal）→ 应该汇总成一句人话
    #    注意不用"连不上端口"来模拟：网络错属于"换家也没用"，代码会直接返回
    #    那一家的具体错误，这是有意的（比笼统的"都不可用"信息量大）。
    MODE["bad"] = 403
    install(bad_port, good_port)
    L.PROVIDERS["good"]["base"] = "http://127.0.0.1:%d/v1" % bad_port
    L.PROVIDERS["good"]["model"] = "bad-model"      # 让它也 403
    L.DEFAULT_ORDER = ["bad", "good"]
    with L._COOL_LOCK:
        L._COOL.clear()
    L._LAST_OK["id"] = None
    text, usage, cost, err = ask()
    check("两家都没额度 → 汇总成人话错误",
          err is not None and "所有模型都不可用" in (err or ""),
          (err or "")[:70])

    # 8：本机模型绕开系统代理（模拟开着代理的情况）
    os.environ["HTTP_PROXY"] = "http://127.0.0.1:19999"    # 一个必然不通的代理
    os.environ["HTTPS_PROXY"] = "http://127.0.0.1:19999"
    os.environ["NO_PROXY"] = ""
    install(bad_port, good_port)
    L.PROVIDERS = {"good": {
        "name": "假服务-正常", "base": "http://127.0.0.1:%d/v1" % good_port,
        "model": "good-model", "key_env": [], "key": "fake",
        "thinking": None, "concurrency": 4, "local": True,
    }}
    L.DEFAULT_ORDER = ["good"]
    with L._COOL_LOCK:
        L._COOL.clear()
    # 走 local 分支（trust_env=False）应该照样通
    text, usage, cost, err = ask()
    check("本地模型绕开系统代理", err is None, err or "代理开着也能连 127.0.0.1")

    print("\n".join(L_root))
    print()
    print("  —— 通过 %d/%d" % (passed, total))
    return 0 if passed == total else 1


if __name__ == "__main__":
    sys.exit(main())
