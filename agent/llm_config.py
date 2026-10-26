# -*- coding: utf-8 -*-
"""
llm_config.py —— 全项目唯一的模型配置 + 调用入口

2026-10-25：从 DeepSeek 官方 API 换成「阿里云百炼 MaaS」的 OpenAI 兼容接口。
之前 parse_intent / extract / extract_ops / support / llm_test 五个文件各写一份
地址和密钥，换平台得挨个改；现在统一从这里取，换模型只改本文件。

【第一步：填密钥】三种方式，优先级从高到低
  ① 环境变量 DASHSCOPE_API_KEY
  ② 文件     agent/llm_key.txt（一行，只放 key；已在 .gitignore 里，不会上传）
  ③ 本文件   下面的 API_KEY

【第二步：自检】填完跑一句，能通就会打印"可用"：
  python agent/llm_config.py
"""
import os
import sys
import threading
import time
from pathlib import Path

import requests

# 输出兜底：终端表示不了的字符就替换，不让程序崩
try:
    sys.stdout.reconfigure(errors="replace")
except Exception:
    pass

# ================== 配置区（只改这里） ==================
# 百炼的 OpenAI 兼容地址（华北2-北京）。换区域就把 cn-beijing 换掉：
#   北京   https://cn-beijing.maas.qianwenaiapi.com/compatible-mode/v1
#   通用   https://maas.qianwenaiapi.com/compatible-mode/v1
# 临时试别的模型不用改代码：设环境变量 A23_MODEL / A23_LLM_BASE_URL 即可。
BASE_URL = (os.environ.get("A23_LLM_BASE_URL")
            or "https://cn-beijing.maas.qianwenaiapi.com/compatible-mode/v1")
MODEL = os.environ.get("A23_MODEL") or "deepseek-v4.1-flash"
# 若报 404「模型不存在」：去百炼控制台 → 模型广场，复制准确的模型名填到上面。

# 本文件要进 git，所以这里留空。key 请写到 agent/llm_key.txt（一行，已 gitignore），
# 或设环境变量 DASHSCOPE_API_KEY。两者的读取顺序见下方 _key()。
API_KEY = ""

# 关思考（速度）：这个模型默认"边想边答"，解析一句简单指令也要先想几十~两百个
# reasoning token。2026-10-25 用模块1 真实 prompt 实测（6906 字，各 2 次）：
#   不传            4.20s  reasoning=81/190
#   enable_thinking 1.81s  不产生 reasoning  ← 采用
#   thinking 写法   2.04s
#   reasoning_effort 2.34s
# 可选值："enable_thinking"（默认）｜"thinking"｜"none"（不传，最兼容）
# 如果对方以后不认这个字段，chat() 会自动去掉参数重试一次，不会因此失败。
THINKING_OFF_STYLE = "enable_thinking"
# ========================================================

API_URL = BASE_URL.rstrip("/") + "/chat/completions"
KEY_FILE = Path(__file__).resolve().parent / "llm_key.txt"

_HINT = ("还没填密钥：打开 agent/llm_config.py 把 API_KEY 填上，"
         "或新建 agent/llm_key.txt 把 key 写进去（一行）")

_thread_local = threading.local()


def _session():
    """按线程各存一个 Session：复用连接省掉每次的 DNS+TLS 握手（约 0.2~0.4s），
    又不跨线程共用一个连接池（requests 的 Session 不建议多线程共用）。"""
    s = getattr(_thread_local, "s", None)
    if s is None:
        s = requests.Session()
        _thread_local.s = s
    return s


def api_key() -> str:
    """按 环境变量 → llm_key.txt → 本文件常量 的顺序取 key。"""
    k = (os.environ.get("DASHSCOPE_API_KEY") or "").strip()
    if not k and KEY_FILE.exists():
        try:
            lines = KEY_FILE.read_text(encoding="utf-8").splitlines()
            k = next((ln.strip() for ln in lines if ln.strip()), "")
        except Exception:
            k = ""
    return k or (API_KEY or "").strip()


def has_key() -> bool:
    return bool(api_key())


def describe() -> str:
    return "%s @ %s" % (MODEL, BASE_URL)


def _friendly(status: int, text: str) -> str:
    """把 HTTP 状态码翻成人话（这些码是新手最常撞的）。"""
    tail = (text or "").strip().replace("\n", " ")[:200]
    m = {
        400: "请求参数不对（400）",
        401: "密钥不对或没带上（401）——确认填的是百炼的 key，不是 DeepSeek 官方的",
        402: "额度/余额不足（402）",
        403: "没有权限或欠费（403）——去百炼控制台看这个模型是否已开通",
        404: "地址或模型名不对（404）——检查 BASE_URL 和 MODEL",
        429: "请求太频繁（429），稍等再试",
    }
    if status in m:
        return m[status] + ("：" + tail if tail else "")
    if 500 <= status < 600:
        return "服务端出错（%d），稍后重试：%s" % (status, tail)
    return "模型返回 %d：%s" % (status, tail)


_warned = False


def chat(messages, *, json_mode=False, temperature=None, max_tokens=None,
         timeout=90, session=None, retries=1):
    """调一次模型。

    messages  —— OpenAI 格式：[{"role": "system"/"user", "content": "..."}]
    json_mode —— True 时带 response_format=json_object（要模型只回 JSON 时用）
    返回 (text, usage, cost, err)：
        成功 → (文本, usage 字典, 耗时秒, None)
        失败 → (None, None, 耗时秒, 一句人话错误)
    """
    global _warned
    key = api_key()
    if not key:
        return None, None, 0.0, _HINT

    headers = {"Authorization": "Bearer " + key, "Content-Type": "application/json"}
    body = {"model": MODEL, "messages": messages}
    opts = {}
    if json_mode:
        opts["response_format"] = {"type": "json_object"}
    if temperature is not None:
        opts["temperature"] = temperature
    if max_tokens:
        opts["max_tokens"] = int(max_tokens)
    if THINKING_OFF_STYLE == "enable_thinking":
        opts["enable_thinking"] = False
    elif THINKING_OFF_STYLE == "thinking":
        opts["thinking"] = {"type": "disabled"}

    sess = session if session is not None else _session()
    t0 = time.time()
    last_err = ""
    degrade = False
    # 第一轮带全部可选参数；遇到 400/422 说明对方不认其中某个，
    # 退成"只有 model + messages"再试一次（JSON 靠 prompt + 宽容解析兜底）。
    # 注意：网络异常重试不算降级，参数照旧带上。
    for attempt in (0, 1):
        payload = dict(body)
        if not degrade:
            payload.update(opts)
        try:
            resp = sess.post(API_URL, headers=headers, json=payload, timeout=timeout)
        except requests.exceptions.Timeout:
            last_err = "请求超时（%d 秒）——网络慢或模型排队，重试一次" % timeout
            if attempt < retries:
                continue
            return None, None, time.time() - t0, last_err
        except Exception as e:
            last_err = "连不上模型服务：%s" % str(e)[:120]
            if attempt < retries:
                continue
            return None, None, time.time() - t0, last_err

        cost = time.time() - t0
        if resp.status_code == 200:
            try:
                data = resp.json()
                text = data["choices"][0]["message"]["content"]
            except Exception:
                return None, None, cost, "返回的内容不是预期的 JSON，可能是网关页面"
            return text, data.get("usage", {}), cost, None

        last_err = _friendly(resp.status_code, resp.text)
        if resp.status_code in (400, 422) and opts and not degrade:
            degrade = True
            if not _warned:
                print("【模型】服务端不认部分可选参数，已自动降级重试（只发 model + messages）")
                _warned = True
            continue
        return None, None, cost, last_err

    return None, None, time.time() - t0, last_err


if __name__ == "__main__":
    print("当前模型：", describe())
    print("密钥：", "已读到（%s…%s）" % (api_key()[:6], api_key()[-4:]) if has_key() else "还没填")
    if not has_key():
        print(_HINT)
        raise SystemExit(1)
    text, usage, cost, err = chat(
        [{"role": "user", "content": "只回两个字：可用"}], timeout=40)
    if err:
        print("调用失败：", err)
        raise SystemExit(1)
    print("回答：", (text or "").strip())
    print("耗时：%.2f 秒" % cost)
    print("用量：", usage)
