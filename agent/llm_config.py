# -*- coding: utf-8 -*-
"""
llm_config.py —— 全项目唯一的模型配置 + 调用入口

2026-10-25：从 DeepSeek 官方 API 换成「阿里云百炼 MaaS」的 OpenAI 兼容接口。
2026-11-02：改成一个平台变成「一串平台」—— 免费额度用完了自动换下一家，
            不再因为某一家没额度就整个系统停摆。

【为什么改成多家】
免费额度是会用完的。以前的写法是 base_url 和模型名写死一个，
用完了只能改代码；现在配了哪几家就按顺序试，第一家 403 就跳第二家。

【三种填 key 的方式】优先级从高到低
  ① 环境变量：ZHIPU_API_KEY / SILICONFLOW_API_KEY / DASHSCOPE_API_KEY / A23_API_KEY
  ② 文件 agent/llm_key.txt（一行一家，已 gitignore，不会被上传）：
         zhipu=xxxxxxxx
         siliconflow=sk-xxxxxxxx
     只有一行、没有 "=" 的老格式，会当成百炼（dashscope）的 key —— 兼容旧文件。
  ③ 本文件下面的 API_KEY 常量（不建议，因为本文件要进 git）

【自检】能通就会打印"可用"
  python agent/llm_config.py            # 只测首选那家
  python agent/llm_config.py --test     # 把配了 key 的全部测一遍，告诉你哪家能用
  python agent/llm_config.py --list     # 只看有哪些可选，不发请求

【切换】默认顺序见 DEFAULT_ORDER；想指定首选就设环境变量
  A23_PROVIDER=zhipu      （Windows: set A23_PROVIDER=zhipu）
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

_HERE = Path(__file__).resolve().parent
KEY_FILE = _HERE / "llm_key.txt"
PROVIDERS_FILE = _HERE / "llm_providers.json"   # 可选：自己加平台，不用改本文件

# ================== 内置平台预设（只改这里） ==================
# 字段说明：
#   base        OpenAI 兼容地址（不带 /chat/completions）
#   model       模型名（各家不一样，写错会 404，然后自动跳下一家）
#   key_env     从哪个环境变量读 key（按顺序找，找到就用）
#   thinking    怎么关"边想边答"。None=不发；("字段名", 值) 形式最通用
#   free        True=有免费额度。只影响提示文案，不影响逻辑
#   local       True=本机模型，不需要 key，永远不冷却
#   concurrency 并发上限。免费档常常是 1，超过会 429，用信号量排队而不是报错
#   ctx         上下文窗口（token），只用于提示
PROVIDERS = {
    "zhipu": {
        "name": "智谱 GLM-5.3-Flash（新用户赠 1500 万 token）",
        "base": "https://open.bigmodel.cn/api/paas/v4",
        "model": "glm-5.3-flash",
        "key_env": ["ZHIPU_API_KEY", "A23_API_KEY"],
        # 这个模型**始终思考**，不支持 {"thinking":{"type":"disabled"}}（会 400），
        # 只能调档：reasoning_effort = low / high / max。
        # 实测 low：2.1s、46 token；不传：3.4s、92 token —— 抽取任务用 low 足够。
        # 若换回 glm-4.7-flash，要改成 ("thinking", {"type": "disabled"})，它不认 reasoning_effort。
        "thinking": ("reasoning_effort", "low"),
        "free": False, "concurrency": 1, "ctx": 200000,
        "note": "新用户赠 1500 万 token；始终思考、可用 reasoning_effort 降档；国内直连；单并发，高峰会 429",
        "signup": "https://open.bigmodel.cn/usercenter/apikeys",
    },
    "siliconflow": {
        "name": "硅基流动 SiliconFlow（部分模型永久免费）",
        "base": "https://api.siliconflow.cn/v1",
        "model": "Qwen/Qwen3-8B-Instruct",
        "key_env": ["SILICONFLOW_API_KEY"],
        "thinking": None,
        "free": True, "concurrency": 4, "ctx": 131072,
        "note": "注册送额度 + 部分模型永久免费；模型名可换，见控制台",
        "signup": "https://cloud.siliconflow.cn",
    },
    "dashscope": {
        "name": "阿里云百炼 MaaS（现用，额度已用尽）",
        "base": "https://cn-beijing.maas.qianwenaiapi.com/compatible-mode/v1",
        "model": "deepseek-v4.1-flash",
        "key_env": ["DASHSCOPE_API_KEY"],
        # 这一家用 enable_thinking 关思考，2026-10-25 实测比不传快 2.3 倍
        "thinking": ("enable_thinking", False),
        "free": False, "concurrency": 8, "ctx": 128000,
        "note": "原主力；免费额度用尽后仍可当付费兜底",
        "signup": "https://bailian.console.aliyun.com",
    },
    "ollama": {
        "name": "本机 Ollama（完全免费、无限量、断网也能用）",
        "base": "http://127.0.0.1:11434/v1",
        "model": "qwen2.5:7b-instruct",
        "key_env": [],
        "key": "ollama",          # 本地服务不校验，但不能空着
        "thinking": None,
        "free": True, "local": True, "concurrency": 1, "ctx": 32768,
        "note": "要先装 Ollama 并 ollama pull 模型；准确率不如云端，适合跑批量回归",
        "signup": "https://ollama.com/download",
    },
}

# 没指定 A23_PROVIDER 时的尝试顺序。本地排最后：慢但永远不会没额度。
DEFAULT_ORDER = ["zhipu", "siliconflow", "dashscope", "ollama"]

# 发现某家不可用时的冷却时长（秒）。冷却期内不重试，避免每次都撞墙。
COOL_DOWN = {401: 1800, 404: 1800, 402: 600, 403: 600, 429: 30}

# 被限流（429）时在同一家等待重试的间隔（秒）。用完这几档还不行才换家。
# 免费档普遍是"每分钟 N 次"，隔几秒就能缓过来；直接换家会把可用的一家误判成不可用。
RATE_WAIT = (2.0, 6.0, 15.0)

# 本文件要进 git，所以这里留空。key 请写到 agent/llm_key.txt（已 gitignore），
# 或设上面的环境变量。
API_KEY = ""
# ========================================================

_HINT = ("还没填密钥：新建 agent/llm_key.txt，写一行如  zhipu=你的key\n"
         "          或设环境变量 ZHIPU_API_KEY。免费 key 申请：%s"
         % PROVIDERS["zhipu"]["signup"])

_thread_local = threading.local()

# 每家一个信号量，把超过并发上限的请求改成排队（免费档并发=1 时全靠它）
_SEMS = {}
_SEMS_LOCK = threading.Lock()
# 冷却表：provider -> 冷却截止时间（time.monotonic）
_COOL = {}
_COOL_LOCK = threading.Lock()
# 上一次真正用上的那家，优先复用，避免每次都从第一家开始试
_LAST_OK = {"id": None}
_warned_degrade = False
_warned_429 = False


# ---------------- 配置装配 ----------------

def _load_keys():
    """解析 llm_key.txt → {provider: key}。

    支持：zhipu=xxx / zhipu: xxx / zhipu xxx；# 开头是注释。
    没有 "=" 的单行老格式，归给 dashscope（兼容 10-25 之前写的文件）。
    """
    out = {}
    if not KEY_FILE.exists():
        return out
    try:
        lines = KEY_FILE.read_text(encoding="utf-8").splitlines()
    except Exception:
        return out
    for raw in lines:
        ln = raw.strip()
        if not ln or ln.startswith("#"):
            continue
        pid, k = None, ln
        for sep in ("=", ":", " "):
            if sep in ln:
                left, right = ln.split(sep, 1)
                cand = left.strip().lower()
                if cand in PROVIDERS or cand.isascii():
                    pid, k = cand, right.strip()
                    break
        # 一行只有 key（旧格式）：当前默认就是百炼
        out[pid or "dashscope"] = k
    return out


def _providers():
    """内置预设 + 可选的外部 JSON 覆盖（agent/llm_providers.json）。"""
    cfg = {k: dict(v) for k, v in PROVIDERS.items()}
    if PROVIDERS_FILE.exists():
        try:
            import json
            extra = json.loads(PROVIDERS_FILE.read_text(encoding="utf-8"))
            for pid, v in (extra or {}).items():
                cfg.setdefault(pid, {}).update(v)
        except Exception as e:
            print("【模型】llm_providers.json 读不了，已忽略：%s" % str(e)[:80])
    # 环境变量兜底覆盖首选那家（临时试别的模型不用改文件）
    if os.environ.get("A23_LLM_BASE_URL"):
        pid = os.environ.get("A23_PROVIDER") or "dashscope"
        cfg.setdefault(pid, {})["base"] = os.environ["A23_LLM_BASE_URL"]
    if os.environ.get("A23_MODEL"):
        pid = os.environ.get("A23_PROVIDER") or "dashscope"
        cfg.setdefault(pid, {})["model"] = os.environ["A23_MODEL"]
    return cfg


def _key_for(pid, p):
    """按 环境变量 → llm_key.txt → 内置常量 的顺序取 key。"""
    for env in (p.get("key_env") or []):
        v = (os.environ.get(env) or "").strip()
        if v:
            return v
    k = (_load_keys().get(pid) or "").strip()
    if k:
        return k
    if p.get("key"):
        return p["key"]
    return (API_KEY or "").strip() if pid == "dashscope" else ""


def _usable(pid, p):
    """这家能不能用：本地模型不看 key，但要看服务起没起；其余必须有 key。"""
    if p.get("local"):
        return _local_up(p)
    return bool(_key_for(pid, p))


_LOCAL_CACHE = {}
_LOCAL_TTL = 60.0


def _local_up(p):
    """本机服务探活（127.0.0.1 上连一下端口，毫秒级）。结果缓存 60 秒，
    不然没装 Ollama 时每次调用都要先撞一次连接超时。"""
    import socket
    import urllib.parse
    try:
        u = urllib.parse.urlparse(str(p["base"]))
        host, port = u.hostname or "127.0.0.1", u.port or 80
    except Exception:
        return False
    ck = "%s:%s" % (host, port)
    hit = _LOCAL_CACHE.get(ck)
    if hit and time.monotonic() - hit[0] < _LOCAL_TTL:
        return hit[1]
    ok = False
    try:
        with socket.create_connection((host, port), timeout=0.4):
            ok = True
    except Exception:
        ok = False
    _LOCAL_CACHE[ck] = (time.monotonic(), ok)
    return ok


def order():
    """当前实际会尝试的顺序：首选 → 上次成功过的 → 默认顺序 → 其它配了 key 的。"""
    prefer = (os.environ.get("A23_PROVIDER") or "").strip().lower()
    cfg = _providers()
    seq = []
    for pid in ([prefer] if prefer else []) + [_LAST_OK["id"]] + DEFAULT_ORDER + list(cfg):
        if pid and pid in cfg and pid not in seq and _usable(pid, cfg[pid]):
            seq.append(pid)
    return seq


def _sem(pid, p):
    n = max(1, int(p.get("concurrency") or 1))
    with _SEMS_LOCK:
        s = _SEMS.get(pid)
        if s is None:
            s = _SEMS[pid] = threading.Semaphore(n)
        return s


def _is_cooling(pid):
    with _COOL_LOCK:
        until = _COOL.get(pid)
        return until is not None and time.monotonic() < until


def _mark_cool(pid, status, p):
    if p.get("local"):
        return
    sec = COOL_DOWN.get(status)
    if not sec:
        return
    with _COOL_LOCK:
        _COOL[pid] = time.monotonic() + sec


def _clear_cool(pid):
    with _COOL_LOCK:
        _COOL.pop(pid, None)


def _session():
    """按线程各存一个 Session：复用连接省掉每次的 DNS+TLS 握手（约 0.2~0.4s），
    又不跨线程共用一个连接池（requests 的 Session 不建议多线程共用）。"""
    s = getattr(_thread_local, "s", None)
    if s is None:
        s = requests.Session()
        _thread_local.s = s
    return s


def _local_session():
    """本机服务专用的 Session：强制不走系统代理。

    2026-11-02：这项目以前踩过这个坑 —— 开着系统代理时，连 127.0.0.1 的请求
    会被代理劫持，POST 看似通、GET 却回假 404。本地模型必须绕开。
    """
    s = getattr(_thread_local, "sl", None)
    if s is None:
        s = requests.Session()
        s.trust_env = False
        s.proxies = {}
        _thread_local.sl = s
    return s


# ---------------- 对外信息 ----------------

def api_key() -> str:
    """首选那家的 key（给只想拿一个 key 的旧代码用）。"""
    cfg = _providers()
    seq = order()
    return _key_for(seq[0], cfg[seq[0]]) if seq else ""


def has_key() -> bool:
    return bool(order())


def describe() -> str:
    seq = order()
    if not seq:
        return "没有可用模型（一家都没配 key）"
    cfg = _providers()
    p = cfg[seq[0]]
    return "%s @ %s" % (p.get("model"), p.get("base"))


def status():
    """给界面/接口看的分家状态。"""
    cfg = _providers()
    rows = []
    for pid in order() + [x for x in cfg if x not in order()]:
        p = cfg[pid]
        rows.append({
            "id": pid,
            "name": p.get("name") or pid,
            "model": p.get("model"),
            "base": p.get("base"),
            "ready": _usable(pid, p),
            "local": bool(p.get("local")),
            "free": bool(p.get("free")),
            "cooling": _is_cooling(pid),
            "last_ok": _LAST_OK["id"] == pid,
            "note": p.get("note", ""),
            "signup": p.get("signup", ""),
        })
    seq = order()
    return {"primary": seq[0] if seq else None, "order": seq, "providers": rows}


# ---------------- 调用 ----------------

def _friendly(status: int, text: str, pid="") -> str:
    """把 HTTP 状态码翻成人话（这些码是新手最常撞的）。"""
    tail = (text or "").strip().replace("\n", " ")[:160]
    m = {
        400: "请求参数不对（400）",
        401: "密钥不对或没带上（401）",
        402: "额度/余额不足（402）",
        403: "没有权限或额度用尽（403）——去控制台看该模型是否已开通/是否还有免费额度",
        404: "地址或模型名不对（404）——检查 base 和 model",
        429: "请求太频繁（429）——免费档常见，稍等再试",
    }
    head = ("%s：" % pid if pid else "") + m.get(status, "模型返回 %d" % status)
    return head + ("：%s" % tail if tail else "")


def _one(pid, p, messages, json_mode, temperature, max_tokens, timeout, session):
    """向一家发请求。返回 (text, usage, cost, err, fatal)。

    并发上限靠调用方的信号量排队（见 _sem），这里不再管。
    fatal=True 表示"这家不能用了"（没额度/没权限/key 错/模型名错/限流），
    应该换下一家；fatal=False 表示"这次请求本身有问题"，换家也没用。
    """
    key = _key_for(pid, p)
    url = str(p["base"]).rstrip("/") + "/chat/completions"
    headers = {"Authorization": "Bearer " + key, "Content-Type": "application/json"}
    body = {"model": p["model"], "messages": messages}
    opts = {}
    if json_mode:
        opts["response_format"] = {"type": "json_object"}
    if temperature is not None:
        opts["temperature"] = temperature
    if max_tokens:
        opts["max_tokens"] = int(max_tokens)
    th = p.get("thinking")
    if th:
        opts[th[0]] = th[1]

    sess = session if session is not None else (_local_session() if p.get("local") else _session())
    t0 = time.time()
    last_err, degrade = "", False
    rl = 0                      # 429 已等待重试次数
    conn = 0                    # 连接层重试次数（DNS 抖动/握手失败）
    global _warned_degrade, _warned_429
    while True:
        payload = dict(body)
        if not degrade:
            payload.update(opts)
        try:
            resp = sess.post(url, headers=headers, json=payload, timeout=timeout)
        except requests.exceptions.Timeout:
            return None, None, time.time() - t0, _friendly(408, "", pid), False
        except Exception as e:
            # 连接层错误多数是网络抖动（2026-11-02 实测碰到过一次：整列抽取因为
            # 一次握手失败就全废）。同家再试一次，别把偶发抖动当"这家挂了"。
            if conn < 1:
                conn += 1
                time.sleep(1.0)
                continue
            return None, None, time.time() - t0, "连不上 %s：%s" % (pid, str(e)[:110]), False

        cost = time.time() - t0
        if resp.status_code == 200:
            try:
                data = resp.json()
                text = data["choices"][0]["message"]["content"]
            except Exception:
                return None, None, cost, "%s 返回的不是预期 JSON，可能是网关页面" % pid, False
            return text, data.get("usage", {}), cost, None, False

        last_err = _friendly(resp.status_code, resp.text, pid)

        # 429 是"先歇会儿"，不是"这家废了"：免费档限流很常见，隔几秒就好了。
        # 如果只靠换家，一旦别家也欠费，整个系统就停摆 —— 而它本来只是被限流。
        if resp.status_code == 429 and rl < len(RATE_WAIT):
            wait = RATE_WAIT[rl]
            rl += 1
            if not _warned_429:
                print("【模型】%s 被限流，等 %.0fs 后重试（免费档常见，不是故障）" % (pid, wait))
                _warned_429 = True
            time.sleep(wait)
            continue

        # 400/422 说明对方不认某个可选参数，退成"只有 model + messages"再来一次
        # （JSON 靠 prompt + 宽容解析兜底）。
        if resp.status_code in (400, 422) and opts and not degrade:
            degrade = True
            if not _warned_degrade:
                print("【模型】%s 不认部分可选参数，已自动降级重试（只发 model + messages）" % pid)
                _warned_degrade = True
            continue

        # 这几类换一家就能救回来
        fatal = resp.status_code in (401, 402, 403, 404, 429)
        return None, None, cost, last_err, fatal


def chat(messages, *, json_mode=False, temperature=None, max_tokens=None,
         timeout=90, session=None, retries=1, provider=None, strict=False):
    """调一次模型。失败会自动换下一家（额度用尽不再让整个系统停摆）。

    messages  —— OpenAI 格式：[{"role": "system"/"user", "content": "..."}]
    json_mode —— True 时带 response_format=json_object（要模型只回 JSON 时用）
    provider  —— 指定只走这一家（不自动切换）。默认 None = 按 order() 试
    strict    —— True 时全部失败也不做"忽略冷却重试"
    返回 (text, usage, cost, err)：
        成功 → (文本, usage 字典, 耗时秒, None)
        失败 → (None, None, 耗时秒, 一句人话错误)
    usage 里额外带 "provider"，标明这次是哪家答的。
    """
    cfg = _providers()
    seq = [provider] if provider else order()
    if provider and provider not in cfg:
        return None, None, 0.0, "没有这个平台：%s（可选：%s）" % (provider, "、".join(cfg))
    if not seq:
        return None, None, 0.0, _HINT

    # 冷却期内的先跳过；但如果全在冷却，就忽略冷却全试一遍
    # （否则"等了 30 秒的 429 冷却"会让系统看起来彻底不可用）
    fresh = [x for x in seq if not _is_cooling(x)]
    plan = fresh or seq

    t_all = time.time()
    errs = []
    for pid in plan:
        p = cfg[pid]
        # 免费档并发上限常常是 1：这里排队等，比让服务端回 429 更省事
        with _sem(pid, p):
            text, usage, cost, err, fatal = _one(
                pid, p, messages, json_mode, temperature, max_tokens, timeout, session)
        if err is None:
            _clear_cool(pid)
            _LAST_OK["id"] = pid
            if isinstance(usage, dict):
                usage = dict(usage)
                usage["provider"] = pid
                usage["model"] = p["model"]
            return text, usage, cost, None

        errs.append(err)
        _mark_cool(pid, _status_of(err), p)
        if not fatal:
            return None, None, time.time() - t_all, err
        # fatal：这家救不回来，换下一家
        if len(plan) > 1:
            print("【模型】%s 不可用，自动换下一家（%s）" % (pid, err[:90]))

    return None, None, time.time() - t_all, "所有模型都不可用 → " + " ｜ ".join(errs[:3])


def _status_of(err: str) -> int:
    """从错误话里把状态码抠回来，用于决定冷却多久。"""
    for code in (401, 402, 403, 404, 429):
        if ("（%d）" % code) in err:
            return code
    return 0


# ---------------- 兼容旧代码的模块级常量 ----------------
# parse_intent / extract / support 还在 `from llm_config import API_KEY, API_URL, MODEL`
# （主要用于打印"这次用哪个模型"）。放在文件末尾是因为它们要调用上面定义的函数。
# 注意：只代表"当前首选那家"，真正发请求一律走 chat()，可能中途换到别家。
def _primary_cfg():
    cfg = _providers()
    seq = order()
    return cfg[seq[0]] if seq else {}


MODEL = _primary_cfg().get("model", "")
API_URL = (str(_primary_cfg().get("base", "")).rstrip("/") + "/chat/completions") \
    if _primary_cfg() else ""


# ---------------- 自检 ----------------

def _probe(pid):
    cfg = _providers()
    p = cfg[pid]
    t0 = time.time()
    text, usage, cost, err = chat(
        [{"role": "user", "content": "只回两个字：可用"}], timeout=40, provider=pid)
    if err:
        return False, err, cost
    return True, (text or "").strip().replace("\n", " ")[:20], cost


if __name__ == "__main__":
    args = set(sys.argv[1:])
    st = status()
    print("当前首选：", describe())
    if "--list" in args or not (has_key() or "--test" in args):
        print()
        print("%-13s %-9s %s" % ("平台", "状态", "说明"))
        print("-" * 78)
        for r in st["providers"]:
            mark = "✔ 已配" if r["ready"] else "— 未配"
            flag = "（冷却中）" if r["cooling"] else ""
            print("%-13s %-9s %s" % (r["id"], mark + flag, r["name"]))
            if r["note"]:
                print("%-13s %-9s %s" % ("", "", r["note"]))
        print()
        print("配 key 方法：agent/llm_key.txt 写一行  zhipu=你的key")
        print("免费 key：", PROVIDERS["zhipu"]["signup"])
        raise SystemExit(0 if has_key() else 1)

    if not has_key():
        print(_HINT)
        raise SystemExit(1)

    targets = order() if "--test" in args else [order()[0]]
    if "--test" in args:
        # 探测时清空冷却，否则"刚失败过的"会被跳过，看不出真实情况
        with _COOL_LOCK:
            _COOL.clear()
        print("逐家测试（共 %d 家）……" % len(targets))
        print()
    ok_any = False
    for pid in targets:
        ok, info, cost = _probe(pid)
        ok_any = ok_any or ok
        print("%-13s %s  %.2fs  %s" % (pid, "✔ 可用  " if ok else "✘ 不可用", cost, info))
    print()
    if ok_any:
        print("至少一家可用。实际调用会按顺序自动挑能用的那家。")
    else:
        print("一家都不通，检查 key 和网络。")
    raise SystemExit(0 if ok_any else 1)
