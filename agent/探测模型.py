# -*- coding: utf-8 -*-
"""
探测模型.py —— 拿到 key 后，自动试出你账号里到底哪个模型名能用

为什么要这个：模型名会随版本变（glm-4.5-flash → 4.7-flash → 以后 5.x），
文档写的和账号实际开通的可能不一致。与其猜，不如逐个试一遍。

它会测三件事（项目实际依赖这三件）：
  1. 能不能调通（model 名对不对、key 有没有效）
  2. 认不认"关思考"参数（不关的话简单指令也要多花几十~两百个 token 的时间）
  3. 认不认 response_format=json_object（抽取要模型只回 JSON）

用法：
  python agent/探测模型.py              # 测智谱（默认）
  python agent/探测模型.py --provider siliconflow
"""
import os
import sys
import time
from pathlib import Path

import requests

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
import llm_config as L  # noqa: E402

# 平台上值得一试的模型名（免费的打 ★）
CANDIDATES = {
    "zhipu": [
        ("glm-5.3-flash", "★ 现在主用（新用户赠 1500 万 token，始终思考）"),
        ("glm-5.3", "同代收费版"),
        ("glm-4.7-flashX", "控制台显示的写法（11-02 实测可用，比 4.7-flash 慢）"),
        ("glm-4.7-flash", "上一代（免费）"),
        ("glm-4.5-flash", "★ 更早一代免费"),
        ("glm-4-flash", "老版免费名（有些账号仍可用）"),
        ("glm-4.6v-flash", "★ 视觉版免费（本项目用不到）"),
        ("glm-4.7", "收费，质量更好，当备选"),
        ("glm-5", "收费，最强，贵"),
    ],
    "siliconflow": [
        ("Qwen/Qwen3-8B-Instruct", "免费档"),
        ("Qwen/Qwen2.5-7B-Instruct", "免费档"),
        ("THUDM/glm-4-9b-chat", "免费档"),
    ],
    "dashscope": [
        ("deepseek-v4.1-flash", "原主力"),
        ("qwen-plus", "备选"),
    ],
}


def probe_one(pid, model, cfg):
    """返回 (ok, 说明, 耗时)"""
    p = dict(cfg[pid])
    p["model"] = model
    key = L._key_for(pid, p)
    if not key:
        return False, "没有 key", 0
    url = str(p["base"]).rstrip("/") + "/chat/completions"
    s = requests.Session()
    s.trust_env = False
    t0 = time.time()

    # 第一步：最基本的一问一答
    try:
        r = s.post(url, timeout=30,
                   headers={"Authorization": "Bearer " + key,
                            "Content-Type": "application/json"},
                   json={"model": model,
                         "messages": [{"role": "user", "content": "只回两个字：可用"}]})
    except Exception as e:
        return False, "连不上：%s" % str(e)[:70], time.time() - t0

    if r.status_code != 200:
        try:
            msg = r.json().get("error", {}).get("message", "")
        except Exception:
            msg = r.text[:80]
        return False, "HTTP %d：%s" % (r.status_code, msg[:70]), time.time() - t0
    cost = time.time() - t0
    try:
        txt = r.json()["choices"][0]["message"]["content"]
    except Exception:
        return False, "返回不是预期结构", cost

    # 第二步：项目依赖的两个可选参数
    extras = []
    th = p.get("thinking")
    if th:
        try:
            r2 = s.post(url, timeout=30,
                        headers={"Authorization": "Bearer " + key,
                                 "Content-Type": "application/json"},
                        json={"model": model,
                              "messages": [{"role": "user", "content": "只回两个字：可用"}],
                              th[0]: th[1]})
            extras.append("思考档✓" if r2.status_code == 200 else "思考档✗(%d)" % r2.status_code)
        except Exception:
            extras.append("思考档✗")
    try:
        r3 = s.post(url, timeout=30,
                    headers={"Authorization": "Bearer " + key,
                             "Content-Type": "application/json"},
                    json={"model": model,
                          "messages": [{"role": "user",
                                        "content": '只回 JSON：{"ok":true}'}],
                          "response_format": {"type": "json_object"}})
        extras.append("JSON✓" if r3.status_code == 200 else "JSON✗(%d)" % r3.status_code)
    except Exception:
        extras.append("JSON✗")

    note = (txt or "").strip().replace("\n", " ")[:8]
    return True, "答「%s」  %s" % (note, " ".join(extras)), cost


def main():
    args = sys.argv[1:]
    pid = "zhipu"
    if "--provider" in args:
        pid = args[args.index("--provider") + 1]

    cfg = L._providers()
    if pid not in cfg:
        print("没有这个平台：%s" % pid)
        return 1
    if not L._key_for(pid, cfg[pid]):
        print("还没配 %s 的 key。" % pid)
        print("写进 agent/llm_key.txt 一行：%s=你的key" % pid)
        print("免费 key：%s" % cfg[pid].get("signup", ""))
        return 1

    print("平台：%s（%s）" % (cfg[pid].get("name", pid), cfg[pid]["base"]))
    print("=" * 72)
    good = []
    for model, tip in CANDIDATES.get(pid, []):
        ok, info, cost = probe_one(pid, model, cfg)
        print("%-26s %-8s %5.2fs  %s" % (model, "✔" if ok else "✘", cost, info))
        print("%-26s %s%s" % ("", " " * 8, tip))
        if ok:
            good.append(model)
    print("=" * 72)
    if not good:
        print("一个都不通。检查 key 是否正确、账号是否开通了模型。")
        return 1
    print("能用：%s" % "、".join(good))
    print()
    print("建议把第一个填进 agent/llm_config.py 的 PROVIDERS['%s']['model']：" % pid)
    print("    \"model\": \"%s\"," % good[0])
    return 0


if __name__ == "__main__":
    sys.exit(main())
