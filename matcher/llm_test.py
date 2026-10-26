# -*- coding: utf-8 -*-
"""
llm_test.py —— 第 4 步的第一步：跑通你的第一次大模型调用

跑法：VSCode 打开 matcher 文件夹 → 打开本文件 → 按 F5（或点右上角 ▶）。

2026-10-25：换成了阿里云百炼 MaaS（OpenAI 兼容接口）。
地址、模型名、密钥都统一在 agent/llm_config.py 里，这里不再重复写；
填密钥 / 换模型只改那一个文件，改完回这里按 F5 就能验证通不通。
"""
import sys
import time
from pathlib import Path

# 让本文件能找到 agent/llm_config.py（从 matcher 目录跑也能找到）
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "agent"))

from llm_config import MODEL, describe, has_key, chat as _llm_chat  # noqa: E402


def ask(question: str):
    """向模型提问一次，返回 (回答, 用量, 耗时秒数)。"""
    answer, usage, cost, err = _llm_chat(
        [{"role": "system", "content": "你是一个严谨的文档助手，回答简洁、只依据事实。"},
         {"role": "user", "content": question}],
        timeout=60,
    )
    if err:
        # 401=key 不对；402/403=额度或权限；404=地址/模型名不对；429=请求太频繁
        print("请求失败！" + err)
        return None, None, cost
    return answer, usage or {}, cost


if __name__ == "__main__":
    if not has_key():
        print("先别急着跑 —— 你还没填 key。")
        print("打开 agent/llm_config.py，把 API_KEY 填成百炼的 key，保存后重新运行。")
        print("（也可以新建 agent/llm_key.txt，把 key 写一行进去。）")
        raise SystemExit

    print("当前模型：", describe())
    question = "用一句话说明：什么是'从文档里提取信息'？"
    answer, usage, cost = ask(question)

    print("=" * 52)
    print("问题：", question)
    print("-" * 52)
    print("回答：", answer)
    print("=" * 52)
    print("本次耗时：%.2f 秒（记录它！赛题要求每文档 ≤ 90 秒）" % cost)
    print("本次用量：", usage)
