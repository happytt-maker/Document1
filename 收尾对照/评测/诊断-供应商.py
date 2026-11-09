# -*- coding: utf-8 -*-
"""
诊断-供应商.py —— 只读：为什么模板1「供应商」抽不到。

背景（2026-11-02）：定源把模板1 锁在《设备采购合同》内（4/5 列列头精确覆盖），
但「供应商」这一列在合同内字面分 <20（=源内 weak），跑出来 0 条，模板1 因此漏填 4 格。
本脚本把合同的所有块、以及「供应商」的打分逐条打出来，判断到底是
  ① 源里根本没有（该留空）  还是  ② 有但换了说法（该救）
预期答案：② —— 合同里写的是「乙方（供货方）：泰隆机电」。

用法：python 收尾对照/评测/诊断-供应商.py
输出：收尾对照/评测/诊断-供应商.txt
"""
import os
import sys
from pathlib import Path

BASE = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(BASE / "matcher"))
sys.path.insert(0, str(BASE / "agent"))
os.environ.setdefault("A23_OUTPUT_JSON", str(BASE / "doc_reader" / "output.json"))

import extract                                      # noqa: E402
from retrieve import locate_text, score_blocks      # noqa: E402

OUT = Path(__file__).resolve().parent / "诊断-供应商.txt"
lines = []


def say(t=""):
    lines.append(t)


def main():
    blocks = extract.load_pool()
    doc = [b for b in blocks if b.get("file", "").startswith("设备采购合同")]
    say("合同块数：%d" % len(doc))
    say("")
    say("=== 合同全部块 ===")
    for i, b in enumerate(doc, 1):
        loc = b.get("locator") or {}
        say("[%d] kind=%s loc=%s" % (i, loc.get("kind"), locate_text(b)))
        say("    text=%r" % (b.get("text") or "")[:120])
        say("    colheads=%r" % ((b.get("_colheads") or [])[:5],))

    say("")
    say("=== 对「供应商」打分（含语义召回）===")
    for s, b in score_blocks(doc, "供应商")[:12]:
        say("%4d 分  via_sem=%s ｜ %s ｜ %s"
            % (s, b.get("_via_sem"), locate_text(b), (b.get("text") or "")[:50]))

    say("")
    say("=== 含「乙方」「供货」的块 ===")
    n = 0
    for b in doc:
        t = b.get("text") or ""
        if "乙方" in t or "供货" in t:
            n += 1
            say("%s ｜ %r" % (locate_text(b), t[:100]))
    if not n:
        say("（无）")

    OUT.write_text("\n".join(lines), encoding="utf-8")
    print("（已存 %s）" % OUT)


if __name__ == "__main__":
    main()
