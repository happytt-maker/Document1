# -*- coding: utf-8 -*-
"""
定源-验证.py —— 只验证"定源选得对不对"，不调用大模型、不花钱、不覆盖任何抽取结果。

为什么单独有这个脚本：
  定源定错 = 整表报废（比不定源严重）。所以上线前先纯本地跑一遍，
  确认 5 张模板的主源都选对、放开列也符合预期，再拿真抽取验成绩。

用法：python 收尾对照/评测/定源-验证.py
输出：屏幕 + 收尾对照/评测/定源-验证.txt
"""
import json
import os
import sys
from pathlib import Path

BASE = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(BASE / "matcher"))
sys.path.insert(0, str(BASE / "agent"))
os.environ.setdefault("A23_OUTPUT_JSON", str(BASE / "doc_reader" / "output.json"))

import anchor                      # noqa: E402
import extract                     # noqa: E402

STRUCT = BASE / "template_reader" / "模板表结构.json"
OUT = Path(__file__).resolve().parent / "定源-验证.txt"

# 期望的主源（人工按"这张表该用哪份文档填"核对出来的，不是系统自证）
EXPECT = {
    "模板1": "设备采购合同",
    "模板2": "员工信息表",
    "模板3": "项目推进会会议纪要",
    "模板4": "客户投诉处理记录",
    "模板5": "库存盘点表",
}
# 预期"源内字面认不出、要靠语义或留空"的列（anchor 里叫 weak）
EXPECT_WEAK = {
    "模板1": {"供应商"},                 # 合同里叫「乙方（供货方）」，换个说法
    "模板3": {"备注"},                   # 会议纪要真没有备注列 → 应诚实留空
    "模板5": {"库存数量", "是否紧缺"},     # 库存数量＝计算列，是否紧缺是主观列
}

lines = []


def say(t=""):
    lines.append(t)
    print(t)


def main():
    blocks = extract.load_pool()
    data = json.loads(STRUCT.read_text(encoding="utf-8"))
    ok = 0
    total = 0

    say("=" * 78)
    say("定源验证 ｜ 池：%s ｜ %d 块" % (extract.OUTPUT_JSON.name, len(blocks)))
    say("=" * 78)

    for t in data:
        if not t.get("sheets"):
            continue
        cols = [c["title"] for c in t["sheets"][0].get("columns", [])
                if c.get("title") and c["title"] != "序号"]
        if not cols:
            continue
        name = t.get("file", "?")
        key = name.split("-")[0]

        say("")
        say("-" * 78)
        say("%s ｜ %d 列" % (name, len(cols)))
        say("-" * 78)
        p = anchor.plan(blocks, cols)
        for ln in p["lines"]:
            say("  " + ln)

        total += 1
        want = EXPECT.get(key)
        got = p["file"]
        hit = bool(got and want and want in got)
        say("  结论：%s（期望含「%s」，实得《%s》）"
            % ("✅ 选对" if hit else "❌ 不对", want, got))
        if hit:
            ok += 1

        # weak 列对照（源内字面认不出的列）
        weak = set(p.get("weak") or [])
        want_weak = EXPECT_WEAK.get(key, set())
        say("  源内认不出：%s" % ("、".join(sorted(weak)) if weak else "（无）"))
        if want_weak or weak:
            say("  期望结果：%s → %s"
                % ("、".join(sorted(want_weak)) or "（无）",
                   "✅ 一致" if weak == want_weak else "❌ 不一致"))

    say("")
    say("=" * 78)
    say("主源选中：%d/%d" % (ok, total))
    say("=" * 78)
    OUT.write_text("\n".join(lines), encoding="utf-8")
    print("（已存 %s）" % OUT)
    return 0 if ok == total else 1


if __name__ == "__main__":
    sys.exit(main())
