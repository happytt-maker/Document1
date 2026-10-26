# -*- coding: utf-8 -*-
"""_查看候选.py —— 重现某列在大池上的 TOP 候选（用于诊断模型纠结的根源）"""
import json
import sys
import time
from pathlib import Path

ROOT = Path(r"C:\Users\覃敬涛\Desktop\研一项目")
sys.path.insert(0, str(ROOT / "matcher"))
import retrieve   # noqa: E402
from retrieve import locate_text, strip_unit   # noqa: E402

POOL = ROOT / "压测-大文档" / "结果" / "大池-output.json"
OUT = Path(__file__).resolve().parent / "_候选诊断.txt"

COLS = ["物料名称", "单位", "物料编码", "存放库位", "规格"]


def main():
    lines = []

    def say(t=""):
        lines.append(t)
        print(t)

    t0 = time.time()
    data = json.loads(POOL.read_text(encoding="utf-8"))
    blocks = [b for d in data.get("documents", []) for b in d.get("blocks", [])]
    retrieve.build_column_contexts(blocks)
    say("池：%d 块 ｜ 载入 %.1fs" % (len(blocks), time.time() - t0))

    for col in COLS:
        core = strip_unit(col)
        t0 = time.time()
        scored = retrieve.score_blocks(blocks, core, 20)
        dt = time.time() - t0
        say("")
        say("=" * 64)
        say("【%s】候选 %d 条 ｜ 打分 %.2fs ｜ TOP 15：" % (col, len(scored), dt))
        for i, (s, b) in enumerate(scored[:15], start=1):
            txt = (b.get("text") or "").replace("\n", " ⏎ ")
            if len(txt) > 46:
                txt = txt[:46] + "……"
            head = (b.get("_head") or "").replace("\n", " ")
            say("[%2d] %3d分 ｜ %s ｜ %s ｜ 列头:%s ｜ %s"
                % (i, s, b.get("file"), locate_text(b), head[:14], txt))

    OUT.write_text("\n".join(lines) + "\n", encoding="utf-8")
    print("（已存：%s）" % OUT.name)


if __name__ == "__main__":
    main()
