# -*- coding: utf-8 -*-
"""
定源-诊断.py —— 只读、不调大模型，量"每份文档对每张模板能承接几列"。

为什么要有这个脚本（2026-11-02）：
  reaudit() 是"跑完再投票纠偏"，根因却是**抽取时模型同时看到两份文档的同名列**，
  于是自己把两处都抽出来（模板3「负责人」曾抽出 9 条：会议纪要 4 + 项目总览 5）。
  要改成"事前定源"，就得先知道：本地信号到底能不能把正确的那份文档选出来。
  不先量、直接拍阈值，等于又造一个"改了代码但行为没变"的坑。

量什么（每份文档 × 每列）：
  best   该文档内这一列的字面最高分（retrieve.score_blocks）
  exact  该文档里有没有"列头与目标列名完全相等"的列（colhead_hits，强信号）
  cols   该文档一共认出多少个列头（真表格还是零散段落）
  rows   该文档的块数

用法：
    python 收尾对照/评测/定源-诊断.py            # 全部模板
    python 收尾对照/评测/定源-诊断.py 模板3       # 只诊断一张

输出：屏幕 + 收尾对照/评测/定源-诊断.txt
"""
import json
import os
import sys
from collections import Counter
from pathlib import Path

BASE = Path(__file__).resolve().parents[2]          # = 研一项目/
sys.path.insert(0, str(BASE / "matcher"))
sys.path.insert(0, str(BASE / "agent"))

os.environ.setdefault("A23_OUTPUT_JSON", str(BASE / "doc_reader" / "output.json"))

import extract                                     # noqa: E402
from retrieve import score_blocks, strip_unit      # noqa: E402

STRUCT = BASE / "template_reader" / "模板表结构.json"
OUT_TXT = Path(__file__).resolve().parent / "定源-诊断.txt"

lines = []


def say(t=""):
    lines.append(t)
    print(t)


def templates():
    """[(模板名, [列名...]), ...] —— 序号是结构列，不参与。"""
    data = json.loads(STRUCT.read_text(encoding="utf-8"))
    out = []
    for t in data:
        if not t.get("sheets"):
            continue
        cols = [c["title"] for c in t["sheets"][0].get("columns", [])
                if c.get("title") and c["title"] != "序号"]
        if cols:
            out.append((t.get("file", "?"), cols))
    return out


def main():
    want = sys.argv[1] if len(sys.argv) > 1 else ""
    blocks = extract.load_pool()
    by_file = {}
    for b in blocks:
        by_file.setdefault(b.get("file") or "?", []).append(b)

    say("=" * 78)
    say("定源诊断 ｜ 池：%s" % extract.OUTPUT_JSON)
    say("文档数 %d ｜ 块数 %d" % (len(by_file), len(blocks)))
    say("=" * 78)

    for name, cols in templates():
        if want and want not in name:
            continue
        say("")
        say("-" * 78)
        say("%s ｜ %d 列：%s" % (name, len(cols), "、".join(cols)))
        say("-" * 78)

        # 表头：每份文档认出多少列头
        say("%-34s %6s %6s" % ("文档", "块数", "列头数"))
        for f, bs in sorted(by_file.items(), key=lambda kv: -len(kv[1])):
            heads = set()
            for b in bs:
                for h in (b.get("_colheads") or [])[:5]:
                    heads.add(strip_unit(h))
            say("%-34s %6d %6d" % (f[:32], len(bs), len(heads)))

        # 逐列：每份文档能不能承接
        say("")
        say("逐列承接（best=字面最高分，★=列头精确相等，·=字面≥40）")
        hdr = "%-14s" % "列"
        rank = []
        for f, bs in by_file.items():
            hdr += " %-12s" % f[:10]
            rank.append(f)
        say(hdr)
        cover = Counter()
        for col in cols:
            core = strip_unit(col)
            row = "%-14s" % col
            for f in rank:
                bs = by_file[f]
                scored = score_blocks(bs, core, use_semantic=False)
                best = scored[0][0] if scored else 0
                exact = any(b.get("_colheads_norm") and core in b["_colheads_norm"]
                            for b in bs)
                mark = "★" if exact else ("·" if best >= 40 else (" " if best < 20 else "~"))
                row += " %-12s" % ("%s%d" % (mark, best))
                if exact:
                    cover[f] += 1
            say(row)
        say("")
        say("列头精确覆盖数：")
        for f, n in cover.most_common():
            say("   %-34s %d/%d 列" % (f[:32], n, len(cols)))

    OUT_TXT.write_text("\n".join(lines), encoding="utf-8")
    print("\n（已存 %s）" % OUT_TXT)


if __name__ == "__main__":
    main()
