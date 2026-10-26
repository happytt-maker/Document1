# -*- coding: utf-8 -*-
"""_回归与提速验证.py —— 压测优化（池缓存 + 上下文缓存 + 预计算）的正确性与提速验证
================================================================
改了什么（2026-10-14 压测后优化）：
  retrieve.py：build_column_contexts 加缓存 + 统一挂 _ctx + 预归一 _colheads_norm；
               score_blocks 支持 contexts 参数 + prefixes/bigrams 预计算；
  extract.py： load_pool() 池缓存（按 mtime 失效）+ 解析后预热列上下文。

本脚本用「旧实现」（改前逻辑的忠实复刻）对同一份大池逐列对比：
  ① 行为一致性：候选条数 + 全量逐项（分数、块身份）完全一致才算过
  ② 提速量化：同池同列，旧 vs 新 耗时对比

跑法：python _回归与提速验证.py
产出：_回归验证.txt
"""
import json
import re
import sys
import time
from pathlib import Path

ROOT = Path(r"C:\Users\覃敬涛\Desktop\研一项目")
sys.path.insert(0, str(ROOT / "matcher"))
import retrieve   # 新版（已优化）  # noqa: E402

POOL = ROOT / "压测-大文档" / "结果" / "大池-output.json"
OUT = Path(__file__).resolve().parent / "_回归验证.txt"

TESTS = ["责任判定", "规格", "库存数量", "投诉编号", "单价（元）", "备注"]


# ================== 旧实现（改前逻辑忠实复刻，勿改） ==================
def _old_strip_unit(name):
    return re.sub(r"[（(].*?[)）]", "", name).strip()


def _old_bigrams(text):
    return [text[i:i + 2] for i in range(len(text) - 1)]


def _old_score_block(text, core):
    score = 0
    prefixes = [core[:i] for i in range(3, len(core) + 1)]
    if any((p + "：" in text) or (p + ":" in text) for p in prefixes):
        score += 90
    if core in text:
        score += 40
    parts = _old_bigrams(core)
    if parts:
        hit = sum(1 for p in parts if p in text)
        score += int(40 * hit / len(parts))
    return score


def _old_build_column_contexts(blocks):
    cols = {}
    for i, b in enumerate(blocks):
        loc = b.get("locator") or {}
        text = (b.get("text") or "").strip()
        if not text:
            continue
        kind = loc.get("kind")
        if kind == "table_cell":
            key = ("d", b.get("file"), loc.get("table_index"), loc.get("col"))
        elif kind == "cell":
            key = ("x", b.get("file"), loc.get("sheet"), loc.get("col"))
        else:
            continue
        cols.setdefault(key, []).append((loc.get("row") or 0, i, text))

    contexts = [""] * len(blocks)
    for items in cols.values():
        items.sort()
        top = items[:3]
        head = top[0][2] if top else ""
        heads5 = [t for (_, _, t) in items[:5]]
        for (_, j, _t) in items:
            contexts[j] = " ".join(t for (_, jj, t) in top if jj != j)
            blocks[j]["_ctx"] = contexts[j]
            blocks[j]["_head"] = head
            blocks[j]["_colheads"] = heads5
    return contexts


def _old_colhead_hits(b, core):
    for h in (b.get("_colheads") or []):
        if _old_strip_unit(h) == core:
            return True
    return False


def _old_score_blocks(blocks, core, min_score=20):
    contexts = _old_build_column_contexts(blocks)
    scored = []
    for b, ctx in zip(blocks, contexts):
        text = b.get("text") or ""
        probe = (ctx + " " + text).strip() if ctx else text
        s = _old_score_block(probe, core)
        if _old_colhead_hits(b, core):
            s += 60
        if s >= min_score:
            scored.append((s, b))
    scored.sort(key=lambda t: -t[0])
    return scored
# =====================================================================


def compare(old_scored, new_scored):
    if len(old_scored) != len(new_scored):
        return False, "条数不同：旧 %d / 新 %d" % (len(old_scored), len(new_scored))
    for i, ((s1, b1), (s2, b2)) in enumerate(zip(old_scored, new_scored)):
        if s1 != s2 or b1 is not b2:
            return False, "第 %d 项不同：旧 (%d, %s) / 新 (%d, %s)" % (
                i + 1, s1, b1.get("file"), s2, b2.get("file"))
    return True, "%d 条全量逐项一致" % len(old_scored)


def main():
    lines = []

    def say(t=""):
        lines.append(t)
        print(t)

    say("=" * 64)
    say("回归与提速验证（旧实现 vs 新实现）· %s" % time.strftime("%Y-%m-%d %H:%M"))
    say("=" * 64)

    data = json.loads(POOL.read_text(encoding="utf-8"))
    blocks = [b for d in data.get("documents", []) for b in d.get("blocks", [])]
    say("池：%s ｜ %d 块 ｜ %.1f MB" % (POOL.name, len(blocks), POOL.stat().st_size / 1e6))
    say("")
    say("【1】行为一致性 + 提速（逐列）")
    say("-" * 64)

    all_same = True
    t_old_sum = t_new_sum = 0.0
    detail = []
    for col in TESTS:
        core = _old_strip_unit(col)
        t0 = time.time()
        old_s = _old_score_blocks(blocks, core, 20)
        t_old = time.time() - t0
        t0 = time.time()
        new_s = retrieve.score_blocks(blocks, core, 20)
        t_new = time.time() - t0
        ok, msg = compare(old_s, new_s)
        all_same = all_same and ok
        t_old_sum += t_old
        t_new_sum += t_new
        detail.append((col, t_old, t_new, len(old_s), ok))
        say("  %-8s 旧 %7.2fs → 新 %7.2fs ｜ %s ｜ %s" % (
            col, t_old, t_new, "一致 ✓" if ok else "不一致 ✗", msg))

    say("")
    say("【2】汇总")
    say("-" * 64)
    say("  6 列耗时合计：旧 %.2fs → 新 %.2fs（提速 %.1f 倍）"
        % (t_old_sum, t_new_sum, (t_old_sum / t_new_sum) if t_new_sum else 0))
    say("  逐列平均：旧 %.2fs → 新 %.2fs" % (t_old_sum / len(TESTS), t_new_sum / len(TESTS)))
    say("  行为一致性：%s" % ("6/6 列全量逐项一致 ✓" if all_same else "存在不一致 ✗（见上）"))
    say("  说明：旧列耗时含「每列重建列上下文」（改前行为）；新列耗时中第 1 列含首次构建，"
        "2 列起命中缓存。")
    say("=" * 64)

    OUT.write_text("\n".join(lines) + "\n", encoding="utf-8")
    print("（已存：%s）" % OUT)


if __name__ == "__main__":
    main()
