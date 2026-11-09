# -*- coding: utf-8 -*-
"""并发压测：N 张模板同时跑，看三件事 —— 会不会互相污染、耗时要多久、准确率退不退化。

【为什么要这个】
以前 api_server 用一把全局锁把「读取→抽取→写回」整条流水线串起来，
并发能力等于 0，`--workers` 也就无从谈起（开了也没用，反正排队）。
根因不是锁写得不好，是中间产物共享文件名：抽取结果、模板结构、成品、
验收报告全写在固定位置，两个任务同时跑必然互相覆盖。
现在每个任务一套独立目录（A23_MATCHER_DIR / A23_OUT_DIR / A23_STRUCTURE_JSON），
这个脚本就是用来证明隔离真的生效。

【数据渠道】
  small（默认）：测试数据/源文档 + 测试数据/导入文档 → doc_reader/output.json
                 （16 份文档、702 块；就是"90 秒原来是在这个池上测的"那个池）
  big          ：压测-大文档/结果/大池-output.json（16 份大文档、约 50 万块、194MB）
  模板          ：测试数据/模板表格（模板1/4/5）

【口径】
  · 这里报的耗时 = 一个任务从进程启动到产出成品的墙钟。
  · 用 --skip-read，即"池已就绪"的前提下计时（赛题是 16 份文档上传一次、5 张表分别填）。
  · 同一时刻最多 concurrency 个任务在跑，其余排队 —— 排队时间也算进各自的墙钟，
    因为用户看到的等待就是这个。
"""
import argparse
import json
import os
import re
import subprocess
import sys
import threading
import time
from pathlib import Path

BASE = Path(__file__).resolve().parents[2]          # 收尾对照/评测/xxx.py → 研一项目
ROOT = BASE / "测试数据" / "并发压测-1103"
TEMPLATE_DIR = BASE / "测试数据" / "模板表格"
SMALL_POOL = BASE / "doc_reader" / "output.json"
BIG_POOL = BASE / "压测-大文档" / "结果" / "大池-output.json"
OUT_TXT = Path(__file__).resolve().parent / "_并发压测.txt"
OUT_JSON = Path(__file__).resolve().parent / "_并发压测.json"


def find_tpl(n):
    hits = sorted(TEMPLATE_DIR.glob("模板%d-*.xlsx" % n))
    return hits[0].name if hits else None


def run_one(tpl_n, pool, sem, results, lock):
    name = "模板%d" % tpl_n
    tpl = find_tpl(tpl_n)
    if tpl is None:
        with lock:
            results[name] = {"成功": False, "错误": "模板不存在"}
        return

    # 目录名必须叫「结果-模板N」：评分器按这个约定找成品
    work = ROOT / ("结果-" + name)
    (work / "matcher").mkdir(parents=True, exist_ok=True)

    env = os.environ.copy()
    env.update({
        "PYTHONIOENCODING": "utf-8",
        "A23_TEMPLATE_DIR": str(TEMPLATE_DIR),
        "A23_MATCHER_DIR": str(work / "matcher"),
        "A23_OUT_DIR": str(work),                  # 成品直接落在这里，评分器才找得到
        "A23_STRUCTURE_JSON": str(work / "structure.json"),
        "A23_OUTPUT_JSON": str(pool),
    })

    with sem:
        t0 = time.time()
        try:
            proc = subprocess.run(
                [sys.executable, str(BASE / "pipeline.py"),
                 "--template", tpl, "--skip-read"],
                cwd=str(BASE), capture_output=True, env=env, timeout=900)
            el = time.time() - t0
            text = proc.stdout.decode("utf-8", "replace")
            err = proc.stderr.decode("utf-8", "replace")
            rc = proc.returncode
        except subprocess.TimeoutExpired:
            el = time.time() - t0
            text, err, rc = "", "超时（>900 秒）", -1

    (work / "日志.txt").write_text(text + ("\n--- stderr ---\n" + err if err.strip() else ""),
                                   encoding="utf-8")
    m = re.search(r"总耗时：([\d.]+) 秒", text)
    stages = {k: float(v) for k, v in
              re.findall(r"【(.+?)】完成 ｜ ([\d.]+) 秒", text)}
    ok = rc == 0 and "交卷文件" in text
    with lock:
        results[name] = {
            "成功": ok,
            "返回码": rc,
            "墙钟秒": round(el, 2),
            "流水线总耗时": float(m.group(1)) if m else None,
            "阶段": stages,
            "目录": str(work),
            "错误": (err.strip()[-300:] if not ok else None),
        }


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--templates", default="1,4,5", help="逗号分隔的模板编号")
    ap.add_argument("--concurrency", type=int, default=3, help="同时跑几个任务")
    ap.add_argument("--pool", default="small", choices=["small", "big"])
    args = ap.parse_args()

    pool = SMALL_POOL if args.pool == "small" else BIG_POOL
    if not pool.exists():
        raise SystemExit("池不存在：%s（先跑 doc_reader 或压测）" % pool)
    pool_mb = round(pool.stat().st_size / 1024 / 1024, 1)

    nums = [int(x) for x in args.templates.split(",") if x.strip()]
    ROOT.mkdir(parents=True, exist_ok=True)

    # 池里有多少块，报告里要写清楚（"702 块"这个数字要有出处）
    try:
        data = json.loads(pool.read_text(encoding="utf-8"))
        n_blocks = sum(len(d.get("blocks") or []) for d in data.get("documents", []))
        n_docs = len(data.get("documents", []))
    except Exception:
        n_blocks, n_docs = -1, -1

    print("=" * 62)
    print("并发压测 ｜ 池：%s（%d 份文档 / %d 块 / %.1f MB）"
          % (pool.name, n_docs, n_blocks, pool_mb))
    print("模板：%s ｜ 并发：%d ｜ 输出：%s" % (nums, args.concurrency, ROOT))
    print("=" * 62)

    results, lock = {}, threading.Lock()
    sem = threading.Semaphore(max(1, args.concurrency))
    t_all = time.time()
    threads = [threading.Thread(target=run_one,
                                args=(n, pool, sem, results, lock))
               for n in nums]
    for t in threads:
        t.start()
    for t in threads:
        t.join()
    wall_all = time.time() - t_all

    lines = []
    lines.append("并发压测报告 · %s" % time.strftime("%Y-%m-%d %H:%M"))
    lines.append("池：%s（%d 份文档 / %d 块 / %.1f MB）"
                 % (pool.name, n_docs, n_blocks, pool_mb))
    lines.append("并发度：%d ｜ 模板：%s" % (args.concurrency, nums))
    lines.append("")
    lines.append("%-8s %-6s %10s %12s %10s" % ("任务", "结果", "墙钟(秒)", "流水线(秒)", "红线"))
    for n in nums:
        name = "模板%d" % n
        r = results.get(name, {})
        w = r.get("墙钟秒", 0)
        flag = "✓ ≤90" if (r.get("成功") and w <= 90) else ("✗ 超时" if w > 90 else "✗ 失败")
        lines.append("%-8s %-6s %10.2f %12s %10s"
                     % (name, "成功" if r.get("成功") else "失败", w,
                        r.get("流水线总耗时"), flag))
    lines.append("")
    lines.append("总墙钟：%.2f 秒（%d 个任务）" % (wall_all, len(nums)))
    lines.append("")
    lines.append("各阶段耗时：")
    for n in nums:
        name = "模板%d" % n
        st = results.get(name, {}).get("阶段") or {}
        lines.append("  %s：%s" % (name, "、".join(
            "%s %.2fs" % (k, v) for k, v in st.items()) or "（没抓到）"))
    lines.append("")
    bad = [n for n in nums if not results.get("模板%d" % n, {}).get("成功")]
    over = [n for n in nums
            if results.get("模板%d" % n, {}).get("成功")
            and (results.get("模板%d" % n, {}).get("墙钟秒") or 0) > 90]
    if bad:
        verdict = "失败 %s" % bad
    elif over:
        # 以前这里只判"成功/失败"，超红线也照样报"全部成功"（11-03 模板5 126 秒撞上过）
        verdict = "全部产出成功，但 %s 超 90 秒红线" % over
    else:
        verdict = "全部成功且都在 90 秒内"
    lines.append("结论：%s" % verdict)

    text = "\n".join(lines)
    print(text)
    OUT_TXT.write_text(text, encoding="utf-8")
    OUT_JSON.write_text(json.dumps({"pool": str(pool), "blocks": n_blocks,
                                    "documents": n_docs,
                                    "concurrency": args.concurrency,
                                    "templates": nums,
                                    "wall_all": round(wall_all, 2),
                                    "results": results},
                                   ensure_ascii=False, indent=2), encoding="utf-8")


if __name__ == "__main__":
    main()
