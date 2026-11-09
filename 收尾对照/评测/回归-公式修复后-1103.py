# -*- coding: utf-8 -*-
"""
回归-公式修复后-1103.py
-----------------------
池重建（含源文档 SUM 求值）+ fill.py 的合计行/补行改动之后，
五张模板重跑一遍，确认准确率与耗时都没退化。

为什么不直接重跑「并发压测-1103.py」：那个脚本写死在 测试数据/并发压测-1103/，
里面是并发那组证据，不该被覆盖。这里单独落 测试数据/09-复测-公式修复后/。

用法：
    python 收尾对照\\评测\\回归-公式修复后-1103.py            # 串行跑模板1~5
    python 收尾对照\\评测\\回归-公式修复后-1103.py --templates 1,4
"""
import argparse
import json
import os
import re
import subprocess
import sys
import time
from pathlib import Path

BASE = Path(__file__).resolve().parents[2]
ROOT = BASE / "测试数据" / "09-复测-公式修复后"
TEMPLATE_DIR = BASE / "测试数据" / "模板表格"
POOL = BASE / "doc_reader" / "output.json"
HERE = Path(__file__).resolve().parent


def find_tpl(n):
    hits = sorted(TEMPLATE_DIR.glob("模板%d-*.xlsx" % n))
    return hits[0].name if hits else None


def run_one(n):
    name = "模板%d" % n
    tpl = find_tpl(n)
    if tpl is None:
        return {"成功": False, "错误": "模板不存在"}
    work = ROOT / ("结果-" + name)
    (work / "matcher").mkdir(parents=True, exist_ok=True)
    env = os.environ.copy()
    env.update({
        "PYTHONIOENCODING": "utf-8",
        "A23_TEMPLATE_DIR": str(TEMPLATE_DIR),
        "A23_MATCHER_DIR": str(work / "matcher"),
        "A23_OUT_DIR": str(work),
        "A23_STRUCTURE_JSON": str(work / "structure.json"),
        "A23_OUTPUT_JSON": str(POOL),
    })
    t0 = time.time()
    proc = subprocess.run([sys.executable, str(BASE / "pipeline.py"),
                           "--template", tpl, "--skip-read"],
                          cwd=str(BASE), capture_output=True, env=env, timeout=900)
    el = time.time() - t0
    text = proc.stdout.decode("utf-8", "replace")
    err = proc.stderr.decode("utf-8", "replace")
    (work / "日志.txt").write_text(
        text + ("\n--- stderr ---\n" + err if err.strip() else ""), encoding="utf-8")
    m = re.search(r"总耗时：([\d.]+) 秒", text)
    ok = proc.returncode == 0 and "交卷文件" in text
    return {"成功": ok, "墙钟秒": round(el, 2),
            "流水线总耗时": float(m.group(1)) if m else None,
            "错误": None if ok else (err.strip()[-300:] or text.strip()[-300:])}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--templates", default="1,2,3,4,5")
    args = ap.parse_args()
    nums = [int(x) for x in args.templates.split(",") if x.strip()]

    data = json.loads(POOL.read_text(encoding="utf-8"))
    n_blocks = sum(len(d.get("blocks") or []) for d in data.get("documents", []))
    print("池：%s（%d 块）｜模板：%s" % (POOL.name, n_blocks, nums))

    res = {}
    t_all = time.time()
    for n in nums:
        r = run_one(n)
        res["模板%d" % n] = r
        print("  模板%d  %-4s %7.2f 秒  流水线 %s"
              % (n, "成功" if r["成功"] else "失败", r["墙钟秒"], r.get("流水线总耗时")))

    wall = time.time() - t_all
    out = HERE / "_回归-公式修复后.txt"
    lines = ["回归（公式修复后）· %s" % time.strftime("%Y-%m-%d %H:%M"),
             "池 %d 块 ｜ 串行跑 %s ｜ 总墙钟 %.2f 秒" % (n_blocks, nums, wall), ""]
    for n in nums:
        r = res["模板%d" % n]
        lines.append("模板%d  %-4s 墙钟 %.2f 秒  %s"
                     % (n, "成功" if r["成功"] else "失败", r["墙钟秒"],
                        "" if r["成功"] else (r.get("错误") or "")))
    out.write_text("\n".join(lines), encoding="utf-8")
    print("\n[已写] %s" % out)
    print("下一步评分：$env:A23_FILL_ROOT=\"%s\"; python 收尾对照\\评测\\评分器.py" % ROOT)


if __name__ == "__main__":
    main()
