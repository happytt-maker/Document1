# -*- coding: utf-8 -*-
"""
run_all.py
----------
一键跑完：读取 → 验收 → 输出报告。

用法：在 VSCode 里打开这个文件，按 F5 选「一键全流程」，
或者终端里 python run_all.py

做完会生成三份报告：
    output.json      读取结果（给程序用）
    验收报告.txt      位置回溯验收（给人看）
    汇总报告.txt      本次全流程的大盘数据
"""

import io
import json
import subprocess
import sys
from pathlib import Path

BASE = Path(__file__).parent


def run_step(name, script):
    """跑一个脚本，把它的输出收集起来。"""
    print("\n" + "=" * 66)
    print("▶ " + name)
    print("=" * 66)

    proc = subprocess.run(
        [sys.executable, str(BASE / script)],
        cwd=str(BASE),
        capture_output=True,
    )
    stdout = proc.stdout.decode("utf-8", errors="replace")
    stderr = proc.stderr.decode("utf-8", errors="replace")

    print(stdout)
    if stderr.strip():
        print("--- 错误输出 ---")
        print(stderr)

    return stdout, stderr, proc.returncode


def summarize():
    """从 output.json 汇总四种格式的成果。"""
    p = BASE / "output.json"
    if not p.exists():
        return None

    data = json.loads(p.read_text(encoding="utf-8"))

    lines = []
    lines.append("=" * 66)
    lines.append("汇总报告")
    lines.append("=" * 66)
    lines.append("")

    total_files = 0
    total_blocks = 0
    by_type = {}

    for doc in data["documents"]:
        if "error" in doc:
            continue
        t = doc["type"]
        n = doc.get("block_count", 0)
        total_files += 1
        total_blocks += n

        if t not in by_type:
            by_type[t] = {"files": 0, "blocks": 0, "formulas": 0}
        by_type[t]["files"] += 1
        by_type[t]["blocks"] += n

        # xlsx 的公式数在 sheets 里
        for s in doc.get("sheets", []):
            by_type[t]["formulas"] += s.get("formula_count", 0)

    lines.append("{0:<8} {1:>6} {2:>10} {3:>10}".format("格式", "文件数", "内容块", "公式数"))
    lines.append("-" * 40)
    for t in sorted(by_type):
        d = by_type[t]
        lines.append(
            "{0:<8} {1:>6} {2:>10} {3:>10}".format(
                t, d["files"], d["blocks"], d["formulas"]
            )
        )
    lines.append("-" * 40)
    lines.append(
        "{0:<8} {1:>6} {2:>10}".format("合计", total_files, total_blocks)
    )
    lines.append("")
    lines.append("读取失败的文件：")

    fails = [d for d in data["documents"] if "error" in d]
    if fails:
        for d in fails:
            lines.append("  - {0}：{1}".format(d["file"], d["error"]))
    else:
        lines.append("  无。全部读通。")

    lines.append("")
    lines.append("每个文件明细：")
    lines.append("")
    for doc in data["documents"]:
        if "error" in doc:
            lines.append("  [失败] {0}".format(doc["file"]))
            continue
        lines.append(
            "  [{0}] {1}  —— {2} 块".format(
                doc["type"], doc["file"], doc.get("block_count", 0)
            )
        )
        if doc["type"] == "xlsx":
            for s in doc.get("sheets", []):
                lines.append(
                    "        工作表「{0}」 {1}行×{2}列  {3}块  公式{4}".format(
                        s["name"], s["max_row"], s["max_column"],
                        s["block_count"], s["formula_count"],
                    )
                )

    return "\n".join(lines)


def main():
    # 第一步：读取
    out, err, code = run_step("第 1 步 / 读取所有文档", "main.py")
    if code != 0:
        print("\n读取失败，后面步骤跳过。")
        return

    # 第二步：验收
    out, err, code = run_step("第 2 步 / 验收位置回溯", "verify.py")
    if code != 0:
        print("\n验收脚本执行出错。")

    # 第三步：汇总
    text = summarize()
    if text:
        print("\n" + text)
        (BASE / "汇总报告.txt").write_text(text, encoding="utf-8")
        print("\n已写出：汇总报告.txt")
    else:
        print("\n没有找到 output.json，无法汇总。")


if __name__ == "__main__":
    main()
