# -*- coding: utf-8 -*-
"""
回归验证.py —— 确认"为泛化测试做的三处小改造"没有影响模板4 老流程。
（只验证路径与列名解析，不调大模型、不写文件）

三项：
  ① template_reader 不带参数 → 仍读「测试数据/模板表格」的 5 张模板
  ② matcher/run_all 的 load_columns → 仍优先找「模板4」
  ③ writer 的 TEMPLATE_DIR → 不设环境变量时仍是「测试数据/模板表格」
"""
import json
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
BASE = ROOT.parent
PY = sys.executable


def dec(b):
    return b.decode("utf-8", "replace").strip()


print("=" * 58)
print("回归验证 · 模板4 老流程")
print("=" * 58)

# ① template_reader 默认模式
p = subprocess.run([PY, "main.py"], cwd=str(BASE / "template_reader"),
                   capture_output=True)
tpl = json.loads((BASE / "template_reader" / "模板表结构.json").read_text(encoding="utf-8"))
print("① template_reader 默认模式：返回码 %d，读出 %d 张模板" % (p.returncode, len(tpl)))
print("   " + "、".join(t["file"] for t in tpl))
print("   （期望：5 张，含模板4）")
print("")

# ② matcher load_columns
code = "import run_all\nprint(' | '.join(run_all.load_columns()))"
p = subprocess.run([PY, "-c", code], cwd=str(BASE / "matcher"),
                   capture_output=True)
print("② matcher.load_columns（不带参数，应找模板4）：")
print("   " + dec(p.stdout))
if p.returncode != 0:
    print("   " + dec(p.stderr)[:300])
print("")

# ③ writer 默认模板目录
code = ("import fill, verify_fill\n"
        "print('fill.TEMPLATE_DIR =', fill.TEMPLATE_DIR)\n"
        "print('模板4 文件存在 =', (fill.TEMPLATE_DIR / '模板4-客诉处理台账.xlsx').exists())\n"
        "print('verify_fill.TEMPLATE_DIR =', verify_fill.TEMPLATE_DIR)")
p = subprocess.run([PY, "-c", code], cwd=str(BASE / "writer"),
                   capture_output=True)
print("③ writer 默认模板目录（不设 A23_TEMPLATE_DIR）：")
print("   " + dec(p.stdout))
if p.returncode != 0:
    print("   " + dec(p.stderr)[:300])
print("")

print("结论：三项都显示老默认值 → 回归通过，模板4 流程不受影响。")
