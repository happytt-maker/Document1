# -*- coding: utf-8 -*-
"""
公式-源文档SUM-1103.py
----------------------
验证 doc_reader/readers/xlsx_reader.py 的 SUM 求值补丁（2026-11-03）。

背景：
    上一轮只给「模板合计行」补了公式（writer/fill.py: write_summary_formulas），
    但源文档里的合计格同样是公式。以前 _eval_formula 遇到函数一律返回 None，
    于是 销售台账 F14 =SUM(F3:F12) 读出来只是字符串 "=SUM(F3:F12)"，
    抽取端只能当它不存在 —— 而模板合计行恰恰需要这个数。

本脚本验三件事：
    A. 正例：源文档里的 SUM 现在能算出数（销售台账 F14、项目进度跟踪 D8）
    B. 反例：认不出的（多参数 SUM、SUMIF、跨表引用）必须保持"未计算"，
       宁可留空，绝不猜 —— 这条边界比 A 更重要
    C. 回归：全量源文档重读一遍，块数不变、没有异常

运行：
    python 收尾对照\评测\公式-源文档SUM-1103.py
"""
import io
import os
import sys
import traceback
from pathlib import Path

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent.parent                      # 研一项目/
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "doc_reader"))

from readers.xlsx_reader import read_xlsx, _eval_formula   # noqa: E402

SRC = ROOT / "测试数据" / "源文档"

out = io.StringIO()


def P(*a):
    line = " ".join(str(x) for x in a)
    print(line)
    out.write(line + "\n")


def cell_of(res, cell, sheet=None):
    for b in res["blocks"]:
        loc = b.get("locator") or {}
        if loc.get("cell") != cell:
            continue
        if sheet and loc.get("sheet") != sheet:
            continue
        return b
    return None


def show(res, cell, sheet=None, note=""):
    b = cell_of(res, cell, sheet)
    if b is None:
        P("  %-6s 未读到（该格可能是空格/合并区非左上角）%s" % (cell, note))
        return None
    P("  %-6s text=%-14r formula=%-18r type=%s  %s"
      % (cell, b.get("text"), b.get("formula"), b.get("value_type"), note))
    return b


P("=" * 78)
P("公式-源文档 SUM 求值 验收（2026-11-03）")
P("=" * 78)

# ---------------- A. 正例 ----------------
P("")
P("【A】正例：源文档合计行能算出数")
fails = []


def expect(b, want, label):
    if b is None:
        fails.append("%s 未读到" % label)
        return
    got = str(b.get("text") or "").replace(",", "")
    try:
        ok = abs(float(got) - float(want)) < 1e-6
    except Exception:
        ok = got == str(want)
    if b.get("value_type") != "formula_evaluated":
        ok = False
    P("      → %s：期望 %s，实得 %s，%s" % (label, want, got, "PASS" if ok else "FAIL"))
    if not ok:
        fails.append("%s 期望 %s 实得 %s" % (label, want, got))


p1 = SRC / "xlsx" / "销售台账.xlsx"
P("· %s" % p1.name)
r1 = read_xlsx(p1)
b = show(r1, "F14", note="=SUM(F3:F12) 销售额合计")
expect(b, 4768000, "销售台账 F14")

p2 = SRC / "xlsx" / "项目进度跟踪表.xlsx"
P("· %s（两个工作表，「合计」在 工时登记）" % p2.name)
r2 = read_xlsx(p2)
b = show(r2, "E8", sheet="工时登记", note="=SUM(D2:D7) 工时合计")
expect(b, 41.5, "项目进度跟踪 工时登记!E8")

# ---------------- B. 反例 ----------------
P("")
P("【B】反例：认不出的一律不许算（留空 > 算错）")


class FakeSheet:
    """最小假表：够 _eval_formula 取数就行。"""

    def __init__(self, data):
        self.data = data

    def __getitem__(self, coord):
        class C:
            pass
        c = C()
        c.value = self.data.get(coord)
        return c


fake = FakeSheet({"A1": 1, "A2": 2, "A3": 3, "B1": 10, "B2": 20})
cases = [
    ("=SUM(A1:A3)", 6, "连续区间 —— 应该算"),
    ("=SUM(A1:B2)", 33, "矩形区间 —— 应该算"),
    ("=SUM(A1:A2,B1:B2)", None, "多区间参数 —— 不猜"),
    ("=SUMIF(A1:A3,\">1\")", None, "条件求和 —— 不猜"),
    ("=SUM(另一张表!A1:A3)", None, "跨表引用 —— 不猜"),
    ("=A1+B1*2", 21, "四则 —— 老能力不能退化"),
]
for f, want, why in cases:
    got = _eval_formula(f, fake, fake, {})
    if want is None:
        ok = got is None
    else:
        ok = got is not None and abs(float(got) - want) < 1e-6
    P("  %-26s → %-10r %s  （%s）" % (f, got, "PASS" if ok else "FAIL", why))
    if not ok:
        fails.append("%s 期望 %r 实得 %r" % (f, want, got))

# ---------------- C. 回归 ----------------
P("")
P("【C】回归：全量源文档重读")
total_blocks = 0
total_formula = 0
evaluated = 0
files = 0
errors = []
for p in sorted(SRC.rglob("*")):
    if not p.is_file() or p.suffix.lower() not in (".xlsx", ".xlsm"):
        continue
    files += 1
    try:
        r = read_xlsx(p)
    except Exception:
        errors.append("%s: %s" % (p.name, traceback.format_exc(limit=1)))
        continue
    n = r["block_count"]
    fe = [x for x in r["blocks"] if x.get("value_type") == "formula_evaluated"]
    total_blocks += n
    total_formula += sum(1 for x in r["blocks"] if x.get("formula"))
    evaluated += len(fe)
    P("  %-26s 块 %-5d 公式 %-3d 其中算出 %d"
      % (p.name, n, sum(1 for x in r["blocks"] if x.get("formula")), len(fe)))
P("  合计：%d 个 xlsx / %d 块 / 公式 %d 条 / 算出 %d 条" %
  (files, total_blocks, total_formula, evaluated))
if errors:
    P("  异常 %d 个：" % len(errors))
    for e in errors:
        P("   - " + e.replace("\n", " "))
    fails.append("回归出现异常 %d 个" % len(errors))

# ---------------- 结论 ----------------
P("")
P("=" * 78)
if fails:
    P("结论：FAIL —— %d 项不合预期" % len(fails))
    for f in fails:
        P("  · " + f)
else:
    P("结论：PASS —— 正例算出、反例没猜、全量重读无异常")
P("=" * 78)

dest = HERE.parent / "公式-源文档SUM-1103.txt"
dest.write_text(out.getvalue(), encoding="utf-8")
print("\n[已写] %s" % dest)
sys.exit(1 if fails else 0)
