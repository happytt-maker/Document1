# -*- coding: utf-8 -*-
"""
COVID-19 智能填表 · 独立对账（不复用 matcher 的任何代码）
================================================================
为什么另写一份：系统自己的"值一致 15252/15252"只证明**写回没抄错抽取结果**，
不证明**抽取结果对**。真正的对账要绕开整条流水线，直接从源表重算一遍期望值。

对账口径（全部可机械判定）：
  1. 区间内源行数 = 输出数据行数
  2. 逐行逐格：输出第 k 行 == 源表第 k 条命中行的对应列
  3. 输出里每一行都来自区间内（1:1 对上就等于证明了没有混进区间外的行）
  4. 国家覆盖：命中国家数 × 每国天数 == 总行数（没有整段漏掉）
  5. 空值一致：源里空的格，输出也必须空（不许编）
  6. 表头一个字没动
  7. 数值型：人均GDP/人口/病例数 都应是数字（不是文本）
"""
import datetime
import random
import sys
from collections import Counter, defaultdict

from openpyxl import load_workbook

SRC = r"C:\Users\覃敬涛\Desktop\研一项目\测试数据\真实数据\07-COVID-19\源文档\COVID-19全球数据集（节选）.xlsx"
OUT = r"C:\Users\覃敬涛\Desktop\研一项目\测试数据\真实数据\07-COVID-19\结果\COVID-19 模板-已填写.xlsx"
TPL = r"C:\Users\覃敬涛\Desktop\研一项目\测试数据\真实数据\07-COVID-19\模板\COVID-19 模板.xlsx"

LO, HI = datetime.date(2020, 7, 1), datetime.date(2020, 8, 31)
# 模板列（表头原文）→ 源表列（对账用，独立写死，不用 direct.py 的匹配结果）
# 注意模板表头写的是「国家/地区」（结构读取器把斜线拆成了 国家 + 地区，标题记作「国家」），
# 对账按表头原文走，不按结构里的 title —— 少一层转换，少一处能错的地方。
MAP = {"国家/地区": "国家/地区", "大洲": "大洲", "人均GDP": "人均GDP",
       "人口": "人口", "每日检测数": "每日检测数", "病例数": "病例数"}

ok, bad = [], []


def check(cond, msg):
    (ok if cond else bad).append(msg)
    print(("  [通过] " if cond else "  [失败] ") + msg)


def norm(v):
    if v is None:
        return ""
    if isinstance(v, float) and v == int(v):
        return str(int(v))
    return str(v).strip()


def main():
    # ---------- 从源表独立重算期望值 ----------
    wb = load_workbook(SRC, read_only=True, data_only=True)
    ws = wb["Data"]
    rows = list(ws.iter_rows(min_row=1, values_only=True))
    hdr = [str(c).strip() for c in rows[0]]
    idx = {k: hdr.index(v) for k, v in MAP.items()}
    di = hdr.index("日期")

    expect = []          # [(源行号, {列: 值})]
    for i, r in enumerate(rows[1:], start=2):
        v = r[di]
        if isinstance(v, datetime.datetime):
            d = v.date()
        else:
            continue
        if LO <= d <= HI:
            expect.append((i, {k: r[j] for k, j in idx.items()}))
    wb.close()
    print("区间内源记录：%d 条\n" % len(expect))

    # ---------- 读输出 ----------
    wb_f = load_workbook(OUT)
    ws_f = wb_f.active
    tpl = load_workbook(TPL).active
    cols = [c.value for c in tpl[1]]
    print("原模板表头：", cols)
    print("输出表头：  ", [c.value for c in ws_f[1]])
    print()

    print("【1】表头未被改动")
    check([c.value for c in ws_f[1]] == cols, "表头 6 列与原模板完全一致")

    print("【2】行数 = 区间内源记录数")
    n_out = ws_f.max_row - 1
    check(n_out == len(expect), "输出 %d 行 ｜ 期望 %d 行" % (n_out, len(expect)))

    print("【3】逐行逐格比对（%d 行 × 6 列）" % n_out)
    col_pos = {cols.index(c) + 1: c for c in MAP}
    diff, n_cell = [], 0
    for k, (src_row, vals) in enumerate(expect):
        r = k + 2
        for cpos, title in col_pos.items():
            got = ws_f.cell(row=r, column=cpos).value
            exp = vals[title]
            n_cell += 1
            if norm(got) != norm(exp):
                diff.append((r, title, norm(exp), norm(got), src_row))
    check(not diff, "%d 格逐格一致（含空值）" % n_cell)
    for d in diff[:8]:
        print("      第 %d 行 %s：期望 %r ｜ 实际 %r（源行 %d）" % d)

    print("【4】来源完整性：国家数 × 天数")
    c = Counter(norm(v["国家/地区"]) for _r, v in expect)
    check(len(c) == 41, "覆盖 %d 个国家/地区" % len(c))
    check(set(c.values()) == {62}, "每个国家都是 62 天（2020-07-01 ~ 08-31 全月无缺）")

    print("【5】空值一致（源里空的，输出也必须空）")
    n_blank = sum(1 for _r, v in expect if norm(v["每日检测数"]) == "")
    blank_ok = all(norm(ws_f.cell(row=k + 2, column=cols.index("每日检测数") + 1).value) == ""
                   for k, (_r, v) in enumerate(expect) if norm(v["每日检测数"]) == "")
    check(blank_ok, "「每日检测数」源里空 %d 格，输出同样留空（没编数）" % n_blank)

    print("【6】数值列写成了数字，不是文本")
    for title in ("人均GDP", "人口", "病例数"):
        cpos = cols.index(title) + 1
        vals = [ws_f.cell(row=r, column=cpos).value for r in range(2, min(n_out + 2, 402))]
        n_num = sum(1 for v in vals if isinstance(v, (int, float)))
        check(n_num == len(vals), "%s：前 %d 行全是数字（%d/%d）" % (title, len(vals), n_num, len(vals)))

    print("【7】溯源抽检（随机 5 行，手工核对源单元格）")
    random.seed(7)
    for k in sorted(random.sample(range(len(expect)), 5)):
        src_row, vals = expect[k]
        r = k + 2
        line = "  · 输出第 %d 行 ← 源第 %d 行 ｜ " % (r, src_row)
        line += " ｜ ".join("%s=%s" % (t, norm(vals[t])) for t in ("国家/地区", "大洲", "病例数"))
        line += " ｜ 输出：" + " ｜ ".join(
            "%s=%s" % (t, norm(ws_f.cell(row=r, column=cols.index(t) + 1).value))
            for t in ("国家/地区", "大洲", "病例数"))
        print(line)

    print()
    print("=" * 66)
    print("对账结果：通过 %d 项，失败 %d 项" % (len(ok), len(bad)))
    return 0 if not bad else 2


if __name__ == "__main__":
    sys.exit(main())
