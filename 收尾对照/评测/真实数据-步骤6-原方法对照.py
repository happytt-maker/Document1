# -*- coding: utf-8 -*-
"""真实数据 · 消融对照：以「人工核对全对」的成品为金标准，算原方法准确率。

金标准 = 04-试跑产物（当前方法：叙述体全量召回 + 整表抽取 + 动态补行），
         已由使用者人工核对，答案全对。
对照   = 05-对照-原方法（关闭三项改动，回到改动前的逻辑）

三种口径，含义完全不同，必须分开报：
  ① 严格逐格：第 r 行第 c 列 vs 金标准第 r 行第 c 列（表能不能直接用）
  ② 行内自洽：拿这一行的城市名去金标准查真实值，看同行的数据列对不对
                （严格=② 当且仅当行顺序也一致）
  ③ 值召回  ：抽出来的数值里有几个是金标准里真实存在的（忽略位置/归属）
                衡量"读没读对数字"，不衡量"填没填对地方"

分母统一用 100 行 × 5 列 = 500 格（金标准规模）。
"""
import os
from pathlib import Path

from openpyxl import load_workbook

BASE = Path(r"C:\Users\覃敬涛\Desktop\研一项目")
GOLD_XLSX = BASE / "测试数据" / "真实数据" / "04-试跑产物" / "2025年中国城市经济百强全景报告-模板-已填写.xlsx"
# 对照目录可用 A23_CMP_DIR 指定（默认 05-对照-原方法）
CMP_DIR = Path(os.environ.get("A23_CMP_DIR")
               or str(BASE / "测试数据" / "真实数据" / "05-对照-原方法"))
CMP_XLSX = CMP_DIR / "2025年中国城市经济百强全景报告-模板-已填写.xlsx"
OUT = Path(__file__).resolve().parent / (os.environ.get("A23_CMP_OUT") or "_原方法对照.txt")

COLS = ["城市名", "GDP总量（亿元）", "常住人口（万）", "人均GDP（元）", "一般公共预算收入（亿元）"]


def norm(v):
    """数字按数值比（2 位小数容差靠字符串化），文本按去空白比。"""
    if v is None:
        return ""
    if isinstance(v, bool):
        return str(v)
    if isinstance(v, (int, float)):
        return "%.2f" % round(float(v), 2)
    s = str(v).strip().replace(" ", "").replace(",", "").replace("，", "")
    try:
        return "%.2f" % round(float(s), 2)
    except Exception:
        return s


def read_table(path):
    """返回 [(row_no, {列名: 值})]，跳过完全空行。"""
    ws = load_workbook(path, data_only=True).worksheets[0]
    rows = []
    for r in range(2, ws.max_row + 1):
        rec = {}
        for ci, name in enumerate(COLS, start=1):
            v = ws.cell(row=r, column=ci).value
            if v is not None and str(v).strip() != "":
                rec[name] = v
        if rec:
            rows.append((r, rec))
    return rows


def main():
    out = []
    def say(s=""):
        out.append(s)
        print(s)

    gold_rows = read_table(GOLD_XLSX)
    cmp_rows = read_table(CMP_XLSX)

    say("=" * 66)
    say("真实数据 · 消融对照（金标准 = 04-试跑产物，已人工核对全对）")
    say("=" * 66)
    say("金标准：%d 行" % len(gold_rows))
    say("对照版：%d 行（原方法：候选按相关度截断 15 块 + 逐列抽取）" % len(cmp_rows))
    say("")

    # 金标准按城市建索引
    gold_by_city = {}
    for _r, rec in gold_rows:
        c = norm(rec.get("城市名"))
        if c:
            gold_by_city[c] = rec

    # 金标准里的"真实值集合"，用于口径③
    gold_values = set()
    for _r, rec in gold_rows:
        for name in COLS[1:]:
            v = norm(rec.get(name))
            if v:
                gold_values.add(v)

    total_cells = len(gold_rows) * len(COLS)

    # ---------- ① 严格逐格 ----------
    strict_ok = 0
    strict_by_col = {c: [0, 0] for c in COLS}
    strict_bad_examples = []
    for i, (_r, grec) in enumerate(gold_rows):
        # 对照版的第 i 行（数据行顺序对应，都是从第 2 行开始往下写的）
        crec = cmp_rows[i][1] if i < len(cmp_rows) else {}
        for name in COLS:
            gv = norm(grec.get(name))
            cv = norm(crec.get(name))
            strict_by_col[name][1] += 1
            if gv and gv == cv:
                strict_ok += 1
                strict_by_col[name][0] += 1
            elif gv and len(strict_bad_examples) < 6 and cv:
                strict_bad_examples.append(
                    "第 %d 行 %s：金标准=%s ｜ 原方法=%s"
                    % (i + 2, name, gv, cv))

    # ---------- ② 行内自洽（用城市名反查） ----------
    consist_ok = 0
    consist_den = 0
    consist_by_col = {c: [0, 0] for c in COLS[1:]}
    orphan_rows = []          # 有数据但查不到城市名
    unknown_city = []         # 城市名不在金标准里
    row_mismatch = []         # 城市名在，但同行数据与金标准不符
    for r, rec in cmp_rows:
        c = norm(rec.get("城市名"))
        data_cols = [n for n in COLS[1:] if norm(rec.get(n))]
        if not c:
            if data_cols:
                orphan_rows.append((r, len(data_cols)))
            continue
        if c not in gold_by_city:
            unknown_city.append((r, c))
            continue
        grec = gold_by_city[c]
        bad = []
        for name in data_cols:
            consist_den += 1
            consist_by_col[name][1] += 1
            if norm(grec.get(name)) == norm(rec.get(name)):
                consist_ok += 1
                consist_by_col[name][0] += 1
            else:
                bad.append("%s(对=%s/实=%s)" % (name, norm(grec.get(name)), norm(rec.get(name))))
        if bad and len(row_mismatch) < 8:
            row_mismatch.append("第 %d 行 %s：%s" % (r, c, "，".join(bad)))

    # ---------- ③ 值召回 ----------
    recall_ok = 0
    recall_den = 0
    for _r, rec in cmp_rows:
        for name in COLS[1:]:
            v = norm(rec.get(name))
            if not v:
                continue
            recall_den += 1
            if v in gold_values:
                recall_ok += 1

    # ---------- 报告 ----------
    say("【口径①严格逐格】第 r 行对第 r 行，直接比")
    say("  正确 %d / %d 格 = %.1f%%" % (strict_ok, total_cells, 100.0 * strict_ok / total_cells))
    for name in COLS:
        ok, den = strict_by_col[name]
        say("    %-22s %3d/%3d = %5.1f%%" % (name, ok, den, 100.0 * ok / den if den else 0))
    say("")

    say("【口径②行内自洽】用这一行的城市名反查金标准，看数据列对不对")
    say("  （表能不能用，看这个数；分母 = 原方法实际写出的数据格数）")
    if consist_den:
        say("  正确 %d / %d 格 = %.1f%%" % (consist_ok, consist_den,
                                          100.0 * consist_ok / consist_den))
        for name in COLS[1:]:
            ok, den = consist_by_col[name]
            if den:
                say("    %-22s %3d/%3d = %5.1f%%" % (name, ok, den, 100.0 * ok / den))
    else:
        say("  无可判定格")
    say("")

    say("【口径③值召回】抽出的数值里有几个是原文真实值（不管填在哪一行）")
    if recall_den:
        say("  %d / %d = %.1f%%" % (recall_ok, recall_den,
                                  100.0 * recall_ok / recall_den))
    say("")

    say("【记录覆盖】")
    say("  金标准 %d 条城市记录 ｜ 原方法产出 %d 条（含 %d 行只有数据无城市名）"
        % (len(gold_rows), len(cmp_rows), len(orphan_rows)))
    say("  覆盖率 %.1f%%" % (100.0 * len([1 for _r, rec in cmp_rows if norm(rec.get('城市名'))])
                           / len(gold_rows)))
    say("")

    if strict_bad_examples:
        say("【严格口径下的错格样例（同一行的值，来自不同城市）】")
        for s in strict_bad_examples:
            say("  " + s)
        say("")
    if row_mismatch:
        say("【行内自洽口径下对不上的行】")
        for s in row_mismatch:
            say("  " + s)
        say("")
    if unknown_city:
        say("【城市名不在金标准里的行】")
        for r, c in unknown_city:
            say("  第 %d 行 %s" % (r, c))
        say("")
    if orphan_rows:
        say("【只有数据、没有城市名的行（无法归属）】")
        for r, n in orphan_rows:
            say("  第 %d 行：%d 个数据格" % (r, n))
        say("")

    OUT.write_text("\n".join(out), encoding="utf-8")


if __name__ == "__main__":
    main()
