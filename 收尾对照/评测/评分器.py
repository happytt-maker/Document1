# -*- coding: utf-8 -*-
"""
准确率评分器 —— 规则定义见同目录「口径规范.md」，两边必须一致。

用法：
    python 收尾对照/评测/评分器.py
    python 收尾对照/评测/评分器.py --table 模板1
    python 收尾对照/评测/评分器.py --verbose

输入：
    标准答案   收尾对照/评测/标准答案/*.json
    已填写     五表连测/结果-模板N/*-已填写.xlsx

输出：
    收尾对照/评测/评分报告.txt  +  屏幕
"""
import argparse
import json
import os
import re
import sys
import unicodedata
from datetime import date
from pathlib import Path

from openpyxl import load_workbook

HERE = Path(__file__).resolve().parent
BASE = HERE.parents[1]                       # 研一项目
ANS_DIR = HERE / "标准答案"
OUT = HERE / "评分报告.txt"
# 成品去哪找：默认五表连测/；并发压测时用 A23_FILL_ROOT 指到那一批的产物根目录
# （目录结构保持一致：结果-模板N/xxx-已填写.xlsx）
FILL_ROOT = Path(os.environ.get("A23_FILL_ROOT") or (BASE / "五表连测"))

UNIT_RE = re.compile(r"[（(]\s*(元|万元|天|台|件|个|人|小时|分钟|米|吨|公斤|kg|%)\s*[)）]")
DATE_RE = re.compile(r"^(\d{4})[年\-/.](\d{1,2})[月\-/.](\d{1,2})日?$")
NUM_RE = re.compile(r"^-?\d+(?:\.\d+)?$")
TRUE_SET = {"是", "y", "yes", "√", "✓", "对", "有", "1", "true"}
FALSE_SET = {"否", "n", "no", "×", "✗", "错", "无", "0", "false"}


# ---------------------------------------------------------------- 归一化
def norm(v):
    """规范第 2 节 N1~N7。"""
    if v is None:
        return ""
    s = str(v)
    s = unicodedata.normalize("NFKC", s)      # N2 全角→半角；顺带统一中英标点
    s = s.strip().strip("「」『』\"'“”‘’")      # N5 去包裹符
    s = re.sub(r"\s+", " ", s)                # N1 压空白
    s = s.replace(" ", "")
    s = re.sub(r"(?<=\d),(?=\d{3})", "", s)   # N4 千分位
    s = s.rstrip("。；;，,、.")                # N7 去句末标点（断句习惯，不是内容）
    return s.lower()                          # N6 英文小写


def to_num(v):
    s = norm(v)
    if not s:
        return None
    try:
        f = float(s)
    except ValueError:
        return None
    return f


def to_num_strip_unit(v):
    """剥掉末尾单位后再试数字（仅用于 L2）。"""
    s = norm(v)
    s = re.sub(r"(元|万元|天|台|件|个|人|小时|分钟|米|吨|公斤|kg|%)$", "", s)
    try:
        return float(s)
    except ValueError:
        return None


def to_date(v):
    m = DATE_RE.match(norm(v))
    if not m:
        return None
    y, mo, d = (int(x) for x in m.groups())
    try:
        return date(y, mo, d)
    except ValueError:
        return None


def to_bool(v):
    s = norm(v)
    if s in TRUE_SET:
        return True
    if s in FALSE_SET:
        return False
    return None


def lev(a, b):
    if a == b:
        return 0
    if not a or not b:
        return max(len(a), len(b))
    prev = list(range(len(b) + 1))
    for i, ca in enumerate(a, 1):
        cur = [i]
        for j, cb in enumerate(b, 1):
            cur.append(min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + (ca != cb)))
        prev = cur
    return prev[-1]


# ---------------------------------------------------------------- 判定
def judge(expect, actual, column):
    """
    返回 (严格命中, 宽松命中, 标签)
    标签：一致 / 数值等价 / 单位差异 / 日期等价 / 布尔等价 / 简称 / 近似 / 不一致 / 漏填
    """
    e, a = norm(expect), norm(actual)
    if e == a:                                   # 含"两边都空"
        return True, True, "一致"
    if not a:
        return False, False, "漏填"

    ne, na = to_num(expect), to_num(actual)
    if ne is not None and na is not None and abs(ne - na) < 1e-6:
        return False, True, "数值等价"                        # L1

    if UNIT_RE.search(column):                               # L2 仅列名带单位时
        se, sa = to_num_strip_unit(expect), to_num_strip_unit(actual)
        if se is not None and sa is not None and abs(se - sa) < 1e-6:
            return False, True, "单位差异"

    de, da = to_date(expect), to_date(actual)                # L3
    if de and da and de == da:
        return False, True, "日期等价"

    be, ba = to_bool(expect), to_bool(actual)                # L4
    if be is not None and be == ba:
        return False, True, "布尔等价"

    short, long = sorted([e, a], key=len)                    # L5
    if len(short) >= 4 and short in long:
        return False, True, "简称"

    L = max(len(e), len(a))                                  # L6
    if L >= 6 and lev(e, a) <= L * 0.10:
        return False, True, "近似"

    return False, False, "不一致"


# ---------------------------------------------------------------- 读表
def cell_text(ws, row, col):
    v = ws.cell(row=row, column=col).value
    if v is None:
        return ""
    if isinstance(v, float) and v == int(v):
        return str(int(v))
    return str(v).strip()


def score_table(ans_path, verbose=False):
    ans = json.loads(ans_path.read_text(encoding="utf-8"))
    name = ans["template"]
    short = name.split("-")[0]                     # 模板1

    cands = sorted(FILL_ROOT.glob("结果-%s/*已填写.xlsx" % short))
    if not cands:
        return None, "找不到「%s」的已填写文件（先跑五表连测）" % short
    xlsx = cands[0]

    wb = load_workbook(xlsx, data_only=True)
    ws = wb[ans["sheet"]] if ans.get("sheet") in wb.sheetnames else wb.active

    cols = ans["columns"]
    header_row = ans["header_row"]
    fill_rows = list(ans["fill_rows"])
    blank_rows = [r for r in range(ans["fill_start"], ans["fill_end"] + 1)
                  if r not in fill_rows]
    subjective = set(ans.get("subjective", []))

    # 表头核对（结构维度）
    head_bad = []
    for i, c in enumerate(cols, 1):
        got = norm(cell_text(ws, header_row, i))
        if got and norm(c) not in got and got not in norm(c):
            head_bad.append("%s→%s" % (c, cell_text(ws, header_row, i)))

    stat = {
        "template": name, "xlsx": xlsx.name, "columns": cols,
        "fill_cells": 0, "blank_cells": 0,
        "strict": 0, "loose": 0, "miss": 0, "wrong": 0, "blank_ok": 0,
        "tags": {}, "detail": [], "subjective_cells": 0, "subjective_ok": 0,
        "head_bad": head_bad, "overflow": 0,
    }

    for i, r in enumerate(fill_rows):
        rec = ans["rows"][i] if i < len(ans["rows"]) else {}
        for j, c in enumerate(cols, 1):
            actual = cell_text(ws, r, j)

            if c in subjective:
                stat["subjective_cells"] += 1
                if actual:
                    stat["subjective_ok"] += 1
                continue

            expect = rec.get(c, "")
            stat["fill_cells"] += 1
            s_ok, l_ok, tag = judge(expect, actual, c)
            if s_ok:
                stat["strict"] += 1
                stat["loose"] += 1
            elif l_ok:
                stat["loose"] += 1
                stat["tags"][tag] = stat["tags"].get(tag, 0) + 1
            elif tag == "漏填":
                stat["miss"] += 1
            elif not expect:
                # 答案也是空 → 填了东西就是错填
                stat["wrong"] += 1
                tag = "错填"
            stat["tags"][tag] = stat["tags"].get(tag, 0) + 1

            if verbose and not s_ok:
                stat["detail"].append(
                    "    行%-3d %-12s 期望「%s」 实际「%s」  %s"
                    % (r, c, expect, actual, tag))

    for r in blank_rows:                            # 应留空的行
        for j, c in enumerate(cols, 1):
            if c in subjective:
                continue
            stat["blank_cells"] += 1
            if cell_text(ws, r, j):
                stat["wrong"] += 1
                stat["tags"]["错填"] = stat["tags"].get("错填", 0) + 1
                if verbose:
                    stat["detail"].append(
                        "    行%-3d %-12s 应留空 实际「%s」  错填"
                        % (r, c, cell_text(ws, r, j)))
            else:
                stat["blank_ok"] += 1                   # 留空留对了，也是分

    total = stat["fill_cells"] + stat["blank_cells"]
    stat["total"] = total
    stat["acc_strict"] = (stat["strict"] + stat["blank_ok"]) / total if total else 0
    stat["acc_loose"] = (stat["loose"] + stat["blank_ok"]) / total if total else 0
    stat["acc_fill"] = stat["strict"] / stat["fill_cells"] if stat["fill_cells"] else 0
    stat["pass80"] = stat["acc_strict"] >= 0.80
    return stat, None


# ---------------------------------------------------------------- 报告
def render(stats, errs):
    L = []
    L.append("=" * 64)
    L.append("准确率评分报告（口径见 评测/口径规范.md v1）")
    L.append("=" * 64)
    L.append("")
    L.append("口径说明：严格=归一化后逐字符相等（下界）｜宽松=再承认数值/日期/单位/简称/近似（上界）")
    L.append("          整体准确率 = (严格命中 + 正确留空) / 表格总格数")
    L.append("")

    for s in stats:
        L.append("-" * 64)
        L.append("%s  [%s]" % (s["template"], s["xlsx"]))
        L.append("-" * 64)
        if s["head_bad"]:
            L.append("  结构：表头有 %d 处对不上 → %s" % (len(s["head_bad"]), "、".join(s["head_bad"])))
        else:
            L.append("  结构：表头 %d 列全部对上" % len(s["columns"]))
        L.append("  分母：应填 %d 格 + 应留空 %d 格 = %d 格"
                 % (s["fill_cells"], s["blank_cells"], s["total"]))
        L.append("  严格命中 %d ｜ 宽松命中 %d ｜ 正确留空 %d ｜ 漏填 %d ｜ 错填 %d"
                 % (s["strict"], s["loose"], s["blank_ok"], s["miss"], s["wrong"]))
        L.append("  整体准确率：严格 %.1f%% ～ 宽松 %.1f%%   %s"
                 % (s["acc_strict"] * 100, s["acc_loose"] * 100,
                    "达标(≥80%)" if s["pass80"] else "★未达标"))
        L.append("  应填格准确率（诊断）：%.1f%%" % (s["acc_fill"] * 100))
        if s["subjective_cells"]:
            L.append("  主观列（不计分）：%d 格有值 / %d 格"
                     % (s["subjective_ok"], s["subjective_cells"]))
        tags = {k: v for k, v in s["tags"].items() if k != "一致"}
        if tags:
            L.append("  未严格命中的原因：" + "、".join("%s %d" % (k, v) for k, v in tags.items()))
        if s["detail"]:
            L.append("  逐格明细：")
            L.extend(s["detail"])
        L.append("")

    if stats:
        n = len(stats)
        a_s = sum(s["acc_strict"] for s in stats) / n
        a_l = sum(s["acc_loose"] for s in stats) / n
        L.append("=" * 64)
        L.append("五表平均：严格 %.1f%% ～ 宽松 %.1f%%   达标表数 %d/%d"
                 % (a_s * 100, a_l * 100, sum(1 for s in stats if s["pass80"]), n))
        L.append("=" * 64)
    for e in errs:
        L.append("  ⚠ " + e)
    return "\n".join(L)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--table", default=None, help="只评某张表，如 模板1")
    ap.add_argument("--verbose", action="store_true", help="打印逐格明细")
    args = ap.parse_args()

    files = sorted(ANS_DIR.glob("*.json"))
    if not files:
        print("标准答案目录是空的：%s" % ANS_DIR)
        return 2
    if args.table:
        files = [f for f in files if args.table in f.name]

    stats, errs = [], []
    for f in files:
        s, err = score_table(f, verbose=args.verbose)
        if err:
            errs.append(err)
        else:
            stats.append(s)

    text = render(stats, errs)
    OUT.write_text(text, encoding="utf-8")
    try:
        print(text)
    except UnicodeEncodeError:
        sys.stdout.buffer.write(text.encode("utf-8", "replace"))
    print("\n（已写入 %s）" % OUT)
    return 0


if __name__ == "__main__":
    sys.exit(main())
