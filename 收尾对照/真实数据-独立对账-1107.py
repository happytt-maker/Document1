# -*- coding: utf-8 -*-
"""真实数据-独立对账-1107.py —— 不用本系统的任何中间结果，直接从源表反查成品。

思路：把源 xlsx 按「城市 + 站点名称」建索引（只留 2025-11-25 09:00 那一刻），
再逐格核对已填写 docx 的三张表。这样即使抽取/写回整条链路一起错，
也不可能两边同时错成一样 —— 能对上的就是真对上了。

用法：python 真实数据-独立对账-1107.py <已填写.docx> [源.xlsx]
"""
import sys
from pathlib import Path

import docx
from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "导入文档" / "山东省环境空气质量监测数据信息202512171921_0.xlsx"
DOC = Path(sys.argv[1]) if len(sys.argv) > 1 else (
    ROOT / "收尾对照" / "_run1107" / "out" /
    "2025山东省环境空气质量监测数据信息-模板-已填写.docx")

COLS = ["城市", "区", "站点名称", "空气质量指数",
        "PM10监测值", "PM2.5监测值", "首要污染物", "污染类型"]


def norm_time(v):
    s = str(v or "").strip().replace("T", " ")
    return s[:16]          # 2025-11-25 09:00:00.0 → 2025-11-25 09:00


def load_source():
    # 注意：这份 xlsx 没有 dimension 记录，read_only 模式下 iter_rows 只会吐出半行表头
    # （表头被读成 ('序号',)），必须用普通模式读。
    import warnings
    warnings.filterwarnings("ignore")
    wb = load_workbook(SRC, data_only=True)
    ws = wb.worksheets[0]
    rows = ws.iter_rows(values_only=True)
    head = [str(c or "").strip() for c in next(rows)]
    idx = {h: head.index(h) for h in head if h}
    want_time = None
    table = {}
    for r in rows:
        cell = lambda name: r[idx[name]] if name in idx else None
        if want_time is None and norm_time(cell("监测时间")):
            want_time = norm_time(cell("监测时间"))
        if norm_time(cell("监测时间")) != want_time:
            continue
        key = (str(cell("城市") or "").strip(), str(cell("站点名称") or "").strip())
        if not key[1]:
            continue          # 该时刻也有"城市汇总行"没有站点名，跳过
        table[key] = {c: ("" if cell(c) is None else str(cell(c)).strip()) for c in COLS}
    wb.close()
    return table, want_time


def num_eq(a, b):
    try:
        return abs(float(a) - float(b)) < 1e-9
    except (TypeError, ValueError):
        return str(a).strip() == str(b).strip()


def main():
    src, t0 = load_source()
    print("源表：%s" % SRC.name)
    print("取时刻：%s ｜ 该时刻站点数：%d" % (t0, len(src)))

    d = docx.Document(str(DOC))
    print("成品：%s（%d 张表）" % (DOC.name, len(d.tables)))
    ok = bad = missing = 0
    for ti, t in enumerate(d.tables, 1):
        head = [c.text.strip() for c in t.rows[0].cells]
        cols = [head.index(c) if c in head else -1 for c in COLS]
        n_table = 0
        for row in t.rows[1:]:
            vals = [row.cells[i].text.strip() if i >= 0 else "" for i in cols]
            if not any(vals):
                continue
            key = (vals[0], vals[2])
            hit = src.get(key)
            if hit is None:
                missing += 1
                print("  ✗ 表%d %s 在源表该时刻找不到" % (ti, key))
                continue
            n_table += 1
            for c, got in zip(COLS, vals):
                if num_eq(got, hit[c]):
                    ok += 1
                else:
                    bad += 1
                    print("  ✗ 表%d %s · %s：成品「%s」 vs 源表「%s」"
                          % (ti, key, c, got, hit[c]))
        print("  表%d：核对 %d 条" % (ti, n_table))

    print("-" * 56)
    print("一致 %d 格 ｜ 不一致 %d 格 ｜ 源表查无 %d 行" % (ok, bad, missing))
    print("结论：" + ("全对，可以交卷" if bad == 0 and missing == 0 else "有问题，别交"))
    return 0 if bad == 0 and missing == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
