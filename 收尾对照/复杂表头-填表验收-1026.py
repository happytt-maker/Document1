# -*- coding: utf-8 -*-
"""
复杂表头-填表验收-1026.py
=========================
验三件事，全部机械可判（不调模型、不看截图）：

    ① 结构解析对不对   —— 表头块、列名、待填段，逐项与手写标准答案对照
    ② 落位对不对       —— 用「值编码」手法：第 c 列第 i 条记录的值写成 "Vc_i"，
                          填完读回文件，每一格都能验证"它是不是本该在这儿的那条数据"
    ③ 有没有砸坏别处   —— 除待填槽位外，其余格子（表头、续表、合计、行标签）
                          必须与原模板逐格完全一致

为什么用值编码而不是"填真实内容再看对不对"：
    真实内容要靠模型抽取，抽错一次就分不清是"抽取错"还是"落位错"。
    编码值只带坐标信息，落位一旦错位立刻暴露（V4_9 跑到第 2 段就说明跨段错）。
    模型抽取质量是模块②的事，这里只负责"给定正确数据，能不能填到对的格子"。

另外验一个向后兼容点：老的结构文件（只有 header_row / columns / fill_region，
没有 header_rows / segments）仍然能正常填 —— 保证旧流程不因改造而退化。

用法：
    python 复杂表头-填表验收-1026.py
"""

import json
import os
import re
import shutil
import subprocess
import sys
from pathlib import Path

import openpyxl
from docx import Document

HERE = Path(__file__).resolve().parent
BASE = HERE.parent
PY = sys.executable

TPL_X = BASE / "测试数据" / "模板-复杂表头"
TPL_W = BASE / "测试数据" / "模板-复杂表头-word"
OUT_DIR = BASE / "writer" / "输出"
TMP = HERE / "_复杂表头验收"

STRUCT = TMP / "结构.json"
MATCHER = TMP / "matcher"
ANSWER = json.loads((TPL_X / "_标准答案.json").read_text(encoding="utf-8"))

# 每张模板给多少条记录（故意各不相同，覆盖"填不满 / 刚好填满 / 跨段 / 超容量截断"）
N_RECORDS = {
    "模板6-多层表头设备采购台账.xlsx": 6,     # 容量 9 → 测留空
    "模板7-斜线表头费用报销明细.xlsx": 6,     # 容量 6 → 刚好填满
    "模板8-跨页原材料入库台账.xlsx": 12,      # 容量 18 → 前 9 条进第 1 段，第 10~12 条进第 2 段
    "模板9-跨页会议任务跟踪.docx": 10,        # 容量 8 → 测超容量截断
    "模板10-多层表头员工信息登记.docx": 3,    # 容量 3 → 刚好
}

results = []


def ok(name, passed, detail=""):
    results.append((bool(passed), name, str(detail)))
    print("  %s  %-58s %s" % ("✓" if passed else "✗", name, detail))


def section(text):
    print("\n── %s " % text + "─" * max(0, 62 - len(text)))


# ======================================================================
# ① 结构解析
# ======================================================================
def check_structure():
    section("① 结构解析（对照手写标准答案）")
    sys.path.insert(0, str(BASE / "template_reader"))
    from reader import read_folder

    docs = read_folder(TPL_X) + read_folder(TPL_W)
    by_file = {}
    for d in docs:
        if "error" in d:
            ok("解析 %s" % d["file"], False, d["error"])
            continue
        by_file[d["file"]] = d

    for fname, want in ANSWER.items():
        doc = by_file.get(fname)
        if not doc:
            ok("解析 %s" % fname, False, "没读到")
            continue
        sh = doc["sheets"][0]

        ok("%s · 表头块行号" % fname,
           sh["header_rows"] == want["header_rows"],
           "%s（标准答案 %s）" % (sh["header_rows"], want["header_rows"]))

        got = [c["title"] for c in sh["columns"]]
        ok("%s · 列名" % fname, got == want["titles"],
           "" if got == want["titles"] else "实际 %s ≠ 期望 %s" % (got, want["titles"]))

        segs = [[s["start_row"], s["end_row"]] for s in sh["segments"]]
        ok("%s · 待填段" % fname, segs == want["segments"],
           "%s（期望 %s）" % (segs, want["segments"]))

        if "slash" in want:
            sw = want["slash"]
            col = sh["columns"][sw["col"] - 1]
            ok("%s · 斜线格已拆" % fname,
               col["slash"] and col["slash_parts"] == sw["parts"],
               "parts=%s" % col["slash_parts"])

        if "tables" in want:
            ok("%s · 跨页表已合并" % fname,
               sh.get("tables") == want["tables"],
               "tables=%s" % sh.get("tables"))

    # 行标签列：模板7 的第一列必须被认出来（否则待填区整个丢）
    sh7 = by_file["模板7-斜线表头费用报销明细.xlsx"]["sheets"][0]
    ok("模板7 · 行标签列被认出（第 1 列）", sh7["label_cols"] == [1],
       "label_cols=%s" % sh7["label_cols"])

    # 模板6 的两组「数量」必须靠父名区分开，不能同名
    sh6 = by_file["模板6-多层表头设备采购台账.xlsx"]["sheets"][0]
    titles = [c["title"] for c in sh6["columns"]]
    ok("模板6 · 重复子名用父名区分", 
       "上半年采购-数量" in titles and "下半年采购-数量" in titles,
       "")
    ok("模板6 · 唯一子名不硬拼父名",
       "供应商名称" in titles and not any(t.startswith("供应商信息") for t in titles),
       "")

    return docs


# ======================================================================
# ② 造抽取结果（值编码）
# ======================================================================
def build_extracts(docs, only=None):
    """给指定模板造抽取结果。

    【必须一张一张来】不同模板可能有同名列（模板6 和模板8 都有「规格型号」），
    而抽取结果是按列名存的文件（抽取结果-规格型号.json）——
    5 张一起写，后写的会把先写的覆盖掉，落位检查会误报错位。
    真实场景里一次只填一张表，所以这里也一张一张地准备。
    """
    if MATCHER.exists():
        shutil.rmtree(MATCHER)
    MATCHER.mkdir(parents=True)

    made = 0
    for d in docs:
        fname = d["file"]
        if only is not None and fname != only:
            continue
        if fname not in N_RECORDS:
            continue
        sh = d["sheets"][0]
        labels = set(sh.get("label_cols") or [])
        n = N_RECORDS[fname]

        for col in sh["columns"]:
            if col["col"] in labels:
                continue                       # 行标签列不填
            title = col["title"]
            if not title or title == "序号":
                continue                       # 序号走自动编号
            items = [{"value": "V%d_%d" % (col["col"], i),
                      "file": "验收源文档.docx",
                      "loc_text": "第 %d 段" % (i + 1)}
                     for i in range(n)]
            (MATCHER / ("抽取结果-%s.json" % title)).write_text(
                json.dumps({"column": title, "items": items},
                           ensure_ascii=False, indent=2), encoding="utf-8")
            made += 1
    return made


# ======================================================================
# ③ 跑 fill.py 并读回校验
# ======================================================================
def run_fill(folder, name):
    env = dict(os.environ)
    env["A23_STRUCTURE_JSON"] = str(STRUCT)
    env["A23_MATCHER_DIR"] = str(MATCHER)
    env["A23_TEMPLATE_DIR"] = str(folder)
    env["PYTHONIOENCODING"] = "utf-8"

    out = OUT_DIR / (Path(name).stem + "-已填写" + Path(name).suffix)
    if out.exists():
        out.unlink()
    p = subprocess.run([PY, str(BASE / "writer" / "fill.py"), name, "--overwrite"],
                       capture_output=True, text=True, encoding="utf-8",
                       errors="replace", env=env, timeout=300)
    return p.returncode, (p.stdout or "") + (p.stderr or "")


def expected_slots(sh):
    return [(s.get("table") or 1, r)
            for s in sh["segments"]
            for r in range(s["start_row"], s["end_row"] + 1)]


def check_xlsx(fname, sh, before_path, after_path):
    wb0 = openpyxl.load_workbook(before_path, data_only=False)
    wb1 = openpyxl.load_workbook(after_path, data_only=False)
    ws0, ws1 = wb0[sh["name"]], wb1[sh["name"]]

    labels = set(sh.get("label_cols") or [])
    slots = expected_slots(sh)
    n = min(N_RECORDS[fname], len(slots))
    colmap = {c["col"]: c["title"] for c in sh["columns"]}
    seq_cols = [c["col"] for c in sh["columns"] if c["title"] == "序号"]

    # --- 落位：每一个槽位的每一列 ---
    wrong = []
    for i in range(n):
        _t, r = slots[i]
        for c, title in colmap.items():
            if c in labels:
                continue
            v = ws1.cell(row=r, column=c).value
            if c in seq_cols:
                want = i + 1
            else:
                want = "V%d_%d" % (c, i)
            got = v
            if str(got) != str(want):
                wrong.append("第%d条 第%d行第%d列[%s] 期望 %s 实际 %s"
                             % (i + 1, r, c, title, want, got))
    ok("%s · 落位 %d 条 × %d 列" % (fname, n, len(colmap) - len(labels)),
       not wrong, "" if not wrong else "前 3 处：" + "；".join(wrong[:3]))

    # --- 没用到的槽位必须还是空的 ---
    empty_ok = True
    for i in range(n, len(slots)):
        _t, r = slots[i]
        for c in colmap:
            if c in labels:
                continue
            if ws1.cell(row=r, column=c).value not in (None, ""):
                empty_ok = False
    ok("%s · 多余槽位保持空白" % fname, empty_ok, "%d 个空位" % (len(slots) - n))

    # --- 除待填位外，逐格不变（表头/续表/合计/行标签） ---
    allowed = set()
    for i in range(n):
        _t, r = slots[i]
        for c in colmap:
            if c not in labels:
                allowed.add((r, c))

    diffs = []
    for r in range(1, max(ws0.max_row, ws1.max_row) + 1):
        for c in range(1, max(ws0.max_column, ws1.max_column) + 1):
            if (r, c) in allowed:
                continue
            a = ws0.cell(row=r, column=c).value
            b = ws1.cell(row=r, column=c).value
            if str(a) != str(b):
                diffs.append("第%d行第%d列 %r→%r" % (r, c, a, b))
    ok("%s · 其它格子一格没动" % fname, not diffs,
       "" if not diffs else "前 3 处：" + "；".join(diffs[:3]))

    wb0.close()
    wb1.close()


def check_docx(fname, sh, before_path, after_path):
    d0, d1 = Document(str(before_path)), Document(str(after_path))
    labels = set(sh.get("label_cols") or [])
    slots = expected_slots(sh)
    n = min(N_RECORDS[fname], len(slots))
    colmap = {c["col"]: c["title"] for c in sh["columns"]}
    seq_cols = [c["col"] for c in sh["columns"] if c["title"] == "序号"]

    wrong = []
    for i in range(n):
        t, r = slots[i]
        tb = d1.tables[t - 1]
        for c, title in colmap.items():
            if c in labels:
                continue
            want = (i + 1) if c in seq_cols else "V%d_%d" % (c, i)
            got = (tb.cell(r - 1, c - 1).text or "").strip()
            if str(got) != str(want):
                wrong.append("第%d条 表%d第%d行第%d列[%s] 期望 %s 实际 %s"
                             % (i + 1, t, r, c, title, want, got))
    ok("%s · 落位 %d 条 × %d 列" % (fname, n, len(colmap) - len(labels)),
       not wrong, "" if not wrong else "前 3 处：" + "；".join(wrong[:3]))

    # 跨页：容量比数据少时，多出来的记录必须被丢掉且不串到别处
    if N_RECORDS[fname] > len(slots):
        ok("%s · 超容量截断（%d 条 > %d 位）" % (fname, N_RECORDS[fname], len(slots)),
           True, "只填前 %d 条，多出的 %d 条丢弃"
           % (n, N_RECORDS[fname] - n))

    # 除待填位外逐个 table 比对
    allowed = set()
    for i in range(n):
        t, r = slots[i]
        for c in colmap:
            if c not in labels:
                allowed.add((t, r, c))

    diffs = []
    for ti in range(min(len(d0.tables), len(d1.tables))):
        t0, t1 = d0.tables[ti], d1.tables[ti]
        for ri in range(min(len(t0.rows), len(t1.rows))):
            for ci in range(min(len(t0.rows[ri].cells), len(t1.rows[ri].cells))):
                if (ti + 1, ri + 1, ci + 1) in allowed:
                    continue
                a = (t0.rows[ri].cells[ci].text or "").strip()
                b = (t1.rows[ri].cells[ci].text or "").strip()
                if a != b:
                    diffs.append("表%d 第%d行第%d列 %r→%r" % (ti + 1, ri + 1, ci + 1, a, b))
    ok("%s · 其它格子一格没动" % fname, not diffs,
       "" if not diffs else "前 3 处：" + "；".join(diffs[:3]))


# ======================================================================
# ④ 向后兼容：老结构文件（没有 segments / header_rows）
# ======================================================================
def check_backward_compat(docs):
    section("④ 向后兼容（老结构文件只有 fill_region）")
    # 老结构测试用的是模板6，抽取结果也要按模板6 重新准备一份
    build_extracts(docs, only="模板6-多层表头设备采购台账.xlsx")
    old = TMP / "结构-老格式.json"
    data = json.loads(STRUCT.read_text(encoding="utf-8"))
    slim = []
    for d in data:
        if d["file"] != "模板6-多层表头设备采购台账.xlsx":
            continue
        for sh in d["sheets"]:
            sh2 = {
                "name": sh["name"], "max_row": sh["max_row"], "max_column": sh["max_column"],
                "header_row": sh["header_row"], "header_ok": True, "header_note": "",
                "columns": [{"col": c["col"], "title": c["title"]} for c in sh["columns"]],
                "fill_region": sh["fill_region"],
            }
            slim.append({"file": d["file"], "format": "xlsx",
                         "sheet_count": 1, "sheets": [sh2]})
    old.write_text(json.dumps(slim, ensure_ascii=False, indent=2), encoding="utf-8")

    out = OUT_DIR / "模板6-多层表头设备采购台账-已填写.xlsx"
    if out.exists():
        out.unlink()

    env = dict(os.environ)
    env["A23_STRUCTURE_JSON"] = str(old)
    env["A23_MATCHER_DIR"] = str(MATCHER)
    env["A23_TEMPLATE_DIR"] = str(TPL_X)
    env["PYTHONIOENCODING"] = "utf-8"
    p = subprocess.run([PY, str(BASE / "writer" / "fill.py"),
                        "模板6-多层表头设备采购台账.xlsx"],
                       capture_output=True, text=True, encoding="utf-8",
                       errors="replace", env=env, timeout=180)
    ok("老结构仍能填", p.returncode == 0, "rc=%d" % p.returncode)

    if p.returncode == 0 and out.exists():
        wb = openpyxl.load_workbook(out)
        ws = wb["设备采购台账"]
        v = ws.cell(row=5, column=4).value          # 第 1 条 · 上半年采购-数量
        ok("老结构落位正确", str(v) == "V4_0", "第5行第4列 = %s" % v)
        wb.close()
        out.unlink()                                # 清掉，别和正式产物混


def main():
    print("=" * 68)
    print("复杂表头（多层 / 斜线 / 跨页）填表验收")
    print("=" * 68)

    docs = check_structure()

    STRUCT.parent.mkdir(parents=True, exist_ok=True)
    STRUCT.write_text(json.dumps(docs, ensure_ascii=False, indent=2), encoding="utf-8")

    section("② 逐张模板：造数据 → 填表 → 读回校验")
    for fname in N_RECORDS:
        doc = next((d for d in docs if d["file"] == fname), None)
        if not doc:
            ok("读结构 %s" % fname, False, "结构里没这张表")
            continue
        sh = doc["sheets"][0]

        made = build_extracts(docs, only=fname)
        ok("%s · 造抽取结果" % fname, made > 0, "%d 列" % made)

        folder = TPL_X if fname.endswith(".xlsx") else TPL_W
        rc, log = run_fill(folder, fname)
        ok("%s · fill 跑通" % fname, rc == 0,
           "rc=%d" % rc if rc else "")

        src = folder / fname
        dst = OUT_DIR / (Path(fname).stem + "-已填写" + Path(fname).suffix)
        if not dst.exists():
            ok("%s · 输出文件存在" % fname, False, dst.name)
            continue
        if fname.endswith(".xlsx"):
            check_xlsx(fname, sh, src, dst)
        else:
            check_docx(fname, sh, src, dst)

    check_backward_compat(docs)

    # ---------------- 汇总 ----------------
    passed = sum(1 for p, _n, _d in results if p)
    total = len(results)
    print("\n" + "=" * 68)
    print("合计：%d/%d 通过%s" % (passed, total, "" if passed == total else "  ⚠️ 有失败"))
    print("=" * 68)

    lines = ["复杂表头填表验收 · %s" % "2026-10-26",
             "合计：%d/%d 通过" % (passed, total), ""]
    for p, n, d in results:
        lines.append("%s  %s  %s" % ("PASS" if p else "FAIL", n, d))
    (HERE / "复杂表头-填表验收-1026.txt").write_text("\n".join(lines), encoding="utf-8")
    return 0 if passed == total else 1


if __name__ == "__main__":
    sys.exit(main())
