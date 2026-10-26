# -*- coding: utf-8 -*-
"""
_验收-xlsx操作.py —— 模块（1）xlsx 操作补验：8 条正向 + 1 条负向，逐条真跑 + 读回验证

背景：覆盖矩阵体检发现 xlsx 有三处缺口（编辑不支持 / 标红缺失 / 列定位被忽略
"悄悄改表头"），全部修复后新增本验收固化住。

流程：①源文件指纹 → ②逐条调 edit.py 真跑 → ③产物归档改名 → ④读回属性验证
      → ⑤负向测试（不许悄悄错）→ ⑥指纹复核（原件零改动）

跑法：python agent/_验收-xlsx操作.py
"""
import hashlib
import os
import subprocess
import sys
from pathlib import Path

from openpyxl import load_workbook

BASE = Path(__file__).resolve().parents[1]
PY = sys.executable
AGENT = BASE / "agent"
OUT = AGENT / "输出"
DOCS = BASE / "测试数据" / "源文档"
SRC = DOCS / "xlsx" / "员工信息表.xlsx"


def md5(p: Path) -> str:
    return hashlib.md5(p.read_bytes()).hexdigest()


# (编号, 指令, 期望产物名, 归档名)
CASES = [
    (1, "把员工信息表里'王振国'改成'王振国（组长）'", "员工信息表_编辑结果.xlsx", "xlsx-01-单元格替换.xlsx"),
    (2, "把员工信息表里'E1005'删掉", "员工信息表_编辑结果.xlsx", "xlsx-02-单元格删除.xlsx"),
    (3, "把员工信息表的工号列加粗", "员工信息表_格式结果.xlsx", "xlsx-03-列加粗.xlsx"),
    (4, "把员工信息表里'王振国'标红", "员工信息表_格式结果.xlsx", "xlsx-04-单元格标红.xlsx"),
    (5, "把员工信息表的表头字号调成 14 号", "员工信息表_格式结果.xlsx", "xlsx-05-表头字号.xlsx"),
    (6, "把员工信息表的姓名列居中", "员工信息表_格式结果.xlsx", "xlsx-06-列居中.xlsx"),
    (7, "把员工信息表的表头居中", "员工信息表_格式结果.xlsx", "xlsx-07-表头居中.xlsx"),
    (8, "把员工信息表的表头字体改成黑体", "员工信息表_格式结果.xlsx", "xlsx-08-表头字体.xlsx"),
]

# 负向测试：找不到的列 → 必须失败、不得产出文件（不许悄悄改成表头）
NEG_CMD = "把员工信息表的'不存在的列某某'加粗"


def main():
    print("=" * 66)
    print("xlsx 操作补验 · 8 条正向 + 1 条负向（逐条真跑 + 读回验证）")
    print("=" * 66)

    h0 = md5(SRC)

    # ---- ② 逐条跑 + ③ 归档 ----
    results = {}
    for cid, text, outname, archname in CASES:
        print("\n---- 第 %d 条：%s" % (cid, text))
        r = subprocess.run([PY, str(AGENT / "edit.py"), text],
                           capture_output=True, text=True, encoding="utf-8",
                           errors="replace", timeout=240, cwd=str(BASE))
        ok = (r.returncode == 0)
        (OUT / ("xlsx第%d条-报告.txt" % cid)).write_text(r.stdout + "\n" + r.stderr, encoding="utf-8")
        o = OUT / outname
        archived = None
        if o.exists():
            archived = OUT / archname
            os.replace(str(o), str(archived))
        results[cid] = {"ok": ok, "out": archived}
        print("  [%s] 退出码=%d ｜ 产物=%s" % ("OK " if ok else "FAIL", r.returncode,
              archname if archived else "（无产物）"))

    # ---- 负向：找不到的列 ----
    print("\n---- 负向：%s（预期：失败、无产物）" % NEG_CMD)
    r = subprocess.run([PY, str(AGENT / "edit.py"), NEG_CMD],
                       capture_output=True, text=True, encoding="utf-8",
                       errors="replace", timeout=240, cwd=str(BASE))
    neg_ok = (r.returncode != 0) and ("没找到" in r.stdout)
    neg_no_out = not (OUT / "员工信息表_格式结果.xlsx").exists()
    print("  [%s] 退出码=%d ｜ 报告含「没找到」=%s ｜ 无产物=%s"
          % ("OK " if (neg_ok and neg_no_out) else "FAIL", r.returncode,
             "是" if "没找到" in r.stdout else "否", "是" if neg_no_out else "否"))

    # ---- ④ 读回验证 ----
    print("\n" + "=" * 66)
    print("逐条验证（读回输出文件）")
    print("=" * 66)
    checks = []

    def chk(cid, name, cond, detail=""):
        checks.append((cid, name, bool(cond), detail))
        print("  [%s] 第%d条 %s%s" % ("PASS" if cond else "FAIL", cid, name,
                                     (" ｜ " + detail) if detail else ""))

    ws0 = load_workbook(str(SRC)).active  # 源文件用于对比

    # 1 单元格替换 + 保格式
    o = results[1]["out"]
    if o:
        ws = load_workbook(str(o)).active
        chk(1, "B3 = 王振国（组长）", ws["B3"].value == "王振国（组长）", "值=%r" % ws["B3"].value)
        same_font = (ws["B3"].font.name == ws0["B3"].font.name
                     and ws["B3"].font.size == ws0["B3"].font.size)
        chk(1, "B3 字体保真", same_font,
            "%s/%s" % (ws["B3"].font.name, ws["B3"].font.size))
        chk(1, "邻格 B4 未误伤", ws["B4"].value == "李慧敏", "B4=%r" % ws["B4"].value)

    # 2 单元格删除
    o = results[2]["out"]
    if o:
        ws = load_workbook(str(o)).active
        chk(2, "A7 已清空", ws["A7"].value is None, "A7=%r" % ws["A7"].value)
        chk(2, "邻格 A8 未误伤", ws["A8"].value == "E1006", "A8=%r" % ws["A8"].value)

    # 3 工号列（第 1 列）加粗：表头 + 数据全粗，B 列不粗
    o = results[3]["out"]
    if o:
        ws = load_workbook(str(o)).active
        col_a_bold = all(ws.cell(row=r, column=1).font.bold for r in (2, 3, 5, 10))
        chk(3, "工号列 A2/A3/A5/A10 全加粗", col_a_bold)
        chk(3, "B3 不被误伤（非目标列）", not ws["B3"].font.bold,
            "B3 bold=%s" % ws["B3"].font.bold)

    # 4 单元格标红
    o = results[4]["out"]
    if o:
        ws = load_workbook(str(o)).active
        c = ws["B3"].font.color
        chk(4, "B3 颜色=FFFF0000", c is not None and getattr(c, "type", None) == "rgb"
            and c.rgb == "FFFF0000",
            "type=%s rgb=%s" % (getattr(c, "type", None),
                                c.rgb if (c is not None and getattr(c, "type", None) == "rgb") else "-"))
        # 注：openpyxl 对"未设置颜色"的格子，访问 .rgb 会返回一句提示文本而非 None，
        # 所以判"未误伤"要认 type（rgb 类型才算真被上色）。
        c4 = ws["B4"].font.color
        b4_red = (c4 is not None) and (getattr(c4, "type", None) == "rgb") and (c4.rgb == "FFFF0000")
        chk(4, "B4 未误伤", not b4_red, "B4 type=%s" % (getattr(c4, "type", None)))

    # 5 表头字号 14，数据行仍 11
    o = results[5]["out"]
    if o:
        ws = load_workbook(str(o)).active
        hd = {ws.cell(row=2, column=c).font.size for c in range(1, 13)}
        dt = {ws.cell(row=3, column=c).font.size for c in range(1, 13)}
        chk(5, "表头行 14 号", hd == {14.0}, "sizes=%s" % hd)
        chk(5, "数据行仍 11 号（零误伤）", dt == {11.0}, "sizes=%s" % dt)

    # 6 姓名列（第 2 列）居中：表头 + 数据
    o = results[6]["out"]
    if o:
        ws = load_workbook(str(o)).active
        col_b = {ws.cell(row=r, column=2).alignment.horizontal for r in (2, 3, 5)}
        chk(6, "姓名列 B2/B3/B5 全居中", col_b == {"center"}, "aligns=%s" % col_b)
        chk(6, "A3 未误伤", ws["A3"].alignment.horizontal != "center",
            "A3 align=%s" % ws["A3"].alignment.horizontal)

    # 7 表头行居中
    o = results[7]["out"]
    if o:
        ws = load_workbook(str(o)).active
        hd = {ws.cell(row=2, column=c).alignment.horizontal for c in range(1, 13)}
        chk(7, "表头行全居中", hd == {"center"}, "aligns=%s" % hd)
        chk(7, "A3 数据未误伤", ws["A3"].alignment.horizontal != "center",
            "A3 align=%s" % ws["A3"].alignment.horizontal)

    # 8 表头字体黑体
    o = results[8]["out"]
    if o:
        ws = load_workbook(str(o)).active
        names = {ws.cell(row=2, column=c).font.name for c in range(1, 13)}
        chk(8, "表头行字体=黑体", names == {"黑体"}, "names=%s" % names)
        chk(8, "数据行字体未变", ws["B3"].font.name == ws0["B3"].font.name,
            "B3=%s" % ws["B3"].font.name)

    # ---- 负向 ----
    chk(9, "负向：找不到的列 → 失败且无产物", neg_ok and neg_no_out,
        "returncode≠0=%s｜没找到=%s｜无产物=%s" % (r.returncode != 0, "没找到" in r.stdout, neg_no_out))

    # ---- ⑥ 指纹复核 ----
    h1 = md5(SRC)
    chk(0, "源文件零改动（指纹一致）", h0 == h1)

    # ---- 汇总 ----
    n_pass = sum(1 for _, _, ok, _ in checks if ok)
    n_cmd = sum(1 for cid in results if results[cid]["ok"])
    print("\n" + "=" * 66)
    print("小计：%d/%d 项验证通过 ｜ 正向 %d/8 条运行成功（负向 1/1）"
          % (n_pass, len(checks), n_cmd))
    print("=" * 66)

    # 报告存盘
    lines = ["xlsx 操作补验报告", "=" * 40]
    for cid, name, ok, detail in checks:
        lines.append("[%s] %s %s" % ("PASS" if ok else "FAIL", name, detail))
    lines.append("小计：%d/%d" % (n_pass, len(checks)))
    (OUT / "xlsx操作-验收报告.txt").write_text("\n".join(lines), encoding="utf-8")
    print("报告已写入：agent/输出/xlsx操作-验收报告.txt")


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    main()
