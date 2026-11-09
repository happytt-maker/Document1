# -*- coding: utf-8 -*-
"""
复杂表头-素材生成-1026.py
=========================
造 5 张"不规整表头"的模板，专门打原来没测过的三种形态：

    ① 多层表头   —— 父列合并跨若干子列（上半年采购：数量/金额）
    ② 斜线表头   —— 一个格子里塞两个维度（"费用类别＼月份"）
    ③ 跨页表     —— 同一张表被分页切开，第二页重复表头，中间夹「续表」

原有 5 张模板的表头区一个合并单元格都没有（已核实），所以这三类
在改造前 100% 走不通（父列名丢失、斜线格整串当一个列名、第二页当新表）。
本脚本造素材 + 手写标准答案，供结构验收逐项对照。

输出：
    测试数据/模板-复杂表头/*.xlsx
    测试数据/模板-复杂表头-word/*.docx
    测试数据/模板-复杂表头/_标准答案.json
"""

import json
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Border, Font, Side
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt

HERE = Path(__file__).resolve().parent
BASE = HERE.parent
OUT_X = BASE / "测试数据" / "模板-复杂表头"
OUT_W = BASE / "测试数据" / "模板-复杂表头-word"

THIN = Side(style="thin", color="000000")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)


# ----------------------------------------------------------------------
# xlsx 辅助
# ----------------------------------------------------------------------
def new_xlsx(title, n_cols, rows):
    """建一个 n_cols 列、rows 行的空簿子，全部加细边框。"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = title
    for r in range(1, rows + 1):
        ws.row_dimensions[r].height = 22
        for c in range(1, n_cols + 1):
            cell = ws.cell(row=r, column=c)
            cell.border = BOX
            cell.alignment = CENTER
    ws.column_dimensions["A"].width = 12
    for c in range(2, n_cols + 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(c)].width = 14
    return wb, ws


def put(ws, row, col, text, bold=False, size=11):
    cell = ws.cell(row=row, column=col, value=text)
    cell.alignment = CENTER
    cell.border = BOX
    if bold:
        cell.font = Font(bold=True, size=size)
    return cell


def title_row(ws, n_cols, text, row=1):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=n_cols)
    c = put(ws, row, 1, text, bold=True, size=14)
    c.alignment = CENTER
    ws.row_dimensions[row].height = 30


# ----------------------------------------------------------------------
# ① 多层表头（xlsx）
# ----------------------------------------------------------------------
def make_multi_header():
    """9 列 × 2 层。
    关键点：「数量」「金额（元）」两个子名在两组合并父列下重复出现
            → 必须用父名区分为「上半年采购-数量」「下半年采购-数量」；
              「设备名称」「供应商名称」全局唯一 → 保持简洁名，不硬拼父名。
    """
    wb, ws = new_xlsx("设备采购台账", 9, 14)
    title_row(ws, 9, "2026 年度设备采购台账")

    put(ws, 2, 1, "采购单位：")
    ws.merge_cells("A2:C2")
    put(ws, 2, 4, "制表日期：")
    ws.merge_cells("D2:F2")
    put(ws, 2, 7, "金额单位：元")
    ws.merge_cells("G2:I2")

    # 父层（第 3 行）
    for c in (1, 2, 3):
        ws.merge_cells(start_row=3, start_column=c, end_row=4, end_column=c)
        put(ws, 3, c, {1: "序号", 2: "设备名称", 3: "规格型号"}[c], bold=True)
    for c0, c1, name in [(4, 5, "上半年采购"), (6, 7, "下半年采购"), (8, 9, "供应商信息")]:
        ws.merge_cells(start_row=3, start_column=c0, end_row=3, end_column=c1)
        put(ws, 3, c0, name, bold=True)

    # 子层（第 4 行）
    for c, name in [(4, "数量"), (5, "金额（元）"), (6, "数量"), (7, "金额（元）"),
                    (8, "供应商名称"), (9, "联系方式")]:
        put(ws, 4, c, name, bold=True)

    # 待填区 5~13（全空）
    # 结尾
    ws.merge_cells("A14:C14")
    put(ws, 14, 1, "合计", bold=True)
    put(ws, 14, 5, "=SUM(E5:E13)", bold=True)
    put(ws, 14, 7, "=SUM(G5:G13)", bold=True)

    p = OUT_X / "模板6-多层表头设备采购台账.xlsx"
    wb.save(p)
    return p


# ----------------------------------------------------------------------
# ② 斜线表头（xlsx）
# ----------------------------------------------------------------------
def make_slash_header():
    """7 列。左上角 A3:A4 合并成一个斜线格，格子里是两行文字
    「费用类别」+「月份」—— 传统写法，一个格子同时当行维度标题和列维度标题。
    原来的读取器会把它整串（"费用类别 月份"）当成第一列的列名。
    """
    wb, ws = new_xlsx("费用报销明细", 7, 12)
    title_row(ws, 7, "2026 年上半年费用报销明细表")

    put(ws, 2, 1, "报销人：")
    ws.merge_cells("A2:B2")
    put(ws, 2, 3, "部门：")
    ws.merge_cells("C2:D2")
    put(ws, 2, 5, "填报日期：")
    ws.merge_cells("E2:G2")

    # 斜线格：A3:A4 合并，两行文字（用换行体现斜线两侧的两个维度）
    ws.merge_cells("A3:A4")
    sc = put(ws, 3, 1, "费用类别\n月份", bold=True)
    sc.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # 父层：2026 年度
    ws.merge_cells("B3:G3")
    put(ws, 3, 2, "2026 年度", bold=True)

    # 子层：1~6 月
    for i, m in enumerate(["1 月", "2 月", "3 月", "4 月", "5 月", "6 月"]):
        put(ws, 4, 2 + i, m, bold=True)

    # 行标签（第一列是费用类别名，属于表格结构不是待填数据）
    for i, cat in enumerate(["差旅费", "办公费", "招待费", "交通费", "其他"]):
        put(ws, 5 + i, 1, cat, bold=True)

    ws.merge_cells("A11:G11")
    put(ws, 11, 1, "合计", bold=True)

    p = OUT_X / "模板7-斜线表头费用报销明细.xlsx"
    wb.save(p)
    return p


# ----------------------------------------------------------------------
# ③ 跨页表（xlsx）
# ----------------------------------------------------------------------
def make_paged_xlsx():
    """同一 sheet 里两段：第 1 页表头 + 9 行 →「续表」→ 表头重复 + 9 行。
    这是 Excel 里做跨页打印最土也最常见的做法（手动复制表头）。
    原来的读取器只认第一段，第二段的 9 行永远填不到。
    """
    cols = ["序号", "物料编码", "物料名称", "规格型号", "入库数量",
            "单位", "入库日期", "供应商"]
    n = len(cols)
    wb, ws = new_xlsx("原材料入库台账", n, 26)
    title_row(ws, n, "2026 年 9 月原材料入库台账")

    put(ws, 2, 1, "仓库：")
    ws.merge_cells("A2:B2")
    put(ws, 2, 3, "保管员：")
    ws.merge_cells("C2:D2")
    put(ws, 2, 5, "制表日期：")
    ws.merge_cells("E2:H2")

    # 第一页表头
    for i, c in enumerate(cols):
        put(ws, 3, 1 + i, c, bold=True)
    # 待填 4~12（9 行空）

    # 续表分隔
    ws.merge_cells("A13:H13")
    put(ws, 13, 1, "续表", bold=True)

    # 第二页重复表头
    for i, c in enumerate(cols):
        put(ws, 14, 1 + i, c, bold=True)
    # 待填 15~23（9 行空）

    # 注意：合计行不能整行合并（E24 要留公式格），否则 E24 成了只读的 MergedCell
    ws.merge_cells("A24:C24")
    put(ws, 24, 1, "合计", bold=True)
    put(ws, 24, 5, "=SUM(E4:E12)+SUM(E15:E23)", bold=True)
    put(ws, 24, 6, "—", bold=True)

    p = OUT_X / "模板8-跨页原材料入库台账.xlsx"
    wb.save(p)
    return p


# ----------------------------------------------------------------------
# docx 辅助
# ----------------------------------------------------------------------
def center_all(table):
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for row in table.rows:
        for cell in row.cells:
            for p in cell.paragraphs:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.paragraph_format.space_before = Pt(0)
                p.paragraph_format.space_after = Pt(0)


def doc_title(doc, text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(text)
    r.bold = True
    r.font.size = Pt(16)


def doc_note(doc, text, align=WD_ALIGN_PARAGRAPH.CENTER):
    p = doc.add_paragraph()
    p.alignment = align
    r = p.add_run(text)
    r.font.size = Pt(10.5)


# ----------------------------------------------------------------------
# ④ 跨页表（docx，两张独立表格 + 续表）
# ----------------------------------------------------------------------
def make_paged_docx():
    """Word 跨页：第 1 张表 + 一段「续表」文字 + 第 2 张表（表头重复）。
    在 Word 里表格跨页通常是同一个 table，但实际收到的模板常常是
    「复制表头再新建一张」——这种就是两个 table，业务上是一张表。
    """
    doc = Document()
    doc_title(doc, "会议任务跟踪表")
    doc_note(doc, "会议名称：                        "
                  "会议时间：                        ")

    cols = ["序号", "任务内容", "负责人", "完成时限", "当前进度", "备注"]
    n = len(cols)

    t1 = doc.add_table(rows=1 + 4, cols=n)
    t1.style = "Table Grid"
    for i, c in enumerate(cols):
        t1.rows[0].cells[i].text = c
    center_all(t1)

    doc_note(doc, "")
    doc_note(doc, "续表")
    doc_note(doc, "")

    t2 = doc.add_table(rows=1 + 4, cols=n)
    t2.style = "Table Grid"
    for i, c in enumerate(cols):
        t2.rows[0].cells[i].text = c
    center_all(t2)

    p = OUT_W / "模板9-跨页会议任务跟踪.docx"
    doc.save(str(p))
    return p


# ----------------------------------------------------------------------
# ⑤ 多层 + 斜线（docx）
# ----------------------------------------------------------------------
def make_multi_docx():
    """Word 多层表头：基本信息(姓名/性别) + 联系方式(手机/邮箱)，
    以及一列纵向合并的「部门」。用来验证 docx 的 gridSpan / vMerge。
    """
    doc = Document()
    doc_title(doc, "员工信息登记表")

    cols = ["序号", "姓名", "性别", "手机", "邮箱", "部门"]
    n = len(cols)
    table = doc.add_table(rows=2 + 3, cols=n)
    table.style = "Table Grid"

    # 父层：基本信息跨 2 列、联系方式跨 2 列
    a = table.cell(0, 1).merge(table.cell(0, 2))
    a.text = "基本信息"
    b = table.cell(0, 3).merge(table.cell(0, 4))
    b.text = "联系方式"

    # 纵向合并：序号、部门 上下两层各占一格
    for c in (0, 5):
        m = table.cell(0, c).merge(table.cell(1, c))
        m.text = cols[c]

    # 子层
    table.cell(1, 1).text = "姓　名"
    table.cell(1, 2).text = "性别"
    table.cell(1, 3).text = "手机号码"
    table.cell(1, 4).text = "电子邮箱"
    center_all(table)

    p = OUT_W / "模板10-多层表头员工信息登记.docx"
    doc.save(str(p))
    return p


# ----------------------------------------------------------------------
def main():
    OUT_X.mkdir(parents=True, exist_ok=True)
    OUT_W.mkdir(parents=True, exist_ok=True)

    made = []
    for fn in (make_multi_header, make_slash_header, make_paged_xlsx,
               make_paged_docx, make_multi_docx):
        p = fn()
        made.append((p.name, p.stat().st_size))
        print("生成 %-46s %6.1f KB" % (p.name, p.stat().st_size / 1024))

    # ---------------- 标准答案（手写，不由被测代码生成） ----------------
    answer = {
        "模板6-多层表头设备采购台账.xlsx": {
            "header_rows": [3, 4],
            "header_row": 4,
            "titles": ["序号", "设备名称", "规格型号",
                       "上半年采购-数量", "上半年采购-金额（元）",
                       "下半年采购-数量", "下半年采购-金额（元）",
                       "供应商名称", "联系方式"],
            "segments": [[5, 13]],
        },
        "模板7-斜线表头费用报销明细.xlsx": {
            "header_rows": [3, 4],
            "header_row": 4,
            "titles": ["费用类别", "1 月", "2 月", "3 月", "4 月", "5 月", "6 月"],
            "slash": {"col": 1, "parts": ["费用类别", "月份"]},
            "segments": [[5, 10]],
        },
        "模板8-跨页原材料入库台账.xlsx": {
            "header_rows": [3],
            "header_row": 3,
            "titles": ["序号", "物料编码", "物料名称", "规格型号",
                       "入库数量", "单位", "入库日期", "供应商"],
            "segments": [[4, 12], [15, 23]],
        },
        "模板9-跨页会议任务跟踪.docx": {
            "header_rows": [1],
            "header_row": 1,
            "titles": ["序号", "任务内容", "负责人", "完成时限", "当前进度", "备注"],
            # 注意：Word 里第 2 段和第 1 段是两张独立的 table，行号各算各的，
            # 所以两段都是 2~5，靠 segments[].table 区分（不是 7~10 的连续行号）
            "segments": [[2, 5], [2, 5]],
            "tables": [1, 2],
        },
        "模板10-多层表头员工信息登记.docx": {
            "header_rows": [1, 2],
            "header_row": 2,
            "titles": ["序号", "姓　名", "性别", "手机号码", "电子邮箱", "部门"],
            "segments": [[3, 5]],
            "tables": [1],
        },
    }
    (OUT_X / "_标准答案.json").write_text(
        json.dumps(answer, ensure_ascii=False, indent=2), encoding="utf-8")

    print("\n共 %d 份，标准答案：%s" % (len(made), OUT_X / "_标准答案.json"))


if __name__ == "__main__":
    main()
