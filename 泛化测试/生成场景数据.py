# -*- coding: utf-8 -*-
"""
生成场景数据.py —— 造 3 个"比赛中可能出"的场景（模板 + 源文档）
================================================================

为什么要造新场景：
  之前所有测试都围绕同一批数据（模板4 + 那 18 份文档），成绩再好也只是
  "练过的那套"。比赛是随机抽题 —— 必须拿"没见过的新文档"再测一遍，
  才知道系统是真会了，还是只是背住了。

三个场景（每个 = 1 张新模板 + 4~5 份源文档，主来源文档不同）：
  场景1-费用报销   主力是 txt  → 考验文本源 + 行号出处
  场景2-招聘初筛   主力是 xlsx → 考验表格源 + 格子坐标出处
  场景3-设备报修   主力是 txt，干扰很凶（设备台账里有同名设备、同名人）

生成的东西（与现有测试数据同构）：
  泛化测试/场景X-XXX/源文档/{txt,md,docx,xlsx}/...
  泛化测试/场景X-XXX/模板表格/模板N-XXX.xlsx
  泛化测试/生成自检.txt   ← 关键句→行号 / 关键格→坐标 对照表（写标准答案用）

跑法：
  python 生成场景数据.py
"""
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from docx import Document

ROOT = Path(__file__).resolve().parent

# ---------- 通用样式 ----------
THIN = Side(style="thin", color="999999")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
FILL_HEAD = PatternFill("solid", fgColor="D9D9D9")


def write_txt(path: Path, lines):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def write_md(path: Path, text: str):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8")


def write_docx(path: Path, title: str, paras):
    path.parent.mkdir(parents=True, exist_ok=True)
    doc = Document()
    if title:
        doc.add_heading(title, level=1)
    for p in paras:
        doc.add_paragraph(p)
    doc.save(str(path))


def make_template(path: Path, sheet_title: str, big_title: str,
                  meta: list, headers: list, fill_rows: int, footer: str,
                  widths: list = None):
    """造一张"空白待填"模板表。

    meta:    [(列号, 文字), ...]  说明行
    headers: 表头列名（从 A 列开始）
    fill_rows: 待填区行数（表头下一行起，全空 + 细边框）
    footer:  结尾行文字（放 A 列，用来"封住"待填区下沿）
    """
    wb = Workbook()
    ws = wb.active
    ws.title = sheet_title
    ncol = len(headers)
    last = get_column_letter(ncol)

    # 大标题（合并居中）
    ws.merge_cells("A1:%s1" % last)
    c = ws["A1"]
    c.value = big_title
    c.font = Font(bold=True, size=15, name="微软雅黑")
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 30

    # 说明行
    for col, text in meta:
        ws.cell(row=2, column=col, value=text).font = Font(size=11)

    # 表头行（第 3 行）
    for i, h in enumerate(headers, start=1):
        cell = ws.cell(row=3, column=i, value=h)
        cell.font = Font(bold=True)
        cell.fill = FILL_HEAD
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = BORDER

    # 待填区：全空 + 细边框
    r0 = 4
    for r in range(r0, r0 + fill_rows):
        for i in range(1, ncol + 1):
            ws.cell(row=r, column=i).border = BORDER

    # 结尾行
    ws.cell(row=r0 + fill_rows, column=1, value=footer).font = Font(size=11)

    # 列宽
    if widths:
        for i, w in enumerate(widths, start=1):
            ws.column_dimensions[get_column_letter(i)].width = w

    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(str(path))
    return {"sheet": sheet_title, "header_row": 3,
            "fill_start": r0, "fill_end": r0 + fill_rows - 1,
            "headers": headers}


def make_data_xlsx(path: Path, sheet_title: str, big_title: str,
                   headers: list, rows: list, header_row: int = 2,
                   widths: list = None):
    """造一张"有数据"的源表格（主力/干扰都用它）。"""
    wb = Workbook()
    ws = wb.active
    ws.title = sheet_title
    ncol = len(headers)
    last = get_column_letter(ncol)

    ws.merge_cells("A1:%s1" % last)
    c = ws["A1"]
    c.value = big_title
    c.font = Font(bold=True, size=13)
    c.alignment = Alignment(horizontal="center", vertical="center")

    for i, h in enumerate(headers, start=1):
        cell = ws.cell(row=header_row, column=i, value=h)
        cell.font = Font(bold=True)
        cell.fill = FILL_HEAD
        cell.border = BORDER

    for ri, row in enumerate(rows):
        for ci, v in enumerate(row, start=1):
            cell = ws.cell(row=header_row + 1 + ri, column=ci, value=v)
            cell.border = BORDER

    if widths:
        for i, w in enumerate(widths, start=1):
            ws.column_dimensions[get_column_letter(i)].width = w

    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(str(path))


def line_of(lines, needle):
    """关键句在第几行（1 起）。"""
    for i, ln in enumerate(lines, 1):
        if needle in ln:
            return i
    return -1


# ==================================================================
# 场景1 · 费用报销（主力：txt）
# ==================================================================
def scene1():
    base = ROOT / "场景1-费用报销"
    log = ["", "=" * 60, "场景1-费用报销", "=" * 60]

    # ---- 模板6 ----
    info = make_template(
        base / "模板表格" / "模板6-费用报销台账.xlsx",
        sheet_title="费用报销台账",
        big_title="星锐科技有限公司 · 2026 年 3 月费用报销台账",
        meta=[(1, "统计周期：2026 年 3 月"), (4, "制表部门：财务部")],
        headers=["报销单号", "报销人", "所属部门", "提交日期",
                 "费用明细", "报销金额（元）", "审批状态"],
        fill_rows=6,
        footer="财务负责人签字：________________",
        widths=[16, 10, 10, 14, 44, 14, 12],
    )
    log.append("模板：%s（表头第3行，待填第%d~%d行）"
               % ("模板6-费用报销台账.xlsx", info["fill_start"], info["fill_end"]))

    # ---- 主力：费用报销登记簿.txt ----
    main_lines = [
        "星锐科技有限公司",                           # 1
        "2026 年 3 月费用报销登记簿（财务部制）",       # 2
        "本登记簿按提交时间先后记录，每笔报销一条。",    # 3
        "",                                          # 4
        "【报销记录一】",                             # 5
        "报销单号：BX-2026-0312",                     # 6
        "报销人：周敏",                               # 7
        "所属部门：市场部",                            # 8
        "提交日期：2026年3月5日",                      # 9
        "费用明细：3 月 2 日至 4 日赴上海拜访客户产生的差旅费用，含高铁票及住宿。",  # 10
        "报销金额：1860 元",                          # 11
        "审批状态：已通过",                            # 12
        "",                                          # 13
        "【报销记录二】",                             # 14
        "报销单号：BX-2026-0318",                     # 15
        "报销人：李志强",                             # 16
        "所属部门：研发部",                            # 17
        "提交日期：2026年3月9日",                      # 18
        "费用明细：研发实验室用示波器探头采购费用，共 3 支。",  # 19
        "报销金额：2350 元",                          # 20
        "审批状态：审批中",                            # 21
        "",                                          # 22
        "【报销记录三】",                             # 23
        "报销单号：BX-2026-0324",                     # 24
        "报销人：陈雅婷",                             # 25
        "所属部门：行政部",                            # 26
        "提交日期：2026年3月12日",                     # 27
        "费用明细：办公区绿植租摆季度费用（1 月至 3 月）。",     # 28
        "报销金额：1200 元",                          # 29
        "审批状态：已通过",                            # 30
    ]
    write_txt(base / "源文档" / "txt" / "费用报销登记簿.txt", main_lines)

    # ---- 干扰：财务报销管理规定.txt ----
    rule_lines = [
        "星锐科技有限公司财务报销管理规定",
        "（2025 年修订版）",
        "",
        "第一章 总则",
        "第一条 为规范费用报销流程、加强费用管控，制定本规定。",
        "第二条 报销范围包括差旅费、办公费、采购费、招待费等。",
        "",
        "第二章 报销流程",
        "第三条 报销单号由财务系统自动生成，格式为 BX-年份-流水号。",
        "第四条 单笔报销金额超过 5000 元的，须报总经理审批。",
        "第五条 报销人应在费用发生后 10 个工作日内提交报销单，逾期须附情况说明。",
        "",
        "第三章 附则",
        "第六条 本规定由财务部负责解释。",
    ]
    write_txt(base / "源文档" / "txt" / "财务报销管理规定.txt", rule_lines)

    # ---- 干扰：行政部费用月报.md ----
    write_md(base / "源文档" / "md" / "行政部费用月报.md", """# 行政部 2026 年 3 月费用月报

## 一、总体情况

3 月行政部各项费用合计 4540 元。其中办公用品 1620 元、饮水耗材 880 元、快递费 480 元、其他 1560 元。

## 二、与上月对比

环比上升约 6%，主要原因是办公用品采购集中在本月。

## 三、下月计划

- 推进办公区空调滤网集中更换
- 梳理快递月结供应商账期
""")

    # ---- 干扰：差旅管理细则.docx ----
    write_docx(base / "源文档" / "docx" / "差旅管理细则.docx",
               "差旅管理细则",
               ["第一条 员工因公出差，应事先填具出差申请单，经部门负责人核准后方可出行。",
                "第二条 差旅费用包括城市间交通费、住宿费、市内交通费及出差补贴。",
                "第三条 出差结束后，应于十个工作日内办理报销手续，并将票据附于报销单后。",
                "第四条 未按规定办理报销手续的，财务部有权暂缓受理。"])

    # ---- 干扰：各部门费用统计.xlsx ----
    make_data_xlsx(
        base / "源文档" / "xlsx" / "各部门费用统计.xlsx",
        sheet_title="各部门费用",
        big_title="星锐科技 · 3 月各部门费用统计（财务部）",
        headers=["部门", "办公费", "差旅费", "其他", "合计"],
        rows=[["市场部", 620, 1980, 120, 2720],
              ["研发部", 430, 560, 780, 1770],
              ["行政部", 1230, 0, 860, 2090]],
        widths=[12, 10, 10, 10, 10],
    )

    # ---- 关键行号/坐标对照 ----
    log.append("【主力】费用报销登记簿.txt 关键行号：")
    for k in ["报销单号：BX-2026-0312", "报销人：周敏", "所属部门：市场部",
              "提交日期：2026年3月5日", "费用明细：3 月 2 日至", "报销金额：1860",
              "审批状态：已通过（记录一）"]:
        pass
    keys = [
        ("记录一·报销单号", "报销单号：BX-2026-0312"),
        ("记录一·报销人", "报销人：周敏"),
        ("记录一·所属部门", "所属部门：市场部"),
        ("记录一·提交日期", "提交日期：2026年3月5日"),
        ("记录一·费用明细", "费用明细：3 月 2 日至"),
        ("记录一·报销金额", "报销金额：1860"),
        ("记录二·报销单号", "报销单号：BX-2026-0318"),
        ("记录二·报销人", "报销人：李志强"),
        ("记录二·所属部门", "所属部门：研发部"),
        ("记录二·提交日期", "提交日期：2026年3月9日"),
        ("记录二·费用明细", "费用明细：研发实验室"),
        ("记录二·报销金额", "报销金额：2350"),
        ("记录三·报销单号", "报销单号：BX-2026-0324"),
        ("记录三·报销人", "报销人：陈雅婷"),
        ("记录三·所属部门", "所属部门：行政部"),
        ("记录三·提交日期", "提交日期：2026年3月12日"),
        ("记录三·费用明细", "费用明细：办公区绿植"),
        ("记录三·报销金额", "报销金额：1200"),
        ("记录一·审批状态", "审批状态：已通过"),   # 第一条
    ]
    for label, needle in keys:
        n = line_of(main_lines, needle)
        log.append("  %s → 第 %d 行" % (label, n))
    # 审批状态记录三 单独算（第 3 个"已通过"）
    hits = [i for i, ln in enumerate(main_lines, 1) if ln == "审批状态：已通过"]
    log.append("  记录三·审批状态 → 第 %d 行（三个'已通过'中的最后一个：%s）"
               % (hits[-1], hits))
    hits2 = [i for i, ln in enumerate(main_lines, 1) if ln == "审批状态：审批中"]
    log.append("  记录二·审批状态 → 第 %d 行" % hits2[0])

    return log


# ==================================================================
# 场景2 · 招聘初筛（主力：xlsx）
# ==================================================================
def scene2():
    base = ROOT / "场景2-招聘初筛"
    log = ["", "=" * 60, "场景2-招聘初筛", "=" * 60]

    # ---- 模板7 ----
    info = make_template(
        base / "模板表格" / "模板7-候选人初筛表.xlsx",
        sheet_title="候选人初筛",
        big_title="星锐科技 · 2026 春招候选人初筛表",
        meta=[(1, "招聘批次：2026 春招 · 研发中心"), (4, "初筛负责人：田蕾")],
        headers=["姓名", "应聘岗位", "学历", "工作年限（年）",
                 "期望薪资（元）", "联系电话", "初筛结论"],
        fill_rows=6,
        footer="备注：初筛结论须与面试记录一致，存档备查。",
        widths=[12, 20, 10, 14, 14, 16, 12],
    )
    log.append("模板：模板7-候选人初筛表.xlsx（表头第3行，待填第%d~%d行）"
               % (info["fill_start"], info["fill_end"]))

    # ---- 主力：候选人来访登记表.xlsx（注意表头在第2行）----
    headers = ["序号", "姓名", "应聘岗位", "学历", "工作年限（年）",
               "期望薪资（元）", "联系电话", "初筛结论", "备注"]
    rows = [
        [1, "王一鸣", "后端开发工程师", "本科", 3, 15000, "138-6123-4507",
         "进入面试", "拟安排 3 月 20 日面试"],
        [2, "刘思远", "后端开发工程师", "硕士", 1, 18000, "139-2208-7734",
         "待定", "等待研发中心复筛"],
        [3, "赵子墨", "测试开发工程师", "本科", 5, 16000, "137-5541-9026",
         "进入面试", "拟安排 3 月 21 日面试"],
    ]
    make_data_xlsx(
        base / "源文档" / "xlsx" / "候选人来访登记表.xlsx",
        sheet_title="来访登记",
        big_title="星锐科技 · 2026 春招候选人现场登记",
        headers=headers, rows=rows, header_row=2,
        widths=[6, 12, 20, 8, 14, 14, 16, 12, 22],
    )
    # 坐标对照（数据从第3行开始，B~H 列）
    cols = "BCDEFGH"
    names = ["王一鸣", "刘思远", "赵子墨"]
    labels = ["姓名", "应聘岗位", "学历", "工作年限", "期望薪资", "联系电话", "初筛结论"]
    log.append("【主力】候选人来访登记表.xlsx 关键坐标（工作表「来访登记」）：")
    for ri, nm in enumerate(names):
        for ci, lb in enumerate(labels):
            log.append("  %s·%s → %s%d" % (nm, lb, cols[ci], 3 + ri))

    # ---- 干扰：招聘岗位需求.md ----
    write_md(base / "源文档" / "md" / "招聘岗位需求.md", """# 星锐科技 2026 春招 · 研发中心岗位需求

## 后端开发工程师（2 人）

- 学历要求：本科及以上
- 经验要求：3 年及以上相关经验
- 薪资范围：15-20K/月
- 岗位职责：负责服务端模块设计、开发与维护，参与系统架构优化。

## 测试开发工程师（1 人）

- 学历要求：本科及以上
- 经验要求：5 年及以上
- 薪资范围：16-22K/月
- 岗位职责：负责自动化测试框架建设与维护。
""")

    # ---- 干扰：面试安排通知.txt ----
    note_lines = [
        "星锐科技 2026 春招面试安排通知（第一批）",
        "",
        "一、面试时间",
        "3 月 20 日（周五）上午 10:00 —— 后端开发工程师专场",
        "3 月 21 日（周六）下午 14:00 —— 测试开发工程师专场",
        "",
        "二、面试地点",
        "公司总部 B 座 3 层会议室",
        "",
        "三、注意事项",
        "1. 请携带身份证及纸质简历，提前 15 分钟到场签到。",
        "2. 面试结果将于 5 个工作日内电话通知。",
    ]
    write_txt(base / "源文档" / "txt" / "面试安排通知.txt", note_lines)

    # ---- 干扰：用人部门申请单.docx ----
    write_docx(base / "源文档" / "docx" / "用人部门申请单.docx",
               "用人部门招聘申请单",
               ["申请部门：研发中心",
                "招聘岗位：后端开发工程师、测试开发工程师",
                "需求人数：3 人",
                "期望到岗时间：2026 年 4 月",
                "申请理由：项目交付规模扩大，现有团队人力紧张，需补充开发与测试人员。",
                "申请人：田蕾"])

    return log


# ==================================================================
# 场景3 · 设备报修（主力：txt；干扰很凶）
# ==================================================================
def scene3():
    base = ROOT / "场景3-设备报修"
    log = ["", "=" * 60, "场景3-设备报修", "=" * 60]

    # ---- 模板8 ----
    info = make_template(
        base / "模板表格" / "模板8-设备报修工单登记表.xlsx",
        sheet_title="报修工单登记",
        big_title="星锐科技 · 设备报修工单登记表",
        meta=[(1, "登记周期：2026 年 2 月"), (4, "归口部门：设备动力科")],
        headers=["工单号", "设备名称", "故障描述", "报修人",
                 "报修日期", "故障等级", "处理状态"],
        fill_rows=6,
        footer="备注：故障等级分“一般 / 紧急”两档；处理状态分“待维修 / 维修中 / 已修复”三种。",
        widths=[16, 20, 42, 10, 14, 10, 10],
    )
    log.append("模板：模板8-设备报修工单登记表.xlsx（表头第3行，待填第%d~%d行）"
               % (info["fill_start"], info["fill_end"]))

    # ---- 主力：设备报修记录.txt ----
    main_lines = [
        "星锐科技 设备动力科 · 2 月设备报修记录",     # 1
        "",                                        # 2
        "【报修单 一】",                             # 3
        "工单号：WX-2026-0071",                     # 4
        "设备名称：数控车床 CK6140",                 # 5
        "故障描述：主轴启动时有明显异响，无法正常加工。",  # 6
        "报修人：孙建国",                            # 7
        "报修日期：2026年2月10日",                    # 8
        "故障等级：紧急",                            # 9
        "处理状态：已修复",                          # 10
        "",                                        # 11
        "【报修单 二】",                             # 12
        "工单号：WX-2026-0076",                     # 13
        "设备名称：加工中心 XK714",                  # 14
        "故障描述：刀库换刀动作卡滞，偶发报警。",        # 15
        "报修人：周慧",                              # 16
        "报修日期：2026年2月14日",                    # 17
        "故障等级：一般",                            # 18
        "处理状态：维修中",                          # 19
        "",                                        # 20
        "【报修单 三】",                             # 21
        "工单号：WX-2026-0083",                     # 22
        "设备名称：激光切割机 LC-3015",              # 23
        "故障描述：切割边缘毛刺明显，需校准光路。",       # 24
        "报修人：吴峰",                              # 25
        "报修日期：2026年2月21日",                    # 26
        "故障等级：一般",                            # 27
        "处理状态：待维修",                          # 28
    ]
    write_txt(base / "源文档" / "txt" / "设备报修记录.txt", main_lines)

    # ---- 干扰：设备资产台账.xlsx（同名设备 + 同名人，故意凶）----
    make_data_xlsx(
        base / "源文档" / "xlsx" / "设备资产台账.xlsx",
        sheet_title="设备清单",
        big_title="星锐科技 · 生产设备资产台账",
        headers=["设备编号", "设备名称", "型号", "购置日期", "存放车间", "责任人"],
        rows=[["SB-001", "数控车床", "CK6140", "2023-05", "一号车间", "孙建国"],
              ["SB-002", "加工中心", "XK714", "2022-11", "一号车间", "周慧"],
              ["SB-003", "激光切割机", "LC-3015", "2024-02", "二号车间", "吴峰"],
              ["SB-004", "数控铣床", "XK5032", "2021-08", "二号车间", "郑海涛"]],
        widths=[10, 14, 10, 12, 10, 10],
    )

    # ---- 干扰：安全生产月报.md ----
    write_md(base / "源文档" / "md" / "安全生产月报.md", """# 设备动力科 2026 年 2 月安全生产月报

## 一、设备运行概况

本月共收到报修 3 起，其中紧急故障 1 起，均已按流程登记处理。设备平均修复时长 1.5 天。

## 二、主要工作

- 完成一号车间设备季度保养
- 组织安全操作培训 1 次，参加 12 人
- 排查消防隐患 2 处并完成整改

## 三、下月计划

- 推进二号车间老旧线路改造
- 开展应急处置演练
""")

    # ---- 干扰：设备维修管理办法.docx ----
    write_docx(base / "源文档" / "docx" / "设备维修管理办法.docx",
               "设备维修管理办法",
               ["第一条 设备发生故障后，操作人员应在 30 分钟内报告设备动力科，并填写报修单。",
                "第二条 紧急故障指影响生产线运行、存在安全隐患的故障，应优先安排处理。",
                "第三条 一般故障由当班维修工按序处理，原则上 3 个工作日内完成。",
                "第四条 维修完成后，维修工应在报修单上填写处理结果并签字确认。"])

    # ---- 关键行号对照 ----
    log.append("【主力】设备报修记录.txt 关键行号：")
    keys = [
        ("记录一·工单号", "工单号：WX-2026-0071"),
        ("记录一·设备名称", "设备名称：数控车床"),
        ("记录一·故障描述", "故障描述：主轴启动时"),
        ("记录一·报修人", "报修人：孙建国"),
        ("记录一·报修日期", "报修日期：2026年2月10日"),
        ("记录一·故障等级", "故障等级：紧急"),
        ("记录一·处理状态", "处理状态：已修复"),
        ("记录二·工单号", "工单号：WX-2026-0076"),
        ("记录二·设备名称", "设备名称：加工中心"),
        ("记录二·故障描述", "故障描述：刀库换刀"),
        ("记录二·报修人", "报修人：周慧"),
        ("记录二·报修日期", "报修日期：2026年2月14日"),
        ("记录二·故障等级", "故障等级：一般（第18行）"),
        ("记录二·处理状态", "处理状态：维修中"),
        ("记录三·工单号", "工单号：WX-2026-0083"),
        ("记录三·设备名称", "设备名称：激光切割机"),
        ("记录三·故障描述", "故障描述：切割边缘"),
        ("记录三·报修人", "报修人：吴峰"),
        ("记录三·报修日期", "报修日期：2026年2月21日"),
        ("记录三·处理状态", "处理状态：待维修"),
    ]
    for label, needle in keys:
        if "第18行" in needle:
            n = [i for i, ln in enumerate(main_lines, 1) if ln == "故障等级：一般"][0]
        else:
            n = line_of(main_lines, needle)
        log.append("  %s → 第 %d 行" % (label, n))
    hits = [i for i, ln in enumerate(main_lines, 1) if ln == "故障等级：一般"]
    log.append("  （两个'一般'分别在第 %s 行）" % hits)

    return log


# ==================================================================
def main():
    logs = []
    logs += scene1()
    logs += scene2()
    logs += scene3()
    logs += ["", "=" * 60,
             "自检说明：上面是「关键句 → 行号 / 关键格 → 坐标」对照表，",
             "写标准答案时照着抄；跑测试时脚本还会从落盘文件里复核一遍。",
             "=" * 60]

    text = "\n".join(logs) + "\n"
    (ROOT / "生成自检.txt").write_text(text, encoding="utf-8")
    print(text)
    print("已写出：生成自检.txt")
    print("场景数据根目录：%s" % ROOT)


if __name__ == "__main__":
    main()
