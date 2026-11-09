# -*- coding: utf-8 -*-
"""
行业测试集-生成-1105.py —— 造 6 个行业的文档 + 人工标注（给统一知识抽取做验收）

为什么自己造：手上真实的行业文档只有采购/人事/政务几类，
要证明"换行业也能抽"，必须有医疗、法律、物流、教育、制造、环保这些完全不同的版式。
这 6 份是**按真实单据的字段和写法**造的（不是随便填空），
标注文件只写"人看一眼认为该抽到什么"，不写任何系统内部结构。

产出：
  测试数据/行业测试集/01-医疗-检验报告.docx
  测试数据/行业测试集/02-法律-民事判决书.docx
  测试数据/行业测试集/03-物流-运单台账.xlsx
  测试数据/行业测试集/04-教育-期末成绩单.xlsx
  测试数据/行业测试集/05-制造-物料清单.xlsx
  测试数据/行业测试集/06-环保-监测报告.docx
  测试数据/行业测试集/标准答案.json
"""
import json
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Pt
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

BASE = Path(__file__).resolve().parents[1]
OUT = BASE / "测试数据" / "行业测试集"
OUT.mkdir(parents=True, exist_ok=True)


def _font(run, size=10.5, bold=False):
    run.font.name = "宋体"
    run.font.size = Pt(size)
    run.bold = bold
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")


def para(doc, text, size=10.5, bold=False, center=False):
    p = doc.add_paragraph()
    if center:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _font(p.add_run(text), size, bold)
    return p


def table(doc, header, rows, widths=None):
    t = doc.add_table(rows=1, cols=len(header))
    t.style = "Table Grid"
    for i, h in enumerate(header):
        cell = t.rows[0].cells[i]
        cell.text = ""
        _font(cell.paragraphs[0].add_run(h), 10, True)
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    for r in rows:
        cells = t.add_row().cells
        for i, v in enumerate(r):
            cells[i].text = ""
            _font(cells[i].paragraphs[0].add_run(str(v)), 10)
    return t


def fill_sheet(ws, header, rows, title=None):
    thin = Side(style="thin", color="999999")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    start = 1
    if title:
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(header))
        c = ws.cell(row=1, column=1, value=title)
        c.font = Font(name="宋体", size=13, bold=True)
        c.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 26
        start = 2
    for j, h in enumerate(header, start=1):
        c = ws.cell(row=start, column=j, value=h)
        c.font = Font(name="宋体", size=10.5, bold=True)
        c.fill = PatternFill("solid", fgColor="DDEBF7")
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = border
    for i, row in enumerate(rows, start=start + 1):
        for j, v in enumerate(row, start=1):
            c = ws.cell(row=i, column=j, value=v)
            c.font = Font(name="宋体", size=10)
            c.alignment = Alignment(vertical="center", wrap_text=True)
            c.border = border
    widths = [12, 18, 12, 12, 12, 12, 14, 12]
    for j in range(1, len(header) + 1):
        ws.column_dimensions[chr(64 + j)].width = widths[(j - 1) % len(widths)]


answer = {}

# ---------------------------------------------------------------- 1 医疗
d = Document()
para(d, "××市第一人民医院", 16, True, center=True)
para(d, "生化检验报告单", 14, True, center=True)
para(d, "姓名：张伟    性别：男    年龄：45    科室：心内科")
para(d, "样本号：20260712-0831    样本类型：静脉血    采样日期：2026-07-12    报告日期：2026-07-12")
lab = [
    ("白细胞计数", "6.8", "10^9/L", "3.5-9.5", "", "血常规"),
    ("红细胞计数", "4.52", "10^12/L", "4.3-5.8", "", "血常规"),
    ("血红蛋白", "118", "g/L", "130-175", "↓", "血常规"),
    ("血小板计数", "236", "10^9/L", "125-350", "", "血常规"),
    ("谷丙转氨酶", "56", "U/L", "9-50", "↑", "肝功能"),
    ("总胆红素", "14.2", "μmol/L", "3.4-20.5", "", "肝功能"),
]
table(d, ["检验项目", "结果", "单位", "参考区间", "提示", "检验组"], lab)
para(d, "检验结论：血红蛋白偏低，谷丙转氨酶轻度升高，建议一周后复查肝功能。")
para(d, "检验者：陈晓兰    审核者：周敏    报告时间：2026-07-12 15:40")
d.save(OUT / "01-医疗-检验报告.docx")
answer["01-医疗-检验报告.docx"] = {
    "industry": "医疗",
    "objects": ["检验项目"],
    "attrs": ["结果", "单位", "参考区间", "提示"],
    "values": [
        ("白细胞计数", "结果", "6.8"),
        ("血红蛋白", "结果", "118"),
        ("血红蛋白", "提示", "↓"),
        ("谷丙转氨酶", "结果", "56"),
        ("谷丙转氨酶", "提示", "↑"),
        ("总胆红素", "参考区间", "3.4-20.5"),
    ],
    "single": [("样本号", "20260712-0831"), ("姓名", "张伟"), ("报告日期", "2026-07-12")],
}

# ---------------------------------------------------------------- 2 法律
d = Document()
para(d, "江苏省南京市中级人民法院", 16, True, center=True)
para(d, "民 事 判 决 书", 15, True, center=True)
para(d, "（2026）苏01民初1234号", 11, center=True)
para(d, "原告：江苏华建工程有限公司，住所地江苏省南京市建邺区江东中路188号。")
para(d, "法定代表人：李国强，该公司董事长。")
para(d, "被告：南京恒达物流有限公司，住所地江苏省南京市江宁区秣陵街道39号。")
para(d, "法定代表人：赵永康，该公司总经理。")
para(d, "诉讼请求：一、判令被告支付拖欠的运输服务费1,280,000元；二、判令被告承担逾期付款违约金86,400元；三、本案诉讼费由被告承担。")
para(d, "本院认为，原被告签订的《运输服务合同》合法有效，被告未按约定期限支付服务费，构成违约，应当承担继续履行及违约责任。")
para(d, "判决如下：一、被告南京恒达物流有限公司于本判决生效之日起十日内支付原告江苏华建工程有限公司运输服务费1,280,000元；二、被告于同期支付逾期付款违约金86,400元；三、驳回原告其他诉讼请求。")
para(d, "案件受理费17,460元，减半收取8,730元，由被告负担。")
para(d, "审判长：王海燕    审判员：刘志远    审判员：孙倩")
para(d, "二〇二六年六月三十日")
para(d, "书记员：周文静")
d.save(OUT / "02-法律-民事判决书.docx")
answer["02-法律-民事判决书.docx"] = {
    "industry": "法律",
    "objects": ["原告", "被告", "诉讼请求", "判决结果"],
    "attrs": ["案号", "诉讼请求", "判决结果", "审判长"],
    "values": [
        ("原告", "原告", "江苏华建工程有限公司"),
        ("被告", "被告", "南京恒达物流有限公司"),
        ("原告", "法定代表人", "李国强"),
        ("被告", "法定代表人", "赵永康"),
        ("判决结果", "判决结果", "1,280,000元"),
        ("判决结果", "判决结果", "86,400元"),
    ],
    "single": [("案号", "（2026）苏01民初1234号"), ("审判长", "王海燕")],
}

# ---------------------------------------------------------------- 3 物流
wb = Workbook()
ws = wb.active
ws.title = "运单"
rows = [
    ["YD20260701001", "苏州精诚自动化", "杭州华锐电子", "工业机器人本体", 320, "0.86", "德邦物流", "2026-07-01"],
    ["YD20260701002", "无锡恒信精密", "南京华建工程", "精密轴承", 120, "0.35", "顺丰速运", "2026-07-01"],
    ["YD20260702003", "上海远洋贸易", "合肥智造科技", "伺服电机", 80, "0.52", "德邦物流", "2026-07-02"],
    ["YD20260702004", "常州天合材料", "宁波海天机械", "铝合金型材", 1500, "3.20", "安能物流", "2026-07-02"],
    ["YD20260703005", "杭州华锐电子", "苏州精诚自动化", "控制主板", 60, "0.18", "顺丰速运", "2026-07-03"],
    ["YD20260703006", "南京恒达物流", "无锡恒信精密", "包装耗材", 400, "1.05", "中通快运", "2026-07-03"],
]
fill_sheet(ws, ["运单号", "寄件人", "收件人", "货物名称", "数量", "重量（吨）", "承运商", "发运日期"],
           rows, title="2026 年 7 月公路货运运单台账")
wb.save(OUT / "03-物流-运单台账.xlsx")
answer["03-物流-运单台账.xlsx"] = {
    "industry": "物流",
    "objects": ["运单"],
    "attrs": ["寄件人", "收件人", "货物名称", "数量", "重量（吨）", "承运商", "发运日期"],
    "values": [
        ("YD20260701001", "寄件人", "苏州精诚自动化"),
        ("YD20260701001", "收件人", "杭州华锐电子"),
        ("YD20260701001", "重量（吨）", "0.86"),
        ("YD20260702004", "承运商", "安能物流"),
        ("YD20260703005", "货物名称", "控制主板"),
    ],
    "single": [("运单总数", "6")],
}

# ---------------------------------------------------------------- 4 教育
wb = Workbook()
ws = wb.active
ws.title = "成绩"
rows = [
    ["20260101", "陈思远", "高一(1)班", 92, 88, 95, 275, 12],
    ["20260102", "林雨桐", "高一(1)班", 85, 91, 89, 265, 25],
    ["20260103", "王浩然", "高一(1)班", 78, 72, 80, 230, 88],
    ["20260104", "赵一鸣", "高一(2)班", 96, 94, 92, 282, 3],
    ["20260105", "周可欣", "高一(2)班", 88, 86, 90, 264, 29],
    ["20260106", "吴俊杰", "高一(2)班", 73, 68, 75, 216, 132],
]
fill_sheet(ws, ["学号", "姓名", "班级", "语文", "数学", "英语", "总分", "年级排名"], rows,
           title="××中学 2025-2026 学年第二学期期末成绩单")
wb.save(OUT / "04-教育-期末成绩单.xlsx")
answer["04-教育-期末成绩单.xlsx"] = {
    "industry": "教育",
    "objects": ["学生"],
    "attrs": ["姓名", "班级", "语文", "数学", "英语", "总分", "年级排名"],
    "values": [
        ("20260101", "姓名", "陈思远"),
        ("20260101", "总分", "275"),
        ("20260104", "数学", "94"),
        ("20260104", "年级排名", "3"),
        ("20260106", "英语", "75"),
    ],
    "single": [],
}

# ---------------------------------------------------------------- 5 制造 BOM
wb = Workbook()
ws = wb.active
ws.title = "BOM"
rows = [
    ["WL-1001", "伺服电机", "台", 2, "苏州精诚自动化", 4200, "A", "2026-08-15"],
    ["WL-1002", "减速机", "台", 2, "无锡恒信精密", 1850, "A", "2026-08-15"],
    ["WL-1003", "直线导轨", "根", 8, "南京华建工程", 620, "B", "2026-08-20"],
    ["WL-1004", "滚珠丝杠", "根", 4, "南京华建工程", 980, "B", "2026-08-20"],
    ["WL-1005", "控制主板", "块", 1, "杭州华锐电子", 2600, "A", "2026-08-10"],
    ["WL-1006", "铝合金机架", "套", 1, "常州天合材料", 5400, "C", "2026-09-01"],
]
fill_sheet(ws, ["物料编码", "物料名称", "单位", "单台用量", "供应商", "单价（元）", "关键等级", "要求到货日期"],
           rows, title="VMC-850B 数控加工中心 物料清单（BOM）")
wb.save(OUT / "05-制造-物料清单.xlsx")
answer["05-制造-物料清单.xlsx"] = {
    "industry": "制造",
    "objects": ["物料"],
    "attrs": ["物料名称", "单位", "单台用量", "供应商", "单价（元）", "关键等级", "要求到货日期"],
    "values": [
        ("WL-1001", "物料名称", "伺服电机"),
        ("WL-1001", "单价（元）", "4200"),
        ("WL-1005", "供应商", "杭州华锐电子"),
        ("WL-1006", "关键等级", "C"),
        ("WL-1004", "单台用量", "4"),
    ],
    "single": [],
}

# ---------------------------------------------------------------- 6 环保
d = Document()
para(d, "江苏鸿泰环境检测有限公司", 15, True, center=True)
para(d, "固定污染源废气监测报告", 14, True, center=True)
para(d, "报告编号：HT-2026-Q3-0271    委托单位：南京华建工程有限公司")
para(d, "监测类别：有组织废气    监测日期：2026-07-08    监测点位：1# 排气筒")
env = [
    ("颗粒物", "18.6", "mg/m³", "30", "达标"),
    ("二氧化硫", "42", "mg/m³", "100", "达标"),
    ("氮氧化物", "126", "mg/m³", "150", "达标"),
    ("非甲烷总烃", "68", "mg/m³", "60", "超标"),
]
table(d, ["监测因子", "浓度", "单位", "标准限值", "达标情况"], env)
para(d, "评价结论：除非甲烷总烃外，其余因子均满足《大气污染物综合排放标准》(GB 16297-1996) 要求；建议对活性炭吸附装置进行更换后复测。")
para(d, "编制人：许文静    审核人：高鹏    签发日期：2026-07-10")
d.save(OUT / "06-环保-监测报告.docx")
answer["06-环保-监测报告.docx"] = {
    "industry": "环保",
    "objects": ["监测因子"],
    "attrs": ["浓度", "单位", "标准限值", "达标情况"],
    "values": [
        ("颗粒物", "浓度", "18.6"),
        ("氮氧化物", "浓度", "126"),
        ("非甲烷总烃", "达标情况", "超标"),
        ("非甲烷总烃", "标准限值", "60"),
    ],
    "single": [("报告编号", "HT-2026-Q3-0271"), ("监测点位", "1# 排气筒")],
}

(OUT / "标准答案.json").write_text(json.dumps(answer, ensure_ascii=False, indent=2), encoding="utf-8")
print("已生成 %d 份文档 → %s" % (len(answer), OUT))
for k in answer:
    print("  ", k)
