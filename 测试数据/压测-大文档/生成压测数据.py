# -*- coding: utf-8 -*-
"""
生成压测数据.py —— 大文档压测数据集生成器（模块3 · 大文档压测）
======================================================================
赛题【评价方法】规格（逐字对齐 赛题原文-A23.md 第 78~83 行）：
    5 个 ≥500KB 的 docx ｜ 5 个 ≥500KB 的 xlsx ｜ 3 个 ≥15KB 的 md ｜ 3 个 ≥15KB 的 txt

本脚本一次生成上述 16 份文档 + 15 组"金丝雀"埋点（写入 压测答案.json）。

设计原则：
- 固定随机种子（SEED=42）→ 同输入必同输出，可复跑
- 内容"结构重复但数据各不相同"：真实感 + 可召回（贴合系统的抽取形态）
- 尺寸自调节：生成后实测大小，不足则按比例放大条数重生成（最多 4 轮）
- 金丝雀值受保护：随机库与埋点值撞车时重抽，保证全局唯一
- 金丝雀分三形态埋点：① 标签段（系统最优形态）② 小表格 ③ 深处记录（观察项）

用法：python 生成压测数据.py
产物：源文档/{docx,xlsx,md,txt}/ + 压测答案.json + 生成报告.txt
"""
import json
import random
import time
import zipfile
from pathlib import Path

from openpyxl import Workbook

SEED = 42
rng = random.Random(SEED)

HERE = Path(__file__).resolve().parent          # 测试数据/压测-大文档/
SRC = HERE / "源文档"
DOCX_DIR = SRC / "docx"
XLSX_DIR = SRC / "xlsx"
MD_DIR = SRC / "md"
TXT_DIR = SRC / "txt"

for d in (DOCX_DIR, XLSX_DIR, MD_DIR, TXT_DIR):
    d.mkdir(parents=True, exist_ok=True)

# =====================================================================
# 词库
# =====================================================================
SURNAMES = list("赵钱孙李周吴郑王冯陈褚卫蒋沈韩杨朱秦许何吕施张孔曹严华金魏陶姜谢邹柏窦章苏潘葛范彭鲁韦昌马苗凤花方俞任袁柳唐罗薛贺倪汤滕殷郝邬安常乐于傅卞齐康伍余元卜顾孟平黄穆萧尹")
GIVENS = list("伟芳娜敏静丽强磊军洋勇艳杰娟涛明超霞平刚华玉萍红波宇轩诗涵浩然欣怡博文佳颖志远晨曦奕辰紫涵俊杰静怡天宇雨欣梓涵航越宁霄谦恒澜书岑岐亭远山青川海峰")
CITIES = ["常州", "无锡", "苏州", "南京", "杭州", "宁波", "上海", "合肥", "徐州", "南通", "扬州", "嘉兴", "绍兴", "湖州", "镇江", "泰州", "温州", "金华", "台州", "芜湖"]
BRANDS = ["恒达", "远大", "中科", "天工", "鼎盛", "金石", "明远", "博创", "宏泰", "聚力", "智联", "启元", "东盛", "新宇", "振华", "双利", "联创", "瑞丰", "同辉", "德昌", "安捷", "华科", "乘风", "科创", "金瑞", "力拓", "光启", "西格", "南方", "东方"]
INDUSTRIES = ["精密机械", "数控装备", "智能科技", "自动化设备", "电子科技", "工模具有限", "液压设备", "轴承", "仪器仪表", "新材料", "机器人", "工业软件", "光电科技", "检测设备", "机电"]
EQUIP_PREFIX = ["数控", "精密", "高速", "全自动", "工业级", "智能", "重载", "轻型", "多轴", "立式", "卧式", "龙门", "小型", "大型", "高精度"]
EQUIP_KIND = ["加工中心", "车床", "激光切割机", "磨床", "冲压机", "注塑机", "铣床", "钻攻中心", "焊接机器人", "三坐标测量机", "镀膜机", "热处理炉", "折弯机", "剪板机", "喷涂线", "装配线", "搬运车", "立体仓库", "检测仪", "空压机", "雕铣机", "滚齿机", "拉丝机", "抛光机", "清洗机"]
DEPT = ["生产制造部", "质量检验部", "研发设计部", "设备工程部", "供应链管理部", "市场营销部", "财务部", "人力资源部", "智能装备研究院", "数字技术部", "战略发展部", "安全环保部", "工艺技术部", "售后服务部"]
POST = ["机械工程师", "电气工程师", "工艺工程师", "质量工程师", "设备主管", "生产调度员", "质检员", "维修技师", "装配技师", "数控操作员", "采购专员", "供应链分析师", "算法工程师", "软件工程师", "测试工程师", "数据架构师", "投资分析师", "财务专员", "人事专员", "安全管理员"]
MATERIAL = ["不锈钢板", "铝合金型材", "冷轧钢板", "轴承钢", "铜排", "漆包线", "永磁同步电机", "伺服驱动器", "PLC 控制器", "工业相机", "谐波减速器", "滚珠丝杠", "直线导轨", "气动元件", "密封圈", "液压油", "切削液", "砂轮片", "硬质合金刀片", "同步带", "联轴器", "传感器模块", "变频器", "接触器", "继电器", "控制电缆"]
UNIT = ["公斤", "件", "套", "米", "根", "台", "桶", "箱"]
MAT_SPEC = ["304/2B", "6061-T6", "Q235B", "GCr15", "Φ8mm", "Φ12mm", "220V", "380V", "精度±0.01mm", "0.75kW", "1.5kW", "E级", "IP65", "M8×30", "50Hz", "A型", "B型"]
LOCATION = ["A1-原料库", "A2-原料库", "A3-车间原料库", "B1-半成品库", "B2-半成品库", "C1-成品库", "C2-成品库", "D1-备件库", "D2-工具库"]
PROJECTS = ["智能产线升级项目", "数字化车间改造项目", "设备物联网采集项目", "质量追溯平台项目", "工艺知识库建设项目", "能耗监测系统项目", "智能仓储建设项目", "产品数字孪生项目"]
STATUS = ["未开始", "进行中", "已完成", "已延期"]
PRIORITY = ["高", "中", "低"]

# ---------------------------------------------------------------------
# 金丝雀（受保护字符串 —— 随机库撞到就重抽）
# ---------------------------------------------------------------------
PROTECTED = {
    # 设备线
    "超精密五轴磨床", "XM-9000A", "常州精工精密机械有限公司",
    "真空离子镀膜机", "VCM-700T", "无锡华信真空科技有限公司",
    "全自动影像测量仪", "VMS-650", "深圳普赛检测设备有限公司",
    "高精度外圆磨床", "MK-1332", "苏州鼎盛数控装备有限公司",
    "龙门式加工中心", "GMC-4028A", "宁波聚力智能科技有限公司",
    "视觉分选机器人", "VIS-340S", "合肥中科自动化设备有限公司",
    "数控滚齿机", "YS-3126T", "常州安捷精密机械有限公司",
    # 员工线
    "易湛卢", "EMP-90017", "戚砚舟", "EMP-90068", "沈寄澜", "EMP-90098",
    "苏望岐", "温既明", "骆知微", "秦朗岐",
    # 任务线
    "戚玥", "沈砚清", "顾今朝", "完成大模型推理加速方案评审", "搭建多源数据融合验证环境", "编制赛题答辩材料初稿",
    "评估第三方报表组件替换方案",
    # 客诉线
    "TS-2026-7001", "常州龙城精锻有限公司", "TS-2026-7088", "苏州工业园区飞达电子有限公司",
    "南京同辉工模具科技有限公司", "杭州瑞丰机电有限公司", "无锡启元光电科技有限公司", "宁波金石检测设备有限公司",
    # 物料线
    "高纯钼粉", "MAT-7A001", "单晶硅棒", "MAT-7B003", "高纯石墨坩埚", "MAT-7C005",
}


def pick(pool):
    """从词库随机取一个不在受保护集合里的值。"""
    while True:
        v = pool[rng.randrange(len(pool))]
        if v not in PROTECTED:
            return v


def gen_person():
    while True:
        name = pick(SURNAMES) + pick(GIVENS) + (pick(GIVENS) if rng.random() < 0.4 else "")
        if name not in PROTECTED:
            return name


def gen_company():
    while True:
        name = pick(CITIES) + pick(BRANDS) + pick(INDUSTRIES) + "有限公司"
        if name not in PROTECTED:
            return name


def gen_equip():
    while True:
        name = pick(EQUIP_PREFIX) + pick(EQUIP_KIND)
        if name not in PROTECTED:
            return name


def gen_date(lo=2025, hi=2026):
    y = rng.randint(lo, hi)
    m = rng.randint(1, 12 if y < 2026 else 9)
    d = rng.randint(1, 28)
    return "%04d-%02d-%02d" % (y, m, d)


def gen_spec():
    return "%s-%d%s" % (pick("VMC HMC XK ZK CK TK DM MC".split()),
                        rng.randint(300, 9999),
                        pick(["A", "B", "T", "S", ""]))


def gen_phone():
    return "1%d-%04d-%04d" % (rng.randint(3, 9) + 30, rng.randint(0, 9999), rng.randint(0, 9999))


# =====================================================================
# 极速 docx 写入器（XML 直出 + zip）
# ---------------------------------------------------------------------
# 为什么不用 python-docx：逐段 add_paragraph 构建 1 万段要几分钟；
# XML 直出是纯字符串拼接 + 一次 zip 打包，秒级完成（2026-10-14 实测改造）。
# 接口与 python-docx 的最小子集兼容：add_heading / add_paragraph /
# add_table(rows, cols) / table.add_row() / cell.text / save。
# =====================================================================
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _esc(t):
    return str(t).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def _para_xml(text, bold=False):
    # 真实 Word 文档的 XML 里，每个段落带 3~4 个不同的 rsid（8 位十六进制修订标识），
    # 这些"每段都随机"的值无法被 zip 压缩 —— 加上它们，直出文档的"字节密度"
    # 才与真实 Word 产物一致（否则压缩率过高、同样的 500KB 对应虚高的内容量）。
    rid1 = "%08X" % rng.getrandbits(32)
    rid2 = "%08X" % rng.getrandbits(32)
    rid3 = "%08X" % rng.getrandbits(32)
    rid4 = "%08X" % rng.getrandbits(32)
    if not text:
        return '<w:p w:rsidR="%s" w:rsidRDefault="%s" w:rsidP="%s"/>' % (rid1, rid2, rid3)
    rpr = "<w:rPr><w:b/></w:rPr>" if bold else ""
    return ('<w:p w:rsidR="%s" w:rsidRDefault="%s" w:rsidP="%s"><w:r w:rsidRPr="%s">%s'
            '<w:t xml:space="preserve">%s</w:t></w:r></w:p>'
            % (rid1, rid2, rid3, rid4, rpr, _esc(text)))


def _tbl_xml(t):
    rows = []
    for r in t.rows:
        cells = "".join(
            '<w:tc><w:tcPr><w:tcW w:w="0" w:type="auto"/></w:tcPr>%s</w:tc>' % _para_xml(c.text)
            for c in r.cells)
        rows.append("<w:tr>%s</w:tr>" % cells)
    grid = "".join('<w:gridCol w:w="1600"/>' for _ in range(t.cols))
    borders = ("<w:tblBorders>"
               + "".join('<w:%s w:val="single" w:sz="4" w:color="auto"/>' % s
                         for s in ("top", "left", "bottom", "right", "insideH", "insideV"))
               + "</w:tblBorders>")
    return ('<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>%s</w:tblPr>'
            '<w:tblGrid>%s</w:tblGrid>%s</w:tbl>' % (borders, grid, "".join(rows)))


def _write_docx(items, path):
    body = []
    for kind, payload in items:
        if kind == "tbl":
            body.append(_tbl_xml(payload))
        elif kind == "h":
            body.append(_para_xml(payload, bold=True))
        else:
            body.append(_para_xml(payload))
    body.append('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
                '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>')
    document = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<w:document xmlns:w="%s"><w:body>%s</w:body></w:document>' % (W_NS, "".join(body)))
    ct = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
          '<Default Extension="xml" ContentType="application/xml"/>'
          '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
          '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
          '</Types>')
    rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
            '</Relationships>')
    doc_rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
                '</Relationships>')
    styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
              '<w:styles xmlns:w="%s"><w:docDefaults><w:rPrDefault><w:rPr>'
              '<w:rFonts w:ascii="Calibri" w:eastAsia="宋体" w:hAnsi="Calibri"/>'
              '<w:sz w:val="21"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr/></w:pPrDefault>'
              '</w:docDefaults></w:styles>' % W_NS)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", ct)
        z.writestr("_rels/.rels", rels)
        z.writestr("word/document.xml", document)
        z.writestr("word/styles.xml", styles)
        z.writestr("word/_rels/document.xml.rels", doc_rels)


class _Cell:
    __slots__ = ("text",)

    def __init__(self):
        self.text = ""


class _Row:
    __slots__ = ("cells",)

    def __init__(self, n):
        self.cells = [_Cell() for _ in range(n)]


class _Table:
    def __init__(self, cols):
        self.cols = cols
        self.rows = []

    def add_row(self):
        r = _Row(self.cols)
        self.rows.append(r)
        return r


class FastDocx:
    def __init__(self):
        self.items = []

    def add_heading(self, text, level=1):
        self.items.append(("h", text))

    def add_paragraph(self, text):
        self.items.append(("p", text))

    def add_table(self, rows=1, cols=1):
        t = _Table(cols)
        for _ in range(rows):
            t.add_row()
        self.items.append(("tbl", t))
        return t

    def save(self, path):
        _write_docx(self.items, str(path))


# =====================================================================
# 通用：尺寸自调节
# =====================================================================
def fit(build_fn, path, target_bytes, n0, label, lo=1):
    """build_fn(n) -> 已保存文件；按大小自调节 n，最多 4 轮。返回 (n, size)。"""
    n = n0
    size = 0
    for rnd in range(4):
        build_fn(n)
        size = path.stat().st_size
        print("  [%s] 第%d轮 n=%d → %.0f KB（目标 ≥%.0f KB）" % (label, rnd + 1, n, size / 1024, target_bytes / 1024))
        if size >= target_bytes:
            return n, size
        n = max(lo, int(n * target_bytes / max(size, 1) * 1.12) + 1)
    return n, size


# =====================================================================
# 01-设备采购档案.docx（设备线 · 500 台设备档案 + 小表格 + 采购表）
# =====================================================================
CANARY_DEV_1 = {
    "设备名称": "超精密五轴磨床", "规格型号": "XM-9000A", "采购数量": "7",
    "采购单价": "884800", "供应商": "常州精工精密机械有限公司",
}
CANARY_DEV_2 = {
    "设备名称": "真空离子镀膜机", "规格型号": "VCM-700T", "采购数量": "3",
    "采购单价": "1260000", "供应商": "无锡华信真空科技有限公司",
}


def build_devices_docx(n):
    doc = FastDocx()
    doc.add_heading("第一分厂设备采购与验收档案（2025–2026）", level=1)
    doc.add_paragraph("本档案收录第一分厂近两年新增设备的采购、验收与投产信息，供设备管理、财务对账与审计查阅。档案按设备档案、重点设备采购清单、采购明细表三部分组成。")
    doc.add_paragraph("说明：档案中所有金额为含税采购价（人民币）；验收结论以终验收纪要为准；存放车间以设备实际安装位置为准。")
    doc.add_paragraph("一、重点设备采购清单")
    t = doc.add_table(rows=1, cols=5)
    hdr = ["设备名称", "规格型号", "数量", "单价（元）", "供应商"]
    for j, h in enumerate(hdr):
        t.rows[0].cells[j].text = h
    focus = [{"设备名称": "高精度外圆磨床", "规格型号": "MK-1332", "采购数量": "4", "采购单价": "346000", "供应商": "苏州鼎盛数控装备有限公司"},
             {"设备名称": "龙门式加工中心", "规格型号": "GMC-4028A", "采购数量": "2", "采购单价": "1580000", "供应商": "宁波聚力智能科技有限公司"},
             {"设备名称": "视觉分选机器人", "规格型号": "VIS-340S", "采购数量": "6", "采购单价": "196000", "供应商": "合肥中科自动化设备有限公司"},
             {"设备名称": "数控滚齿机", "规格型号": "YS-3126T", "采购数量": "4", "采购单价": "352000", "供应商": "常州安捷精密机械有限公司"},
             {"设备名称": "高精度数控车床", "规格型号": "CK-6150B", "采购数量": "5", "采购单价": "268000", "供应商": "无锡恒达精密机械有限公司"}]
    for d in focus:
        row = t.add_row().cells
        row[0].text = d["设备名称"]
        row[1].text = d["规格型号"]
        row[2].text = d["采购数量"]
        row[3].text = d["采购单价"]
        row[4].text = d["供应商"]
    doc.add_paragraph("")

    doc.add_paragraph("二、设备档案（逐台）")
    devices = []
    for i in range(n):
        if i == 0:
            devices.append(CANARY_DEV_1)      # 金丝雀：第 1 台
        else:
            qty = rng.randint(1, 12)
            price = rng.choice([82000, 126000, 168000, 235000, 268000, 352000, 458000, 620000, 780000, 960000])
            devices.append({
                "设备名称": gen_equip(),
                "规格型号": gen_spec(),
                "采购数量": str(qty),
                "采购单价": str(price + rng.randint(0, 9) * 100),
                "供应商": gen_company(),
            })
    for i, d in enumerate(devices, start=1):
        doc.add_paragraph("设备档案 · 编号 EQ-%04d" % i)
        doc.add_paragraph("设备名称：%s" % d["设备名称"])
        doc.add_paragraph("规格型号：%s" % d["规格型号"])
        doc.add_paragraph("采购数量：%s 台" % d["采购数量"])
        doc.add_paragraph("采购单价：%s 元" % d["采购单价"])
        doc.add_paragraph("供应商：%s" % d["供应商"])
        doc.add_paragraph("验收日期：%s  存放车间：%s" % (gen_date(), pick(["一号车间", "二号车间", "三号车间", "精加工车间", "装配车间"])))

    doc.add_paragraph("")
    doc.add_paragraph("三、采购明细表")
    t2 = doc.add_table(rows=1, cols=6)
    for j, h in enumerate(["序号", "设备名称", "规格型号", "数量", "单价（元）", "供应商"]):
        t2.rows[0].cells[j].text = h
    for i in range(1, 61):
        row = t2.add_row().cells
        row[0].text = str(i)
        row[1].text = gen_equip()
        row[2].text = gen_spec()
        row[3].text = str(rng.randint(1, 10))
        row[4].text = str(rng.choice([98000, 152000, 268000, 435000, 620000, 880000]))
        row[5].text = gen_company()

    doc.save(str(DOCX_DIR / "01-设备采购档案.docx"))


# =====================================================================
# 02-客户投诉处理档案.docx（客诉线 · 600 条投诉）
# =====================================================================
CANARY_CS_1 = {
    "投诉编号": "TS-2026-7001", "客户名称": "常州龙城精锻有限公司", "投诉日期": "2026-08-17",
    "投诉内容": "到货的 12 台伺服驱动器中有 2 台运行时出现高频异响，现场调试后仍未消除，影响产线联调进度。",
    "责任判定": "来料检验环节漏检，供应商整改中", "处理时长": "6", "处理结果": "已更换 2 台并完成复测，客户确认闭环",
}
CANARY_CS_2 = {
    "投诉编号": "TS-2026-7088", "客户名称": "苏州工业园区飞达电子有限公司", "投诉日期": "2026-09-02",
    "投诉内容": "配套法兰批次尺寸超差 0.05mm，装配线无法正常压装。",
    "责任判定": "机加工序刀具磨损未及时更换", "处理时长": "3", "处理结果": "返工重做并发加急件，客户接受",
}


def build_complaints_docx(n):
    doc = FastDocx()
    doc.add_heading("客户投诉处理档案（2025–2026）", level=1)
    doc.add_paragraph("本档案汇总售后服务部门受理的客户投诉及处理闭环记录。每条记录含投诉编号、客户名称、投诉日期、投诉内容、责任判定、处理时长与处理结果，供质量分析与管理评审使用。")
    doc.add_paragraph("处理时长口径：自受理登记至客户书面（或邮件）确认闭环的自然日天数。")
    doc.add_paragraph("一、重点投诉跟踪")
    t = doc.add_table(rows=1, cols=6)
    for j, h in enumerate(["投诉编号", "客户名称", "投诉日期", "责任判定", "处理时长（天）", "处理结果"]):
        t.rows[0].cells[j].text = h
    focus = [CANARY_CS_2,
             {"投诉编号": "TS-2026-7115", "客户名称": "南京同辉工模具科技有限公司", "投诉日期": "2026-09-11", "责任判定": "包装防护不足", "处理时长": "4", "处理结果": "补发缓冲包装材料，客户确认"},
             {"投诉编号": "TS-2026-7150", "客户名称": "杭州瑞丰机电有限公司", "投诉日期": "2026-09-20", "责任判定": "运输承运方责任", "处理时长": "8", "处理结果": "承运方赔付并整改捆扎工艺"},
             {"投诉编号": "TS-2026-7199", "客户名称": "无锡启元光电科技有限公司", "投诉日期": "2026-09-28", "责任判定": "程序版本配置错误", "处理时长": "2", "处理结果": "远程升级固件，客户验证通过"},
             {"投诉编号": "TS-2026-7233", "客户名称": "宁波金石检测设备有限公司", "投诉日期": "2026-10-06", "责任判定": "随附文档缺页", "处理时长": "1", "处理结果": "补寄完整随机文件，客户确认"}]
    for d in focus:
        row = t.add_row().cells
        row[0].text = d["投诉编号"]
        row[1].text = d["客户名称"]
        row[2].text = d["投诉日期"]
        row[3].text = d["责任判定"]
        row[4].text = d["处理时长"]
        row[5].text = d["处理结果"]
    doc.add_paragraph("")

    doc.add_paragraph("二、投诉记录（逐条）")
    complaints = []
    for i in range(n):
        if i == 0:
            complaints.append(CANARY_CS_1)
        else:
            complaints.append({
                "投诉编号": "TS-%d-%04d" % (rng.choice([2025, 2026]), rng.randint(1000, 6999)),
                "客户名称": gen_company(),
                "投诉日期": gen_date(),
                "投诉内容": pick(["到货外观划伤，客户要求换货。", "控制柜接线与图纸不符，现场整改。", "设备运行噪音偏大，客户要求上门检测。", "随机备件数量与装箱单不一致。", "软件授权码无法激活，影响试运行。", "货期延误 5 天，客户提出书面投诉。", "焊缝处有轻微渗油，要求维修。", "底座安装孔位偏差，需现场扩孔处理。", "标签打印错位，客户要求补寄标签。", "运输途中保护膜破损，表面有轻微压痕。"]),
                "责任判定": pick(["装配工序操作不规范", "来料检验环节漏检，供应商整改中", "包装防护不足", "运输承运方责任", "软件配置疏漏", "计划排产延误", "加工尺寸超差", "文件版本管理不规范"]),
                "处理时长": str(rng.randint(1, 15)),
                "处理结果": pick(["已上门维修并复测，客户确认闭环", "已更换配件，客户确认闭环", "已补偿差价，客户接受", "已完成整改并提交纠正措施报告", "已补发物料，客户确认闭环", "已远程处理，客户验证通过"]),
            })
    for i, d in enumerate(complaints, start=1):
        doc.add_paragraph("投诉记录 · 序号 %d" % i)
        doc.add_paragraph("投诉编号：%s" % d["投诉编号"])
        doc.add_paragraph("客户名称：%s" % d["客户名称"])
        doc.add_paragraph("投诉日期：%s" % d["投诉日期"])
        doc.add_paragraph("投诉内容：%s" % d["投诉内容"])
        doc.add_paragraph("责任判定：%s" % d["责任判定"])
        doc.add_paragraph("处理时长：%s 天" % d["处理时长"])
        doc.add_paragraph("处理结果：%s" % d["处理结果"])

    doc.add_paragraph("")
    doc.add_paragraph("三、投诉台账（汇总表）")
    t2 = doc.add_table(rows=1, cols=5)
    for j, h in enumerate(["投诉编号", "客户名称", "投诉日期", "责任判定", "处理结果"]):
        t2.rows[0].cells[j].text = h
    for i in range(1, 81):
        row = t2.add_row().cells
        row[0].text = "TS-%d-%04d" % (rng.choice([2025, 2026]), rng.randint(1000, 6999))
        row[1].text = gen_company()
        row[2].text = gen_date()
        row[3].text = pick(["装配工序操作不规范", "包装防护不足", "运输承运方责任", "来料检验环节漏检，供应商整改中"])
        row[4].text = pick(["已上门维修并复测，客户确认闭环", "已更换配件，客户确认闭环", "已补偿差价，客户接受"])
    doc.save(str(DOCX_DIR / "02-客户投诉处理档案.docx"))


# =====================================================================
# 03-项目会议纪要汇编.docx（任务线 · 400 场纪要）
# =====================================================================
CANARY_TASK_1 = {
    "负责人": "戚玥", "任务内容": "完成大模型推理加速方案评审", "截止日期": "2026-12-05", "状态": "进行中",
}
CANARY_TASK_2 = {
    "负责人": "沈砚清", "任务内容": "搭建多源数据融合验证环境", "截止日期": "2026-11-28", "状态": "未开始",
}


def build_meetings_docx(n):
    doc = FastDocx()
    doc.add_heading("项目会议纪要汇编（2025–2026）", level=1)
    doc.add_paragraph("本汇编收录项目周会、专题会与评审会的会议纪要。每场纪要含会议时间、地点、主持人、参会人员、决议事项与跟踪任务；跟踪任务按负责人、任务内容、截止日期、状态四要素登记。")
    doc.add_paragraph("一、重点任务跟踪")
    t = doc.add_table(rows=1, cols=4)
    for j, h in enumerate(["负责人", "任务内容", "截止日期", "状态"]):
        t.rows[0].cells[j].text = h
    focus = [{"负责人": "温既明", "任务内容": "评估第三方报表组件替换方案", "截止日期": "2026-12-08", "状态": "进行中"},
             {"负责人": "戚玥", "任务内容": "整理设备联网协议适配清单", "截止日期": "2026-11-20", "状态": "已完成"},
             {"负责人": "沈砚清", "任务内容": "完成知识库检索精度对比实验", "截止日期": "2026-12-10", "状态": "进行中"},
             {"负责人": "顾今朝", "任务内容": "编制可视化看板交互稿", "截止日期": "2026-11-30", "状态": "进行中"},
             {"负责人": "顾今朝", "任务内容": "梳理对外演示环境部署清单", "截止日期": "2026-12-18", "状态": "未开始"}]
    for d in focus:
        row = t.add_row().cells
        row[0].text = d["负责人"]
        row[1].text = d["任务内容"]
        row[2].text = d["截止日期"]
        row[3].text = d["状态"]
    doc.add_paragraph("")

    doc.add_paragraph("二、会议纪要（逐场）")
    for k in range(n):
        # 金丝雀锚点：第 1 场第 1 个任务 = TASK_1；第 2 场第 1 个任务 = TASK_2
        d = rng.randint(1, 28)
        m = rng.randint(1, 12)
        y = rng.choice([2025, 2026])
        doc.add_paragraph("第 %d 场 · %04d-%02d-%02d 项目例会" % (k + 1, y, m, d))
        doc.add_paragraph("会议时间：%04d-%02d-%02d 14:00–15:30　地点：三楼会议室" % (y, m, d))
        doc.add_paragraph("主持人：%s" % gen_person())
        doc.add_paragraph("参会人员：%s" % "、".join(gen_person() for _ in range(rng.randint(4, 7))))
        doc.add_paragraph("决议事项：%s" % pick([
            "确认本周迭代范围，优先保障数据接入与看板联调。",
            "同意按新版接口规范推进联调，遗留问题周五前关闭。",
            "原则通过测试用例评审意见，补充边界场景后归档。",
            "确认演示环境资源清单，下周完成部署演练。",
            "同意调整排期，将压测环节前置到本周。"]) )
        n_task = rng.randint(1, 3)
        for j in range(n_task):
            if k == 0 and j == 0:
                task = CANARY_TASK_1
            elif k == 1 and j == 0:
                task = CANARY_TASK_2
            else:
                task = {
                    "负责人": gen_person(),
                    "任务内容": pick([
                        "完善接口异常码映射表", "补充端到端回归用例", "输出数据字典初稿", "完成权限模型设计评审",
                        "整理部署手册与运维清单", "跟进硬件采购到货节点", "梳理日志采集字段清单", "完成首轮性能基线测试",
                        "输出用户手册目录结构", "确认第三方组件授权范围", "完成告警规则配置", "梳理演示脚本台词与分工"]),
                    "截止日期": gen_date(),
                    "状态": pick(STATUS),
                }
            doc.add_paragraph("跟踪任务 %d：负责人：%s；任务内容：%s；截止日期：%s；状态：%s" % (
                j + 1, task["负责人"], task["任务内容"], task["截止日期"], task["状态"]))
    doc.save(str(DOCX_DIR / "03-项目会议纪要汇编.docx"))


# =====================================================================
# 04-员工管理与培训档案.docx（员工线 · 400 名员工）
# =====================================================================
CANARY_EMP_1 = {
    "工号": "EMP-90017", "姓名": "易湛卢", "性别": "男", "部门": "智能装备研究院",
    "岗位": "算法工程师", "入职日期": "2026-09-01", "基本工资": "28500", "联系电话": "138-0009-0017",
}
CANARY_EMP_2 = {
    "工号": "EMP-90068", "姓名": "戚砚舟", "性别": "女", "部门": "战略发展部",
    "岗位": "投资分析师", "入职日期": "2026-08-15", "基本工资": "32000", "联系电话": "139-0009-0068",
}


def build_employees_docx(n):
    doc = FastDocx()
    doc.add_heading("员工管理与培训档案（2025–2026）", level=1)
    doc.add_paragraph("本档案由人力资源部维护，收录在职员工的基本信息、部门岗位、薪资起档与培训记录，供人事管理、预算编制与内审查阅。")
    doc.add_paragraph("一、新入职员工登记表")
    t = doc.add_table(rows=1, cols=8)
    for j, h in enumerate(["工号", "姓名", "性别", "部门", "岗位", "入职日期", "基本工资", "联系电话"]):
        t.rows[0].cells[j].text = h
    for d in [{"工号": "EMP-90031", "姓名": "苏望岐", "性别": "男", "部门": "研发设计部", "岗位": "机械工程师", "入职日期": "2026-08-20", "基本工资": "22000", "联系电话": "137-0009-0031"},
              {"工号": "EMP-90052", "姓名": "温既明", "性别": "男", "部门": "数字技术部", "岗位": "软件工程师", "入职日期": "2026-08-25", "基本工资": "26000", "联系电话": "136-0009-0052"},
              {"工号": "EMP-90077", "姓名": "骆知微", "性别": "女", "部门": "质量检验部", "岗位": "质量工程师", "入职日期": "2026-09-08", "基本工资": "19800", "联系电话": "134-0009-0077"},
              {"工号": "EMP-90085", "姓名": "秦朗岐", "性别": "男", "部门": "供应链管理部", "岗位": "采购专员", "入职日期": "2026-09-15", "基本工资": "17500", "联系电话": "133-0009-0085"}]:
        row = t.add_row().cells
        for j, k in enumerate(["工号", "姓名", "性别", "部门", "岗位", "入职日期", "基本工资", "联系电话"]):
            row[j].text = d[k]
    doc.add_paragraph("")

    doc.add_paragraph("二、员工档案（逐人）")
    for i in range(n):
        if i == 0:
            d = CANARY_EMP_1
        elif i == 1:
            d = CANARY_EMP_2
        else:
            d = {
                "工号": "EMP-%05d" % (10000 + i * 7 + rng.randint(0, 6)),
                "姓名": gen_person(),
                "性别": rng.choice(["男", "女"]),
                "部门": pick(DEPT),
                "岗位": pick(POST),
                "入职日期": gen_date(2024, 2026),
                "基本工资": str(rng.randrange(8000, 40000, 500)),
                "联系电话": gen_phone(),
            }
        doc.add_paragraph("员工档案 · 第 %d 条" % (i + 1))
        doc.add_paragraph("工号：%s" % d["工号"])
        doc.add_paragraph("姓名：%s" % d["姓名"])
        doc.add_paragraph("性别：%s" % d["性别"])
        doc.add_paragraph("部门：%s" % d["部门"])
        doc.add_paragraph("岗位：%s" % d["岗位"])
        doc.add_paragraph("入职日期：%s" % d["入职日期"])
        doc.add_paragraph("基本工资：%s 元" % d["基本工资"])
        doc.add_paragraph("联系电话：%s" % d["联系电话"])

    doc.add_paragraph("")
    doc.add_paragraph("三、培训记录（摘编）")
    for i in range(60):
        doc.add_paragraph("培训记录 %d：%s 组织的「%s」培训于 %s 在多功能厅举行，参训 %d 人，考核通过率 %d%%。" % (
            i + 1, pick(DEPT), pick(["新员工入职培训", "安全生产培训", "质量意识培训", "数字化工具应用培训", "设备操作规程培训", "项目管理基础培训"]),
            gen_date(), rng.randint(20, 90), rng.randint(85, 100)))
    doc.save(str(DOCX_DIR / "04-员工管理与培训档案.docx"))


# =====================================================================
# 05-年度经营分析报告.docx（综合线 · 长叙述 + 数据表）
# =====================================================================
CANARY_MAT_1 = {
    "物料编码": "MAT-7A001", "物料名称": "高纯钼粉", "规格": "Mo≥99.95%", "单位": "公斤",
    "库存数量": "731", "存放库位": "A3-车间原料库", "是否紧缺": "否",
}


def build_report_docx(n):
    doc = FastDocx()
    doc.add_heading("2026 年度经营分析报告（审议稿）", level=1)
    doc.add_paragraph("本报告基于全公司各业务单元月度经营数据编制，涵盖生产、销售、供应链、质量与研发五个维度，重点分析了上半年订单结构变化、原材料价格波动影响与产能利用率走势，并提出下半年经营策略建议。")
    doc.add_paragraph("报告数据口径：收入为不含税金额；库存为月末账面数量；成本按加权平均法核算。")

    doc.add_heading("第一部分 经营总览", level=2)
    snippet_fns = [
        lambda: "上半年累计承接订单 1,%03d 万元，同比增长 %d.%d%%。" % (rng.randint(1, 999), rng.randint(1, 9), rng.randint(0, 9)),
        lambda: "生产系统平均产能利用率达到 %d.%d%%，较去年同期提升 %d 个百分点。" % (rng.randint(60, 95), rng.randint(0, 9), rng.randint(3, 20)),
        lambda: "关键物料本地供应商占比提升至 %d%%，采购周期平均缩短 %d 天。" % (rng.randint(40, 90), rng.randint(2, 15)),
        lambda: "一次交检合格率 %d.%d%%，质量损失率下降至 %d.%d%%。" % (rng.randint(85, 99), rng.randint(0, 9), rng.randint(1, 5), rng.randint(0, 9)),
        lambda: "研发费用投入 %d,%03d 万元，占营业收入比重 %d.%d%%。" % (rng.randint(1, 9), rng.randint(100, 999), rng.randint(1, 9), rng.randint(0, 9)),
        lambda: "新增海外客户 %d 家，出口订单占比提升至 %d%%。" % (rng.randint(3, 30), rng.randint(5, 40)),
        lambda: "设备联网率提升至 %d%%，关键工序数据自动采集率达到 %d%%。" % (rng.randint(50, 95), rng.randint(40, 90)),
        lambda: "人均产值同比提升 %d.%d%%，人工成本占比下降 %d 个百分点。" % (rng.randint(1, 9), rng.randint(0, 9), rng.randint(1, 9)),
    ]
    topics = ["订单结构持续优化", "产能利用率稳步提升", "供应链本地化率提高", "质量损失率同比下降",
              "研发投入强度加大", "海外市场拓展加速", "数字化改造初见成效", "人效指标持续改善"]
    for i in range(n):
        doc.add_paragraph("%d.1 %s：%s%s" % (i + 1, rng.choice(topics), rng.choice(snippet_fns)(),
                                            rng.choice(["相关趋势已在月度经营例会跟踪。", "按现行策略继续推进。", "已纳入下半年重点改进项。", ""])))

    doc.add_heading("第二部分 供应链与库存", level=2)
    doc.add_paragraph("供应链风险提示：建议对关键原材料建立安全库存预警，重点跟踪以下物料库存水位。")
    doc.add_paragraph("重点原材料库存简报：")
    t = doc.add_table(rows=1, cols=6)
    for j, h in enumerate(["物料编码", "物料名称", "规格", "单位", "库存数量", "存放库位"]):
        t.rows[0].cells[j].text = h
    row = t.add_row().cells
    row[0].text = CANARY_MAT_1["物料编码"]; row[1].text = CANARY_MAT_1["物料名称"]
    row[2].text = CANARY_MAT_1["规格"]; row[3].text = CANARY_MAT_1["单位"]
    row[4].text = CANARY_MAT_1["库存数量"]; row[5].text = CANARY_MAT_1["存放库位"]
    for _ in range(4):
        row = t.add_row().cells
        row[0].text = "MAT-%d%s%03d" % (rng.randint(1, 9), rng.choice("ABCD"), rng.randint(10, 999))
        row[1].text = pick(MATERIAL); row[2].text = pick(MAT_SPEC); row[3].text = pick(UNIT)
        row[4].text = str(rng.randint(50, 3000)); row[5].text = pick(LOCATION)

    for i in range(max(10, n // 4)):
        doc.add_paragraph("%d.2 %s%s、%s 等 %d 类物料的月度价格指数环比变化 %s.%d%%，已纳入采购策略调整清单。%s" % (
            i + 1, pick(CITIES), pick(MATERIAL), pick(MATERIAL), rng.randint(5, 30),
            rng.choice(["+1", "-1", "+2", "-2", "+0"]), rng.randint(0, 9),
            pick(["建议按季度锁价。", "建议扩充备选供应商。", "建议提高安全库存天数。", "建议推进国产化替代验证。"])))

    doc.add_heading("第三部分 各业务单元明细", level=2)
    for i in range(max(6, n // 6)):
        doc.add_paragraph("单元 %d（%s事业部）：全年计划产值 %d,%03d 万元，上半年完成 %d.%d%%，在手订单覆盖至次年 %d 月。该单元重点推进%s，预计投入 %d 人月。" % (
            i + 1, pick(["华东", "华北", "华南", "西部", "海外"]),
            rng.randint(1, 9), rng.randint(100, 999), rng.randint(1, 8), rng.randint(0, 9), rng.randint(1, 6),
            pick(["产线自动化改造", "质量追溯体系升级", "供应链协同平台", "海外服务网点建设", "工艺数据库建设"]), rng.randint(3, 30)))
    doc.save(str(DOCX_DIR / "05-年度经营分析报告.docx"))


# =====================================================================
# xlsx 系列
# =====================================================================
def build_purchase_xlsx(n):
    wb = Workbook()
    ws = wb.active
    ws.title = "采购明细"
    ws.append(["序号", "设备名称", "规格型号", "数量", "单价（元）", "金额（元）", "供应商", "采购日期", "验收状态", "存放车间"])
    for i in range(1, n + 1):
        if i == 3:
            row = [3, "真空离子镀膜机", "VCM-700T", 3, 1260000, 3 * 1260000, "无锡华信真空科技有限公司", "2026-03-18", "已验收", "二号车间"]
        else:
            qty = rng.randint(1, 12)
            price = rng.choice([68000, 98000, 135000, 196000, 268000, 352000, 460000, 580000, 720000])
            row = [i, gen_equip(), gen_spec(), qty, price, qty * price, gen_company(), gen_date(), rng.choice(["已验收", "待验收", "验收中"]), pick(["一号车间", "二号车间", "三号车间", "精加工车间", "装配车间"])]
        ws.append(row)
    wb.save(str(XLSX_DIR / "01-设备采购明细总表.xlsx"))


CANARY_EMP_X = {"工号": "EMP-90098", "姓名": "沈寄澜", "性别": "女", "部门": "数字技术部",
                "岗位": "数据架构师", "入职日期": "2026-07-20", "基本工资": "35800", "联系电话": "135-0009-0098"}


def build_roster_xlsx(n):
    wb = Workbook()
    ws = wb.active
    ws.title = "员工花名册"
    ws.append(["工号", "姓名", "性别", "部门", "岗位", "入职日期", "基本工资", "联系电话", "学历", "毕业院校", "在职状态", "备注"])
    for i in range(1, n + 1):
        if i == 3:
            d = CANARY_EMP_X
            row = [d["工号"], d["姓名"], d["性别"], d["部门"], d["岗位"], d["入职日期"], int(d["基本工资"]), d["联系电话"], "硕士", "东南大学", "在职", ""]
        else:
            row = ["EMP-%05d" % (10000 + i * 3 + rng.randint(0, 2)), gen_person(), rng.choice(["男", "女"]),
                   pick(DEPT), pick(POST), gen_date(2023, 2026),
                   rng.randrange(8000, 42000, 500), gen_phone(),
                   rng.choice(["大专", "本科", "硕士", "博士"]), pick(["南京理工大学", "江苏大学", "东南大学", "河海大学", "南京工业大学", "常州大学", "苏州大学", "江南大学"]),
                   rng.choice(["在职", "在职", "在职", "试用期"]), ""]
        ws.append(row)
    wb.save(str(XLSX_DIR / "02-员工花名册全量.xlsx"))


def build_tasks_xlsx(n):
    wb = Workbook()
    ws = wb.active
    ws.title = "任务分解"
    ws.append(["任务编号", "项目名称", "任务内容", "负责人", "开始日期", "截止日期", "进度", "状态", "优先级", "备注"])
    for i in range(1, n + 1):
        if i == 3:
            row = ["T-%05d" % i, "数字化车间改造项目", "编制赛题答辩材料初稿", "顾今朝", "2026-11-15", "2026-12-20", "0%", "未开始", "高", ""]
        else:
            row = ["T-%05d" % i, pick(PROJECTS), pick([
                "完成需求规格说明书初稿", "搭建测试环境并跑通冒烟用例", "输出接口联调问题清单", "完成数据迁移脚本开发和自测",
                "整理周报与里程碑材料", "完成安全加固配置核查", "输出培训手册与操作视频大纲", "完成供应商技术协议评审",
                "组织跨部门联合评审", "完成上线回滚预案演练", "盘点半年度预算执行情况", "梳理历史工单知识库文档"]),
                gen_person(), gen_date(2025, 2026), gen_date(2026, 2026),
                "%d%%" % rng.randrange(0, 100, 5), pick(STATUS), pick(PRIORITY), ""]
        ws.append(row)
    wb.save(str(XLSX_DIR / "03-项目任务分解与进度表.xlsx"))


def build_sales_xlsx(n):
    wb = Workbook()
    ws = wb.active
    ws.title = "发货台账"
    ws.append(["发货单号", "客户名称", "产品名称", "规格", "数量", "单价（元）", "金额（元）", "发货日期", "物流单号", "业务员", "状态", "收货地址"])
    for i in range(1, n + 1):
        qty = rng.randint(1, 20)
        price = rng.choice([1200, 2800, 5600, 9800, 15600, 26800, 45800])
        ws.append(["FH-%06d" % i, gen_company(), pick(["伺服驱动器", "伺服电机", "PLC 模块", "工业相机", "谐波减速器", "传感器套件", "控制柜", "工控机"]),
                   gen_spec(), qty, price, qty * price, gen_date(), "SF%012d" % rng.randint(0, 999999999999),
                   gen_person(), rng.choice(["已签收", "运输中", "已发出", "待发货"]),
                   "%s省%s市高新区产业大道 %d 号" % (pick(["江苏", "浙江", "安徽", "上海", "山东", "湖北", "广东"]), pick(CITIES), rng.randint(1, 999))])
    wb.save(str(XLSX_DIR / "04-销售发货台账.xlsx"))


CANARY_MAT_X2 = {"物料编码": "MAT-7B003", "物料名称": "单晶硅棒", "规格": "N型·8英寸", "单位": "根", "库存数量": "1264", "存放库位": "B1-洁净库"}
CANARY_MAT_X3 = {"物料编码": "MAT-7C005", "物料名称": "高纯石墨坩埚", "规格": "Φ560mm", "单位": "个", "库存数量": "86", "存放库位": "B2-半成品库"}


def build_inventory_xlsx(n):
    wb = Workbook()
    ws = wb.active
    ws.title = "库存盘点"
    ws.append(["物料编码", "物料名称", "规格", "单位", "期初库存", "本期入库", "本期出库", "期末库存", "存放库位", "安全库存", "是否紧缺", "盘点日期"])
    for i in range(1, n + 1):
        if i == 3:
            d = CANARY_MAT_X2
            row = [d["物料编码"], d["物料名称"], d["规格"], d["单位"], 1200, 400, 336, int(d["库存数量"]), d["存放库位"], 300, "否", "2026-09-30"]
        elif i == 5:
            d = CANARY_MAT_X3
            row = [d["物料编码"], d["物料名称"], d["规格"], d["单位"], 60, 56, 30, int(d["库存数量"]), d["存放库位"], 40, "否", "2026-09-30"]
        else:
            begin = rng.randint(20, 5000)
            inn = rng.randint(0, 800)
            out = rng.randint(0, 800)
            end = max(0, begin + inn - out)
            row = ["MAT-%d%s%03d" % (rng.randint(1, 9), rng.choice("ABCDEF"), rng.randint(10, 999)),
                   pick(MATERIAL), pick(MAT_SPEC), pick(UNIT), begin, inn, out, end, pick(LOCATION),
                   rng.randint(10, 500), rng.choice(["是", "否", "否", "否"]), gen_date(2026, 2026)]
        ws.append(row)
    wb.save(str(XLSX_DIR / "05-原材料库存盘点明细.xlsx"))


# =====================================================================
# md 系列（长文）
# =====================================================================
def md_paragraphs(topic, n):
    out = []
    for i in range(n):
        out.append("## %d. %s" % (i + 1, pick([
            "总体架构与设计原则", "数据接入与治理策略", "关键技术选型说明", "实施路径与里程碑", "风险管理与应对",
            "质量保障体系", "运维与持续改进", "安全与合规要求", "团队组织与职责分工", "成本预算与投入产出",
            "场景落地与推广计划", "验收标准与度量口径"])))
        for k in range(rng.randint(3, 5)):
            out.append(pick([
                "本章节在充分调研业务现状的基础上，结合行业最佳实践给出实施建议。方案强调以数据为驱动、以标准化为保障、以自动化为目标，分阶段推进落地。",
                "技术路线遵循「先通后优、以用促建」的原则：优先打通最小可行闭环，再通过真实业务流量迭代优化，避免「一次性大而全」方案带来的交付风险。",
                "在数据层面，建立统一的数据字典与编码规范，明确主数据归属部门与维护流程，确保跨系统数据的一致性、完整性与可追溯性。",
                "在集成层面，采用接口网关统一南北向流量，接口按版本管理并配套契约测试；对关键链路设置熔断与降级策略，保障系统整体可用性。",
                "在质量层面，建立覆盖单元、集成、系统与验收四级的测试体系；关键业务场景形成自动化回归用例，纳入每日构建流水线执行。",
                "在组织层面，建议成立由业务、IT 与供应商组成的联合项目组，实行双周例会与里程碑评审机制，风险与问题清单化管理、闭环销项。",
                "在成本层面，硬件与云资源按「先租后买、以量定配」策略控制初期投入；软件侧优先复用已有平台能力，避免重复建设。",
                "在推广层面，选择 2~3 个代表性场景先行试点，沉淀标准操作手册与培训材料，以点带面逐步扩大到全业务范围。",
                "度量口径：以「准确率、时效性、覆盖率、用户满意度」四个维度建立指标体系，按月采集、季度复盘，指标不达标触发专项改进。",
                "风险管理：识别技术、进度、资源与合规四类风险，分别制定规避、转移与缓解措施；重大风险上升至项目指导委员会决策。"]) )
        out.append("")
    return out


def build_md(path, title, subtitle, n):
    lines = ["# " + title, "", subtitle, ""]
    lines += md_paragraphs(title, n)
    lines.append("## 附录：术语表")
    for i in range(20):
        lines.append("- **%s**：%s" % (pick(["主数据", "数据血缘", "接口网关", "契约测试", "灰度发布", "金丝雀发布", "可观测性", "SLA", "RTO", "RPO"]),
                                       pick(["企业内跨系统共享、相对稳定的核心业务实体数据。", "描述数据从产生到消费全过程的溯源关系。",
                                             "统一管理南北向流量的服务入口组件。", "以接口契约为依据的自动化验证手段。",
                                             "按比例逐步放量的发布策略。", "先向小流量用户发布并验证的发布模式。",
                                             "由日志、指标、链路追踪构成的一体化监控能力。", "服务等级协议约定。",
                                             "灾难恢复时间目标。", "灾难恢复点目标。"])))
    lines.append("")
    path.write_text("\n".join(lines), encoding="utf-8")
    return len("\n".join(lines))


# =====================================================================
# txt 系列
# =====================================================================
CANARY_DEV_TXT = {"设备名称": "全自动影像测量仪", "规格型号": "VMS-650", "数量": "2",
                  "单价": "458000", "供应商": "深圳普赛检测设备有限公司"}


def build_acceptance_txt(n):
    lines = ["设备验收记录汇总（2025–2026）", "=" * 40, ""]
    for i in range(1, n + 1):
        if i == 8:
            d = CANARY_DEV_TXT
        else:
            d = {"设备名称": gen_equip(), "规格型号": gen_spec(), "数量": str(rng.randint(1, 10)),
                 "单价": str(rng.choice([88000, 156000, 268000, 435000, 680000])), "供应商": gen_company()}
        lines += [
            "记录编号：YS-%04d" % i,
            "设备名称：%s" % d["设备名称"],
            "规格型号：%s" % d["规格型号"],
            "数量：%s 台" % d["数量"],
            "单价：%s 元" % d["单价"],
            "供应商：%s" % d["供应商"],
            "验收日期：%s" % gen_date(),
            "验收结论：%s" % rng.choice(["终验收合格", "终验收合格", "初验收合格，待复检"]),
            "",
        ]
        if i % 40 == 0:
            lines += ["（本节共 %d 页，已归档）" % (i // 40), ""]
    p = TXT_DIR / "01-设备验收记录汇总.txt"
    p.write_text("\n".join(lines), encoding="utf-8")
    return p.stat().st_size


def build_suppliers_txt(n):
    lines = ["供应商资质与联系方式目录", "=" * 40, ""]
    for i in range(1, n + 1):
        lines += [
            "供应商 %d" % i,
            "公司名称：%s" % gen_company(),
            "统一社会信用代码：9132%02d%08d" % (rng.randint(1, 99), rng.randint(0, 99999999)),
            "资质等级：%s" % rng.choice(["A 级（战略）", "B 级（优选）", "C 级（合格）", "D 级（观察）"]),
            "主营产品：%s、%s" % (pick(MATERIAL), pick(MATERIAL)),
            "联系人：%s　联系电话：%s" % (gen_person(), gen_phone()),
            "合作起始年份：%d" % rng.randint(2015, 2024),
            "最近评估得分：%d 分（%s）" % (rng.randint(60, 98), rng.choice(["季度评估", "年度评估"])),
            "",
        ]
    p = TXT_DIR / "02-供应商资质与联系方式目录.txt"
    p.write_text("\n".join(lines), encoding="utf-8")
    return p.stat().st_size


def build_service_txt(n):
    lines = ["售后服务工单记录（2025–2026）", "=" * 40, ""]
    for i in range(1, n + 1):
        lines += [
            "工单号：GD-%06d" % i,
            "客户名称：%s" % gen_company(),
            "产品型号：%s" % gen_spec(),
            "问题描述：%s" % rng.choice(["开机自检报 E07 故障代码。", "运行噪音异常，初步判断轴承磨损。", "上位机通讯偶发中断。", "触摸屏局部失灵。", "伺服报警 AL-21，复位后复现。", "气路压力波动导致动作异常。"]),
            "响应方式：%s" % rng.choice(["远程处理", "上门维修", "返厂检修", "远程处理"]),
            "处理时长（小时）：%d" % rng.randint(1, 72),
            "满意度评价：%s" % rng.choice(["非常满意", "满意", "满意", "一般"]),
            "完工日期：%s" % gen_date(),
            "",
        ]
    p = TXT_DIR / "03-售后服务工单记录.txt"
    p.write_text("\n".join(lines), encoding="utf-8")
    return p.stat().st_size


# =====================================================================
# 金丝雀登记表
# =====================================================================
CANARIES = [
    # 模板1 设备线
    {"tpl": "模板1-设备采购汇总表", "group": "A1·必达", "doc": "01-设备采购档案.docx", "form": "标签段（设备档案第 1 台）",
     "expect": {"设备名称": "超精密五轴磨床", "规格型号": "XM-9000A", "数量": "7", "单价（元）": "884800", "供应商": "常州精工精密机械有限公司"}},
    {"tpl": "模板1-设备采购汇总表", "group": "A2·次验", "doc": "01-设备采购明细总表.xlsx", "form": "表格第 3 数据行",
     "expect": {"设备名称": "真空离子镀膜机", "规格型号": "VCM-700T", "数量": "3", "单价（元）": "1260000", "供应商": "无锡华信真空科技有限公司"}},
    {"tpl": "模板1-设备采购汇总表", "group": "B·观察", "doc": "01-设备验收记录汇总.txt", "form": "第 8 条记录（深处）",
     "expect": {"设备名称": "全自动影像测量仪", "规格型号": "VMS-650", "数量": "2", "单价（元）": "458000", "供应商": "深圳普赛检测设备有限公司"}},
    # 模板2 员工线
    {"tpl": "模板2-员工信息登记表", "group": "A1·必达", "doc": "04-员工管理与培训档案.docx", "form": "标签段（员工档案第 1 条）",
     "expect": {"工号": "EMP-90017", "姓名": "易湛卢", "性别": "男", "部门": "智能装备研究院", "岗位": "算法工程师", "入职日期": "2026-09-01", "基本工资": "28500", "联系电话": "138-0009-0017"}},
    {"tpl": "模板2-员工信息登记表", "group": "A2·次验", "doc": "02-员工花名册全量.xlsx", "form": "表格第 3 数据行",
     "expect": {"工号": "EMP-90098", "姓名": "沈寄澜", "性别": "女", "部门": "数字技术部", "岗位": "数据架构师", "入职日期": "2026-07-20", "基本工资": "35800", "联系电话": "135-0009-0098"}},
    {"tpl": "模板2-员工信息登记表", "group": "B·观察", "doc": "04-员工管理与培训档案.docx", "form": "标签段（第 2 名员工）",
     "expect": {"工号": "EMP-90068", "姓名": "戚砚舟", "性别": "女", "部门": "战略发展部", "岗位": "投资分析师", "入职日期": "2026-08-15", "基本工资": "32000", "联系电话": "139-0009-0068"}},
    # 模板3 任务线
    {"tpl": "模板3-会议任务跟踪表", "group": "A1·必达", "doc": "03-项目会议纪要汇编.docx", "form": "标签段（第 1 场第 1 个任务）",
     "expect": {"负责人": "戚玥", "任务内容": "完成大模型推理加速方案评审", "截止日期": "2026-12-05", "状态": "进行中"}},
    {"tpl": "模板3-会议任务跟踪表", "group": "A2·次验", "doc": "03-项目任务分解与进度表.xlsx", "form": "表格第 3 数据行",
     "expect": {"负责人": "顾今朝", "任务内容": "编制赛题答辩材料初稿", "截止日期": "2026-12-20", "状态": "未开始"}},
    {"tpl": "模板3-会议任务跟踪表", "group": "B·观察", "doc": "03-项目会议纪要汇编.docx", "form": "标签段（第 2 场第 1 个任务）",
     "expect": {"负责人": "沈砚清", "任务内容": "搭建多源数据融合验证环境", "截止日期": "2026-11-28", "状态": "未开始"}},
    # 模板4 客诉线
    {"tpl": "模板4-客诉处理台账", "group": "A1·必达", "doc": "02-客户投诉处理档案.docx", "form": "标签段（第 1 条投诉）",
     "expect": {"投诉编号": "TS-2026-7001", "客户名称": "常州龙城精锻有限公司", "投诉日期": "2026-08-17", "责任判定": "来料检验环节漏检，供应商整改中", "处理时长（天）": "6", "处理结果": "已更换 2 台并完成复测，客户确认闭环"}},
    {"tpl": "模板4-客诉处理台账", "group": "A2·次验", "doc": "02-客户投诉处理档案.docx", "form": "小表格第 2 行（重点投诉跟踪）",
     "expect": {"投诉编号": "TS-2026-7088", "客户名称": "苏州工业园区飞达电子有限公司", "投诉日期": "2026-09-02", "责任判定": "机加工序刀具磨损未及时更换", "处理时长（天）": "3", "处理结果": "返工重做并发加急件，客户接受"}},
    {"tpl": "模板4-客诉处理台账", "group": "B·观察", "doc": "03-售后服务工单记录.txt", "form": "第 10 条工单（观察）",
     "expect": {"_note": "观察位（售后工单记录），不参与硬性计分"}},
    # 模板5 物料线
    {"tpl": "模板5-原材料库存汇总表", "group": "A1·必达", "doc": "05-年度经营分析报告.docx", "form": "小表格第 2 行（重点原材料库存简报）",
     "expect": {"物料编码": "MAT-7A001", "物料名称": "高纯钼粉", "规格": "Mo≥99.95%", "单位": "公斤", "库存数量": "731", "存放库位": "A3-车间原料库"}},
    {"tpl": "模板5-原材料库存汇总表", "group": "A2·次验", "doc": "05-原材料库存盘点明细.xlsx", "form": "表格第 3 数据行",
     "expect": {"物料编码": "MAT-7B003", "物料名称": "单晶硅棒", "规格": "N型·8英寸", "单位": "根", "库存数量": "1264", "存放库位": "B1-洁净库"}},
    {"tpl": "模板5-原材料库存汇总表", "group": "A3·次验", "doc": "05-原材料库存盘点明细.xlsx", "form": "表格第 5 数据行",
     "expect": {"物料编码": "MAT-7C005", "物料名称": "高纯石墨坩埚", "规格": "Φ560mm", "单位": "个", "库存数量": "86", "存放库位": "B2-半成品库"}},
]


def main():
    t0 = time.time()
    report = []
    print("=" * 66)
    print("压测数据集生成器 · 开始（SEED=%d）" % SEED)
    print("=" * 66)

    targets = [
        ("docx", "01-设备采购档案.docx", build_devices_docx, 500 * 1024, 1900, "设备档案"),
        ("docx", "02-客户投诉处理档案.docx", build_complaints_docx, 500 * 1024, 1900, "客诉档案"),
        ("docx", "03-项目会议纪要汇编.docx", build_meetings_docx, 500 * 1024, 1300, "会议纪要"),
        ("docx", "04-员工管理与培训档案.docx", build_employees_docx, 500 * 1024, 1700, "员工档案"),
        ("docx", "05-年度经营分析报告.docx", build_report_docx, 500 * 1024, 1200, "经营报告"),
        ("xlsx", "01-设备采购明细总表.xlsx", build_purchase_xlsx, 500 * 1024, 5000, "设备采购表"),
        ("xlsx", "02-员工花名册全量.xlsx", build_roster_xlsx, 500 * 1024, 5000, "员工花名册"),
        ("xlsx", "03-项目任务分解与进度表.xlsx", build_tasks_xlsx, 500 * 1024, 5000, "任务表"),
        ("xlsx", "04-销售发货台账.xlsx", build_sales_xlsx, 500 * 1024, 5000, "销售台账"),
        ("xlsx", "05-原材料库存盘点明细.xlsx", build_inventory_xlsx, 500 * 1024, 5000, "库存盘点表"),
    ]
    for kind, name, fn, target, n0, label in targets:
        path = (DOCX_DIR if kind == "docx" else XLSX_DIR) / name
        n, size = fit(lambda k: fn(k), path, target, n0, label)
        report.append("%s/%s ｜ 规模 %d ｜ %.0f KB（目标 ≥%.0f KB）%s" % (
            kind, name, n, size / 1024, target / 1024, "✓" if size >= target else "✗ 未达标"))

    # md 三份
    for name, title, sub, n0 in [
        ("01-企业数字化转型白皮书.md", "企业数字化转型白皮书", "——从流程在线到数据驱动的实践路径（内部版）", 90),
        ("02-项目管理体系建设方案.md", "项目管理体系建设方案", "——面向装备制造企业的落地指南（评审稿）", 90),
        ("03-技术架构设计说明书.md", "技术架构设计说明书", "——多源数据融合平台总体设计（V2.1）", 90),
    ]:
        p = MD_DIR / name
        sz = build_md(p, title, sub, n0)
        while sz < 15 * 1024:
            n0 += 20
            sz = build_md(p, title, sub, n0)
        report.append("md/%s ｜ %.0f KB（目标 ≥15 KB）%s" % (name, sz / 1024, "✓" if sz >= 15 * 1024 else "✗"))

    # txt 三份
    for name, fn, n0 in [("01-设备验收记录汇总.txt", build_acceptance_txt, 240),
                         ("02-供应商资质与联系方式目录.txt", build_suppliers_txt, 210),
                         ("03-售后服务工单记录.txt", build_service_txt, 300)]:
        n = n0
        sz = fn(n)
        while sz < 15 * 1024:
            n += 60
            sz = fn(n)
        report.append("txt/%s ｜ 规模 %d ｜ %.0f KB（目标 ≥15 KB）%s" % (name, n, sz / 1024, "✓" if sz >= 15 * 1024 else "✗"))

    # 压测答案.json
    answer = {
        "generated_at": time.strftime("%Y-%m-%d %H:%M:%S"),
        "seed": SEED,
        "spec": "5×≥500KB docx + 5×≥500KB xlsx + 3×≥15KB md + 3×≥15KB txt（对齐赛题评价方法）",
        "canaries": CANARIES,
    }
    (HERE / "压测答案.json").write_text(json.dumps(answer, ensure_ascii=False, indent=2), encoding="utf-8")

    # 生成报告
    total = time.time() - t0
    header = ["压测数据集生成报告", "生成时间：%s ｜ 随机种子：%d" % (time.strftime("%Y-%m-%d %H:%M:%S"), SEED),
              "依据：赛题原文-A23.md【评价方法】文档样本集规格", "=" * 66, ""]
    body = report + ["", "金丝雀：共 %d 组（写入 压测答案.json）" % len(CANARIES), "总用时：%.1f 秒" % total]
    (HERE / "生成报告.txt").write_text("\n".join(header + body), encoding="utf-8")
    print("\n".join(header + body))
    print("\nDONE in %.1fs" % total)


if __name__ == "__main__":
    main()
