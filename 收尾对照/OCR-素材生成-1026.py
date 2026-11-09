# -*- coding: utf-8 -*-
"""
OCR-素材生成-1026.py —— 造一套「扫描件 / 电子版 / 图片」测试素材

为什么要自己造：
  项目里原来没有扫描版 PDF，也没法拿真实合同（有隐私）。自己造的好处是
  「标准答案就是我写进去的原文」——识别对不对，可以逐字比对，不靠感觉。

生成 5 份，覆盖真实场景里的四种情况：
  1. 电子版-采购合同.pdf     3 页，自带文字层（应该【不】走 OCR，0 误差）
  2. 扫描版-采购合同.pdf     3 页，整页是图（应该全部走 OCR）
  3. 混合版-统计公报.pdf     第 1 页文字层 + 第 2 页扫描图（最容易出错的情况）
  4. 扫描件-设备验收单.png   一张表格式截图（图片 OCR）
  5. 拍照件-公告（歪+噪点）.png  轻微倾斜 + 噪点，模拟手机拍纸（考验鲁棒性）

跑法（项目根目录）：
  python 收尾对照/OCR-素材生成-1026.py
"""

import io
import os
import random
import sys
from pathlib import Path

BASE = Path(__file__).resolve().parents[1]
OUT = BASE / "测试数据" / "OCR样本"

# A4 @ 200dpi ≈ 1654 × 2339；用这个尺寸渲染，跟 pdf_reader 的默认 dpi 对齐
W, H = 1654, 2339
DPI = 200

FONT_CANDIDATES = [
    r"C:\Windows\Fonts\msyh.ttc",      # 微软雅黑
    r"C:\Windows\Fonts\simsun.ttc",    # 宋体
    r"C:\Windows\Fonts\simhei.ttf",    # 黑体
]

# ==================== 测试原文（标准答案） ====================

CONTRACT_P1 = [
    ("title", "设备采购合同"),
    ("small", "合同编号：HD-2026-0317"),
    ("h2", "甲方（采购方）：华东理工大学"),
    ("h2", "乙方（供货方）：苏州精工机械有限公司"),
    ("h2", "签订地点：上海市奉贤区"),
    ("body", "第一条　标的物"),
    ("body", "乙方向甲方供应数控车床 3 台，型号 CK6140，单价 128000 元，"),
    ("body", "合计 384000 元（大写：叁拾捌万肆仟元整）。"),
    ("body", "第二条　交付"),
    ("body", "乙方应于 2026 年 11 月 30 日前将全部设备交付至甲方奉贤校区"),
    ("body", "实验楼 B 座 102 室，运输费用由乙方承担。"),
    ("body", "第三条　验收"),
    ("body", "甲方应在收货后 7 个工作日内完成验收。验收标准按 GB/T 9061-2006"),
    ("body", "执行，不合格的设备乙方应在 15 日内无偿更换。"),
]

CONTRACT_P2 = [
    ("h2", "第四条　付款方式"),
    ("body", "合同签订后 5 个工作日内，甲方支付合同总额的 30% 作为预付款；"),
    ("body", "设备验收合格后 10 个工作日内付清余款 70%。"),
    ("body", "乙方开户行：中国工商银行苏州工业园区支行"),
    ("body", "账号：1102 0231 0900 5566 778"),
    ("body", "第五条　违约责任"),
    ("body", "乙方逾期交货的，每逾期一日按合同总额的 0.5% 向甲方支付违约金，"),
    ("body", "累计不超过合同总额的 10%。"),
]

CONTRACT_P3_TABLE = [
    ["序号", "设备名称", "型号", "数量", "单价（元）", "金额（元）"],
    ["1", "数控车床", "CK6140", "3", "128000", "384000"],
    ["2", "立式铣床", "X5032", "2", "86000", "172000"],
    ["3", "万能磨床", "M1432", "1", "96000", "96000"],
    ["合计", "—", "—", "6", "—", "652000"],
]

CONTRACT_P3_TAIL = [
    ("body", "第六条　其他约定"),
    ("body", "本合同一式肆份，甲乙双方各执贰份，自双方签字盖章之日起生效。"),
    ("small", "甲方代表：____________　　乙方代表：____________"),
    ("small", "签订日期：2026 年 10 月 20 日"),
]

BULLETIN_P1 = [
    ("title", "2021 年民政事业发展统计公报（节选）"),
    ("h2", "一、服务机构与床位"),
    ("body", "截至 2021 年底，全国共有各类民政服务机构 4.2 万个，"),
    ("body", "提供床位 856.3 万张，比上年增长 3.7%。"),
    ("h2", "二、婚姻登记"),
    ("body", "全年共依法办理结婚登记 764.3 万对，离婚登记 213.9 万对。"),
    ("body", "其中涉外及涉港澳台婚姻登记 4.1 万对。"),
]

BULLETIN_P2 = [
    ("h2", "三、社会救助"),
    ("body", "全国城市低保对象 737.8 万人，农村低保对象 3474.4 万人。"),
    ("body", "全年累计支出低保资金 1651.6 亿元。"),
    ("h2", "四、养老服务"),
    ("body", "全国各类养老服务机构和设施 35.8 万个，床位 815.5 万张。"),
    ("body", "其中社区养老服务机构和设施 31.5 万个。"),
]

# 验收要逐条核对的"关键串"（去空白后比对）
EXPECT = {
    "电子版-采购合同.pdf": [
        "设备采购合同", "华东理工大学", "苏州精工机械有限公司", "HD-2026-0317",
        "CK6140", "128000", "384000", "叁拾捌万肆仟元整", "2026年11月30日",
        "GB/T9061-2006", "1102023109005566778", "0.5%",
        "数控车床", "立式铣床", "万能磨床", "M1432", "652000",
    ],
    "扫描版-采购合同.pdf": [
        "设备采购合同", "华东理工大学", "苏州精工机械有限公司", "HD-2026-0317",
        "CK6140", "128000", "384000", "叁拾捌万肆仟元整",
        "数控车床", "立式铣床", "M1432", "652000", "一式肆份",
    ],
    "混合版-统计公报.pdf": [
        "2021年民政事业发展统计公报", "4.2万个", "856.3万张", "3.7%",
        "764.3万对", "213.9万对", "737.8万人", "3474.4万人", "1651.6亿元",
        "35.8万个", "815.5万张",
    ],
    "扫描件-设备验收单.png": [
        "设备验收单", "验收日期", "数控车床", "CK6140", "128000", "384000",
        "立式铣床", "X5032", "86000", "172000", "556000",
    ],
    "拍照件-公告（歪+噪点）.png": [
        "关于开展2026年度设备安全专项检查的通知", "2026年11月1日",
        "各部门", "自查自纠", "11月20日",
    ],
}

NOTICE_LINES = [
    ("title", "关于开展 2026 年度设备安全专项检查的通知"),
    ("small", "校设〔2026〕27 号"),
    ("body", "各部门、各学院："),
    ("body", "为切实加强实验室设备安全管理，消除安全隐患，经研究决定，"),
    ("body", "于 2026 年 11 月 1 日至 11 月 20 日开展设备安全专项检查。"),
    ("body", "一、检查范围"),
    ("body", "全校单价 10 万元以上的教学科研仪器设备，共 1268 台。"),
    ("body", "二、工作安排"),
    ("body", "1. 自查自纠阶段（11 月 1 日至 11 月 10 日）：各部门自行排查。"),
    ("body", "2. 集中检查阶段（11 月 11 日至 11 月 20 日）：资产处组织抽查。"),
    ("body", "三、有关要求"),
    ("body", "请各部门于 11 月 20 日前将自查报告报送资产处设备管理科。"),
    ("small", "资产与实验室管理处"),
    ("small", "2026 年 10 月 24 日"),
]


# ==================== 渲染 ====================

def _font(size):
    from PIL import ImageFont
    for p in FONT_CANDIDATES:
        if os.path.exists(p):
            try:
                return ImageFont.truetype(p, size)
            except Exception:
                continue
    return ImageFont.load_default()


def _draw_lines(draw, lines, x, y, max_w):
    """按 (kind, text) 画一段内容，返回新的 y。"""
    for kind, text in lines:
        if kind == "title":
            f, gap, indent = _font(46), 78, 0
            bb = draw.textbbox((0, 0), text, font=f)
            x0 = (max_w - (bb[2] - bb[0])) / 2 + x
            draw.text((x0, y), text, font=f, fill=(15, 15, 15))
            y += gap
        elif kind == "h2":
            f, gap = _font(31), 56
            draw.text((x + 4, y), text, font=f, fill=(15, 15, 15))
            y += gap
        elif kind == "small":
            f, gap = _font(24), 44
            draw.text((x + 4, y), text, font=f, fill=(60, 60, 60))
            y += gap
        else:
            f, gap = _font(29), 50
            draw.text((x + 4, y), text, font=f, fill=(20, 20, 20))
            y += gap
    return y


def _draw_table(draw, rows, x, y, col_w, row_h=58):
    """画一个带线的表格。"""
    f = _font(25)
    n_col = len(col_w)
    total_w = sum(col_w)
    # 横线
    for i in range(len(rows) + 1):
        yy = y + i * row_h
        draw.line([(x, yy), (x + total_w, yy)], fill=(90, 90, 90), width=2)
    # 竖线
    cx = x
    for c in range(n_col + 1):
        draw.line([(cx, y), (cx, y + len(rows) * row_h)], fill=(90, 90, 90), width=2)
        if c < n_col:
            cx += col_w[c]
    # 文字（居中）
    for r, row in enumerate(rows):
        cx = x
        for c, cell in enumerate(row):
            bb = draw.textbbox((0, 0), cell, font=f)
            tw = bb[2] - bb[0]
            th = bb[3] - bb[1]
            draw.text((cx + (col_w[c] - tw) / 2, y + r * row_h + (row_h - th) / 2 - 4),
                      cell, font=f, fill=(20, 20, 20))
            cx += col_w[c]
    return y + len(rows) * row_h


def render_page(lines, table=None, tail=None, scanned=False, tilt=0.0, noise=0):
    """渲染一页 A4 图，返回 PNG bytes。"""
    from PIL import Image, ImageDraw

    bg = (250, 250, 247) if scanned else (255, 255, 255)
    img = Image.new("RGB", (W, H), bg)
    d = ImageDraw.Draw(img)

    margin_x = 150
    y = 170
    y = _draw_lines(d, lines, margin_x, y, W - margin_x * 2)
    if table:
        y += 26
        y = _draw_table(d, table, margin_x, y, [120, 300, 240, 120, 260, 280])
    if tail:
        y += 40
        y = _draw_lines(d, tail, margin_x, y, W - margin_x * 2)

    if scanned:
        # 模拟纸张纹理：加一点极浅的横纹
        d2 = ImageDraw.Draw(img)
        for yy in range(0, H, 7):
            d2.line([(0, yy), (W, yy)], fill=(248, 248, 244), width=1)

    if noise:
        rnd = random.Random(20261026)
        px = img.load()
        for _ in range(noise):
            xx, yy = rnd.randrange(W), rnd.randrange(H)
            v = rnd.randint(120, 200)
            px[xx, yy] = (v, v, v)

    if tilt:
        img = img.rotate(tilt, resample=Image.BICUBIC, fillcolor=(244, 244, 240), expand=False)

    buf = io.BytesIO()
    img.save(buf, format="PNG", optimize=True)
    return buf.getvalue()


# ==================== 造 PDF ====================

def build_pdf(path, pages):
    """pages = [(png_bytes, use_text_layer_lines_or_None), ...]

    给了 lines → 这一页写成真文字（有文字层）；
    只给图   → 这一页只插图片（扫描页）。
    """
    import fitz

    doc = fitz.open()
    for png, lines, table, tail in pages:
        page = doc.new_page(width=595, height=842)      # A4 点

        if lines is None:
            # 扫描页：整页就是一张图
            page.insert_image(fitz.Rect(0, 0, 595, 842), stream=png)
        else:
            # 电子版：真文字 + 真表格线
            y = 70
            for kind, text in lines:
                size = {"title": 18, "h2": 12, "small": 9.5}.get(kind, 11)
                if kind == "title":
                    tw = fitz.get_text_length(text, fontname="china-s", fontsize=size)
                    page.insert_text((clamp((595 - tw) / 2), y), text,
                                     fontname="china-s", fontsize=size)
                    y += 30
                else:
                    page.insert_text((62, y), text, fontname="china-s", fontsize=size)
                    y += size + 8
            if table:
                y += 8
                cw = [30, 92, 72, 34, 78, 84]
                rh = 24
                x0 = 62
                for i in range(len(table) + 1):
                    yy = y + i * rh
                    page.draw_line(fitz.Point(x0, yy), fitz.Point(x0 + sum(cw), yy),
                                   color=(0.35, 0.35, 0.35), width=0.7)
                cx = x0
                for c in range(len(cw) + 1):
                    page.draw_line(fitz.Point(cx, y), fitz.Point(cx, y + len(table) * rh),
                                   color=(0.35, 0.35, 0.35), width=0.7)
                    if c < len(cw):
                        cx += cw[c]
                for r, row in enumerate(table):
                    cx = x0
                    for c, cell in enumerate(row):
                        page.insert_text((cx + 4, y + r * rh + 16), cell,
                                         fontname="china-s", fontsize=9)
                        cx += cw[c]
                y += len(table) * rh
            if tail:
                y += 12
                for kind, text in tail:
                    size = {"title": 18, "h2": 12, "small": 9.5}.get(kind, 11)
                    page.insert_text((62, y), text, fontname="china-s", fontsize=size)
                    y += size + 8
    doc.save(str(path), garbage=3, deflate=True)
    doc.close()


def clamp(v):
    return max(40, float(v))


# ==================== 主流程 ====================

def main():
    try:
        from PIL import Image  # noqa: F401
    except Exception:
        print("缺 Pillow：pip install pillow")
        sys.exit(1)
    try:
        import fitz  # noqa: F401
    except Exception:
        print("缺 PyMuPDF：pip install pymupdf")
        sys.exit(1)

    OUT.mkdir(parents=True, exist_ok=True)
    made = []

    # 1) 电子版 PDF：三页都是真文字
    p1 = render_page(CONTRACT_P1, scanned=False)
    p2 = render_page(CONTRACT_P2, scanned=False)
    p3 = render_page(CONTRACT_P3_TAIL[:1], table=CONTRACT_P3_TABLE,
                     tail=CONTRACT_P3_TAIL[1:], scanned=False)
    build_pdf(OUT / "电子版-采购合同.pdf", [
        (p1, CONTRACT_P1, None, None),
        (p2, CONTRACT_P2, None, None),
        (p3, CONTRACT_P3_TAIL[:1], CONTRACT_P3_TABLE, CONTRACT_P3_TAIL[1:]),
    ])
    made.append("电子版-采购合同.pdf")

    # 2) 扫描版 PDF：三页全是图（无文字层）
    s1 = render_page(CONTRACT_P1, scanned=True)
    s2 = render_page(CONTRACT_P2, scanned=True)
    s3 = render_page(CONTRACT_P3_TAIL[:1], table=CONTRACT_P3_TABLE,
                     tail=CONTRACT_P3_TAIL[1:], scanned=True)
    build_pdf(OUT / "扫描版-采购合同.pdf", [
        (s1, None, None, None), (s2, None, None, None), (s3, None, None, None),
    ])
    made.append("扫描版-采购合同.pdf")

    # 3) 混合版：第 1 页真文字 + 第 2 页扫描图
    b1 = render_page(BULLETIN_P1, scanned=False)
    b2 = render_page(BULLETIN_P2, scanned=True)
    build_pdf(OUT / "混合版-统计公报.pdf", [
        (b1, BULLETIN_P1, None, None),
        (b2, None, None, None),
    ])
    made.append("混合版-统计公报.pdf")

    # 4) 图片：表格式验收单
    from PIL import Image, ImageDraw
    img = Image.new("RGB", (1400, 1100), (252, 252, 250))
    d = ImageDraw.Draw(img)
    f_t = _font(40)
    tw = d.textbbox((0, 0), "设备验收单", font=f_t)[2]
    d.text(((1400 - tw) / 2, 60), "设备验收单", font=f_t, fill=(10, 10, 10))
    f_s = _font(24)
    d.text((90, 150), "项目名称：机械加工实训中心设备采购", font=f_s, fill=(40, 40, 40))
    d.text((90, 195), "验收日期：2026 年 10 月 18 日", font=f_s, fill=(40, 40, 40))
    d.text((820, 150), "合同编号：HD-2026-0317", font=f_s, fill=(40, 40, 40))
    _draw_table(d, [
        ["序号", "设备名称", "型号", "数量", "单价（元）", "金额（元）"],
        ["1", "数控车床", "CK6140", "3", "128000", "384000"],
        ["2", "立式铣床", "X5032", "2", "86000", "172000"],
        ["合计", "—", "—", "5", "—", "556000"],
    ], 90, 270, [100, 250, 200, 110, 230, 240], row_h=64)
    d.text((90, 700), "验收结论：设备运行正常，符合合同约定，同意验收。", font=f_s, fill=(30, 30, 30))
    d.text((90, 750), "验收人（签字）：张建国　　　　使用部门（盖章）：机械工程学院", font=f_s, fill=(30, 30, 30))
    img.save(OUT / "扫描件-设备验收单.png")
    made.append("扫描件-设备验收单.png")

    # 5) 拍照件：轻微倾斜 + 噪点
    n = render_page(NOTICE_LINES, scanned=True, tilt=-0.6, noise=9000)
    (OUT / "拍照件-公告（歪+噪点）.png").write_bytes(n)
    made.append("拍照件-公告（歪+噪点）.png")

    print("素材已生成到：%s" % OUT)
    for m in made:
        p = OUT / m
        print("  - %-30s %7.1f KB" % (m, p.stat().st_size / 1024))
    print("\n标准答案（验收时逐条比对）：")
    for k, v in EXPECT.items():
        print("  %s → %d 条关键串" % (k, len(v)))


if __name__ == "__main__":
    main()
