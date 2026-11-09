# -*- coding: utf-8 -*-
"""
OCR-手写-素材生成-1103.py —— 造「手写」测试素材（补 1026 那批只有印刷体的空缺）

为什么要造、以及造得有多真：
  真实手写样本拿不到（涉及隐私），但手写 OCR 的难点是**确定的**：
    · 笔画粗细不均、连笔、字形不规范（不像印刷体那样每个字一模一样）
    · 逐字位置会飘（基线不齐、字距不均、轻微旋转）
    · 常常是「印刷表格 + 手写填空」混在一起 —— 这才是企业里最常见的样子
  所以素材用「楷体 / 华文行楷」渲染，并给每个字加随机旋转、缩放、位移、深浅，
  尽量贴近"人一笔一笔写"的不规则性。这不能替代真实手写，但能稳定比较
  不同模型在同一批字上的差距 —— 而比较模型正是这次要解决的问题。

4 张，覆盖 4 种真实形态：
  1. 手写-验收单（印刷表格+手写填值）.png   ← 最常见的形态，手写只占少数格
  2. 手写-会议记录（整篇手写）.png          ← 最难的形态
  3. 手写-报销单（手写数字+涂改）.png       ← 数字最要命，还带划改
  4. 手写-混合（印刷正文+手写批注）.png     ← 批注挤在页边

标准答案写在 GROUND_TRUTH 里，验收脚本逐字比对，不靠肉眼。

跑法（项目根目录）：
  python 收尾对照/OCR-手写-素材生成-1103.py
"""

import io
import random
from pathlib import Path

BASE = Path(__file__).resolve().parents[1]
OUT = BASE / "测试数据" / "OCR样本" / "手写"

# 印刷体 / 手写体字体（按可用性挑，缺了自动退）
PRINT_FONTS = [
    r"C:\Windows\Fonts\msyh.ttc",      # 微软雅黑
    r"C:\Windows\Fonts\simsun.ttc",    # 宋体
]
HAND_FONTS = [
    r"C:\Windows\Fonts\STXINGKA.TTF",  # 华文行楷（最像手写）
    r"C:\Windows\Fonts\STKAITI.TTF",   # 华文楷体
    r"C:\Windows\Fonts\simkai.ttf",    # 楷体
]

# ==================== 标准答案 ====================

# 1) 印刷表格 + 手写填值
FORM_PRINT = [
    ("title", "设备验收单"),
    ("label", "项目名称：机械加工实训中心设备采购"),
    ("label", "合同编号：HD-2026-0317"),
]
FORM_FIELDS = [
    # (标签, 手写值, 标签所在 x, 值所在 x, y)
    ("验收日期：", "2026年10月18日", 110, 330),
    ("验收人：", "张建国", 110, 330),
    ("使用部门：", "机械工程学院", 110, 330),
    ("设备状态：", "运行正常", 800, 1010),
    ("验收结论：", "符合合同约定，同意验收", 800, 1010),
]
FORM_TABLE = [
    ["序号", "设备名称", "型号", "数量", "单价（元）", "金额（元）"],
    ["1", "数控车床", "CK6140", "3", "128000", "384000"],
    ["2", "立式铣床", "X5032", "2", "86000", "172000"],
    ["合计", "—", "—", "5", "—", "556000"],
]
# 表格里手写的是「数量」列（第 3 列，0-based）
FORM_HAND_COL = 3

# 2) 整篇手写会议记录
MEETING_LINES = [
    "设备安全专项检查会议记录",
    "时间：2026年10月24日 下午三点",
    "地点：实验楼B座102室",
    "参会人：张建国、李红梅、王海涛",
    "一、本月完成设备巡检126台，发现隐患3处",
    "二、立式铣床X5032主轴异响，需更换轴承",
    "三、下周组织安全培训，全员参加",
    "记录人：李红梅",
]

# 3) 手写报销单（数字 + 一处涂改）
CLAIM_ROWS = [
    ["日期", "事由", "金额（元）"],
    ["2026-10-08", "差旅费（苏州—上海）", "386.50"],
    ["2026-10-12", "设备配件采购", "1280.00"],
    ["2026-10-18", "快递费", "45.00"],
    ["合计", "", "1711.50"],
]
# 涂改：第 2 行金额写错划掉，改成 1280.00 —— 标准答案按改后的算
CLAIM_CORRECT = (2, 2, "1280.00")

# 4) 印刷正文 + 手写批注
MIXED_BODY = [
    ("title", "关于开展设备安全专项检查的通知"),
    ("body", "各部门、各学院："),
    ("body", "为切实加强实验室设备安全管理，消除安全隐患，经研究决定，"),
    ("body", "于2026年11月1日至11月20日开展设备安全专项检查。"),
    ("body", "一、检查范围：全校单价10万元以上的教学科研仪器设备。"),
    ("body", "二、工作安排：分自查自纠和集中检查两个阶段。"),
    ("body", "三、有关要求：请各部门按时报送自查报告。"),
]
MIXED_NOTES = [
    # (别名, 手写内容, x, y)
    ("页边批注1", "已阅 同意 张建国 10.25", 1230, 260),
    ("页边批注2", "11月20日前必须完成", 1230, 700),
]

GROUND_TRUTH = {
    "手写-验收单（印刷表格+手写填值）.png": [
        "设备验收单", "机械加工实训中心设备采购", "HD-2026-0317",
        "2026年10月18日", "张建国", "机械工程学院", "运行正常",
        "符合合同约定，同意验收",
        "数控车床", "CK6140", "384000", "立式铣床", "X5032", "172000", "556000",
    ],
    "手写-会议记录（整篇手写）.png": [
        "设备安全专项检查会议记录", "2026年10月24日", "实验楼B座102室",
        "张建国", "李红梅", "王海涛",
        "本月完成设备巡检126台", "发现隐患3处",
        "立式铣床X5032主轴异响", "需更换轴承",
        "下周组织安全培训",
    ],
    "手写-报销单（数字+涂改）.png": [
        "差旅费", "386.50", "设备配件采购", "1280.00", "快递费", "45.00", "1711.50",
    ],
    "手写-混合（印刷正文+手写批注）.png": [
        "关于开展设备安全专项检查的通知", "实验室设备安全管理",
        "2026年11月1日至11月20日", "自查自纠", "集中检查",
        "已阅", "同意", "11月20日前必须完成",
    ],
}


# ==================== 基础绘制 ====================

def _font(paths, size):
    from PIL import ImageFont
    for p in paths:
        try:
            return ImageFont.truetype(p, size)
        except Exception:
            continue
    return ImageFont.load_default()


def hand_write(draw, xy, text, size, seed, color=(18, 26, 58), jitter=1.0):
    """把 text 一个字一个字画出来，每个字随机旋转/缩放/偏移 —— 模拟手写。

    为什么逐字画：整行一次 draw.text 出来的是完美排版，OCR 一眼就认出，
    测不出手写模型和印刷模型的区别。逐字加扰动才是有效对比。
    jitter=0 时退化成正常渲染（用来做"印刷体对照"）。
    """
    from PIL import Image

    rnd = random.Random(seed)
    x, y = xy
    for ch in text:
        if jitter <= 0:
            f = _font(HAND_FONTS, size)
            draw.text((x, y), ch, font=f, fill=color)
            x += draw.textbbox((0, 0), ch, font=f)[2]
            continue

        s = int(size * rnd.uniform(0.92, 1.06))
        f = _font(HAND_FONTS, s)
        # 字宽按真实字形量（不能用字号当步进：数字的实际宽度只有字号的一半，
        # 拿字号当步进会把 "2026" 摆成 "2 0 2 6"，那是排版错误，不是手写难）
        bb = f.getbbox(ch)
        gw = max(1, bb[2] - bb[0])

        # 在透明小图上写这个字，再旋转贴回去（PIL 不支持直接旋转文字）
        pad = int(s * 0.6)
        tile = Image.new("RGBA", (gw + pad * 2, s + pad * 2), (0, 0, 0, 0))
        td = __import__("PIL.ImageDraw", fromlist=["ImageDraw"]).Draw(tile)
        td.text((pad - bb[0], pad), ch, font=f, fill=color + (255,))
        tile = tile.rotate(rnd.uniform(-3.5, 3.5), resample=Image.BICUBIC,
                           expand=False, fillcolor=(0, 0, 0, 0))
        dy = int(rnd.uniform(-2.5, 2.5) * size / 30.0)
        target = draw._image
        target.paste(tile, (int(x) - pad, int(y) + dy - pad // 2), tile)
        # 字距：字宽 + 一点点随机间隙（真实手写会在 0~1.5 像素/字间浮动）
        x += gw * rnd.uniform(0.99, 1.10) + rnd.uniform(0.5, 2.2)


def _measure(draw, text, paths, size):
    f = _font(paths, size)
    bb = draw.textbbox((0, 0), text, font=f)
    return bb[2] - bb[0], bb[3] - bb[1]


def draw_table(draw, rows, x, y, col_w, row_h, hand_cols=(), seed0=1, size=26):
    """带竖线的表格；hand_cols 里的列用手写字。"""
    f = _font(PRINT_FONTS, size)
    total = sum(col_w)
    for i in range(len(rows) + 1):
        yy = y + i * row_h
        draw.line([(x, yy), (x + total, yy)], fill=(80, 80, 80), width=2)
    cx = x
    for c in range(len(col_w) + 1):
        draw.line([(cx, y), (cx, y + len(rows) * row_h)], fill=(80, 80, 80), width=2)
        if c < len(col_w):
            cx += col_w[c]
    for r, row in enumerate(rows):
        cx = x
        for c, cell in enumerate(row):
            if not cell:
                cx += col_w[c]
                continue
            if c in hand_cols and r > 0:
                # 手写：左对齐、略高一点，逐字抖动
                hand_write(draw, (cx + 10, y + r * row_h + 12), cell, size,
                           seed0 + r * 17 + c, jitter=1.0)
            else:
                bb = draw.textbbox((0, 0), cell, font=f)
                tw, th = bb[2] - bb[0], bb[3] - bb[1]
                draw.text((cx + (col_w[c] - tw) / 2, y + r * row_h + (row_h - th) / 2 - 4),
                          cell, font=f, fill=(20, 20, 20))
            cx += col_w[c]
    return y + len(rows) * row_h


def paper(w, h, tint=(252, 252, 249), shade=True, seed=7):
    """造一张"纸"：底色 + 极浅的横向纹理 + 四个角压暗。"""
    from PIL import Image, ImageDraw
    img = Image.new("RGB", (w, h), tint)
    d = ImageDraw.Draw(img)
    if shade:
        rnd = random.Random(seed)
        for yy in range(0, h, 9):
            v = rnd.randint(0, 4)
            d.line([(0, yy), (w, yy)], fill=(tint[0] - v, tint[1] - v, tint[2] - v), width=1)
    return img


# ==================== 四张素材 ====================

def make_form():
    """印刷表格 + 手写填值（企业里最常见的形态）。"""
    from PIL import ImageDraw

    W, H = 1654, 1300
    img = paper(W, H)
    d = ImageDraw.Draw(img)

    f_t = _font(PRINT_FONTS, 48)
    bb = d.textbbox((0, 0), "设备验收单", font=f_t)
    d.text(((W - (bb[2] - bb[0])) / 2, 70), "设备验收单", font=f_t, fill=(10, 10, 10))

    f_l = _font(PRINT_FONTS, 27)
    d.text((110, 165), FORM_PRINT[1][1], font=f_l, fill=(35, 35, 35))
    d.text((880, 165), FORM_PRINT[2][1], font=f_l, fill=(35, 35, 35))

    # 左列 3 个字段 + 右列 2 个字段，值都是手写
    y0 = 230
    for i, (label, val, lx, vx) in enumerate(FORM_FIELDS[:3]):
        yy = y0 + i * 62
        d.text((lx, yy), label, font=f_l, fill=(35, 35, 35))
        d.line([(lx, yy + 40), (vx + 420, yy + 40)], fill=(150, 150, 150), width=1)
        hand_write(d, (vx, yy + 4), val, 30, 300 + i)
    for i, (label, val, lx, vx) in enumerate(FORM_FIELDS[3:]):
        yy = y0 + i * 62
        d.text((lx, yy), label, font=f_l, fill=(35, 35, 35))
        hand_write(d, (vx, yy + 4), val, 30, 400 + i)

    draw_table(d, FORM_TABLE, 110, 440, [110, 280, 220, 140, 260, 300],
               row_h=70, hand_cols=(FORM_HAND_COL,), seed0=11, size=28)

    f_s = _font(PRINT_FONTS, 24)
    d.text((110, 880), "验收人（签字）：", font=f_s, fill=(35, 35, 35))
    hand_write(d, (330, 878), "张建国", 34, 900)
    d.text((700, 880), "使用部门（盖章）：", font=f_s, fill=(35, 35, 35))
    hand_write(d, (930, 878), "机械工程学院", 34, 901)

    # 手写签名（连笔，认不出是正常的，标准答案里不要求）
    hand_write(d, (1180, 980), "张建国", 52, 902, jitter=1.4)

    img.save(OUT / "手写-验收单（印刷表格+手写填值）.png")


def make_meeting():
    """整篇手写 —— 最难的一种。"""
    from PIL import ImageDraw

    W, H = 1654, 1300
    img = paper(W, H, tint=(250, 250, 245), seed=13)
    d = ImageDraw.Draw(img)

    y = 110
    for i, ln in enumerate(MEETING_LINES):
        if i == 0:
            w, _ = _measure(d, ln, HAND_FONTS, 46)
            hand_write(d, ((W - w) / 2, y), ln, 46, 500 + i)
            y += 100
        else:
            d.line([(90, y + 62), (W - 90, y + 62)], fill=(228, 226, 218), width=1)
            hand_write(d, (110, y), ln, 36, 500 + i)
            y += 78

    img.save(OUT / "手写-会议记录（整篇手写）.png")


def make_claim():
    """手写报销单：数字 + 一处涂改（划掉重写）。"""
    from PIL import ImageDraw

    W, H = 1654, 900
    img = paper(W, H)
    d = ImageDraw.Draw(img)

    f_t = _font(PRINT_FONTS, 42)
    bb = d.textbbox((0, 0), "费用报销单", font=f_t)
    d.text(((W - (bb[2] - bb[0])) / 2, 70), "费用报销单", font=f_t, fill=(10, 10, 10))
    f_l = _font(PRINT_FONTS, 26)
    d.text((110, 160), "报销人：", font=f_l, fill=(35, 35, 35))
    hand_write(d, (250, 158), "王海涛", 30, 600)
    d.text((880, 160), "部门：", font=f_l, fill=(35, 35, 35))
    hand_write(d, (960, 158), "机械工程学院", 30, 601)

    end_y = draw_table(d, CLAIM_ROWS, 110, 240, [300, 560, 300],
                       row_h=80, hand_cols=(0, 1, 2), seed0=700, size=28)

    # 涂改：把第 2 行金额划掉，旁边重写
    r, c, val = CLAIM_CORRECT
    row_h, y0 = 80, 240
    cx = 110 + 300 + 560
    yy = y0 + r * row_h
    # 先擦掉原来那个手写数字（贴一块纸色），再画横线，再重写
    d.rectangle([cx + 6, yy + 8, cx + 290, yy + row_h - 8], fill=(252, 252, 249))
    d.line([(cx + 20, yy + row_h // 2), (cx + 200, yy + row_h // 2)], fill=(40, 40, 40), width=3)
    hand_write(d, (cx + 20, yy + row_h // 2 - 34), "1180.00", 24, 777)
    hand_write(d, (cx + 190, yy + 12), val, 28, 778)

    f_s = _font(PRINT_FONTS, 24)
    d.text((110, end_y + 70), "审批意见：", font=f_s, fill=(35, 35, 35))
    hand_write(d, (270, end_y + 68), "同意报销", 34, 800)

    img.save(OUT / "手写-报销单（数字+涂改）.png")


def make_mixed():
    """印刷正文 + 页边手写批注。"""
    from PIL import ImageDraw

    W, H = 1654, 1200
    img = paper(W, H, tint=(255, 255, 253), shade=False)
    d = ImageDraw.Draw(img)

    y = 100
    for kind, text in MIXED_BODY:
        if kind == "title":
            f = _font(PRINT_FONTS, 42)
            bb = d.textbbox((0, 0), text, font=f)
            d.text(((W - (bb[2] - bb[0])) / 2, y), text, font=f, fill=(10, 10, 10))
            y += 90
        else:
            f = _font(PRINT_FONTS, 28)
            d.text((140, y), text, font=f, fill=(25, 25, 25))
            y += 62

    # 页边竖排批注（手写，逆时针竖着写）
    from PIL import Image
    for alias, note, nx, ny in MIXED_NOTES:
        tile = Image.new("RGBA", (700, 90), (0, 0, 0, 0))
        td = __import__("PIL.ImageDraw", fromlist=["ImageDraw"]).Draw(tile)
        hand_write(td, (4, 4), note, 30, hash(alias) % 999)
        tile = tile.rotate(90, expand=True)
        img.paste(tile, (nx, ny), tile)

    img.save(OUT / "手写-混合（印刷正文+手写批注）.png")


def main():
    OUT.mkdir(parents=True, exist_ok=True)
    make_form()
    make_meeting()
    make_claim()
    make_mixed()

    lines = []
    for name, keys in GROUND_TRUTH.items():
        p = OUT / name
        lines.append("  - %-40s %6.1f KB ｜ 标准答案 %d 条"
                     % (name, p.stat().st_size / 1024, len(keys)))
    print("素材已生成：%s" % OUT)
    print("\n".join(lines))
    (OUT / "标准答案.txt").write_text(
        "\n".join("%s\n    %s" % (k, " ｜ ".join(v)) for k, v in GROUND_TRUTH.items()),
        encoding="utf-8")


if __name__ == "__main__":
    main()
