# -*- coding: utf-8 -*-
"""
复杂表头-效果图-1026.py
=======================
把「已填写」的两份复杂模板渲染成图片，方便直接看效果。

做法：openpyxl 读值和合并单元格 → 生成一张 HTML 表格（保留合并、加侧栏标出
表头层和第 2 页起点）→ 用本机 Edge 无头模式截图。

输出：
    docs/screenshots/13-多层表头填写.png
    docs/screenshots/14-跨页表填写.png
"""

import os
import subprocess
import sys
from pathlib import Path

import openpyxl

HERE = Path(__file__).resolve().parent
BASE = HERE.parent
OUT_DIR = BASE / "writer" / "输出"
SHOTS = BASE / "docs" / "screenshots"
TMP = HERE / "_效果图"

EDGE = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"

CSS = """
<style>
  body { margin:0; padding:22px 26px; background:#f5f6f8;
         font-family:"Microsoft YaHei","Segoe UI",sans-serif; color:#1b1b1f; }
  h1 { font-size:17px; margin:0 0 4px; font-weight:600; letter-spacing:.2px; }
  .sub { font-size:12px; color:#6b6b76; margin-bottom:14px; }
  .sub b { color:#8A3B2A; font-weight:600; }
  table { border-collapse:collapse; font-size:12.5px; background:#fff;
          box-shadow:0 1px 3px rgba(0,0,0,.12); }
  td { border:1px solid #9aa0a8; padding:5px 9px; height:22px;
       vertical-align:middle; }
  td.hdr { background:#e8eaee; font-weight:600; text-align:center; }
  td.title { background:#d8dbe1; font-weight:700; font-size:14px; text-align:center; }
  td.meta { background:#f0f1f3; color:#4a4a52; }
  td.data { background:#fff; }
  td.cut { background:#fde4c8; }
  td.sep { background:#f3ece2; font-weight:600; text-align:center; color:#8A3B2A; }
  tr.divider td { border-top:2px solid #8A3B2A; }
  .legend { margin-top:12px; font-size:11.5px; color:#6b6b76; }
  .legend span { display:inline-block; margin-right:16px; }
  .sw { display:inline-block; width:10px; height:10px; border:1px solid #9aa0a8;
        vertical-align:-1px; margin-right:4px; }
</style>
"""


def render_html(xlsx_path, title, subtitle, header_rows=(), cut_from=None,
                sep_rows=(), note=""):
    wb = openpyxl.load_workbook(xlsx_path)
    ws = wb.worksheets[0]

    # 合并信息：{(row,col): (rowspan, colspan)} + 被覆盖的格子
    anchors, covered = {}, set()
    for rng in ws.merged_cells.ranges:
        rs = rng.max_row - rng.min_row + 1
        cs = rng.max_col - rng.min_col + 1
        anchors[(rng.min_row, rng.min_col)] = (rs, cs)
        for r in range(rng.min_row, rng.max_row + 1):
            for c in range(rng.min_col, rng.max_col + 1):
                if (r, c) != (rng.min_row, rng.min_col):
                    covered.add((r, c))

    rows_html = []
    # 尾部连续空行不渲染（openpyxl 保存后 max_row 常比实际内容多几行）
    last = ws.max_row
    while last > 1:
        if any((ws.cell(row=last, column=c).value not in (None, ""))
               for c in range(1, ws.max_column + 1)):
            break
        last -= 1

    for r in range(1, last + 1):
        cells = []
        for c in range(1, ws.max_column + 1):
            if (r, c) in covered:
                continue
            v = ws.cell(row=r, column=c).value
            text = "" if v is None else str(v)
            # 公式在 openpyxl 写出的文件里没有缓存值，直接显示会把 "=SUM(...)" 露出来
            if text.startswith("="):
                text = "（自动计算）"

            rs, cs = anchors.get((r, c), (1, 1))
            if r == 1:
                cls = "title"
            elif sep_rows and r in sep_rows:
                cls = "sep"
            elif r in header_rows:
                cls = "hdr"
            elif text.endswith("："):
                cls = "meta"
            elif any(k in text for k in ("合计", "续表", "验收结论", "说明")):
                cls = "meta"
            elif cut_from and r >= cut_from:
                cls = "cut"
            else:
                cls = "data"

            span = ""
            if rs > 1:
                span += ' rowspan="%d"' % rs
            if cs > 1:
                span += ' colspan="%d"' % cs
            cells.append('<td class="%s"%s>%s</td>' % (cls, span, text))
        rows_html.append("<tr>%s</tr>" % "".join(cells))

    total_rows = last
    wb.close()

    html = """<!DOCTYPE html><html><head><meta charset="utf-8">%s</head><body>
<h1>%s</h1><div class="sub">%s</div>
<table>%s</table>
<div class="legend">%s
  <span><i class="sw" style="background:#e8eaee"></i>表头</span>
  <span><i class="sw" style="background:#fde4c8"></i>第 2 页（续表之后）</span>
  <span><i class="sw" style="background:#f3ece2"></i>标题 / 续表 / 合计</span>
</div>
</body></html>""" % (CSS, title, subtitle, "".join(rows_html), note)

    TMP.mkdir(parents=True, exist_ok=True)
    h = TMP / (xlsx_path.stem + ".html")
    h.write_text(html, encoding="utf-8")
    return h, total_rows


def shot(html_path, png_path, height=760):
    png_path.parent.mkdir(parents=True, exist_ok=True)
    if png_path.exists():
        png_path.unlink()
    subprocess.run([EDGE, "--headless=new", "--disable-gpu", "--hide-scrollbars",
                    "--force-device-scale-factor=2",
                    "--screenshot=%s" % png_path,
                    "--window-size=1120,%d" % int(height),
                    html_path.as_uri()],
                   capture_output=True, timeout=120)
    return png_path.exists(), (png_path.stat().st_size if png_path.exists() else 0)


def main():
    jobs = [
        (OUT_DIR / "模板6-多层表头设备采购台账-已填写.xlsx",
         "多层表头 · 设备采购台账（两级表头，父列合并）",
         '表头占 <b>第 3~4 行</b>；「数量」「金额（元）」在两个父列下各有一份，'
         '分别落进 <b>上半年采购-数量</b> / <b>下半年采购-数量</b>，没有错列。',
         dict(header_rows=(1, 3, 4)),
         "13-多层表头填写.png"),
        (OUT_DIR / "模板8-跨页原材料入库台账-已填写.xlsx",
         "跨页表 · 原材料入库台账（第 2 页自动接着填）",
         '表头只有 <b>第 3 行</b>，但表体分两段：第 4~12 行 + 第 15~23 行。'
         '12 条记录 → <b>第 10~12 条落在第 2 段</b>（橙色区），序号连续到 12。',
         dict(header_rows=(1, 3, 14), cut_from=15, sep_rows=(13, 24)),
         "14-跨页表填写.png"),
    ]

    print("=" * 62)
    print("渲染复杂表头填写效果图")
    print("=" * 62)
    for xlsx, title, sub, kw, png_name in jobs:
        if not xlsx.exists():
            print("缺少文件：%s（先跑 复杂表头-演示-1026.py）" % xlsx.name)
            return 1
        html, n_rows = render_html(xlsx, title, sub, **kw)
        height = 250 + n_rows * 30          # 按行数算高度，别把合计行裁掉
        png = SHOTS / png_name
        made, size = shot(html, png, height)
        print("  %s %-40s %6.1f KB  (%d 行)" % (
            "✓" if made else "✗", png.name, size / 1024, n_rows))
        if not made:
            return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
