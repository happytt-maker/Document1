# -*- coding: utf-8 -*-
"""
第 5 步 · 写回验收（verify_fill.py）

对照检查"已填写"文件，回答三个问题：

  1. 值对不对   —— ① 待填区与抽取结果逐格核对
                    ② 待填区以外必须纹丝未动（表头、说明行、标题……全表对比）
  2. 格式保没保 —— 单元格样式（字体/边框/对齐/填充/数字格式）、合并单元格、
                    列宽、行高、数据验证、条件格式，逐项对比
  3. 能不能交卷 —— 汇总结论

用法（VSCode 打开 writer 文件夹后，直接 F5）：
    python verify_fill.py                     # 默认模板4
    python verify_fill.py 模板3-会议任务跟踪表.xlsx

输出：writer/输出/验收报告-<模板名>.txt（控制台同步打印一份）
"""
import json
import os
import sys
import time
from pathlib import Path

from openpyxl import load_workbook

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

HERE = Path(__file__).resolve().parent            # writer/
BASE = HERE.parent                                # 研一项目/
# 产物目录（2026-11-03）：默认仍是 template_reader/、matcher/、writer/输出/；
# 并发跑多个任务时由调用方用环境变量指到各自任务目录 —— 与 fill.py 同一约定。
# 不改的话，两个任务会读到对方的抽取结果，验收结论会串。
STRUCTURE_JSON = Path(os.environ.get("A23_STRUCTURE_JSON")
                      or (BASE / "template_reader" / "模板表结构.json"))
MATCHER_DIR = Path(os.environ.get("A23_MATCHER_DIR") or (BASE / "matcher"))
# 模板目录：默认 测试数据/模板表格；可用环境变量 A23_TEMPLATE_DIR 覆盖
# （泛化测试 / 比赛换数据集时用；老流程不设变量时不受影响）
TEMPLATE_DIR = Path(os.environ.get("A23_TEMPLATE_DIR",
                                   str(BASE / "测试数据" / "模板表格")))
OUT_DIR = Path(os.environ.get("A23_OUT_DIR") or (HERE / "输出"))

# 期望值必须用 fill.py 的真实槽位算法来算（按源行号对齐 + 文档级单值广播）。
# 2026-11-02 踩的坑：模板1「供应商」被广播到 4 行后，验收这边仍按"1 条对 1 槽"
# 算期望，于是报"第 6/7/8 行应该为空"，而实际填得完全正确（评分器 100%）。
# 验收脚本本身的期望算法过时，会把好结果报成问题 —— 比不报更糟。
sys.path.insert(0, str(HERE))
from fill import build_slot_map, fill_slots, writable_columns  # noqa: E402

DEFAULT_TEMPLATE = "模板4-客诉处理台账.xlsx"


def say(text=""):
    print(text)


def load_sheet_info(template_name):
    data = json.loads(STRUCTURE_JSON.read_text(encoding="utf-8"))
    for item in data:
        if item.get("file") == template_name:
            return item["sheets"][0]
    return None


def load_sheet_infos(template_name):
    """这个模板的全部块（docx 多张表 / xlsx 多 sheet）—— 验收要一块一块对。"""
    data = json.loads(STRUCTURE_JSON.read_text(encoding="utf-8"))
    for item in data:
        if item.get("file") == template_name:
            return list(item.get("sheets") or [])
    return []


def load_extract(col_title):
    f = MATCHER_DIR / ("抽取结果-%s.json" % col_title)
    if not f.exists():
        return None
    return json.loads(f.read_text(encoding="utf-8"))


def norm(v):
    """值的规范化：None 和空串都算「空」，其余转字符串。"""
    if v is None:
        return ""
    return str(v).strip()


def style_key(cell):
    """把一个格子的"视觉样式"取成一串可比对的标签。"""
    f = cell.font
    b = cell.border
    a = cell.alignment
    fl = cell.fill
    return (
        f.name, f.size, bool(f.bold), str(f.color),
        b.left.style, b.right.style, b.top.style, b.bottom.style,
        a.horizontal, a.vertical, bool(a.wrap_text),
        str(cell.number_format),
        str(fl.patternType), str(fl.fgColor),
    )


def width_map(ws):
    m = {}
    for k, d in ws.column_dimensions.items():
        m[k] = d.width
    return m


def height_map(ws):
    m = {}
    for k, d in ws.row_dimensions.items():
        m[k] = d.height
    return m


# ============================================================================
# Word（docx）模板的验收（2026-10-14 新增）
# ============================================================================
# 与 xlsx 版同样回答三个问题：
#   1. 值对不对    —— 待填区 vs 抽取结果逐格核对 + 待填区以外纹丝未动（文字层面）
#   2. 格式保没保  —— Word 版的检查项：表格行列数 / 表格样式名 /
#                     待填格段落对齐（写入前后） / 正文段落数
#   3. 能不能交卷

def verify_docx(template_name):
    """Word 版写回验收：值对不对 + 待填区外有没有动 + 段落格式保没保。

    **按块验收**（2026-11-07）：一份 docx 里可能有好几张结构相同的表
    （山东空气质量模板就是德州 / 潍坊 / 临沂 三张），每张表的期望值不同。
    旧版只认 sheets[0] 和 doc.tables[0]，把三张表的期望值按顺序串成一串，
    去和第 1 张表逐格比 —— 于是"文档填得全对"却报出 157 格不一致。
    验收脚本算错期望，比不验收更糟：它会让人去改对的东西。
    """
    from docx import Document

    infos = load_sheet_infos(template_name)
    if not infos:
        say("模板表结构.json 里没有这个模板：%s" % template_name)
        say("先在 template_reader 里把读取器跑一遍（python main.py \"模板目录\"），生成结构信息。")
        return 1

    orig_path = TEMPLATE_DIR / template_name
    filled_path = OUT_DIR / (Path(template_name).stem + "-已填写.docx")
    if not filled_path.exists():
        say("还没有「已填写」文件，先跑 fill.py：%s" % filled_path)
        return 1

    doc_o = Document(str(orig_path))
    doc_f = Document(str(filled_path))
    if not doc_o.tables or not doc_f.tables:
        say("模板或产物里没有表格，无法验收。")
        return 1

    lines = []
    problems = []          # (类别, 描述)

    lines.append("写回验收报告 · %s（Word 模板）" % template_name)
    lines.append("生成时间：%s" % time.strftime("%Y-%m-%d %H:%M"))
    lines.append("原模板：%s" % orig_path)
    lines.append("已填写：%s" % filled_path)
    lines.append("")

    # ---------- 1. 值：逐块逐格核对（与 xlsx 版同规则） ----------
    expected = {}          # (表号, 行, 列) -> 期望值
    fill_cells = []        # [(表号, 行)]，待填区
    value_ok = 0
    value_bad = []
    block_brief = []

    # python-docx 的 table.cell() 每次都要走整张表的 XML；先把行网格取出来复用
    grid_f = {i: [row.cells for row in t.rows]
              for i, t in enumerate(doc_f.tables, start=1)}
    grid_o = {i: [row.cells for row in t.rows]
              for i, t in enumerate(doc_o.tables, start=1)}

    for bi, info in enumerate(infos, start=1):
        columns = writable_columns(info)
        slots = fill_slots(info)
        if not slots:
            continue
        tables_in_block = sorted({t for t, _r in slots})
        fill_cells.extend([(t, r) for t, r in slots])

        per_col_items = {}
        for col in columns:
            res = load_extract(col["title"])
            per_col_items[col["title"]] = (res or {}).get("items", [])
        sub = {c: [it for it in items if int(it.get("block") or 1) == bi]
               for c, items in per_col_items.items()}
        sub = {c: v for c, v in sub.items() if v}
        slot_map = build_slot_map(sub, len(slots))

        n_set = 0
        for col in columns:
            for idx, it in slot_map.get(col["title"], []):
                t, r = slots[idx]
                expected[(t, r, col["col"])] = norm(it.get("value"))
                n_set += 1

        n_rec = 0
        for col in columns:
            seq = slot_map.get(col["title"], [])
            if seq and col["title"] != "序号":
                n_rec = max(n_rec, max(i for i, _ in seq) + 1)
        for col in columns:
            if col["title"] == "序号" and not sub.get("序号"):
                for i in range(n_rec):
                    t, r = slots[i]
                    expected[(t, r, col["col"])] = str(i + 1)
                    n_set += 1

        block_brief.append("块%d（%s）：%d 个位置，期望有值 %d 格"
                           % (bi, info.get("name") or "?", len(slots), n_set))

        for t, r in slots:
            if t not in grid_f or r - 1 >= len(grid_f[t]):
                value_bad.append(("表%d 第%d行" % (t, r), "（整行）", "表不存在", ""))
                continue
            for col in columns:
                exp = expected.get((t, r, col["col"]), "")
                got = norm(grid_f[t][r - 1][col["col"] - 1].text)
                if exp == got:
                    value_ok += 1
                else:
                    value_bad.append(("表%d 第%d行" % (t, r) if len(infos) > 1
                                      else "第%d行" % r, col["title"], exp, got))

    n_fill_cells = sum(len(writable_columns(i)) for i in infos
                       for _t, _r in fill_slots(i))
    lines.append("【1. 值 · 待填区】")
    for b in block_brief:
        lines.append("  " + b)
    lines.append("  与抽取结果一致：%d / %d 格（含未用到、应为空的格子）"
                 % (value_ok, n_fill_cells))
    if value_bad:
        lines.append("  以下格子与期望不一致：")
        for coord, title, exp, got in value_bad[:60]:
            lines.append("    %s [%s] 期望 = %s ｜ 实际 = %s"
                         % (coord, title, exp if exp else "（空）", got if got else "（空）"))
        if len(value_bad) > 60:
            lines.append("    …… 共 %d 处" % len(value_bad))
        problems.append(("值", "待填区有 %d 格不一致" % len(value_bad)))
    lines.append("")

    # ---------- 2. 值：待填区以外，必须纹丝未动 ----------
    # 逐张表比；每张表自己的待填行集合不同（表1 只到 25 行、表2 到 29 行）
    slot_by_table = {}
    for t, r in fill_cells:
        slot_by_table.setdefault(t, set()).add(r)

    diff_other = []
    n_tab = min(len(doc_o.tables), len(doc_f.tables))
    for ti in range(1, n_tab + 1):
        ro, rf = grid_o[ti], grid_f[ti]
        n_rows = max(len(ro), len(rf))
        n_cols = max(len(doc_o.tables[ti - 1].columns),
                     len(doc_f.tables[ti - 1].columns))
        skip = slot_by_table.get(ti, set())
        for r in range(1, n_rows + 1):
            if r in skip or r - 1 >= len(ro) or r - 1 >= len(rf):
                continue
            for c in range(1, n_cols + 1):
                if c - 1 >= len(ro[r - 1]) or c - 1 >= len(rf[r - 1]):
                    continue
                vo = norm(ro[r - 1][c - 1].text)
                vf = norm(rf[r - 1][c - 1].text)
                if vo != vf:
                    diff_other.append(("表%d第%d行第%d列" % (ti, r, c) if n_tab > 1
                                       else "第%d行第%d列" % (r, c), vo, vf))

    lines.append("【2. 值 · 待填区以外（表头/标题/说明行等）】")
    if diff_other:
        lines.append("  发现 %d 处被改动（不应该发生）：" % len(diff_other))
        for coord, vo, vf in diff_other[:30]:
            lines.append("    %s：原 = %s ｜ 现 = %s"
                         % (coord, vo if vo else "（空）", vf if vf else "（空）"))
        if len(diff_other) > 30:
            lines.append("    …… 共 %d 处" % len(diff_other))
        problems.append(("值", "待填区以外有 %d 处被改动" % len(diff_other)))
    else:
        lines.append("  全表其余格子一致，未被改动。")
    lines.append("")

    # ---------- 3. 格式（Word 版检查项） ----------
    fmt_bad = []

    for ti in range(1, n_tab + 1):
        to, tf = doc_o.tables[ti - 1], doc_f.tables[ti - 1]
        size_o = (len(to.rows), len(to.columns))
        size_f = (len(tf.rows), len(tf.columns))
        tag = "表%d " % ti if n_tab > 1 else ""
        if size_o != size_f:
            fmt_bad.append("%s表格行列数变了：原 %d×%d ｜ 现 %d×%d"
                           % (tag, size_o[0], size_o[1], size_f[0], size_f[1]))
        style_o = getattr(to.style, "name", None)
        style_f = getattr(tf.style, "name", None)
        if style_o != style_f:
            fmt_bad.append("%s表格样式变了：原 %s ｜ 现 %s" % (tag, style_o, style_f))

    align_bad = []
    for ti, rows in slot_by_table.items():
        if ti not in grid_o or ti not in grid_f:
            continue
        info = infos[ti - 1] if ti - 1 < len(infos) else infos[0]
        for r in sorted(rows):
            if r - 1 >= len(grid_o[ti]) or r - 1 >= len(grid_f[ti]):
                continue
            for col in writable_columns(info):
                c = col["col"] - 1
                if c >= len(grid_o[ti][r - 1]) or c >= len(grid_f[ti][r - 1]):
                    continue
                po = grid_o[ti][r - 1][c].paragraphs[0].alignment
                pf = grid_f[ti][r - 1][c].paragraphs[0].alignment
                if po != pf:
                    align_bad.append("表%d第%d行第%d列" % (ti, r, col["col"]))
    if align_bad:
        fmt_bad.append("待填格段落对齐变了：%s%s"
                       % ("、".join(align_bad[:10]), "……" if len(align_bad) > 10 else ""))

    n_po = len(doc_o.paragraphs)
    n_pf = len(doc_f.paragraphs)
    if n_po != n_pf:
        fmt_bad.append("正文段落数变了：原 %d ｜ 现 %d" % (n_po, n_pf))

    lines.append("【3. 格式】")
    if fmt_bad:
        for b in fmt_bad:
            lines.append("  " + b)
        problems.append(("格式", "；".join(fmt_bad)))
    else:
        size_o = (len(doc_o.tables[0].rows), len(doc_o.tables[0].columns))
        lines.append("  表格数：一致（%d 张）" % n_tab)
        lines.append("  表格行列数：一致（%d×%d）" % size_o)
        lines.append("  表格样式：一致（%s）"
                     % (getattr(doc_o.tables[0].style, "name", None) or "无样式"))
        lines.append("  待填格段落对齐：一致（写值未动对齐格式）")
        lines.append("  正文段落数：一致（%d 段）" % n_po)
    lines.append("")

    # ---------- 4. 结论 ----------
    lines.append("【4. 结论】")
    if not problems:
        lines.append("  全部通过 —— 值填对了，格式一点没动，可以交卷。")
    else:
        lines.append("  发现 %d 类问题，需要处理：" % len(problems))
        for cat, desc in problems:
            lines.append("    - [%s] %s" % (cat, desc))

    report_path = OUT_DIR / ("验收报告-%s-word.txt" % Path(template_name).stem)
    report_path.write_text("\n".join(lines), encoding="utf-8")

    say("")
    say("\n".join(lines))
    say("")
    say("验收报告：%s" % report_path)
    return 0 if not problems else 2


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    template_name = args[0] if args else DEFAULT_TEMPLATE

    # Word（docx）模板 → 走 Word 版验收路径
    if Path(template_name).suffix.lower() == ".docx":
        return verify_docx(template_name)

    info = load_sheet_info(template_name)
    if not info:
        say("模板表结构.json 里没有这个模板：%s" % template_name)
        return 1

    sheet_name = info["name"]
    columns = info["columns"]
    fr = info.get("fill_region") or {}
    start_row, end_row = fr.get("start_row"), fr.get("end_row")

    orig_path = TEMPLATE_DIR / template_name
    filled_path = OUT_DIR / (Path(template_name).stem + "-已填写.xlsx")
    if not filled_path.exists():
        say("还没有「已填写」文件，先跑 fill.py：%s" % filled_path)
        return 1

    wb_o = load_workbook(orig_path)
    wb_f = load_workbook(filled_path)
    ws_o = wb_o[sheet_name]
    ws_f = wb_f[sheet_name]

    # 模板只有表头时（fill_region = null，真实模板很常见），待填区是填表阶段
    # "往下补行"补出来的 —— 验收这边不能拿 fill_region 当锚（它是 None，取下标直接崩），
    # 改成按已填写文件里真实的最后一行有内容的行来定范围。2026-11-03 COVID 模板踩的。
    deduced = False
    if not start_row:
        hrows = info.get("header_rows") or (
            [info["header_row"]] if info.get("header_row") else [1])
        start_row = max(hrows) + 1
        end_row = start_row - 1
        for r in range(start_row, ws_f.max_row + 1):
            if any(norm(ws_f.cell(row=r, column=c["col"]).value) for c in columns):
                end_row = r
        deduced = True

    lines = []
    problems = []          # (类别, 描述)

    lines.append("写回验收报告 · %s" % template_name)
    lines.append("生成时间：%s" % time.strftime("%Y-%m-%d %H:%M"))
    lines.append("原模板：%s" % orig_path)
    lines.append("已填写：%s" % filled_path)
    lines.append("")
    if deduced:
        lines.append("待填区：模板只给了表头、没留空行 → 按已填写文件实际写到的范围验收"
                     "（第 %d ~ %d 行，共 %d 行）" % (start_row, end_row, end_row - start_row + 1))
        lines.append("")

    # ---------- 1. 值：待填区 vs 抽取结果 ----------
    expected = {}          # (行, 列) -> 期望值
    n_expect = 0
    per_col_items = {}
    for col in columns:
        res = load_extract(col["title"])
        per_col_items[col["title"]] = (res or {}).get("items", [])

    n_slots = end_row - start_row + 1
    slot_map = build_slot_map(per_col_items, n_slots)
    for col in columns:
        for idx, it in slot_map.get(col["title"], []):
            r = start_row + idx
            if r <= end_row:
                expected[(r, col["col"])] = norm(it.get("value"))
                n_expect += 1

    # 「序号」列：fill.py 会按记录数自动编号 1..N —— 验收期望值跟上同一规则
    n_rec = 0
    for col in columns:
        seq = slot_map.get(col["title"], [])
        if seq and col["title"] != "序号":
            n_rec = max(n_rec, max(i for i, _ in seq) + 1)
    for col in columns:
        if col["title"] == "序号" and not per_col_items.get("序号"):
            for i in range(n_rec):
                r = start_row + i
                if r <= end_row:
                    expected[(r, col["col"])] = str(i + 1)
                    n_expect += 1

    value_ok = 0
    value_bad = []
    n_fill_cells = 0
    for r in range(start_row, end_row + 1):
        for col in columns:
            n_fill_cells += 1
            exp = expected.get((r, col["col"]), "")
            got = norm(ws_f.cell(row=r, column=col["col"]).value)
            if exp == got:
                value_ok += 1
            else:
                value_bad.append((r, col["title"], exp, got))

    lines.append("【1. 值 · 待填区】")
    lines.append("  与抽取结果一致：%d / %d 格（含未用到、应为空的格子）" % (value_ok, n_fill_cells))
    if value_bad:
        lines.append("  以下格子与期望不一致：")
        for r, title, exp, got in value_bad:
            lines.append("    第 %d 行 [%s] 期望 = %s ｜ 实际 = %s"
                         % (r, title, exp if exp else "（空）", got if got else "（空）"))
        problems.append(("值", "待填区有 %d 格不一致" % len(value_bad)))
    lines.append("")

    # ---------- 2. 值：待填区以外，必须纹丝未动 ----------
    diff_other = []
    max_r = max(ws_o.max_row, ws_f.max_row)
    max_c = max(ws_o.max_column, ws_f.max_column)
    for r in range(1, max_r + 1):
        if start_row <= r <= end_row:
            continue
        for c in range(1, max_c + 1):
            vo = norm(ws_o.cell(row=r, column=c).value)
            vf = norm(ws_f.cell(row=r, column=c).value)
            if vo != vf:
                coord = ws_o.cell(row=r, column=c).coordinate
                diff_other.append((coord, vo, vf))

    lines.append("【2. 值 · 待填区以外（表头/标题/说明行等）】")
    if diff_other:
        lines.append("  发现 %d 处被改动（不应该发生）：" % len(diff_other))
        for coord, vo, vf in diff_other:
            lines.append("    %s：原 = %s ｜ 现 = %s" % (coord, vo if vo else "（空）", vf if vf else "（空）"))
        problems.append(("值", "待填区以外有 %d 处被改动" % len(diff_other)))
    else:
        lines.append("  全表其余格子一致，未被改动。")
    lines.append("")

    # ---------- 3. 格式 ----------
    # 模板只给表头时，补出来的数据行**原本不存在**，没有"原样式"可比 ——
    # 拿空单元格的默认样式去比对必然报几千处差异（假阳性，2026-11-03 COVID 模板实测
    # 报过 15252 处）。所以：只比原模板真实存在的行；新补行单独说明。
    cmp_rows = ws_o.max_row if deduced else max_r
    style_diffs = []
    for r in range(1, cmp_rows + 1):
        for c in range(1, max_c + 1):
            ko = style_key(ws_o.cell(row=r, column=c))
            kf = style_key(ws_f.cell(row=r, column=c))
            if ko != kf:
                style_diffs.append(ws_o.cell(row=r, column=c).coordinate)

    lines.append("【3. 格式】")
    if deduced:
        n_new = max(0, end_row - max(ws_o.max_row, start_row - 1))
        lines.append("  原模板只有表头（共 %d 行），本次补出 %d 行数据行 —— "
                     "新行没有原样式可比，按默认样式写入（不加粗、不染色）"
                     % (ws_o.max_row, n_new))

    if style_diffs:
        show = "、".join(style_diffs[:10]) + ("……" if len(style_diffs) > 10 else "")
        lines.append("  单元格样式差异：%d 处（%s）" % (len(style_diffs), show))
        problems.append(("格式", "样式差异 %d 处" % len(style_diffs)))
    else:
        lines.append("  单元格样式差异：0 处（字体/边框/对齐/填充/数字格式全一致）")

    # 合并单元格
    merged_o = {str(x) for x in ws_o.merged_cells.ranges}
    merged_f = {str(x) for x in ws_f.merged_cells.ranges}
    if merged_o == merged_f:
        lines.append("  合并单元格：一致（%s）" % ("、".join(sorted(merged_o)) if merged_o else "无"))
    else:
        lines.append("  合并单元格：不一致！原 = %s ｜ 现 = %s" % (sorted(merged_o), sorted(merged_f)))
        problems.append(("格式", "合并单元格不一致"))

    # 列宽 / 行高
    wo, wf = width_map(ws_o), width_map(ws_f)
    if wo == wf:
        lines.append("  列宽：一致")
    else:
        lines.append("  列宽：不一致！原 = %s ｜ 现 = %s" % (wo, wf))
        problems.append(("格式", "列宽不一致"))
    ho, hf = height_map(ws_o), height_map(ws_f)
    if ho == hf:
        lines.append("  行高：一致")
    else:
        lines.append("  行高：不一致！原 = %s ｜ 现 = %s" % (ho, hf))
        problems.append(("格式", "行高不一致"))

    # 数据验证 / 条件格式
    try:
        dv_o = len(ws_o.data_validations.dataValidation)
        dv_f = len(ws_f.data_validations.dataValidation)
        if dv_o == dv_f:
            lines.append("  数据验证（下拉等）：一致（%d 处）" % dv_o)
        else:
            lines.append("  数据验证：不一致！原 %d 处 / 现 %d 处" % (dv_o, dv_f))
            problems.append(("格式", "数据验证数量不一致"))
    except Exception as e:
        lines.append("  数据验证：读取异常（跳过）%s" % e)
    try:
        cf_o = sum(len(rules) for rules in ws_o.conditional_formatting._cf_rules.values())
        cf_f = sum(len(rules) for rules in ws_f.conditional_formatting._cf_rules.values())
        if cf_o == cf_f:
            lines.append("  条件格式：一致（%d 处）" % cf_o)
        else:
            lines.append("  条件格式：不一致！原 %d 处 / 现 %d 处" % (cf_o, cf_f))
            problems.append(("格式", "条件格式数量不一致"))
    except Exception as e:
        lines.append("  条件格式：读取异常（跳过）%s" % e)
    lines.append("")

    # ---------- 4. 结论 ----------
    lines.append("【4. 结论】")
    if not problems:
        lines.append("  全部通过 —— 值填对了，格式一点没动，可以交卷。")
    else:
        lines.append("  发现 %d 类问题，需要处理：" % len(problems))
        for cat, desc in problems:
            lines.append("    - [%s] %s" % (cat, desc))

    report_path = OUT_DIR / ("验收报告-%s.txt" % Path(template_name).stem)
    report_path.write_text("\n".join(lines), encoding="utf-8")

    say("")
    say("\n".join(lines))
    say("")
    say("验收报告：%s" % report_path)
    return 0 if not problems else 2


if __name__ == "__main__":
    sys.exit(main())
