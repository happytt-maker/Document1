# -*- coding: utf-8 -*-
"""
template_reader/reader.py
-------------------------
模板表结构读取器。

目标：给它一张陌生的空白模板表，它输出一份「填表说明书」：
    - 表头在哪几行
    - 每一列叫什么
    - 哪些位置需要填（待填区域的范围，可能不止一段）

═══════════════════════════════════════════════════════════════════
  2026-10-26 改造：从「只认规整表」到「认三种不规整表头」
═══════════════════════════════════════════════════════════════════

改造前的做法（已存档在 _before.json 的失败证据里）：
    ① 找「非空格子数最多的那一行」当表头
    ② 按那一行取列名
    ③ 表头下一段连续空行当待填区

这套做法只对「单层、无合并、不跨页」的表格成立。三种真实形态全会翻车：

  【多层表头】父列是合并格，只有左上角有值，其余是 None
      → 「非空格最多」会选中子层那行，父列名（上半年采购）整个丢掉；
        实测模板6：9 列里 4 列 title = None，填进去必然错位。

  【斜线表头】一个格子里塞两个维度（"费用类别\\n月份"）
      → 整串被当成一个列名，下游拿这个名字去源文档里找，必然抽空；
        而且第一列往往是「行标签列」（差旅费/办公费…），
        它不是待填数据，却会让「连续空行」的判断第一行就断掉
        （实测模板7：待填区只剩 1 行）。

  【跨页表】第二页重复表头，中间夹「续表」
      → 「第一段连续空行」到「续表」就停，第二页的行永远填不到
        （实测模板8：15~23 行整段丢失）；
        Word 里跨页常被做成两张独立的 table，会变成两张互不相干的表
        （实测模板9：分裂成「表格1」「表格2」）。

改造后的四步（结构上仍然只有"格子表"这一个抽象，与文件格式无关）：

  第 1 步 合并铺平：把合并区域的值铺满整个区域 → dense grid
          （父列「上半年采购」在它跨的每一列上都有值了）
  第 2 步 表头块：找「连续若干行都长得像表头」的段，而不是单行；
          再剔掉整行同值的（那是大标题，不是表头层）
  第 3 步 列名合成：每列把各层文字摞起来；子名全局唯一就用子名，
          重复才用「父-子」区分；斜线格拆成主名 + 别名
  第 4 步 分段：待填区可以有多段，「续表」/重复表头负责衔接；
          第一列若是行标签列，判断空行时把它排除

═══════════════════════════════════════════════════════════════════
  输出结构（新增字段都向后兼容：老的下游只读 header_row/columns/fill_region）
═══════════════════════════════════════════════════════════════════

    {
      "file": "模板6-多层表头设备采购台账.xlsx",
      "format": "xlsx",
      "sheet_count": 1,
      "sheets": [
        {
          "name": "设备采购台账",
          "header_rows": [3, 4],       # 表头占哪几行（从上到下）
          "header_row": 4,             # 列名层（= header_rows 最后一行）
          "header_block_note": "2 层表头",
          "columns": [
            {"col": 4, "title": "上半年采购-数量", "aliases": [...],
             "header_parts": ["上半年采购", "数量"],
             "slash": false, "slash_parts": []},
            ...
          ],
          "label_cols": [],            # 行标签列（不填数据）
          "fill_region": {...},        # 第一段（兼容老下游）
          "segments": [                # 全部待填段
            {"start_row": 5, "end_row": 13, "row_count": 9, "table": 1}
          ],
          "tables": [1]                # 用到了 docx 的第几张表（xlsx 恒为 1）
        }
      ]
    }
"""

import math
import re
from pathlib import Path

import openpyxl
from docx import Document

# 列名文字的平均长度超过它 → 更像「大标题」而不是「表头」
MAX_TITLE_AVG_LEN = 14

# 斜线格里拆出来的每一段，最长允许多少字（超过就不当斜线格拆）
MAX_SLASH_SEG_LEN = 10

# 「续表」这类分隔行的标记词
CONTINUE_WORDS = ("续表", "续页", "接上页", "接下页", "续下页", "转下页",
                  "continued", "cont'd")


# =================================================================
# 第一步：把工作表读成「二维格子表」，再把合并区域铺平
# =================================================================

def read_grid(sheet):
    """把工作表读成二维列表（**不**铺平合并格，只左上角有值）。

    返回 [[值或None, ...], ...]，坐标从 0 开始。空字符串算 None。
    """
    grid = []
    for row in sheet.iter_rows():
        line = []
        for cell in row:
            v = cell.value
            if v is None:
                line.append(None)
            elif isinstance(v, str) and v.strip() == "":
                line.append(None)
            else:
                line.append(str(v).strip() if isinstance(v, str) else v)
        grid.append(line)
    return grid


def _pad(rows, n_cols):
    """把锯齿状的行补齐到 n_cols 宽（docx 的行长度可能不一致）。"""
    out = []
    for r in rows:
        line = list(r[:n_cols])
        line += [None] * (n_cols - len(line))
        out.append(line)
    return out


def dense_xlsx(ws, grid):
    """把合并单元格「铺平」：合并区里的每一格都填上左上角的值。

    这是整个改造的地基。举例（模板6 的父层，B3:C3 合并写「采购信息」）：

        原始 grid : [序号, 设备名称, 规格型号, 上半年采购, None, 下半年采购, None, 供应商信息, None]
        dense grid: [序号, 设备名称, 规格型号, 上半年采购, 上半年采购,
                     下半年采购, 下半年采购, 供应商信息, 供应商信息]

    铺平之后，「这一列属于哪个父列」就成了一个普通的取列值问题，
    不需要再到处查合并区间 —— 后面所有逻辑都只面对普通二维数组。

    返回 (dense, spans)
        spans: {(row, col): (r0, c0, r1, c1)}，1-based，供排查用
    """
    n_cols = max(
        len(grid[0]) if grid else 0,
        ws.max_column or 0,
        max((len(r) for r in grid), default=0),
    )
    grid = _pad(grid, n_cols)

    spans = {}
    for rng in ws.merged_cells.ranges:
        r0, c0, r1, c1 = rng.min_row, rng.min_col, rng.max_row, rng.max_col
        for r in range(r0, r1 + 1):
            for c in range(c0, c1 + 1):
                spans[(r, c)] = (r0, c0, r1, c1)

    dense = []
    for r_idx in range(1, len(grid) + 1):
        line = []
        for c_idx in range(1, n_cols + 1):
            v = grid[r_idx - 1][c_idx - 1]
            if v is None:
                sp = spans.get((r_idx, c_idx))
                if sp:
                    v = grid[sp[0] - 1][sp[1] - 1]
            line.append(v)
        dense.append(line)
    return dense, spans


def _norm_cell_text(text):
    """规整 docx 单元格文字：合并格的 text 有时会叠加（'吴晓敏\\n吴晓敏'）。"""
    v = (text or "").strip()
    if "\n" in v:
        parts = [p.strip() for p in v.split("\n") if p.strip()]
        if parts and len(set(parts)) == 1:
            return parts[0]
    return v or None


def dense_docx(table):
    """把 python-docx 表格转成**已铺平**的 dense grid。

    python-docx 的 row.cells 是按「表格网格位置」展开的：
        横向合并（gridSpan）的格子会在该行里重复出现同一个 _tc；
        纵向合并（vMerge）的两行都指向同一个 _tc。
    所以直接遍历 row.cells（**不要**按 _tc 去重）就天然得到铺平效果。

    【踩过的坑】老版本按 cell._tc 去重了 —— 那是为了消掉"同一格重复出现"，
    但代价是行变短、后面的列整体左移，8 列的表会被读成 6 列，写入必然错位。
    """
    rows = []
    for row in table.rows:
        rows.append([_norm_cell_text(cell.text) for cell in row.cells])
    n_cols = max((len(r) for r in rows), default=0)
    return _pad(rows, n_cols)


# =================================================================
# 第二步：找表头块（连续的若干行）
# =================================================================

def _row_values(row, n_cols):
    return [v for v in row[:n_cols] if v is not None and str(v).strip() != ""]


def _looks_like_header(row, n_cols):
    """这一行长得像不像表头。**

    五条判据（都用真实模板校过）：
      1. 铺平后的非空格数 ≥ 列数的一半（合并的父层铺平后自然满足）
      2. 不带冒号（带冒号的是「填报单位：」这类说明栏）
      3. 文字短（平均 ≤ 14 字）——大标题「2026 年 9 月原材料入库台账」被这一条挡掉
      4. 不是纯数字行（防止把数据行当表头）
      5. 去重后不止一个值（整行合并的是大标题 / 「续表」分隔行）
    """
    vals = _row_values(row, n_cols)
    if len(vals) < max(2, math.ceil(n_cols * 0.5)):
        return False
    texts = [str(v).strip() for v in vals]
    if any("：" in t or ":" in t for t in texts):
        return False
    avg = sum(len(t) for t in texts) / len(texts)
    if avg > MAX_TITLE_AVG_LEN:
        return False
    if all(re.fullmatch(r"-?\d+(\.\d+)?", t) for t in texts):
        return False
    if len(set(texts)) == 1:
        return False        # 整行同值 = 合并的大标题，不是表头层
    return True


def find_header_block(dense, n_cols):
    """找表头块，返回 (header_rows, note)。

    header_rows 是从上到下的行号列表（1-based）。单层表头就是长度 1 的列表。

    做法：先标出所有「像表头」的行，再取**最长的那一段连续行**。
    为什么取最长：多层表头的每一层单独看都像表头，连起来才完整；
    而大标题、说明栏即使偶然通过判据，也只会是孤立的一两行，
    拼不过「父层+子层」这种更长的连续段（同长度时取最靠上的）。
    """
    cands = [i for i, row in enumerate(dense, start=1) if _looks_like_header(row, n_cols)]
    if not cands:
        return [], "没找到表头"

    best, cur = [], []
    for i in cands:
        cur = cur + [i] if (cur and i == cur[-1] + 1) else [i]
        if len(cur) > len(best):
            best = list(cur)

    if len(best) > 1:
        note = "%d 层表头（第 %d~%d 行）" % (len(best), best[0], best[-1])
    else:
        note = "单层表头（第 %d 行）" % best[0]
    return best, note


def find_header_row(grid):
    """兼容老接口：只返回表头块的最后一行（= 列名层）。"""
    n_cols = max((len(r) for r in grid), default=0)
    rows, _ = find_header_block(_pad(grid, n_cols), n_cols)
    return rows[-1] if rows else None


# =================================================================
# 第三步：合成列名（多层拼接 + 斜线拆词）
# =================================================================

def split_slash(text):
    """把斜线表头的一格拆成几段。

    真实写法有三种，都要能拆：
        "费用类别\\n月份"     换行（Excel 里 Alt+Enter 敲出来的，最常见）
        "费用类别\\\\月份"     反斜线
        "费用类别/月份"       正斜线

    只在「拆得出 2 段以上、每段都不长」时才认，避免把
    「单价（元）」「完成/未完成」这种正常列名误拆成主名+别名。
    """
    t = str(text).strip()
    if not t:
        return [t]
    segs = []
    for piece in re.split(r"[\r\n]+", t):
        p = piece.strip()
        if not p:
            continue
        # 一个格子里只有一条斜线时，才按斜线再拆
        if p.count("\\") == 1:
            a, b = p.split("\\")
            segs.extend([a.strip(), b.strip()])
        elif p.count("/") == 1 and not re.search(r"\d", p):
            a, b = p.split("/")
            segs.extend([a.strip(), b.strip()])
        else:
            segs.append(p)

    segs = [s for s in segs if s]
    if len(segs) < 2:
        return [t]
    if any(len(s) > MAX_SLASH_SEG_LEN for s in segs):
        return [t]
    # 去重保序
    return list(dict.fromkeys(segs))


def compose_columns(dense, header_rows, n_cols):
    """按表头的每一层，把列名摞起来。

    规则（每条都有对应的真实模板）：
      · 纵向合并且上下文字相同 → 只留一份（模板6 的「序号」）
      · 子名在本表内**唯一** → 直接用子名，不硬拼父名
        （模板6 的「供应商名称」不该变成「供应商信息-供应商名称」）
      · 子名**重复** → 用父名区分（模板6 的「数量」在两组下各有一个，
        变成「上半年采购-数量」「下半年采购-数量」，否则两列同名、
        下游分不清该往哪列填）
      · 斜线格 → 第一段当主名，其余段进 aliases（模板7 的「费用类别\\n月份」）
    """
    raw_parts = []
    for c in range(1, n_cols + 1):
        parts = []
        for r in header_rows:
            v = dense[r - 1][c - 1] if r - 1 < len(dense) else None
            if v is None or str(v).strip() == "":
                continue
            v = str(v).strip()
            if parts and parts[-1] == v:
                continue        # 纵向合并：同一格铺平后上下重复
            parts.append(v)
        raw_parts.append(parts)

    bottom = [p[-1] if p else None for p in raw_parts]
    counter = {}
    for b in bottom:
        if b:
            counter[b] = counter.get(b, 0) + 1

    columns = []
    for i, parts in enumerate(raw_parts, start=1):
        title, aliases, slash, slash_parts = None, [], False, []
        if parts:
            sub = parts[-1]
            segs = split_slash(sub)
            if len(segs) > 1:
                slash, slash_parts = True, segs
                sub = segs[0]
                aliases.extend(segs)

            if counter.get(parts[-1], 0) > 1 and len(parts) >= 2:
                title = "%s-%s" % (parts[-2], sub)
            else:
                title = sub

            aliases.append("-".join(parts))
            aliases.append("".join(parts))
            if parts:
                aliases.append(parts[0])
            aliases = [a for a in dict.fromkeys(aliases) if a and a != title]
        columns.append({
            "col": i,
            "title": title,
            "aliases": aliases,
            "header_parts": parts,
            "slash": slash,
            "slash_parts": slash_parts,
        })
    return columns


def read_columns(grid, header_row):
    """兼容老接口：单行表头 → 列名列表。"""
    n_cols = max((len(r) for r in grid), default=0)
    rows = [header_row] if header_row else []
    return compose_columns(_pad(grid, n_cols), rows, n_cols)


# =================================================================
# 第四步：待填区（可多段）+ 行标签列
# =================================================================

def detect_label_cols(dense, header_end, n_cols):
    """认出「行标签列」——第 1 列从上到下全是分类名（差旅费/办公费…），
    它是表格结构的一部分，不是待填数据。

    为什么要单独认：判断"这一行空不空"时如果不排除它，
    模板7 的待填区第 1 行就会被"差旅费"占住，一段都找不到。

    判据用的是更本质的那个特征：**"只有它自己填了、别的列全空"的行**
    占「有内容的行」的一半以上 —— 这正是"行标签"的形态。

        · 普通数据行：一格有值，同一行其它列也常有值（不会 solo）
        · 合计行：合并铺平后整行都有值（不会 solo）
        · 尾部空行：根本不算"有内容的行"，不会稀释比例

    但光靠 solo 还不够 —— 结尾的「本季度合计投诉数：」「按期结论：」这类
    说明栏也是"只有第一列有值"。所以再加三条，把说明栏挡掉：
        · 分类名至少 3 个（差旅费/办公费/招待费…）
        · 不带冒号（带冒号的是"XX："说明栏）
        · 不超过 12 字

    【踩过的坑】第一版只用「第一列有值数 / 总行数 ≥ 0.8」，
    模板7 因为尾部留了空行被稀释成 0.75，判不出来，待填区整段丢失。
    第二版改用 solo 占比，又踩到模板4：最后两行说明栏被误认成分类名，
    待填区从「4~9」错误地拉长到「4~11」（会把数据填到说明栏上）。
    第三版（本版）加上"至少 3 个 + 无冒号 + 短"，两边的坑都避开。
    """
    rows = list(range(header_end + 1, len(dense) + 1))
    if len(rows) < 2:
        return []

    solo_texts, filled_rows = [], 0
    for r in rows:
        row = dense[r - 1]
        vals = [v for v in row[:n_cols] if v is not None and str(v).strip() != ""]
        if not vals:
            continue                      # 全空行：不参与统计
        filled_rows += 1
        first_has = row[0] is not None and str(row[0]).strip() != ""
        others_empty = all(
            row[c] is None or str(row[c]).strip() == ""
            for c in range(1, n_cols)
        )
        if first_has and others_empty:
            solo_texts.append(str(row[0]).strip())

    if not filled_rows or len(solo_texts) / filled_rows < 0.5:
        return []
    if len(solo_texts) < 3:
        return []                          # 太少，更像结尾说明栏
    if any("：" in t or ":" in t for t in solo_texts):
        return []                          # 带冒号 = 说明栏
    if any(len(t) > 12 for t in solo_texts):
        return []                          # 太长，不像分类名
    return [1]


def _is_blank_row(row, n_cols, ignore_cols):
    for c in range(1, n_cols + 1):
        if c in ignore_cols:
            continue
        v = row[c - 1] if c - 1 < len(row) else None
        if v is not None and str(v).strip() != "":
            return False
    return True


def _is_continue_marker(row, n_cols):
    """这一行是不是「续表」这种分页分隔行。"""
    vals = [str(v).strip() for v in _row_values(row, n_cols)]
    if not vals:
        return False
    joined = " ".join(vals).lower()
    if len(set(vals)) == 1:             # 常见：整行合并写「续表」
        return any(w in joined for w in CONTINUE_WORDS)
    return len(vals) <= 2 and any(w in joined for w in CONTINUE_WORDS)


def _row_matches_titles(row, columns, n_cols):
    """这一行是不是「重复的表头」（跨页续表的第二页表头）。

    判据：非空值里，有 60% 以上能在本表的列名/别名里找到。
    """
    vals = [str(v).strip() for v in _row_values(row, n_cols)]
    if len(vals) < 2:
        return False
    known = set()
    for c in columns:
        if c.get("title"):
            known.add(c["title"])
        for a in c.get("aliases") or []:
            known.add(a)
    hit = sum(1 for v in vals if v in known)
    return hit / len(vals) >= 0.6


def scan_segments(dense, header_rows, columns, n_cols, label_cols):
    """从表头下面开始扫，切出所有待填段。

    三种情况：
        · 空行 → 攒成一段
        · 「续表」→ 跳过它，再把紧随其后的重复表头跳掉，接着攒新的一段
        · 重复表头（没有「续表」字样，比如 Excel 里手动复制的）→ 同样跳过
        · 其它有内容行（合计/备注/盖章）→ 结束
    """
    n_rows = len(dense)
    header_end = header_rows[-1] if header_rows else 0
    n_header_layers = max(1, len(header_rows))

    segs = []
    r = header_end + 1
    while r <= n_rows:
        if _is_blank_row(dense[r - 1], n_cols, label_cols):
            start = r
            while r <= n_rows and _is_blank_row(dense[r - 1], n_cols, label_cols):
                r += 1
            segs.append({"start_row": start, "end_row": r - 1,
                         "row_count": r - start, "table": 1})
            continue

        if _is_continue_marker(dense[r - 1], n_cols):
            r += 1
            # 续表后面通常紧跟重复表头（可能有多层），逐层跳
            skipped = 0
            while (r <= n_rows and skipped < n_header_layers
                   and _row_matches_titles(dense[r - 1], columns, n_cols)):
                r += 1
                skipped += 1
            continue

        if _row_matches_titles(dense[r - 1], columns, n_cols):
            # 没有「续表」字样，但表头原样重复了一遍
            r += n_header_layers
            continue

        break

    return segs


def find_fill_region(grid, header_row):
    """兼容老接口：只返回第一段待填区。"""
    n_cols = max((len(r) for r in grid), default=0)
    dense = _pad(grid, n_cols)
    rows, _ = find_header_block(dense, n_cols)
    if not rows:
        return None
    label_cols = detect_label_cols(dense, rows[-1], n_cols)
    cols = compose_columns(dense, rows, n_cols)
    segs = scan_segments(dense, rows, cols, n_cols, label_cols)
    if not segs:
        return None
    s = segs[0]
    return {"start_row": s["start_row"], "end_row": s["end_row"],
            "row_count": s["row_count"]}


# =================================================================
# 第五步：组装一张表（xlsx / docx 共用后半段）
# =================================================================

def describe_dense(dense, n_cols):
    """dense grid → 填表说明书（这层完全不含文件格式概念）。"""
    header_rows, note = find_header_block(dense, n_cols)
    if not header_rows:
        return {
            "header_rows": [], "header_row": None, "header_block_note": note,
            "columns": [], "label_cols": [], "fill_region": None,
            "segments": [], "tables": [1],
        }

    columns = compose_columns(dense, header_rows, n_cols)
    label_cols = detect_label_cols(dense, header_rows[-1], n_cols)
    segments = scan_segments(dense, header_rows, columns, n_cols, label_cols)

    # 兜底：一段都没扫到时，再试一次「忽略第一列」——
    # 万一第一列是行标签但判据没认出来，也不能让待填区整个丢掉。
    if not segments and n_cols > 1 and label_cols != [1]:
        retry = scan_segments(dense, header_rows, columns, n_cols, [1])
        if retry:
            segments, label_cols = retry, [1]

    return {
        "header_rows": header_rows,
        "header_row": header_rows[-1],
        "header_block_note": note,
        "columns": columns,
        "label_cols": label_cols,
        "fill_region": dict(segments[0]) if segments else None,
        "segments": segments,
        "tables": [1],
    }


def read_template(xlsx_path):
    """读一张 xlsx 模板，返回完整的「填表说明书」。"""
    xlsx_path = Path(xlsx_path)
    wb = openpyxl.load_workbook(str(xlsx_path), data_only=True)

    sheets = []
    for sheet in wb.worksheets:
        grid = read_grid(sheet)
        dense, _spans = dense_xlsx(sheet, grid)
        n_cols = len(dense[0]) if dense else 0

        info = describe_dense(dense, n_cols)
        info.update({
            "name": sheet.title,
            "max_row": sheet.max_row,
            "max_column": n_cols,
            "header_ok": bool(info["header_rows"]),
            "header_note": "文字不带冒号，符合表头特征",
        })
        sheets.append(info)

    wb.close()
    return {"file": xlsx_path.name, "format": "xlsx",
            "sheet_count": len(sheets), "sheets": sheets}


# =================================================================
# docx（Word）模板
# =================================================================

def read_docx_grid(table):
    """兼容老接口：docx 表格 → 二维格子表（已铺平，与 dense 同口径）。"""
    return dense_docx(table)


def _paragraph_maps(doc):
    """返回 (before, between) 两张表：段落分别属于「某张表之前」和「两张表之间」。

    修的是这个错（2026-11-07，山东空气质量模板）：旧的 `_paragraphs_between_tables`
    用的是 0/1 两套下标混着数，`out[k]` 实际存的是「第 k 张表之后」的文字，
    调用方却按「第 k 张表之前」去取 —— 于是判断"这两张表之间有没有文字"时
    读到了下一张表后面的文字。最后一张表后面本来就是空的，就被判成
    「紧邻且无文字」→ 和第 3 张表的表头完全一样 → 合并成"跨页续表"，
    结果 3 张表只剩下 2 个块，潍坊那张（28 行）整块消失。

    现在两件事分开存，下标就是表序号（1 起）：
        before[i]  = 第 i 张表**之前**那段文字（通常是这张表的描述）
        between[i] = 第 i 张表与第 i+1 张表**之间**的文字
    """
    from docx.text.paragraph import Paragraph

    before, between = {}, {}
    pending, t_idx = [], 0
    for child in doc.element.body.iterchildren():
        if child.tag.endswith("}p"):
            try:
                pending.append(Paragraph(child, doc).text.strip())
            except Exception:
                pass
        elif child.tag.endswith("}tbl"):
            t_idx += 1
            texts = [p for p in pending if p]
            before[t_idx] = texts
            if t_idx > 1:
                between[t_idx - 1] = texts      # 上一张表与这一张之间的文字
            pending = []
    if t_idx:
        before[t_idx + 1] = [p for p in pending if p]    # 最后一张表之后
    return before, between


def _titles_overlap(a_cols, b_cols, ratio=0.7):
    """两张表的列名重合度。"""
    ta = {c.get("title") for c in a_cols if c.get("title")}
    tb = {c.get("title") for c in b_cols if c.get("title")}
    if not ta or not tb:
        return False
    inter = len(ta & tb)
    return inter / max(len(ta), len(tb)) >= ratio


def _merge_continued_tables(sheets, between):
    """把「被分页拆开的同一张表」合回一张。

    判据（两条满足其一）：
        · 两张表的列名重合度 ≥ 70%，且它们之间夹着「续表」字样
        · 两张表的列名**完全**重合，且它们之间没有任何文字（紧邻）

    合并不是"看起来像就合" —— 真实模板里一张文档常常有好几张结构相同的表
    （山东空气质量模板：德州 / 潍坊 / 临沂 各一张，表头一模一样），
    每张表前面还有一句"本表记录…"的描述。**只要中间有描述文字，就是不同的表**，
    判成续表会让后面那张表整块消失（旧版真把潍坊那张吞了）。

    合并只改结构描述，不动文件：
        后一张表的待填段追加到前一张的 segments 里，
        segments[].table 记住「该往第几张表写」。
    """
    out = []
    for idx, sh in enumerate(sheets, start=1):
        prev = out[-1] if out else None
        # between 的下标是「前一张表的序号」：第 idx-1 张表与第 idx 张之间
        between_text = " ".join(between.get(idx - 1, [])) if idx > 1 else ""
        has_continue = any(w in between_text.lower() for w in CONTINUE_WORDS)

        if prev is not None:
            overlap = _titles_overlap(prev["columns"], sh["columns"])
            exact = ([c.get("title") for c in prev["columns"]]
                     == [c.get("title") for c in sh["columns"]])
            if overlap and (has_continue or (exact and not between_text)):
                for seg in sh.get("segments") or []:
                    s = dict(seg)
                    s["table"] = idx
                    prev["segments"].append(s)
                prev["tables"].append(idx)
                prev["fill_region"] = dict(prev["segments"][0])
                prev["continued"] = True
                prev["header_note"] = "跨页续表：第 %s 张表是同一张表" % (
                    "、".join(str(t) for t in prev["tables"]))
                continue
        out.append(sh)
    return out


def read_docx_template(docx_path):
    """读一张 docx（Word）模板。多个表格会被识别、必要时合并为跨页续表。"""
    docx_path = Path(docx_path)
    doc = Document(str(docx_path))
    before, between = _paragraph_maps(doc)

    sheets = []
    for ti, table in enumerate(doc.tables, start=1):
        dense = dense_docx(table)
        n_cols = len(dense[0]) if dense else 0

        info = describe_dense(dense, n_cols)
        # describe_dense 不知道"这是第几张表"，段里的 table 一律写 1 —— 多张表时
        # 写回会全往第 1 张里塞。这里按真实表号改回来。
        for seg in info.get("segments") or []:
            seg["table"] = ti
        info["fill_region"] = dict(info["segments"][0]) if info["segments"] else None
        info["tables"] = [ti]
        info["name"] = "表格%d" % ti
        info["max_row"] = len(dense)
        info["max_column"] = n_cols
        info["header_ok"] = bool(info["header_rows"])
        info["header_note"] = "文字不带冒号，符合表头特征"
        # 表前那句描述（"本表记录…德州市各监测站点…"）—— 用户什么都不填时，
        # 这就是"这张表要哪个城市/哪个时刻"的唯一依据，必须带着走。
        info["caption"] = "".join(before.get(ti) or [])[:400]
        sheets.append(info)

    sheets = _merge_continued_tables(sheets, between)

    return {"file": docx_path.name, "format": "docx",
            "sheet_count": len(sheets), "sheets": sheets}


# =================================================================
# 入口：读一个文件夹
# =================================================================

def read_folder(folder, only=None):
    """读一个文件夹里所有模板（xlsx 和 docx），返回说明书列表。

    only：只读这个文件名的模板（忽略同目录下的其它文件）。
          真实数据里源文档常和模板放在一起，全读一遍纯属浪费（见 main.py 的说明）。
    """
    folder = Path(folder)
    results = []

    files = sorted(list(folder.glob("*.xlsx")) + list(folder.glob("*.docx")))
    for p in files:
        if p.name.startswith("~$"):        # Office 打开时产生的临时文件
            continue
        if only and p.name != only:
            continue
        try:
            if p.suffix.lower() == ".docx":
                results.append(read_docx_template(p))
            else:
                results.append(read_template(p))
        except Exception as e:
            results.append({"file": p.name, "error": repr(e)})

    return results
