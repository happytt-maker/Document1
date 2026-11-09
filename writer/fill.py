# -*- coding: utf-8 -*-
"""
第 5 步 · 写回表格（fill.py）

把 matcher 抽出的值，按「第 i 条记录 → 待填区第 i 个格」填回模板，
生成一份"已填写"的新文件 —— 这就是比赛要交的东西。

用法（VSCode 打开 writer 文件夹后，直接在 F5 下拉里选）：
    python fill.py                          # 默认模板4，正常写入
    python fill.py 模板3-会议任务跟踪表.xlsx   # 换别的模板
    python fill.py --overwrite              # 待填区已有内容时，覆盖写入（旧值会记进报告）

三条铁律（写在代码里，也写在报告里）：
  1. 永不改原模板 —— 输出永远是一份新文件（writer/输出/xxx-已填写.xlsx）
  2. 先查后写 —— 发现待填区已有内容时默认停下、列出冲突，由人拍板
  3. 只填值、不碰格式 —— 字体、边框、列宽、合并单元格全部原样保留
     （openpyxl 只改 cell.value 时，样式天然不动；这是它最省心的地方）

═══════════════════════════════════════════════════════════════════
  2026-10-26 改造：待填区从「一段」变成「一串槽位」
═══════════════════════════════════════════════════════════════════

改造前这里只有一个 start_row / end_row，写入方式是
    「第 i 条记录 → 第 start_row + i 行」

这对跨页表直接失效：模板8 第 2 页的 9 行在另一个行号区间（15~23），
第 10 条记录会被写进「续表」那一行，把分隔行冲掉；第 19 条记录直接
被 `if r > end_row: break` 丢掉 —— 9 行容量白白浪费。

现在结构里给的是 segments（可能多段，Word 还可能跨多张表）：

    第 1 条 → 第 1 段第 1 行
    第 9 条 → 第 1 段第 9 行
    第 10 条 → 第 2 段第 1 行     ← 跨段连续编号，不浪费容量
    第 19 条 → 第 2 段第 10 行

`fill_slots()` 就是干这个的：把若干段摊平成一串「(第几张表, 第几行)」，
后面所有写入、查冲突、编号全按这串槽位走，xlsx 和 docx 共用一套逻辑。

另外新增：行标签列（label_cols）一律不写。
    模板7 第一列是「差旅费/办公费/招待费…」，那是表格结构，不是待填数据；
    不排除的话，资料里的公司名会被填到"差旅费"上，整列分类全毁。
"""
import json
import os
import re
import sys
import time
from pathlib import Path

from openpyxl import load_workbook

# 让中文在终端正常显示（失败了也不影响）
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

HERE = Path(__file__).resolve().parent            # writer/
BASE = HERE.parent                                # 研一项目/
# 结构文件与抽取结果目录：默认走项目内的固定位置；
# 用环境变量可指向别处（验收脚本用独立目录，避免污染真实产物；比赛换数据集也用得上）
STRUCTURE_JSON = Path(os.environ.get(
    "A23_STRUCTURE_JSON", str(BASE / "template_reader" / "模板表结构.json")))
MATCHER_DIR = Path(os.environ.get("A23_MATCHER_DIR", str(BASE / "matcher")))
# 模板目录：默认 测试数据/模板表格；可用环境变量 A23_TEMPLATE_DIR 覆盖
# （泛化测试 / 比赛换数据集时用；老流程不设变量时不受影响）
TEMPLATE_DIR = Path(os.environ.get("A23_TEMPLATE_DIR",
                                   str(BASE / "测试数据" / "模板表格")))
# 输出目录：默认 writer/输出；可用 A23_OUT_DIR 指到别处
# （真实数据试跑 / 泛化测试用独立目录 —— 五表连测的"还原"是按固定清单回滚的，
#  放在默认目录里的新产物会被它当垃圾清掉，2026-11-02 实测踩过一次）
OUT_DIR = Path(os.environ.get("A23_OUT_DIR", str(HERE / "输出")))

DEFAULT_TEMPLATE = "模板4-客诉处理台账.xlsx"


def say(text=""):
    print(text)


def load_sheet_info(template_name):
    """从模板表结构.json 里找到这个模板的结构信息（第③层的产物）。"""
    data = json.loads(STRUCTURE_JSON.read_text(encoding="utf-8"))
    for item in data:
        if item.get("file") == template_name:
            return item["sheets"][0]
    return None


def load_sheet_infos(template_name):
    """这个模板的**全部**块（docx 多张表 / xlsx 多 sheet）。

    旧版只取 sheets[0]，于是山东那份三表模板永远只填第一张，
    后两张"有表头没数据"却没人报错 —— 看起来像"填完了"。
    """
    data = json.loads(STRUCTURE_JSON.read_text(encoding="utf-8"))
    for item in data:
        if item.get("file") == template_name:
            return list(item.get("sheets") or [])
    return []


def load_extract(col_title):
    """读取某一列的抽取结果；没有就返回 None。"""
    f = MATCHER_DIR / ("抽取结果-%s.json" % col_title)
    if not f.exists():
        return None
    return json.loads(f.read_text(encoding="utf-8"))


def coerce(value):
    """把抽出的文本值转成合适的类型：纯数字 → 数字；其余原样。
    （这样"处理时长（天）"填进去的是数字 2，而不是文本"2"）"""
    s = str(value).strip()
    if re.fullmatch(r"-?\d+", s):
        return int(s)
    if re.fullmatch(r"-?\d+\.\d+", s):
        return float(s)
    return s


# =================================================================
# 槽位：第 i 条记录该写哪一格
# =================================================================

def fill_slots(info):
    """把待填区摊平成一串槽位，返回 [(第几张表, 行号), ...]。

    优先用新的 segments；老的结构文件只有 fill_region（一段），
    自动当成「只有一段」处理 —— 保证旧流程照跑不误。
    """
    segs = info.get("segments") or []
    if not segs:
        fr = info.get("fill_region") or {}
        if fr.get("start_row"):
            segs = [{"start_row": fr["start_row"], "end_row": fr["end_row"],
                     "table": 1}]

    slots = []
    for s in segs:
        t = s.get("table") or 1
        for r in range(s["start_row"], s["end_row"] + 1):
            slots.append((t, r))
    return slots


MAX_EXCEL_ROW = 1048576     # xlsx 单表最大行数，写超了 openpyxl 直接报错


def ensure_slots(ws, info, slots, need_n):
    """模板**一行空行都没留**（只有表头）时，按记录数往下补行。

    真实的百强城市模板就是这样：A1:E1 只有表头，下面全是空气。
    结构读取器如实报告「没找到空白区」，于是以前整条链路到写回这一步就断了。

    这里的取舍（2026-11-03 修过一次，教训写在这）：
      * **模板留了行，就尊重模板**。模板给了 8 行、数据 15 条 → 只填 8 条，
        多出来的在报告里说明。
        以前是"容量不够就往下补"，结果模板1 的「合计金额」「验收结论」
        两行表尾被当成待填行，既写错了位置，又把合计行挤掉了。
        模板作者留几行，本身就是一种约定。
      * **只在零槽位时才补**，且补之前确认目标区确实是空的 ——
        万一表头下面紧跟合计行，补行同样会把它吃掉。
      * 新行的边框/字体从上一行抄一份，保证看起来像同一张表；
        上一行是表头（加粗+底色）时不抄 —— 否则数据行会变成一片加粗黄底。

    返回 (新槽位表, 补了几行)。
    """
    need_n = max(0, int(need_n))
    if need_n <= len(slots):
        return slots, 0

    if slots:
        # 模板自己留了行 → 不扩容（见上面第二条取舍）
        return slots, 0

    hrows = info.get("header_rows") or (
        [info["header_row"]] if info.get("header_row") else [])
    if not hrows:
        return slots, 0
    table_no, start = 1, max(hrows) + 1

    add = need_n - len(slots)
    if start + add - 1 > MAX_EXCEL_ROW:          # 碰到表底就截断，不硬写
        add = MAX_EXCEL_ROW - start + 1
    if add <= 0:
        return slots, 0

    # 补行前先确认目标区是空的：表头下面如果紧跟「合计/验收结论」这类表尾行，
    # 盲目补行会把它覆盖掉（模板1 就是这种结构）。
    hset = set(hrows)
    for r in range(start, start + add):
        if r in hset:
            continue
        for c in range(1, ws.max_column + 1):
            v = ws.cell(row=r, column=c).value
            if v is not None and str(v).strip() != "":
                return slots, 0        # 有内容 → 不补，宁可少填也不覆盖


    # 抄样式：只抄「上一行不是表头」的那些格
    hrows = set(info.get("header_rows") or (
        [info["header_row"]] if info.get("header_row") else []))
    src_row = (slots[-1][1] if slots else None)
    if src_row and src_row not in hrows:
        from copy import copy as _copy
        for col in info.get("columns", []):
            c = col.get("col")
            if not c:
                continue
            src = ws.cell(row=src_row, column=c)
            for r in range(start, start + add):
                dst = ws.cell(row=r, column=c)
                if src.has_style:
                    dst.font = _copy(src.font)
                    dst.border = _copy(src.border)
                    dst.alignment = _copy(src.alignment)
                    dst.number_format = src.number_format
    elif not slots and hrows:
        # 紧接表头新建的行：给一条细边框，至少看得出是表格的一部分
        from openpyxl.styles import Border, Side
        line = Side(style="thin", color="FFBFBFBF")
        for row in ws.iter_rows(min_row=start, max_row=start + add - 1,
                                min_col=1, max_col=ws.max_column):
            for cell in row:
                cell.border = Border(left=line, right=line, top=line, bottom=line)

    return slots + [(table_no, r) for r in range(start, start + add)], add


def records_needed(per_col_items):
    """这些列的抽取结果里，一共需要几行。

    广播列（整份文档只有一个值、来自表外段落，比如「供应商」）不参与计数 ——
    它自己只有 1 条，但会铺满所有记录行，拿它计数会把表压成 1 行。
    """
    n = 0
    for items in per_col_items.values():
        if len(items) == 1 and items[0].get("from_paragraph"):
            continue
        n = max(n, len(items))
    return n


_LOC_ROW_PATTERNS = [
    re.compile(r"行\s*(\d+)"),            # docx 表格格：「表格1 · 行2·列1」
    re.compile(r"([A-Z]+)(\d+)\s*$"),     # xlsx 单元格：「工作表「库存明细」 J5」
]


def loc_row(loc_text):
    """从 loc_text 析出源行号；析不出返回 None。

    只认「源文档里的行号」，不认模板行号 —— 两者本来是两回事，
    以前没区分，所以只能按顺序硬填。
    """
    if not loc_text:
        return None
    s = str(loc_text)
    for pat in _LOC_ROW_PATTERNS:
        m = pat.search(s)
        if m:
            return int(m.group(1) if pat is _LOC_ROW_PATTERNS[0] else m.group(2))
    return None


def build_slot_map(per_col_items, n_slots):
    """把每条值按「源行号」映射到槽位下标，返回 {列名: [(槽位下标, item), ...]}。

    修的是这个错：备注列只有 5 个非空值，按顺序填就全挤到前 5 行；
    负责人列少了重复的一条，后面每行都往上串一格。

    规则：以所有列里出现的最小源行号为第 1 条记录，行号与它的差 = 记录序号。
    析不出行号的（少数连续文本块）退回原来的顺序位置，保证不比改前更差。

    **筛过行的源表（2026-11-03）**：行号不再连续（COVID 数据集筛日期区间后，
    命中行是 41 个国家各自一段，中间隔着大段没命中的行）。这时候"行号差"
    会算出几千，全部超出槽位数被丢掉 —— 表面现象是"抽取有 2542 条、写回只填了 3 行"。
    所以改成**按名次**：把命中的源行号排序去重，第 k 个对应第 k 条记录。
    行号本来连续时，名次 == 行号差，行为与改前完全一致（这是能直接换的前提）。

    文档级单值广播（2026-11-02）：
      模板1「供应商」—— 合同里「乙方（供货方）：泰隆机电设备有限公司」写在**表外段落**，
      一份合同只有一个乙方，但它适用于这份合同里的每一台设备（4 行）。
      以前这种列只填第 1 行，于是整列漏 3 格（模板1 91.7% 的失分全在这）。
      判据必须是"来源是表外段落"（from_paragraph），**不能**用"只有一条"：
      只有一条也可能是模型漏抽，那样广播等于把漏抽变成 4 份错填。
    """
    def _row(it):
        return loc_row(it.get("loc_text"))

    all_rows = []
    for items in per_col_items.values():
        for it in items:
            r = _row(it)
            if r:
                all_rows.append(r)
    base = min(all_rows) if all_rows else None
    # 源行号 → 名次（0 起）。连续行号时 rank[r] == r - base，等价于老规则。
    rank = {r: k for k, r in enumerate(sorted(set(all_rows)))} if all_rows else {}

    out = {}
    for col, items in per_col_items.items():
        seq, used = [], set()
        for i, it in enumerate(items):
            r = _row(it)
            if r and rank:
                idx = rank.get(r, r - base)
            else:
                idx = i
            if 0 <= idx < n_slots and idx not in used:
                used.add(idx)
                seq.append((idx, it))
        out[col] = seq

    # 记录条数 = 非广播列用到的最大槽位 + 1（广播列不参与，否则自己撑大自己）
    def _is_broadcast(seq):
        return len(seq) == 1 and bool(seq[0][1].get("from_paragraph"))

    n_rec = 0
    for seq in out.values():
        if seq and not _is_broadcast(seq):
            n_rec = max(n_rec, max(i for i, _ in seq) + 1)
    if n_rec > 1:
        for col, seq in list(out.items()):
            if _is_broadcast(seq):
                out[col] = [(i, seq[0][1]) for i in range(n_rec)]
    return out


def segments_brief(info):
    """给人看的待填区描述，例如「第 4~12 行（9 行）+ 第 15~23 行（9 行）」。"""
    segs = info.get("segments") or []
    if not segs:
        fr = info.get("fill_region") or {}
        if not fr.get("start_row"):
            return "没找到空白区"
        return "第 %d ~ %d 行（%d 行）" % (
            fr["start_row"], fr["end_row"], fr["row_count"])

    parts = []
    for s in segs:
        tag = ""
        if (s.get("table") or 1) != 1 and len({x.get("table") for x in segs}) > 1:
            tag = "表%d " % s.get("table")
        parts.append("%s第 %d ~ %d 行（%d 行）" % (
            tag, s["start_row"], s["end_row"], s["row_count"]))
    return "＋".join(parts)


def header_brief(info):
    """表头描述：层数 + 行号。"""
    rows = info.get("header_rows") or ([info["header_row"]] if info.get("header_row") else [])
    if not rows:
        return "（没找到表头）"
    if len(rows) == 1:
        return "第 %d 行（单层）" % rows[0]
    return "第 %d~%d 行（%d 层）" % (rows[0], rows[-1], len(rows))


def writable_columns(info):
    """排除行标签列后的可写列。"""
    labels = set(info.get("label_cols") or [])
    return [c for c in info.get("columns", []) if c.get("col") not in labels]


def is_seq_col(col):
    """是不是「序号」列（结构列，按记录数自动编号，不来自源文档）。"""
    if col.get("title") == "序号":
        return True
    return "序号" in (col.get("aliases") or [])


# =================================================================
# 合计行公式（2026-11-03）
# =================================================================
SUMMARY_WORDS = ("合计", "总计", "小计")
COUNT_WORDS = ("条数", "数量", "个数", "笔数", "家数", "台数", "人数",
               "投诉数", "次数", "件数")


def _pick_col(cols, keys):
    for c in cols:
        t = str(c.get("title") or "")
        if any(k in t for k in keys):
            return c
    return None


def _col_has_numbers(ws, col, first, last):
    n = 0
    for r in range(first, last + 1):
        if isinstance(ws.cell(row=r, column=col).value, (int, float)):
            n += 1
    return n >= 2


def _find_summary_rows(ws):
    """找「标签 + 同行其余全空」的合计行 → [(row, label_col, label)]。

    要求"其余全空"是有意的：模板8 那种已经写好 =SUM(...) 的合计行
    不该被再算一遍（会把用户原来的公式覆盖掉）。
    """
    rows = []
    for r in range(1, ws.max_row + 1):
        for c in (1, 2):
            v = ws.cell(row=r, column=c).value
            if not isinstance(v, str):
                continue
            s = v.strip()
            if not s or not any(k in s for k in SUMMARY_WORDS):
                continue
            rest = [ws.cell(row=r, column=x).value
                    for x in range(1, ws.max_column + 1) if x != c]
            if all(x is None or str(x).strip() == "" for x in rest):
                rows.append((r, c, s))
            break
    return rows


def _label_span(ws, row, col):
    """标签所在的合并区覆盖到第几列 —— 合计值必须避开它。

    模板1 的「合计金额」是 A13:C13 合并，值只能从 D 列往后放。
    """
    cell = ws.cell(row=row, column=col)
    for rng in ws.merged_cells.ranges:
        if rng.min_row <= row <= rng.max_row and cell.coordinate in rng:
            return rng.max_col
    return col


def write_summary_formulas(ws, info, slots):
    """给「合计/总计/小计」行补公式。返回 (written, notes)。

    必须先写完数据行再算，SUM 的区间才有意义。

    只算口径没有歧义的三种，其余留空并明说 —— 填一个"看着像合计、
    其实口径错了"的数字，比空着更糟（看起来像已经核对过的结果）：
      ① 金额 = SUMPRODUCT(数量区, 单价区)   ← 数量、单价两列都在才做
      ② 条数 = COUNTA(记录首列数据区)       ← 标签含 条数/数量/个数/投诉数…
      ③ 其它 = SUM(该列数据区)              ← 只补空格，不碰已有内容
    「率 / 占比 / 环比」要人定口径，不猜，只在报告里列出来。
    """
    from openpyxl.cell.cell import MergedCell
    from openpyxl.utils import get_column_letter as L

    written, notes = [], []
    if not slots:
        return written, notes
    cols = writable_columns(info)
    if not cols:
        return written, notes
    data_rows_all = sorted({r for (_t, r) in slots})

    for sum_row, label_col, label in _find_summary_rows(ws):
        data_rows = [r for r in data_rows_all if r < sum_row]
        if len(data_rows) < 2:
            notes.append("第 %d 行「%s」：上方没有数据区，跳过" % (sum_row, label))
            continue
        first, last = data_rows[0], data_rows[-1]

        # 标签可能占着合并区（模板1 的 A13:C13），可写列必须跳过去
        span_max = _label_span(ws, sum_row, label_col)
        right = [c for c in cols if c["col"] > span_max]
        if not right:
            notes.append("第 %d 行「%s」：标签右边没有可写列，跳过" % (sum_row, label))
            continue

        if any(k in label for k in ("率", "占比", "环比", "同比")):
            notes.append("第 %d 行「%s」：比率口径要人来定，没自动算（留空）"
                         % (sum_row, label))
            continue

        # ① 金额类：优先 数量 × 单价
        if "金额" in label or "总额" in label:
            qty = _pick_col(cols, ("数量",))
            price = _pick_col(cols, ("单价", "价格"))
            # 值放在金额口径最接近的列（单价/价格/金额），没有就取最左可写列。
            # 模板1 实测：放在「数量」列下面会读成"合计数量"，放在单价列才像话。
            target = _pick_col(right, ("单价", "价格", "金额")) or right[0]
            cell = ws.cell(row=sum_row, column=target["col"])
            if isinstance(cell, MergedCell) or cell.value not in (None, ""):
                notes.append("第 %d 行「%s」：%s 写不进去（合并格或已有内容），跳过"
                             % (sum_row, label, target.get("title")))
                continue
            if qty and price:
                f = "=SUMPRODUCT(%s%d:%s%d,%s%d:%s%d)" % (
                    L(qty["col"]), first, L(qty["col"]), last,
                    L(price["col"]), first, L(price["col"]), last)
                cell.value = f
                written.append((cell.coordinate, f,
                                "%s × %s，第 %d~%d 行" % (qty["title"], price["title"],
                                                          first, last)))
                continue
            notes.append("第 %d 行「%s」：找不到「数量」「单价」两列，没算"
                         % (sum_row, label))
            continue

        # ② 计数类
        if any(k in label for k in COUNT_WORDS):
            src = cols[0]
            target = right[0]
            cell = ws.cell(row=sum_row, column=target["col"])
            if isinstance(cell, MergedCell) or cell.value not in (None, ""):
                notes.append("第 %d 行「%s」：%s 写不进去（合并格或已有内容），跳过"
                             % (sum_row, label, target.get("title")))
                continue
            f = "=COUNTA(%s%d:%s%d)" % (L(src["col"]), first, L(src["col"]), last)
            cell.value = f
            written.append((cell.coordinate, f,
                            "按「%s」计数，第 %d~%d 行" % (src["title"], first, last)))
            continue

        # ③ 其它：这一行所有"数据区有数字"的可写列各补一个 SUM
        did = False
        for c in cols:
            cell = ws.cell(row=sum_row, column=c["col"])
            if isinstance(cell, MergedCell) or cell.value not in (None, ""):
                continue
            if not _col_has_numbers(ws, c["col"], first, last):
                continue
            f = "=SUM(%s%d:%s%d)" % (L(c["col"]), first, L(c["col"]), last)
            cell.value = f
            written.append((cell.coordinate, f,
                            "「%s」第 %d~%d 行求和" % (c["title"], first, last)))
            did = True
        if not did:
            notes.append("第 %d 行「%s」：上方没有可求和的数值列，跳过"
                         % (sum_row, label))
    return written, notes


# =================================================================
# 报告（xlsx / docx 共用，保证两份格式一致）
# =================================================================

def write_report(template_name, kind, overwrite, conflicts, written,
                 columns, missing_cols, slot_count, out_path):
    """写填表报告，并在控制台打一份。written 里每项是
    (记录序号, 坐标文字, 列名, 值, 出处)。"""
    lines = []
    title = "填表报告 · %s" % template_name
    if kind == "docx":
        title += "（Word 模板）"
    lines.append(title)
    lines.append("生成时间：%s" % time.strftime("%Y-%m-%d %H:%M"))
    lines.append("模式：%s" % ("覆盖写入（--overwrite）" if (overwrite and conflicts) else "正常写入"))
    lines.append("输出文件：%s" % out_path)
    lines.append("内容来源：matcher/抽取结果-*.json（每一格都能回溯到源文档）")
    lines.append("")

    lines.append("【一、冲突检查】")
    if conflicts:
        lines.append("待填区原有 %d 个格子有内容（%s）："
                     % (len(conflicts), "已被本次覆盖" if overwrite else "未处理"))
        for _i, coord, col_title, old in conflicts:
            lines.append("  %s [%s] 旧值 = %s" % (coord, col_title, old))
    else:
        lines.append("干净：待填区没有任何已有内容。")
    lines.append("")

    lines.append("【二、写入明细】共 %d 格" % len(written))
    by_slot = {}
    for slot_i, coord, col_title, val, src in written:
        by_slot.setdefault(slot_i, []).append((coord, col_title, val, src))
    for slot_i in sorted(by_slot):
        lines.append("第 %d 条记录（%s）" % (slot_i + 1, by_slot[slot_i][0][0]))
        for coord, col_title, val, src in by_slot[slot_i]:
            lines.append("  %s = %s   [%s]   ← 出处：%s" % (col_title, val, coord, src))
        lines.append("")

    lines.append("【三、统计】")
    lines.append("写入 %d 格（%d 列 × 最多 %d 条记录，容量 %d 条）"
                 % (len(written), len(columns) - len(missing_cols),
                    len(by_slot), slot_count))
    if conflicts and overwrite:
        lines.append("覆盖 %d 格（模板里原有的内容）" % len(conflicts))
    if len(by_slot) < slot_count:
        lines.append("还有 %d 个空位没用到（源文档只有 %d 条记录，正常）"
                     % (slot_count - len(by_slot), len(by_slot)))
    lines.append("")
    lines.append("下一步：跑 verify_fill.py 验收（值对不对 + 格式保没保）。")

    text = "\n".join(lines)
    tag = "-word" if kind == "docx" else ""
    report_path = OUT_DIR / ("填表报告-%s%s.txt" % (Path(template_name).stem, tag))
    report_path.write_text(text, encoding="utf-8")

    say("")
    say(text)
    say("")
    say("已生成：%s" % out_path)
    say("填表报告：%s" % report_path)
    return report_path


# =================================================================
# xlsx 模板
# =================================================================

def set_cell_text(cell, text):
    """往 docx 表格单元格写值：清掉旧文字（覆盖模式的残留），保留段落格式。"""
    p = cell.paragraphs[0]
    for run in list(p.runs):
        run.text = ""          # 保留 run 本身（=保留格式），只清文字
    p.add_run(text)


def _read_cell_xlsx(ws, r, c):
    v = ws.cell(row=r, column=c).value
    return "" if v is None else str(v).strip()


def main():
    args = sys.argv[1:]
    overwrite = "--overwrite" in args
    plain = [a for a in args if not a.startswith("--")]
    template_name = plain[0] if plain else DEFAULT_TEMPLATE

    tpl_path = TEMPLATE_DIR / template_name
    if not tpl_path.exists():
        say("找不到模板文件：%s" % tpl_path)
        return 1

    # Word（docx）模板 → 走 Word 版写回路径（规则相同，写入方式不同）
    if tpl_path.suffix.lower() == ".docx":
        return fill_docx(template_name, tpl_path, overwrite)

    info = load_sheet_info(template_name)
    if not info:
        say("模板表结构.json 里没有这个模板：%s" % template_name)
        say("先在 template_reader 里把读取器跑一遍，生成结构信息。")
        return 1

    sheet_name = info["name"]
    columns = writable_columns(info)
    slots = fill_slots(info)

    say("=" * 62)
    say("写回表格 · %s" % template_name)
    say("=" * 62)
    say("工作表：%s ｜ 表头在%s ｜ 可写 %d 列" % (sheet_name, header_brief(info), len(columns)))
    if info.get("label_cols"):
        say("行标签列：第 %s 列不填（那是表格自带的分类）"
            % "、".join(str(c) for c in info["label_cols"]))

    wb = load_workbook(tpl_path)
    if sheet_name not in wb.sheetnames:
        say("模板里没有工作表「%s」，实际有：%s" % (sheet_name, wb.sheetnames))
        return 1
    ws = wb[sheet_name]

    # 抽取结果先收齐 —— 要先知道"有几条记录"，才知道模板该补几行。
    # （以前是查到没空白区就直接放弃；真实模板只给表头，那样整条链就断在这）
    per_col_items = {}
    missing_cols = []
    for col in columns:
        res = load_extract(col["title"])
        if res is None:
            missing_cols.append(col["title"])
            continue
        per_col_items[col["title"]] = res.get("items", [])

    # 模板自带空行不够时，往下补行（只补不缩）
    # A23_NO_EXPAND=1 是消融开关：模拟"改动前"的行为，不补行。
    before = len(slots)
    need_rec = records_needed(per_col_items)
    if os.environ.get("A23_NO_EXPAND") == "1":
        added = 0
        say("待填区：%s ｜ 共 %d 个位置（对照实验：关闭动态补行）"
            % (segments_brief(info), before))
    else:
        slots, added = ensure_slots(ws, info, slots, need_rec)
    if before:
        say("待填区：%s ｜ 共 %d 个位置" % (segments_brief(info), before))
    else:
        say("待填区：模板没留空白行（只有表头）")
    if added:
        say("   → 按抽取到的 %d 条记录，往下补了 %d 行（第 %d ~ %d 行）"
            % (len(slots), added, slots[-added][1], slots[-1][1]))
    elif slots and need_rec > len(slots):
        say("   → 模板留了 %d 行，抽取到 %d 条：按模板容量填前 %d 条"
            "（不往下补，怕吃掉表尾的合计/结论行）"
            % (len(slots), need_rec, len(slots)))

    if not slots:
        say("")
        say("这张表既没有空白区，也没有可用的抽取结果 —— 先跑 matcher/run_all.py。")
        return 1
    say("")

    # ---------- ① 先查 ----------
    conflicts = []
    for si, (_t, r) in enumerate(slots):
        for col in columns:
            old = _read_cell_xlsx(ws, r, col["col"])
            if old:
                conflicts.append((si, "第%d行第%d列" % (r, col["col"]), col["title"], old))

    say("① 检查待填区……")
    if conflicts:
        say("   发现 %d 个格子已有内容：" % len(conflicts))
        for _si, coord, col_title, old in conflicts[:12]:
            say("     %s [%s] 旧值 = %s" % (coord, col_title, old))
        if len(conflicts) > 12:
            say("     …… 共 %d 个" % len(conflicts))
        if not overwrite:
            say("")
            say("已停下，没有写入任何东西。两条路：")
            say("  A. 打开模板，把上面这些内容清掉，然后重跑本脚本")
            say("  B. 确认可以覆盖的话，用 --overwrite 再跑一次（旧值会记进报告）")
            say("")
            say("（默认不覆盖，是怕把好数据冲掉；覆盖与否，由人拍板）")
            return 2
        say("   （--overwrite 模式：将覆盖以上内容，旧值会记进报告）")
    else:
        say("   干净：待填区没有任何已有内容。")
    say("")

    # ---------- ② 再写 ----------
    say("② 写入（第 i 条记录 → 第 i 个位置，跨段连续）……")
    written = []
    missing_cols = []
    per_col_counts = {}

    # 按「源行号」统一对齐（见 build_slot_map）。
    # 以前是逐列 enumerate 填 slots[i]，稀疏列和缺值列会整段错位。
    slot_map = build_slot_map(per_col_items, len(slots))

    for col in columns:
        if col["title"] not in per_col_items:
            continue
        items = per_col_items[col["title"]]
        per_col_counts[col["title"]] = len(items)
        for i, it in slot_map[col["title"]]:
            _t, r = slots[i]
            cell = ws.cell(row=r, column=col["col"])
            cell.value = coerce(it.get("value"))
            src = ("%s %s" % (it.get("file") or "?", it.get("loc_text") or "")).strip()
            written.append((i, cell.coordinate, col["title"], it.get("value"), src))

    # 「序号」列自动编号：按槽位顺序 1..N，跨段连续
    seq_cols = [c for c in columns if is_seq_col(c) and c["title"] not in per_col_counts]
    if seq_cols and per_col_counts:
        n_rec = min(max(per_col_counts.values()), len(slots))
        for col in seq_cols:
            for i in range(n_rec):
                _t, r = slots[i]
                cell = ws.cell(row=r, column=col["col"])
                cell.value = i + 1
                written.append((i, cell.coordinate, col["title"], i + 1, "（序号·自动编号）"))
        say("   「序号」列自动编号 1~%d（结构列，不来自源文档）" % n_rec)

    missing_cols = [m for m in missing_cols if m != "序号"]
    if missing_cols:
        say("   注意：这些列没有抽取结果 → %s（先跑 matcher/run_all.py）"
            % "、".join(missing_cols))

    # ---------- ②.5 合计行公式（2026-11-03） ----------
    # 放在数据写完之后：SUM 的区间要正好盖住刚写进去的那些行。
    sum_written, sum_notes = write_summary_formulas(ws, info, slots)
    if sum_written:
        say("")
        say("   合计行补了 %d 个公式：" % len(sum_written))
        for coord, f, why in sum_written:
            say("     %s  %s   ← %s" % (coord, f, why))
        say("     （公式由 Excel / WPS 打开时计算，文件里存的是公式本身）")
    for n in sum_notes:
        say("   合计行未处理：" + n)

    # ---------- ③ 保存（保存到新文件，原模板不动） ----------
    OUT_DIR.mkdir(exist_ok=True)
    out_path = OUT_DIR / (Path(template_name).stem + "-已填写.xlsx")
    try:
        wb.save(out_path)
    except PermissionError:
        say("")
        say("保存失败：输出文件正被占用，暂时写不进去 ——")
        say("  %s" % out_path)
        say("多半是它正在 Excel / WPS 里开着。关掉那个窗口，然后重跑本脚本。")
        return 1

    # ---------- ④ 报告 ----------
    write_report(template_name, "xlsx", overwrite, conflicts, written,
                 columns, missing_cols, len(slots), out_path)
    return 0


# =================================================================
# Word（docx）模板的写回
# =================================================================
# 赛题评价方法：「再分别上传 5 个模版表格文件（word 或 excel）」
# Word 模板 = docx 里的一张或多张表格。跨页表在 Word 里可能是两张 table，
# 结构层已经把它们的待填段合并到同一个 segments 里（segments[].table 记住
# 该往第几张表写），所以这里只需要按槽位取对应的 table。
#
# 写值方式选"段落里加 run"而不是"cell.text = v"：
#     后者会重建段落、丢掉模板给待填格设的对齐等格式；
#     add_run 保留段落格式，新文字随段落格式走（跟手打进去一样）。

def _docx_coord(table_idx, r, c, multi_table):
    if multi_table:
        return "表%d 第%d行第%d列" % (table_idx, r, c)
    return "第%d行第%d列" % (r, c)


def fill_docx(template_name, tpl_path, overwrite):
    """Word 版写回：与 xlsx 版同规则、同报告格式。**支持一张文档里多张表**。"""
    from docx import Document

    infos = load_sheet_infos(template_name)
    if not infos:
        say("模板表结构.json 里没有这个模板：%s" % template_name)
        say("先在 template_reader 里把读取器跑一遍（python main.py \"模板目录\"），生成结构信息。")
        return 1

    # 每块 = 模板里的一张表（或一张被拆页的表的若干段）
    blocks = []
    for bi, info in enumerate(infos, start=1):
        sl = fill_slots(info)
        if sl:
            blocks.append((bi, info, sl))
    if not blocks:
        say("这张模板没找到可填的空白区 —— 可能已经被填满，或者表头识别错了。")
        return 1

    columns = writable_columns(blocks[0][1])
    total_slots = sum(len(sl) for _b, _i, sl in blocks)

    say("=" * 62)
    say("写回表格（Word 版） · %s" % template_name)
    say("=" * 62)
    say("共 %d 个块 ｜ 可写 %d 列 ｜ 待填位置合计 %d 个"
        % (len(blocks), len(columns), total_slots))
    for bi, info, sl in blocks:
        say("   块%d（%s）：%s ｜ %d 个位置%s"
            % (bi, info.get("name") or "?", segments_brief(info), len(sl),
               ("  ← " + (info.get("caption") or "")[:36] + "…")
               if info.get("caption") else ""))
    if len(blocks) > 1:
        say("   （一张文档多张结构相同的表：各填各的，不会互相串）")
    if blocks[0][1].get("label_cols"):
        say("行标签列：第 %s 列不填（那是表格自带的分类）"
            % "、".join(str(c) for c in blocks[0][1]["label_cols"]))
    say("")

    doc = Document(str(tpl_path))
    if not doc.tables:
        say("这个 docx 里没有表格，无法填写。")
        return 1

    tables_used = sorted({t for _b, _i, sl in blocks for t, _r in sl})
    for t in tables_used:
        if t - 1 >= len(doc.tables):
            say("模板里没有第 %d 张表格（实际只有 %d 张）" % (t, len(doc.tables)))
            return 1
    tables = {t: doc.tables[t - 1] for t in tables_used}
    # python-docx 的 table.cell(r, c) 每次调用都要把整张表的 XML 走一遍，
    # 一格一次调用 = 上千次遍历（三张表实测占 2~3 秒）。这里一次取出行网格复用。
    grid = {t: [row.cells for row in tables[t].rows] for t in tables_used}

    def _cell(t, r, c):
        return grid[t][r - 1][c - 1]

    multi_table = len({t for _b, _i, sl in blocks for t, _r in sl}) > 1 or len(blocks) > 1

    # ---------- ① 先查 ----------
    conflicts = []
    for _bi, _info, slots in blocks:
        for si, (t, r) in enumerate(slots):
            for col in columns:
                v = (_cell(t, r, col["col"]).text or "").strip()
                if v:
                    conflicts.append((si, _docx_coord(t, r, col["col"], multi_table),
                                      col["title"], v))

    say("① 检查待填区……")
    if conflicts:
        say("   发现 %d 个格子已有内容：" % len(conflicts))
        for _si, coord, col_title, old in conflicts[:12]:
            say("     %s [%s] 旧值 = %s" % (coord, col_title, old))
        if len(conflicts) > 12:
            say("     …… 共 %d 个" % len(conflicts))
        if not overwrite:
            say("")
            say("已停下，没有写入任何东西。两条路：")
            say("  A. 打开模板，把上面这些内容清掉，然后重跑本脚本")
            say("  B. 确认可以覆盖的话，用 --overwrite 再跑一次（旧值会记进报告）")
            say("")
            say("（默认不覆盖，是怕把好数据冲掉；覆盖与否，由人拍板）")
            return 2
        say("   （--overwrite 模式：将覆盖以上内容，旧值会记进报告）")
    else:
        say("   干净：待填区没有任何已有内容。")
    say("")

    # ---------- ② 再写 ----------
    say("② 写入（每块内第 i 条记录 → 第 i 个位置）……")
    written = []
    missing_cols = []
    per_col_counts = {}

    per_col_items = {}
    for col in columns:
        res = load_extract(col["title"])
        if res is None:
            missing_cols.append(col["title"])
            continue
        per_col_items[col["title"]] = res.get("items", [])

    rec = 0                       # 全局记录号（报告里"第 N 条"用）
    for bi, info, slots in blocks:
        # 这一块只认带自己 block 号的值；老结果没有 block 字段 → 全归第 1 块
        sub = {c: [it for it in items
                   if int(it.get("block") or 1) == bi]
               for c, items in per_col_items.items()}
        sub = {c: v for c, v in sub.items() if v}
        if not sub:
            say("   块%d：没有对应的抽取结果 → 留空（不从别的块凑）" % bi)
            continue
        slot_map = build_slot_map(sub, len(slots))
        for col in columns:
            if col["title"] not in sub:
                continue
            per_col_counts[col["title"]] = per_col_counts.get(col["title"], 0) + len(sub[col["title"]])
            for i, it in slot_map[col["title"]]:
                t, r = slots[i]
                cell = _cell(t, r, col["col"])
                set_cell_text(cell, str(coerce(it.get("value"))))
                src = ("%s %s" % (it.get("file") or "?", it.get("loc_text") or "")).strip()
                written.append((rec + i, _docx_coord(t, r, col["col"], multi_table),
                                col["title"], it.get("value"), src))
        used = max((max(i for i, _ in seq) + 1) for seq in slot_map.values() if seq) \
            if any(slot_map.values()) else 0
        say("   块%d：写入 %d 条记录 / 容量 %d 条%s"
            % (bi, used, len(slots), "" if used < len(slots) else "（已填满）"))
        rec += used

    # 「序号」列自动编号：与 xlsx 版同规则（结构列，不来自源文档）
    seq_cols = [c for c in columns if is_seq_col(c) and c["title"] not in per_col_counts]
    if seq_cols and per_col_counts:
        base = 0
        for bi, info, slots in blocks:
            sub = {c: [it for it in items if int(it.get("block") or 1) == bi]
                   for c, items in per_col_items.items()}
            used = min(len(slots), max((len(v) for v in sub.values()), default=0))
            for col in seq_cols:
                for i in range(used):
                    t, r = slots[i]
                    set_cell_text(_cell(t, r, col["col"]), str(base + i + 1))
                    written.append((base + i, _docx_coord(t, r, col["col"], multi_table),
                                    col["title"], base + i + 1, "（序号·自动编号）"))
            base += used
        say("   「序号」列自动编号 1~%d（结构列，不来自源文档）" % base)

    missing_cols = [m for m in missing_cols if m != "序号"]
    if missing_cols:
        say("   注意：这些列没有抽取结果 → %s（先跑 matcher/run_all.py）"
            % "、".join(missing_cols))

    # ---------- ③ 保存 ----------
    OUT_DIR.mkdir(exist_ok=True)
    out_path = OUT_DIR / (Path(template_name).stem + "-已填写.docx")
    try:
        doc.save(str(out_path))
    except PermissionError:
        say("")
        say("保存失败：输出文件正被占用，暂时写不进去 ——")
        say("  %s" % out_path)
        say("多半是它正在 Word / WPS 里开着。关掉那个窗口，然后重跑本脚本。")
        return 1

    # ---------- ④ 报告 ----------
    write_report(template_name, "docx", overwrite, conflicts, written,
                 columns, missing_cols, total_slots, out_path)
    return 0


if __name__ == "__main__":
    sys.exit(main())
