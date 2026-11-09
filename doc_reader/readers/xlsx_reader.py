# -*- coding: utf-8 -*-
"""
readers/xlsx_reader.py
----------------------
Excel 表格读取器。

把 xlsx 读成带位置信息的块，输出格式和其他读取器保持一致。

位置记为「工作表名 + 单元格坐标」，例如：
    {"kind": "cell", "sheet": "员工信息", "cell": "B3", "row": 3, "col": 2}

═══════════════════════════════════════════════════════════════════
   四个必须知道的坑
═══════════════════════════════════════════════════════════════════

【坑 1】公式单元格读出来是 None（最重要）

    openpyxl 用 data_only=True 读的是「上次被 Excel 计算并保存时的缓存值」。
    程序生成的、从未被 Excel 打开过的文件，缓存槽位是空的 → 读出 None。

    实测（员工信息表 I3 = "=G3+H3"）：
        data_only=False  →  '=G3+H3'    （公式字符串）
        data_only=True   →  None        （缓存是空的）

    所以必须读两遍：
        第一遍（data_only=False）拿到公式字符串 → 标记出来
        第二遍（data_only=True）拿缓存值 → 如果公式存在但缓存为 None，标记为"未计算"

    2026-11-02 补充：拿到公式后会**自己算一遍**（_eval_formula，只做四则运算 +
    同表单元格引用）。算出值的块 value_type = 'formula_evaluated'，text 就是数值 ——
    下游可以直接用，不用知道它原本是公式。只有算不出来的（SUM/VLOOKUP/跨表引用）
    才保持 'formula_uncached'。

【坑 2】必须遍历所有工作表

    wb.worksheets 是一个列表。只读 wb.active 会漏掉其他 sheet。
    实测：「项目进度跟踪表.xlsx」有两个工作表（项目总览 / 工时登记）。

【坑 3】空单元格和「值为 0」是两回事

    库存盘点表里，M-1006 的「本期入库」是真的没有数据（入库单未到），
    不是 0。判断时要区分 is None 和 == 0。

【坑 4】合并单元格

    Excel 的合并区域，只有左上角那个格子有值，其他格子返回 None。
    可以用 sheet.merged_cells.ranges 拿到所有合并区域。

═══════════════════════════════════════════════════════════════════
"""

import re

import openpyxl
from openpyxl.utils import get_column_letter, column_index_from_string

from .base import make_block


def _cell_ref(row, col):
    """把行列号转成 Excel 坐标，如 (3, 2) -> 'B3'。"""
    return "{0}{1}".format(get_column_letter(col), row)


def _looks_like_formula(value):
    """判断一个值是不是公式（以 = 开头的字符串）。"""
    return isinstance(value, str) and value.startswith("=")


# ===== 公式求值（2026-11-02）=====
# 为什么必须做：程序生成/未在 Excel 里保存过的表，"缓存值"是空的，
# 于是「库存数量 = E+F-G」这类列读出来只有公式字符串 "=E3+F3-G3"。
# 下游抽取拿到它只能按"公式不可用"跳过 —— 模型只好去抄「期初库存」，
# 整列 10 格全错（模板5 实测）。而现实台账里"合计/余额/结存"几乎都是公式列，
# 不解决这个，换任何真实数据都会在同类列上翻车。
#
# 只做最保守的求值：四则运算 + 括号 + 同表单元格引用（可递归，最多 10 层）。
# 认不出的（SUM/VLOOKUP/跨表引用…）一律返回 None，仍然按"未计算"处理 ——
# 宁可不算，也不能算错。
_REF_RE = re.compile(r"\$?([A-Z]{1,3})\$?(\d{1,7})")
_NUM_EXPR = re.compile(r"^[0-9eE\.\+\-\*/\(\)\s]+$")
_MAX_DEPTH = 10

# SUM 的两种形态：函数本身、以及它里面的连续区间
_SUM_RE = re.compile(r"\bSUM\s*\(\s*([^()]{1,80}?)\s*\)", re.I)
_RANGE_RE = re.compile(
    r"^\$?([A-Z]{1,3})\$?(\d{1,7})\s*:\s*\$?([A-Z]{1,3})\$?(\d{1,7})$", re.I)


def _range_values(spec, v_sheet, f_sheet, cache, depth):
    """把 "F3:F12" 展开成每格的数值；认不出返回 None。

    只认「单个连续矩形区间」。SUM(A1:B9, C3) 这种多参数、非连续、
    或者带条件的（SUMIF）一律不猜 —— 返回 None 让它保持"未计算"。
    这条边界很重要：算错一个合计，比留空更糟（空着人一看就知道要补）。
    """
    m = _RANGE_RE.match(spec.strip())
    if not m:
        return None
    c1, r1 = m.group(1).upper(), int(m.group(2))
    c2, r2 = m.group(3).upper(), int(m.group(4))
    i1, i2 = column_index_from_string(c1), column_index_from_string(c2)
    if i1 > i2:
        i1, i2 = i2, i1
    if r1 > r2:
        r1, r2 = r2, r1
    if (i2 - i1 + 1) * (r2 - r1 + 1) > 20000:      # 防呆：别把整张表拖进来算
        return None
    out = []
    for r in range(r1, r2 + 1):
        for ci in range(i1, i2 + 1):
            out.append(_cell_number("%s%d" % (get_column_letter(ci), r),
                                    v_sheet, f_sheet, cache, depth))
    return out


def _cell_number(coord, v_sheet, f_sheet, cache, depth=0):
    """取某格参与计算的数值；取不到返回 None。"""
    if coord in cache:
        return cache[coord]
    if depth > _MAX_DEPTH:
        return None
    try:
        v = v_sheet[coord].value
    except Exception:
        return None
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        cache[coord] = v
        return v
    f = None
    try:
        f = f_sheet[coord].value if f_sheet is not None else None
    except Exception:
        f = None
    if _looks_like_formula(f):
        got = _eval_formula(f, v_sheet, f_sheet, cache, depth + 1)
        if got is not None:
            cache[coord] = got
        return got
    if v is None and f is None:
        # 空单元格在 Excel 的算术里就是 0，不是"算不出来"。
        # 库存盘点表 M-1006「入库单未到」→ 本期入库为空 → 期末 = 期初 + 0 - 出库。
        # 若这里返回 None，整条公式就废了（H8 坚持"未计算"）。
        cache[coord] = 0
        return 0
    return None


def _eval_formula(formula, v_sheet, f_sheet, cache=None, depth=0):
    """把 "=E3+F3-G3"、"=SUM(F3:F12)" 算成数值；算不了返回 None。"""
    if cache is None:
        cache = {}
    if not _looks_like_formula(formula) or depth > _MAX_DEPTH:
        return None
    expr = formula.lstrip("=").strip()

    # SUM(区间)：台账里的"合计"几乎都是这个写法。
    # 2026-11-03 补 —— 以前遇到函数一律返回 None，源文档的合计格读出来
    # 只剩 "=SUM(F3:F12)" 这串字符（销售台账 F14、项目进度跟踪 D8），
    # 抽取端只能当它不存在；而模板里的合计行又需要这个数。
    if _SUM_RE.search(expr):
        def _sub_sum(m):
            vals = _range_values(m.group(1), v_sheet, f_sheet, cache, depth)
            if vals is None:
                return m.group(0)      # 认不出就原样留着，下面会判"算不了"
            nums = [v for v in vals if isinstance(v, (int, float))]
            return repr(float(sum(nums)))
        expr = _SUM_RE.sub(_sub_sum, expr)

    missing = []

    def _sub(m):
        coord = m.group(1) + m.group(2)
        num = _cell_number(coord, v_sheet, f_sheet, cache, depth)
        if num is None:
            missing.append(coord)
            return "0"
        return repr(float(num))

    expr = _REF_RE.sub(_sub, expr)
    if missing or not _NUM_EXPR.match(expr):
        return None        # 引用不到 / 有 SUM 之类不认识的函数
    try:
        val = eval(expr, {"__builtins__": {}}, {})   # 字符集已限死，无可执行内容
    except Exception:
        return None
    if isinstance(val, float) and val.is_integer():
        val = int(val)
    return val if isinstance(val, (int, float)) else None


def _normalize(value):
    """
    把单元格的值整理成干净的形式。

    返回 (文本, 值类型)
      值类型可能是：'text' / 'number' / 'datetime' / 'bool' / 'formula'
    """
    if value is None:
        return "", "empty"

    if isinstance(value, bool):
        return ("TRUE" if value else "FALSE"), "bool"

    # 日期时间（openpyxl 会读成 datetime 对象）
    if hasattr(value, "isoformat") and not isinstance(value, str):
        try:
            return value.isoformat(), "datetime"
        except Exception:
            pass

    if isinstance(value, (int, float)):
        # 整数就不要显示成 100.0
        if isinstance(value, float) and value.is_integer():
            return str(int(value)), "number"
        return str(value), "number"

    return str(value).strip(), "text"


def _merged_ranges(sheet):
    """
    把 sheet 里所有合并区域整理成一张表，方便查某个坐标属于哪个合并区。

    返回 { "A1": "A1:B2", "B1": "A1:B2", ... }
    即：区域内每个坐标 -> 整个区域的字符串
    """
    mapping = {}
    try:
        for rng in sheet.merged_cells.ranges:
            key = str(rng)
            for row in sheet[rng.coord]:
                for cell in row:
                    mapping[cell.coordinate] = key
    except Exception:
        pass
    return mapping


def read_sheet(sheet, file_name, vsheet, _max_cached=None):
    """读一个工作表，返回块列表。

    sheet   ：`data_only=False` 打开的那本（值和公式字符串都在，**主本**）
    vsheet  ：`data_only=True` 的同一张工作表；文件里一个公式都没有时传 None

    为什么改成"主本只开一次"（2026-11-07）：
      旧版不管有没有公式都开两个工作簿。实测山东那份 20000 行 × 18 列的表：
        load(data_only=False) 4.7 秒 + load(data_only=True) 5.6 秒 = 10.3 秒，
        而真正遍历取值只花 0.3 秒 —— 一半时间花在"为了拿公式再加开一遍"上，
      可那份表里一个公式都没有。现在先扫一遍 data_type，确实有公式才开第二本。
      有公式的文件行为与旧版完全一致（同一条取值/求值分支）。
    """
    blocks = []
    merged = _merged_ranges(sheet)

    for row in sheet.iter_rows():
        for cell in row:
            coord = cell.coordinate
            value = cell.value
            is_formula = (cell.data_type == "f")

            # 合并区域里非左上角的格子，值是 None，跳过（避免大量空洞）
            if value is None and not is_formula:
                if coord in merged:
                    continue
                continue

            formula = None
            text = ""

            if is_formula:
                formula = value if isinstance(value, str) else str(value)
                cached = None
                if vsheet is not None:
                    try:
                        cached = vsheet[coord].value
                    except Exception:
                        cached = None

                if cached is not None:
                    text, value_type = _normalize(cached)
                    value_type = "formula_cached"
                    if text == "":
                        continue
                else:
                    # 有公式但读不到缓存值：先试着自己算
                    got = _eval_formula(formula, vsheet if vsheet is not None else sheet,
                                        sheet, {})
                    if got is not None:
                        text, _nt = _normalize(got)
                        value_type = "formula_evaluated"
                    else:
                        text = formula
                        value_type = "formula_uncached"
            else:
                if value is None:
                    continue
                text, value_type = _normalize(value)
                if text == "":
                    continue

            block = make_block(
                file_name=file_name,
                file_type="xlsx",
                text=text,
                locator={
                    "kind": "cell",
                    "sheet": sheet.title,
                    "cell": coord,
                    "row": cell.row,
                    "col": cell.column,
                },
                value_type=value_type,
                formula=formula,
                is_formula=formula is not None,
            )

            # 附上合并区域信息
            if coord in merged:
                block["merged_range"] = merged[coord]

            blocks.append(block)

    return blocks


def read_xlsx(path):
    """
    读一个 xlsx 文件，返回和 read_txt / read_md / read_docx 同构的字典。

    主本用 data_only=False 打开（值 + 公式一次拿全）。
    只有当文件里**确实存在公式**时，才再开一本 data_only=True 去取缓存值 ——
    纯数据表因此省掉一次整本加载（大表实测省 5 秒级，见 read_sheet 的说明）。
    """
    wb_formulas = openpyxl.load_workbook(str(path), data_only=False)

    has_formula = False
    for _ws in wb_formulas.worksheets:
        for _row in _ws.iter_rows():
            for _c in _row:
                if _c.data_type == "f":
                    has_formula = True
                    break
            if has_formula:
                break
        if has_formula:
            break

    wb_values = (openpyxl.load_workbook(str(path), data_only=True)
                 if has_formula else None)

    all_blocks = []
    sheet_info = []

    for sheet in wb_formulas.worksheets:
        vsheet = None
        if wb_values is not None:
            try:
                vsheet = wb_values[sheet.title]
            except Exception:
                vsheet = None
        blocks = read_sheet(sheet, path.name, vsheet)
        all_blocks.extend(blocks)

        formulas = [b for b in blocks if b.get("formula")]
        sheet_info.append(
            {
                "name": sheet.title,
                "max_row": sheet.max_row,
                "max_column": sheet.max_column,
                "block_count": len(blocks),
                "formula_count": len(formulas),
            }
        )

    return {
        "file": path.name,
        "type": "xlsx",
        "sheet_count": len(wb_formulas.worksheets),
        "sheets": sheet_info,
        "block_count": len(all_blocks),
        "blocks": all_blocks,
    }
