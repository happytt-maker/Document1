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

import openpyxl
from openpyxl.utils import get_column_letter, column_index_from_string

from .base import make_block


def _cell_ref(row, col):
    """把行列号转成 Excel 坐标，如 (3, 2) -> 'B3'。"""
    return "{0}{1}".format(get_column_letter(col), row)


def _looks_like_formula(value):
    """判断一个值是不是公式（以 = 开头的字符串）。"""
    return isinstance(value, str) and value.startswith("=")


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


def read_sheet(sheet, file_name, wb_formulas):
    """
    读一个工作表，返回块列表。

    wb_formulas 是「用 data_only=False 打开的同一个工作簿」，
    用来拿公式字符串做对比。
    """
    blocks = []
    merged = _merged_ranges(sheet)

    # 公式版的对应 sheet（可能因为 sheet 名不一致而拿不到，容错处理）
    fsheet = None
    try:
        fsheet = wb_formulas[sheet.title]
    except Exception:
        pass

    # 遍历所有有数据的行列
    for row in sheet.iter_rows():
        for cell in row:
            coord = cell.coordinate
            value = cell.value

            # 合并区域里非左上角的格子，值是 None，跳过（避免大量空洞）
            if value is None and coord in merged:
                continue

            formula = None
            text = ""

            if value is None:
                # 这一格没值 —— 看它是不是「公式但缓存为空」
                fval = None
                if fsheet is not None:
                    fval = fsheet[coord].value

                if _looks_like_formula(fval):
                    # 有公式但读不到值：这是坑 1 的现场
                    formula = fval
                    text = fval                 # 把公式本身当内容存着
                    value_type = "formula_uncached"
                else:
                    # 真的空单元格，跳过
                    continue
            else:
                text, value_type = _normalize(value)
                if text == "":
                    continue

                # 有值但也是公式（缓存已算出）—— 标记出来
                fval = None
                if fsheet is not None:
                    fval = fsheet[coord].value
                if _looks_like_formula(fval):
                    formula = fval
                    value_type = "formula_cached"

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

    会打开两次：
      第一次 data_only=False  → 拿公式字符串
      第二次 data_only=True   → 拿缓存值（可能为 None）
    """
    # 第一遍：拿到公式
    wb_formulas = openpyxl.load_workbook(str(path), data_only=False)

    # 第二遍：拿到值（公式的位置是缓存值，可能为 None）
    wb_values = openpyxl.load_workbook(str(path), data_only=True)

    all_blocks = []
    sheet_info = []

    for sheet in wb_values.worksheets:
        blocks = read_sheet(sheet, path.name, wb_formulas)
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
        "sheet_count": len(wb_values.worksheets),
        "sheets": sheet_info,
        "block_count": len(all_blocks),
        "blocks": all_blocks,
    }
