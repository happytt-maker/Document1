# -*- coding: utf-8 -*-
"""
format_ops.py —— 模块（1）· 第 3 步：格式与排版执行器

链路：一句话 → 解析（parse_intent）→ 定位 → 执行格式设置 → 输出新文件。

支持的十一种操作：
  排版类（段落级）：
    set_align    对齐（居中 / 右对齐 / 左对齐 / 两端对齐）
    set_spacing  行距（1.5 倍 / 2 倍 …）
    set_indent   首行缩进（2 字符 / 1 字符 …）
  格式类（文字级）：
    set_bold      加粗（支持语义目标——"所有金额"；params.off=true 取消）
    set_italic    斜体（params.off=true 取消）
    set_underline 下划线（params.off=true 取消）
    set_strike    删除线（params.off=true 取消）
    set_highlight 高亮 / 突出显示（params.color 选色；params.off=true 取消）
    set_color     颜色（"标红"）
    set_font      字体（"改成黑体"）
    set_size      字号（"调成 14 号"）

三种文档：
  docx    —— python-docx。文字级操作会"切分 run"：把一段文字拆成
             前/中/后三个 run，只给"中"加格式——这样"只加粗金额"才不会
             把整段话都加粗。
  xlsx    —— openpyxl。表头行识别 = 非空格子数最多的行（与模板读取器同一规则）。
  md/txt  —— 纯文本。支持 加粗(** **) / 斜体(* *) / 删除线(~~ ~~) /
             高亮(<mark>) / 标红(<span style="color:red">)；
             对齐、行距这类"版式"在纯文本里没有对应写法，会明确拒绝。

定位线索（locate.hint）怎么变成"具体对象"：
  "标题"  → 样式为 Title 的段落（兜底：第一个非空段落）
  "正文"  → 第一个标题之后的非标题段落
  "表头"  → xlsx 表头行
  具体词  → 含该词的段落 / 该词在段内的每一次出现
  空      → 先看指令原文里的语义目标（"金额"）；
            缩进、行距类默认作用于正文
"""
import copy
import re
from pathlib import Path

import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_COLOR_INDEX
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor
from docx.text.run import Run

from openpyxl import load_workbook
from openpyxl.styles import Alignment, PatternFill
from openpyxl.styles.colors import Color

# ==================== 通用小工具 ====================

_XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"

_ALIGN_MAP = {
    "center": WD_ALIGN_PARAGRAPH.CENTER, "居中": WD_ALIGN_PARAGRAPH.CENTER,
    "right": WD_ALIGN_PARAGRAPH.RIGHT, "右对齐": WD_ALIGN_PARAGRAPH.RIGHT,
    "left": WD_ALIGN_PARAGRAPH.LEFT, "左对齐": WD_ALIGN_PARAGRAPH.LEFT,
    "justify": WD_ALIGN_PARAGRAPH.JUSTIFY, "两端对齐": WD_ALIGN_PARAGRAPH.JUSTIFY,
}

_COLOR_MAP = {
    "红色": "FF0000", "红": "FF0000", "red": "FF0000",
    "蓝色": "0070C0", "蓝": "0070C0", "blue": "0070C0",
    "绿色": "00B050", "绿": "00B050", "green": "00B050",
    "橙色": "ED7D31", "黄色": "FFC000", "紫色": "7030A0", "灰色": "808080",
}

# 高亮色（颜色词 → docx 枚举 / xlsx 填充色，两套取值）
_HIGHLIGHT_WD = {
    "黄色": WD_COLOR_INDEX.YELLOW, "黄": WD_COLOR_INDEX.YELLOW, "yellow": WD_COLOR_INDEX.YELLOW,
    "绿色": WD_COLOR_INDEX.BRIGHT_GREEN, "绿": WD_COLOR_INDEX.BRIGHT_GREEN,
    "青色": WD_COLOR_INDEX.TURQUOISE, "蓝绿": WD_COLOR_INDEX.TURQUOISE,
    "红色": WD_COLOR_INDEX.RED, "红": WD_COLOR_INDEX.RED,
    "蓝色": WD_COLOR_INDEX.BLUE, "蓝": WD_COLOR_INDEX.BLUE,
    "粉色": WD_COLOR_INDEX.PINK, "粉": WD_COLOR_INDEX.PINK,
    "灰色": WD_COLOR_INDEX.GRAY_25, "灰": WD_COLOR_INDEX.GRAY_25,
}

_FILL_COLOR = {
    "黄色": "FFFF00", "黄": "FFFF00", "yellow": "FFFF00",
    "绿色": "92D050", "绿": "92D050", "green": "92D050",
    "红色": "FF0000", "红": "FF0000",
    "蓝色": "00B0F0", "蓝": "00B0F0",
    "青色": "00FFFF", "蓝绿": "00FFFF",
    "灰色": "D9D9D9", "灰": "D9D9D9",
}

# 语义目标识别器：指令里提到这类词、而不是某个具体词时，用识别器出候选区间
_SEMANTIC_RECOGNIZERS = {
    "金额": [
        re.compile(r"[¥￥]\s?[\d,，]+(?:\.\d+)?"),          # ￥1,800,000.00
        re.compile(r"人民币[\u4e00-\u9fa5]+?元整?"),          # 人民币壹佰捌拾万元整
    ],
}


def _clean_hint(h) -> str:
    """去掉定位线索两边的引号、空白。"""
    s = (h or "").strip()
    for ch in "\"'「」『』“”‘’":
        s = s.strip(ch)
    return s.strip()


# "标题"这类方位词指的是文档里的"位置"，不是要找的字面词（与 edit.py 的 TITLE_WORDS 保持同步）
TITLE_WORDS = ("标题", "题目", "标题行", "文档标题", "文章标题", "首行")


def _find_all(text: str, word: str):
    """返回 word 在 text 里每一次出现的 (start, end) 区间。"""
    spans = []
    if not word:
        return spans
    start = 0
    while True:
        i = text.find(word, start)
        if i == -1:
            return spans
        spans.append((i, i + len(word)))
        start = i + len(word)


def _color_rgb(v) -> RGBColor:
    """把'红色'/'FF0000'/'red' 统一变成 RGBColor。"""
    s = (v or "红色").strip()
    hexv = _COLOR_MAP.get(s, s)
    m = re.search(r"([0-9a-fA-F]{6})", hexv)
    h = m.group(1) if m else "FF0000"
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def _highlight_wd(v):
    """'黄色' → WD_COLOR_INDEX.YELLOW（docx 高亮枚举）。"""
    return _HIGHLIGHT_WD.get((v or "黄色").strip(), WD_COLOR_INDEX.YELLOW)


def _op_label(op: str, params: dict) -> str:
    """操作的短标签（报告用）：'加粗' / '取消加粗' / '高亮（黄色）' …"""
    params = params or {}
    base = {"set_bold": "加粗", "set_italic": "斜体", "set_underline": "下划线",
            "set_strike": "删除线",
            "set_highlight": "高亮（%s）" % (params.get("color") or "黄色"),
            "set_color": "标为%s" % (params.get("color") or "红色"),
            "set_font": "字体 = %s" % (params.get("font") or "黑体"),
            "set_size": "字号 = %s" % params.get("size")}.get(op, str(op))
    if params.get("off") and op in ("set_bold", "set_italic", "set_underline",
                                    "set_strike", "set_highlight"):
        return "取消" + base
    return base


def _parse_spacing(v):
    """'1.5 倍' → 1.5；'2倍行距' → 2.0。"""
    m = re.search(r"(\d+(?:\.\d+)?)", str(v or ""))
    return float(m.group(1)) if m else 1.5


def _parse_indent_chars(v) -> int:
    """'首行缩进2字符' → 200（Word 的 firstLineChars 单位：1 字符 = 100）。"""
    m = re.search(r"(\d+(?:\.\d+)?)\s*(?:个)?字", str(v or ""))
    n = float(m.group(1)) if m else 2.0
    return int(round(n * 100))


# ==================== docx · run 切分（文字级格式核心算法） ====================

def _set_elem_text(elem, text: str):
    """把一个 <w:r> 元素的文字设成 text（保留它的 rPr 格式）。"""
    ts = elem.findall(qn("w:t"))
    if ts:
        ts[0].text = text
        ts[0].set(_XML_SPACE, "preserve")
        for extra in ts[1:]:
            elem.remove(extra)
    else:
        t = OxmlElement("w:t")
        t.text = text
        t.set(_XML_SPACE, "preserve")
        elem.append(t)


def _split_run(r, a: int, b: int, setter):
    """把 run r 切成 前/中/后 三段，只对"中"段（[a,b)）应用 setter。

    这是"只加粗金额、不动整段"的关键：金额所在的 run 被拆开后，
    只有金额那一段带上新格式，左右两段保持原样。
    """
    t = r.text or ""
    segments = (t[:a], t[a:b], t[b:])
    elem = r._element
    parent = elem.getparent()
    idx = list(parent).index(elem)
    parent.remove(elem)
    mid_elem = None
    for i, seg in enumerate(segments):
        if seg == "":
            continue
        ne = copy.deepcopy(elem)
        _set_elem_text(ne, seg)
        parent.insert(idx, ne)
        idx += 1
        if i == 1:
            mid_elem = ne
    if mid_elem is not None:
        setter(Run(mid_elem, r._parent))


def _format_one_range(p, s: int, e: int, setter):
    """对段落 p 的字符区间 [s,e) 应用 setter（穿多个 run 时逐个切分）。"""
    runs = list(p.runs)
    pos = 0
    for r in runs:
        t = r.text or ""
        rs, re_ = pos, pos + len(t)
        pos = re_
        a, b = max(s, rs), min(e, re_)
        if a >= b:
            continue
        ra, rb = a - rs, b - rs
        if ra == 0 and rb == len(t):
            setter(r)                     # 整个 run 都在区间内：直接设
        else:
            _split_run(r, ra, rb, setter)  # 部分相交：切分


def _format_ranges(p, spans, setter):
    """对段落 p 的多个区间应用 setter。从后往前处理，防止前面切分后偏移错乱。"""
    for s, e in sorted(spans, key=lambda x: -x[0]):
        _format_one_range(p, s, e, setter)


# ==================== docx · 目标定位 ====================

def _find_title_para(d):
    """标题 = 样式为 Title 的段落；找不到就取第一个非空段落。"""
    for p in d.paragraphs:
        sty = p.style.name if p.style else ""
        if sty == "Title" and p.text.strip():
            return p
    for p in d.paragraphs:
        if p.text.strip():
            return p
    return None


def _find_body_paras(d):
    """正文 = 第一个标题之后、非 Title/Heading 的非空段落。"""
    nonempty = [p for p in d.paragraphs if p.text.strip()]
    first_head = None
    for i, p in enumerate(nonempty):
        sty = p.style.name if p.style else ""
        if sty.startswith("Heading"):
            first_head = i
            break
    start = first_head if first_head is not None else 0
    body = []
    for p in nonempty[start:]:
        sty = p.style.name if p.style else ""
        if sty == "Title" or sty.startswith("Heading"):
            continue
        body.append(p)
    return body


def _resolve_paragraph_targets(d, hint: str, op: str):
    """段落级操作的目标。返回 (段落列表, 描述)。"""
    hint = _clean_hint(hint)
    if hint in TITLE_WORDS:
        t = _find_title_para(d)
        return ([t] if t is not None else []), "标题段落"
    if hint == "正文":
        return _find_body_paras(d), "正文段落"
    if hint:
        hit = [p for p in d.paragraphs if p.text.strip() and hint in p.text]
        return hit, "含「%s」的段落" % hint
    # hint 为空：缩进 / 行距默认作用于正文
    if op in ("set_indent", "set_spacing"):
        return _find_body_paras(d), "正文段落（未指定对象，默认正文）"
    return [], "未指定对象"


def _resolve_text_ranges(d, hint: str, instruction: str):
    """文字级操作的目标。返回 ([(段落, [区间...], [显示文本...])], 描述)。"""
    hint = _clean_hint(hint)
    # A. 语义目标：hint 本身就是语义词，或 hint 为空但指令里出现语义词
    sem = None
    if hint in _SEMANTIC_RECOGNIZERS:
        sem = hint
    elif not hint:
        for k in _SEMANTIC_RECOGNIZERS:
            if k in instruction:
                sem = k
                break
    if sem:
        items = []
        for p in d.paragraphs:
            spans = []
            for pat in _SEMANTIC_RECOGNIZERS[sem]:
                spans.extend(m.span() for m in pat.finditer(p.text))
            if spans:
                spans.sort()
                items.append((p, spans, [p.text[s:e] for s, e in spans]))
        return items, "语义目标「%s」" % sem
    # B. 具体词：每段里该词的每一次出现
    if hint:
        items = []
        for p in d.paragraphs:
            spans = _find_all(p.text, hint)
            if spans:
                items.append((p, spans, [hint] * len(spans)))
        return items, "定位词「%s」" % hint
    return [], "没有可用的目标"


# ==================== docx · setter 工厂 ====================

# 文字级（run 级）操作全集：加粗/斜体/下划线/删除线/高亮/颜色/字体/字号
_TEXT_OPS = ("set_bold", "set_italic", "set_underline", "set_strike",
             "set_highlight", "set_color", "set_font", "set_size")


def _make_text_setter(op: str, params: dict):
    """给文字级操作造一个 setter(run)。params 里 off=true 表示"取消该格式"。"""
    off = bool((params or {}).get("off"))
    if op == "set_bold":
        def f(r):
            r.font.bold = not off
        return f
    if op == "set_italic":
        def f(r):
            r.font.italic = not off
        return f
    if op == "set_underline":
        def f(r):
            r.font.underline = not off
        return f
    if op == "set_strike":
        def f(r):
            r.font.strike = not off
        return f
    if op == "set_highlight":
        if off:
            def f(r):
                r.font.highlight_color = None
        else:
            col = _highlight_wd(params.get("color"))

            def f(r):
                r.font.highlight_color = col
        return f
    if op == "set_color":
        rgb = _color_rgb(params.get("color"))
        def f(r):
            r.font.color.rgb = rgb
        return f
    if op == "set_font":
        name = (params.get("font") or "黑体").strip()
        def f(r):
            r.font.name = name
            rPr = r._element.get_or_add_rPr()
            rFonts = rPr.get_or_add_rFonts()
            rFonts.set(qn("w:eastAsia"), name)   # 中文字体必须设 eastAsia，否则不生效
        return f
    if op == "set_size":
        size = float(params.get("size") or 12)
        def f(r):
            r.font.size = Pt(size)
        return f
    return None


def _apply_docx(src: Path, out_path: Path, intent: dict, instruction: str):
    d = docx.Document(str(src))
    op = intent.get("op")
    params = intent.get("params") or {}
    hint = ((intent.get("locate") or {}).get("hint")) or ""
    changes = []

    # ---------- 排版类（段落级） ----------
    if op in ("set_align", "set_spacing", "set_indent"):
        paras, desc = _resolve_paragraph_targets(d, hint, op)
        if not paras:
            return [], False, "没找到要处理的段落（定位线索：%r）" % hint

        if op == "set_align":
            key = (params.get("align") or "").lower().strip()
            align = _ALIGN_MAP.get(key)
            if align is None:
                return [], False, "不认识的 align 参数：%r" % params.get("align")
            for p in paras:
                p.paragraph_format.alignment = align
            cn = {"center": "居中", "right": "右对齐", "left": "左对齐",
                  "justify": "两端对齐"}.get(key, key)
            changes.append({
                "para": _para_nums(d, paras),
                "before": "对齐 = 默认（继承样式）",
                "after": "对齐 = %s" % cn,
                "how": "%s，共 %d 段：对齐设为「%s」" % (desc, len(paras), cn),
            })

        elif op == "set_spacing":
            sp = _parse_spacing(params.get("spacing"))
            for p in paras:
                p.paragraph_format.line_spacing = sp
            changes.append({
                "para": _para_nums(d, paras),
                "before": "行距 = 默认",
                "after": "行距 = %.1f 倍" % sp,
                "how": "%s，共 %d 段：行距设为 %.1f 倍" % (desc, len(paras), sp),
            })

        elif op == "set_indent":
            chars = _parse_indent_chars(params.get("indent"))
            for p in paras:
                _set_first_line_chars(p, chars)
            # 兜底把精确值也写上（不同阅读器认不同字段）
            changes.append({
                "para": _para_nums(d, paras),
                "before": "首行缩进 = 无",
                "after": "首行缩进 = %.1f 字符" % (chars / 100.0),
                "how": "%s，共 %d 段：首行缩进 %.1f 字符" % (desc, len(paras), chars / 100.0),
            })

        d.save(str(out_path))
        return changes, True, desc

    # ---------- 格式类（文字级） ----------
    if op in _TEXT_OPS:
        setter = _make_text_setter(op, params)
        if setter is None:
            return [], False, "不认识的格式操作：%s" % op

        # 分支机构：定位是"标题/正文"这类结构词 → 整段（全部 run）设
        h = _clean_hint(hint)
        if h in TITLE_WORDS or h == "正文":
            paras, desc = _resolve_paragraph_targets(d, h, op)
            if not paras:
                return [], False, "没找到要处理的段落（定位线索：%r）" % hint
            n_runs = 0
            for p in paras:
                for r in p.runs:
                    setter(r)
                    n_runs += 1
            _after = _op_label(op, params)
            changes.append({
                "para": _para_nums(d, paras),
                "before": "原格式",
                "after": _after,
                "how": "%s，共 %d 段 / %d 个文字块：%s"
                       % (desc, len(paras), n_runs, _after),
            })
            d.save(str(out_path))
            return changes, True, desc

        # 否则：语义识别 or 具体词（词级切分）
        items, desc = _resolve_text_ranges(d, hint, instruction)
        if not items:
            return [], False, "没找到要处理的文字（定位线索：%r）" % hint
        total = 0
        sample_texts = []
        for p, spans, texts in items:
            _format_ranges(p, spans, setter)
            total += len(spans)
            sample_texts.extend(texts)
        label = _op_label(op, params)                  # "加粗" / "取消加粗"
        off = bool((params or {}).get("off"))
        shown = "「" + "」「".join(sample_texts[:4]) + "」"
        if off:
            before_txt = "%s ×%d 处：%s（原有标记，已处理）" % (
                label.replace("取消", ""), total, shown)
            after_txt = "已取消（%s）" % shown
        else:
            before_txt = label + " = 无"
            after_txt = "%s ×%d 处：%s" % (label, total, shown)
        changes.append({
            "para": _para_nums(d, [it[0] for it in items]),
            "before": before_txt,
            "after": after_txt,
            "how": "%s：命中 %d 段、共 %d 处 → %s"
                   % (desc, len(items), total, label),
        })
        d.save(str(out_path))
        return changes, True, desc

    return [], False, ("不认识的格式操作：%s（docx 支持 加粗/斜体/下划线/删除线/"
                       "高亮/颜色/字体/字号 + 对齐/行距/缩进）" % op)


def _para_nums(d, paras) -> str:
    """给报告用：列出涉及的段落编号（最多 8 个）。"""
    try:
        nums = [str(d.paragraphs.index(p)) for p in paras[:8]]
        s = "段落[" + "、".join(nums) + "]"
        if len(paras) > 8:
            s += " 等 %d 段" % len(paras)
        return s
    except ValueError:
        return "段落（含表格内）"


def _set_first_line_chars(p, chars: int):
    """首行缩进：优先用 firstLineChars（按字符，Word 中文档的标准做法），
    同时写 firstLine（磅值兜底，其他阅读器也认）。"""
    pPr = p._p.get_or_add_pPr()
    ind = pPr.find(qn("w:ind"))
    if ind is None:
        ind = OxmlElement("w:ind")
        pPr.append(ind)
    ind.set(qn("w:firstLineChars"), str(chars))
    # 兜底：按 12 磅字号估算（2 字符 ≈ 24 磅 = 480 twips），字号不同也不影响
    # Word 的显示——它会优先用 firstLineChars。
    ind.set(qn("w:firstLine"), str(int(round(chars / 100.0 * 240))))


# ==================== xlsx ====================

def _find_header_row(ws):
    """表头行 = 前 6 行里非空格子数最多的那一行（与模板读取器同一规则）。"""
    best, best_n = 1, -1
    limit = min(6, ws.max_row)
    for r in range(1, limit + 1):
        n = sum(1 for c in range(1, ws.max_column + 1)
                if ws.cell(row=r, column=c).value is not None)
        if n > best_n:
            best, best_n = r, n
    return best


_UNSET = object()   # 哨兵：区分"不改这个属性"（未传）与"把属性设空"（显式传 None）


def _set_cell_font(cell, size=_UNSET, bold=_UNSET, name=_UNSET, color=_UNSET,
                   italic=_UNSET, underline=_UNSET, strike=_UNSET):
    """改单元格字体，同时保留原有其他字体属性（复制后只改要改的）。
    未传的参数 = 保持原样；显式传 None/False = 清空该属性。"""
    f = copy.copy(cell.font)
    if size is not _UNSET:
        f.size = size
    if bold is not _UNSET:
        f.bold = bold
    if italic is not _UNSET:
        f.italic = italic
    if underline is not _UNSET:
        f.underline = underline          # openpyxl：'single'/'double'，None=无
    if strike is not _UNSET:
        f.strike = strike
    if name is not _UNSET:
        f.name = name
    if color is not _UNSET:
        f.color = color
    cell.font = f


def _set_cell_align(cell, horizontal: str):
    """改单元格水平对齐，同时保留原有垂直对齐 / 自动换行等属性。"""
    old = cell.alignment
    cell.alignment = Alignment(horizontal=horizontal,
                               vertical=old.vertical,
                               wrap_text=old.wrap_text)


def _xlsx_targets(ws, header_row: int, hint: str, src_stem: str):
    """把定位线索（hint）变成 xlsx 的目标格子清单。返回 (cells, desc)；
    找不到则返回 (None, None)，由调用方报"没找到"（宁可暴露问题，不许悄悄错）。

    规则（按优先级）：
      ① hint 空 / "表头"类词 / 文档名残影 → 表头行非空格
      ② hint 含"整张表/整表/全表" → 全表非空格
      ③ hint 匹配表头列名 → 该列（表头格 + 数据非空格）
      ④ hint 是表里出现的文字 → 含该文字的格子
      ⑤ 都没命中 → None
    """
    H = _clean_hint(hint)

    # ① 表头（默认目标）
    if (not H or H in ("表头", "表头行", "标题行", "首行")
            or (len(H) >= 3 and (H in src_stem or src_stem in H))):
        cells = [ws.cell(row=header_row, column=c)
                 for c in range(1, ws.max_column + 1)
                 if ws.cell(row=header_row, column=c).value is not None]
        return cells, "表头行（第 %d 行）" % header_row

    # ② 整张表
    if any(k in H for k in ("整张表", "整表", "全表", "整个表")):
        cells = [cell for row in ws.iter_rows() for cell in row if cell.value is not None]
        return cells, "全表（%d 个非空格）" % len(cells)

    # ③ 列名（先精确、后包含）
    h = H.replace("列", "").strip()
    col_idx, col_name = None, ""
    for c in range(1, ws.max_column + 1):
        v = ws.cell(row=header_row, column=c).value
        if v is not None and str(v).strip() == h:
            col_idx, col_name = c, str(v).strip()
            break
    if col_idx is None:
        for c in range(1, ws.max_column + 1):
            v = ws.cell(row=header_row, column=c).value
            if v is not None and (h in str(v) or str(v) in h):
                col_idx, col_name = c, str(v).strip()
                break
    if col_idx is not None:
        cells = [ws.cell(row=r, column=col_idx)
                 for r in range(header_row, ws.max_row + 1)
                 if ws.cell(row=r, column=col_idx).value is not None]
        return cells, "「%s」列（第 %d 列，共 %d 格）" % (col_name, col_idx, len(cells))

    # ④ 文字命中
    cells = [cell for row in ws.iter_rows() for cell in row
             if isinstance(cell.value, str) and H in cell.value]
    if cells:
        coords = "、".join(c.coordinate for c in cells[:5])
        more = " 等" if len(cells) > 5 else ""
        return cells, "含「%s」的格子（%s%s）" % (H, coords, more)

    # ⑤ 没找到
    return None, None


def _apply_xlsx(src: Path, out_path: Path, intent: dict, instruction: str):
    op = intent.get("op")
    params = intent.get("params") or {}
    hint = _clean_hint(((intent.get("locate") or {}).get("hint")) or "")
    wb = load_workbook(str(src))
    ws = wb.active
    header_row = _find_header_row(ws)

    # ① 定位：把 hint 变成目标格子清单（列名 → 整列 / 表头 → 表头行 / 文字 → 命中格）
    targets, desc = _xlsx_targets(ws, header_row, hint, src.stem)
    if targets is None:
        return [], False, ("没找到「%s」对应的位置——可以试试按列名（如「工号列」）"
                           "或「表头」「整张表」来说。" % hint)

    # ② 执行（所有 op 都作用在 targets 上；只动目标、不动别处）
    if op == "set_size":
        size = float(params.get("size") or 14)
        for cell in targets:
            _set_cell_font(cell, size=size)
        before, after = "字号 = 原值", "字号 = %g" % size
    elif op in ("set_bold", "set_italic", "set_underline", "set_strike"):
        kw = {"set_bold": "bold", "set_italic": "italic",
              "set_underline": "underline", "set_strike": "strike"}[op]
        v = not bool(params.get("off"))
        if op == "set_underline":
            v = "single" if v else None      # openpyxl 下划线只认枚举字符串 / None
        for cell in targets:
            _set_cell_font(cell, **{kw: v})
        base = {"set_bold": "加粗", "set_italic": "斜体",
                "set_underline": "下划线", "set_strike": "删除线"}[op]
        before, after = "%s = 原值" % base, _op_label(op, params)
    elif op == "set_highlight":
        if params.get("off"):
            for cell in targets:
                cell.fill = PatternFill(fill_type=None)
        else:
            c = (params.get("color") or "黄色").strip()
            hexv = _FILL_COLOR.get(c, "FFFF00")
            for cell in targets:
                cell.fill = PatternFill(fill_type="solid", start_color="FF" + hexv)
        before, after = "填充 = 原值", _op_label(op, params)
    elif op == "set_font":
        name = (params.get("font") or "黑体").strip()
        for cell in targets:
            _set_cell_font(cell, name=name)
        before, after = "字体 = 原值", "字体 = %s" % name
    elif op == "set_color":
        hexrgb = str(_color_rgb(params.get("color")))
        for cell in targets:
            _set_cell_font(cell, color=Color(rgb="FF" + hexrgb))
        before, after = "颜色 = 原值", "文字颜色 = %s" % (params.get("color") or "红")
    elif op == "set_align":
        amap = {"center": "center", "居中": "center", "right": "right", "右对齐": "right",
                "left": "left", "左对齐": "left", "justify": "justify", "两端对齐": "justify"}
        cn = {"center": "居中", "right": "右对齐", "left": "左对齐", "justify": "两端对齐"}
        hv = amap.get(str(params.get("align") or "center"), "center")
        for cell in targets:
            _set_cell_align(cell, hv)
        before, after = "对齐 = 原值", "对齐 = %s" % cn.get(hv, hv)
    else:
        return [], False, ("xlsx 支持 字号/加粗/斜体/下划线/删除线/高亮底色/字体/"
                           "颜色/对齐（收到：%s）" % op)

    wb.save(str(out_path))
    changes = [{
        "para": desc,
        "before": before,
        "after": after,
        "how": "%s：%s" % (desc, after),
    }]
    return changes, True, desc


# ==================== md / txt ====================

# md/txt 支持的格式 op（对齐/行距这类版式在纯文本里没有写法，会明确拒绝）
_MD_OPS = ("set_bold", "set_italic", "set_strike", "set_highlight", "set_color")

# 各类标记的"成对"正则（取消格式时用来剥壳）
_MD_PAIRS = {
    "set_bold": (r"\*\*", r"\*\*"),
    "set_italic": (r"(?<!\*)\*(?!\*)", r"(?<!\*)\*(?!\*)"),
    "set_strike": (r"~~", r"~~"),
    "set_highlight": (r"<mark>", r"</mark>"),
}


def _md_wrap(body: str, op: str, params: dict):
    """给一段文字套上 md 格式标记。返回 (新文本, 描述)。"""
    if op == "set_bold":
        return "**%s**" % body, "加粗（** **）"
    if op == "set_italic":
        return "*%s*" % body, "斜体（* *）"
    if op == "set_strike":
        return "~~%s~~" % body, "删除线（~~ ~~）"
    if op == "set_highlight":
        return "<mark>%s</mark>" % body, "高亮（<mark>）"
    s = (params.get("color") or "红色").strip()
    hexv = _COLOR_MAP.get(s)
    css = "#%s" % hexv if hexv else s
    return '<span style="color:%s">%s</span>' % (css, body), "标为%s（HTML span）" % s


def _md_unwrap(segment: str, op: str, word: str = None):
    """去掉 segment 里的成对标记（**x** → x）。
    word=None：处理整个 segment 里所有该类标记对（标题行用）；
    word 给定时：只处理"标记 + word + 标记"的精确形式（词级取消用）。
    返回 (新文本, 命中数)。"""
    pair = _MD_PAIRS.get(op)
    if pair is None:
        return segment, 0
    l, r = pair
    inner = re.escape(word) if word else r".+?"
    pat = l + "(" + inner + ")" + r
    new_text, n = re.subn(pat, lambda m: m.group(1), segment)
    return new_text, n


def _apply_md(src: Path, out_path: Path, intent: dict, instruction: str):
    op = intent.get("op")
    params = intent.get("params") or {}
    off = bool(params.get("off"))
    hint = _clean_hint(((intent.get("locate") or {}).get("hint")) or "")

    # ① 能力检查：md/txt 是纯文本，能做 加粗/斜体/删除线/高亮/标红
    if op not in _MD_OPS:
        return [], False, (
            "md/txt 是纯文本格式，支持 加粗 / 斜体 / 删除线 / 高亮 / 标红；"
            "「%s」这类在纯文本里没有对应写法（可换用上面几种，或改用 docx / xlsx 文档）" % op)

    text = src.read_text(encoding="utf-8")
    label = _op_label(op, params)                     # "加粗" / "取消加粗"
    plain = label.replace("取消", "")

    # ② 「标题」类方位词 → 定位到首行（第一个非空行）做格式
    if hint in TITLE_WORDS:
        lines = text.splitlines(keepends=True)
        idx = None
        for i, ln in enumerate(lines):
            if ln.strip():
                idx = i
                break
        if idx is None:
            return [], False, "文档是空的，没有可当标题的行"
        old_line = lines[idx].rstrip("\n")
        m = re.match(r"^(\s*#{1,6}\s+)", old_line)
        prefix = m.group(1) if m else ""
        body = old_line[len(prefix):]
        if not body.strip():
            return [], False, "标题行没有文字内容"
        if off:
            new_body, n = _md_unwrap(body, op)
            if n == 0:
                return [], False, ("标题行上没有找到 %s 标记"
                                   "（可能它本来就没有这个格式）" % plain)
            desc_after = "标题：已取消%s" % plain
        else:
            new_body, d2 = _md_wrap(body, op, params)
            desc_after = "标题" + d2
        lines[idx] = prefix + new_body + ("\n" if lines[idx].endswith("\n") else "")
        out_path.write_text("".join(lines), encoding="utf-8")
        changes = [{
            "para": "第 %d 行（标题）" % (idx + 1),
            "before": old_line, "after": prefix + new_body,
            "how": desc_after,
        }]
        return changes, True, "标题行（第 %d 行）" % (idx + 1)

    # ③ 普通字面词
    if not hint:
        return [], False, "没找到要处理的文字（定位线索为空）"

    if off:
        new_text, n = _md_unwrap(text, op, word=hint)
        if n == 0:
            return [], False, ("没找到「%s」上的 %s 标记——它可能本来就没有这个格式"
                               "（取消只对已带标记的文字生效）" % (hint, plain))
        out_path.write_text(new_text, encoding="utf-8")
        changes = [{
            "para": "全文",
            "before": "「%s」×%d 处 %s标记" % (hint, n, plain),
            "after": "已取消（标记已去除）",
            "how": "文本文件：%d 处「%s」→ 取消%s" % (n, hint, plain),
        }]
        return changes, True, "全文（%d 处命中）" % n

    n = text.count(hint)
    if n == 0:
        return [], False, "文件里找不到「%s」" % hint
    wrapped, desc_after = _md_wrap(hint, op, params)
    # 防重复：先剥掉同类已有标记（幂等），再统一包——避免"加粗两次"变成 ****
    if op in _MD_PAIRS:
        text, _ = _md_unwrap(text, op, word=hint)
    new_text = text.replace(hint, wrapped)
    out_path.write_text(new_text, encoding="utf-8")
    changes = [{
        "para": "全文",
        "before": "「%s」×%d 处 无格式" % (hint, n),
        "after": "「%s」×%d 处 %s" % (hint, n, desc_after),
        "how": "文本文件：%d 处「%s」→ %s" % (n, hint, desc_after),
    }]
    return changes, True, "全文（%d 处命中）" % n


# ==================== 统一入口 ====================

_QUOTED_RE = re.compile(r"['\"“”‘’「『]([^'\"“”‘’」』\n]{1,24})['\"“”‘’」』]")


def _fill_format_target(intent: dict, instruction: str):
    """兜底（模型飘移之二）：模型偶尔漏解析 locate.hint——
    如"把周报里所有的'问题'标红"被解析成 hint 空。此时从指令原文的引号里
    把目标词捡回来（指令原文里用户亲手打的引号，是最可靠的信号）。

    只对"没有合理默认目标"的操作补：文字级四操作 + 对齐；
    行距 / 缩进的空白有"默认正文"的合法行为，不掺和。
    只填空位，绝不覆盖模型已给出的值。
    """
    if intent.get("op") not in ("set_bold", "set_italic", "set_underline", "set_strike",
                                "set_highlight", "set_color", "set_font", "set_size",
                                "set_align"):
        return
    loc = intent.get("locate")
    if not isinstance(loc, dict):
        loc = {}
        intent["locate"] = loc
    if (loc.get("hint") or "").strip():
        return
    m = _QUOTED_RE.search(instruction or "")
    if m:
        loc["hint"] = m.group(1).strip()


# 模型偶尔把"斜体/下划线/删除线/高亮"错塞进 set_font 的 font 参数
# （这些词不可能是字体名，纠正是安全的）
_FONT_MISPLACED = (
    ("斜体", "set_italic"), ("倾斜", "set_italic"), ("italic", "set_italic"),
    ("下划线", "set_underline"), ("底线", "set_underline"), ("underline", "set_underline"),
    ("删除线", "set_strike"), ("strike", "set_strike"),
    ("高亮", "set_highlight"), ("突出显示", "set_highlight"),
    ("加粗", "set_bold"), ("粗体", "set_bold"),
)


def _fix_format_drift(intent: dict, instruction: str):
    """兜底（模型飘移之三）：'把标题设成斜体'这类指令，模型偶尔把"斜体"
    填进 set_font 的 font 参数。这些词不可能是字体名，看到就纠正成对应的
    独立 op——避免"悄悄把字体改错"。
    """
    if intent.get("op") != "set_font":
        return
    f = str(((intent.get("params") or {}).get("font")) or "")
    for w, new_op in _FONT_MISPLACED:
        if w in f:
            intent["op"] = new_op
            intent["params"] = {k: v for k, v in (intent.get("params") or {}).items()
                                if k != "font"}
            return


_OFF_WORDS = ("取消", "去掉", "去除", "消除", "不要", "撤掉", "撤销")


def _fix_off_drift(intent: dict, instruction: str):
    """兜底（模型飘移之四）：指令说"把加粗去掉"，模型偶尔忘带 off 参数。
    在"格式词 + 取消词"紧邻（中间不超过 4 个字）时，给对应 op 补 off=true。
    保守规则：中间隔着逗号/分号就认为不是一个动作（如"给标题加粗，去掉下划线"
    不会误伤"加粗"）。只补空位，不覆盖已有 off。
    """
    op = intent.get("op")
    off_for = {"set_bold": ("加粗", "粗体"), "set_italic": ("斜体", "倾斜"),
               "set_underline": ("下划线",), "set_strike": ("删除线",),
               "set_highlight": ("高亮", "突出显示")}
    if op not in off_for:
        return
    p = intent.get("params") or {}
    if p.get("off"):
        return
    text = instruction or ""
    offw = "|".join(_OFF_WORDS)
    for w in off_for[op]:
        pat1 = r"(%s)[^，。；,;\n]{0,4}%s" % (offw, re.escape(w))
        pat2 = r"%s[^，。；,;\n]{0,4}(%s)" % (re.escape(w), offw)
        if re.search(pat1, text) or re.search(pat2, text):
            intent["params"] = dict(p, off=True)
            return


def apply_format(src: Path, out_path: Path, intent: dict, instruction: str):
    """按文件类型路由。返回 (changes, ok, msg)。"""
    _fix_format_drift(intent, instruction)
    _fix_off_drift(intent, instruction)
    _fill_format_target(intent, instruction)
    suffix = src.suffix.lower()
    if suffix == ".docx":
        return _apply_docx(src, out_path, intent, instruction)
    if suffix == ".xlsx":
        return _apply_xlsx(src, out_path, intent, instruction)
    if suffix in (".md", ".txt"):
        return _apply_md(src, out_path, intent, instruction)
    return [], False, "暂不支持 %s 类型的格式操作（目前支持 docx / xlsx / md / txt）" % suffix
