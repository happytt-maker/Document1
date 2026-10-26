# -*- coding: utf-8 -*-
"""
readers/docx_reader.py
----------------------
Word 文档读取器。

把 docx 读成带位置信息的块，输出格式和 txt_reader / md_reader 保持一致。

产出两类内容：
  1. 段落块 —— 位置记为「第几段」(para_index)
  2. 表格块 —— 位置记为「第几张表 / 第几行 / 第几列」(table_index / row / col)

═══════════════════════════════════════════════════════════════════
   三个必须知道的坑
═══════════════════════════════════════════════════════════════════

【坑 1】合并单元格去重必须用 cell._tc，不能用 id(cell)

    row.cells 每访问一次都会新建一层 _Cell 包装对象，id() 拿到的是
    那层临时壳的地址，不是单元格本身的地址。

    实测（会议纪要，6 行 4 列，应有 23 个唯一单元格）：
        坐标 (1,0)  id=1527604609712  text='吴晓敏\\n吴晓敏'
        坐标 (2,0)  id=1527604666192  text='吴晓敏\\n吴晓敏'
        两者是同一个合并单元格，id 却不同 → 用 id 去重后仍是 24（错）
        用 cell._tc 去重后是 23（对）

【坑 2】doc.paragraphs 和 doc.tables 是分开的

    doc.paragraphs 只含正文段落，不含表格里的段落。
    只读一个 → 另一半内容全丢。必须两个都读。

【坑 3】合并后的 cell.text 会叠加

    合并单元格的 text 是多个段落用 \\n 拼起来的，例如 '吴晓敏\\n吴晓敏'。
    直接拿去填表会出错，要按需要拆开。
"""

from docx import Document
from docx.oxml.ns import qn

from .base import make_block


def _clean(text):
    """把单元格/段落里的文字清理一下：去掉首尾空白，合并内部多余空行。"""
    if text is None:
        return ""
    return text.strip()


def _split_merged(text):
    """
    处理合并单元格里叠加的内容。

    '吴晓敏\n吴晓敏'  →  '吴晓敏'      （两行内容相同，说明是纵向合并，取一份）
    '标题\n内容'      →  '标题 内容'   （内容不同，用空格连起来，不丢信息）
    """
    if "\n" not in text:
        return text

    parts = [p.strip() for p in text.split("\n") if p.strip()]

    if not parts:
        return ""

    # 所有行都一样 → 纵向合并导致重复，只留一份
    if len(set(parts)) == 1:
        return parts[0]

    # 有不同内容 → 用空格连起来（不丢信息）
    return " ".join(parts)


def read_paragraphs(doc, file_name):
    """
    读正文段落。

    位置记为 para_index（第几段，从 1 开始）。
    序号用「原始序号」而不是「有效段落序号」，这样回溯原文时对得上，
    空段落也占一个号。
    """
    blocks = []

    for idx, para in enumerate(doc.paragraphs, start=1):
        text = _clean(para.text)
        if not text:
            continue

        # 记录段落样式名，方便区分标题和正文（标题的样式是 Heading 1/2 等）
        style_name = ""
        try:
            style_name = para.style.name or ""
        except Exception:
            pass

        is_heading = style_name.lower().startswith("heading") or style_name.startswith("标题")

        blocks.append(
            make_block(
                file_name=file_name,
                file_type="docx",
                text=text,
                locator={"kind": "paragraph", "para_index": idx},
                style=style_name if style_name else None,
                is_heading=is_heading,
            )
        )

    return blocks


def read_tables(doc, file_name):
    """
    读表格。

    位置记为 table_index / row / col（都从 1 开始，方便人对着表数）。

    ★ 去重逻辑用 cell._tc —— 这是本文件最关键的一行
    """
    blocks = []

    for t_idx, table in enumerate(doc.tables, start=1):
        seen = set()          # 记录已经处理过的底层单元格元素

        for r_idx, row in enumerate(table.rows, start=1):
            for c_idx, cell in enumerate(row.cells, start=1):
                key = cell._tc            # ← 底层 XML 元素，单元格的真实身份
                if key in seen:
                    continue              # 合并单元格在多个坐标重复出现，跳过
                seen.add(key)

                text = _clean(cell.text)
                if not text:
                    continue

                # 合并后的内容可能叠加，处理一下
                cleaned = _split_merged(text)

                # 取一下合并信息（合并了几列/几行），以后可能有用
                span = None
                vmerge = None
                try:
                    tcPr = cell._tc.tcPr
                    if tcPr is not None:
                        if tcPr.gridSpan is not None:
                            span = tcPr.gridSpan.val
                        if tcPr.vMerge is not None:
                            vmerge = True
                except Exception:
                    pass

                blocks.append(
                    make_block(
                        file_name=file_name,
                        file_type="docx",
                        text=cleaned,
                        locator={
                            "kind": "table_cell",
                            "table_index": t_idx,
                            "row": r_idx,
                            "col": c_idx,
                        },
                        raw_text=text if text != cleaned else None,
                        grid_span=span,
                        v_merge=vmerge,
                    )
                )

    return blocks


def read_docx(path):
    """
    读一个 docx 文件，返回和 read_txt / read_md 同构的字典。
    """
    doc = Document(str(path))

    para_blocks = read_paragraphs(doc, path.name)
    table_blocks = read_tables(doc, path.name)
    blocks = para_blocks + table_blocks

    return {
        "file": path.name,
        "type": "docx",
        "paragraph_count": len(doc.paragraphs),
        "table_count": len(doc.tables),
        "block_count": len(blocks),
        "blocks": blocks,
    }
