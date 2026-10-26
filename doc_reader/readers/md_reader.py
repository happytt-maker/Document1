# -*- coding: utf-8 -*-
"""
readers/md_reader.py
--------------------
Markdown 读取器。

比 txt 多做一件事：记录每个段落处于哪个「标题」之下。
这叫「标题路径」，比单纯的行号更好用。

例：一个段落上面最近的标题是 「## 2. 系统组成」，
    再上面是 「# 文档理解与多源数据融合系统」，
    那它的 section 就是 "文档理解与多源数据融合系统 > 2. 系统组成"

输出格式：
{
  "file": "项目技术方案摘要.md",
  "type": "md",
  "line_count": 120,
  "headings": [{"level": 1, "text": "...", "line": 1}, ...],
  "blocks": [
    {"line_start": 8, "line_end": 10, "text": "...", "section": "标题路径"}
  ]
}
"""

import re

from .base import make_block, read_text_file, split_paragraphs

# 匹配 Markdown 标题：行首 1~6 个 #，后面跟空格和标题文字
HEADING_RE = re.compile(r"^(#{1,6})\s+(.*)$")

# 匹配围栏代码块的开始/结束行：``` 或 ~~~ 开头
FENCE_RE = re.compile(r"^\s*(```|~~~)")


def _code_block_lines(lines):
    """
    找出所有处于「围栏代码块」内部的行号（从 1 开始计）。

    为什么要这个？因为代码块里可能写着：
        ### 错误！id(cell) 不可靠
    这不是真标题，只是代码示例里的注释。如果不排除，
    标题路径就会被污染，分节归错。

    返回一个集合，比如 {21, 22, 23, 24, 25, 26, 27, 28}。
    """
    inside = set()
    fence_open = False   # 当前是否在代码块里

    for idx, line in enumerate(lines, start=1):
        if FENCE_RE.match(line):
            # 遇到 ``` 就在「进代码块」和「出代码块」之间切换
            inside.add(idx)
            fence_open = not fence_open
        elif fence_open:
            inside.add(idx)

    return inside


def _build_section_path(lines):
    """
    先扫一遍文件，标出每一行处于什么标题之下。

    返回一个列表，第 i 个元素是「第 i+1 行所属的标题路径」。
    标题行本身也归到它自己的路径里。

    代码块内部的行会被跳过，不参与标题判断。
    """
    section_of_line = []
    heading_stack = []  # 每个元素是 (层级, 标题文字)
    code_lines = _code_block_lines(lines)

    for idx, line in enumerate(lines, start=1):
        # 代码块里的 ### 是注释，不是标题 —— 直接跳过判断
        if idx not in code_lines:
            m = HEADING_RE.match(line)
            if m:
                level = len(m.group(1))
                title = m.group(2).strip()

                # 弹出所有层级 >= 当前层级的旧标题
                while heading_stack and heading_stack[-1][0] >= level:
                    heading_stack.pop()
                heading_stack.append((level, title))

        path = " > ".join(t for _, t in heading_stack)
        section_of_line.append(path)

    return section_of_line


def _collect_headings(lines):
    """把文件里所有标题收集起来，方便一次性看到全貌。代码块里的不算。"""
    code_lines = _code_block_lines(lines)
    headings = []
    for idx, line in enumerate(lines, start=1):
        if idx in code_lines:
            continue
        m = HEADING_RE.match(line)
        if m:
            headings.append(
                {
                    "level": len(m.group(1)),
                    "text": m.group(2).strip(),
                    "line": idx,
                }
            )
    return headings


def read_md(path):
    """读一个 md 文件，返回带标题路径的字典。"""
    text, _encoding = read_text_file(path)
    lines = text.splitlines()

    section_of_line = _build_section_path(lines)
    code_lines = _code_block_lines(lines)

    blocks = []
    for start, end, para_text in split_paragraphs(lines):
        # 段落所属的标题路径，取它第一行的
        section = section_of_line[start - 1] if start - 1 < len(section_of_line) else ""

        # 判断这个段落本身是不是标题（代码块里的 # 不算）
        first_line = lines[start - 1]
        is_heading = start not in code_lines and bool(HEADING_RE.match(first_line))

        # 这个段落是不是代码块（整段都在围栏里）
        in_code = start in code_lines and end in code_lines

        blocks.append(
            make_block(
                file_name=path.name,
                file_type="md",
                text=para_text,
                locator={"line_start": start, "line_end": end},
                section=section,
                is_heading=is_heading,
                is_code=in_code,
            )
        )

    return {
        "file": path.name,
        "type": "md",
        "line_count": len(lines),
        "headings": _collect_headings(lines),
        "block_count": len(blocks),
        "blocks": blocks,
    }
