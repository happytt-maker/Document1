# -*- coding: utf-8 -*-
"""
readers/base.py
---------------
所有格式读取器的公共约定。

每个读取器只做一件事：把文件读成「块（block）」列表。
每个块都带位置信息，方便以后回溯到原文。
"""


def make_block(file_name, file_type, text, locator, **extra):
    """
    造一个统一的内容块。

    file_name : 文件名，如 "供应商联系方式.txt"
    file_type : 格式，如 "txt"
    text      : 这个块的内容
    locator   : 位置信息（字典），说明这段内容在原文件的哪里
    extra     : 其他要附带的字段（比如是不是公式）
    """
    block = {
        "file": file_name,
        "type": file_type,
        "text": text,
        "locator": locator,
    }
    block.update(extra)
    return block


def read_text_file(path):
    """
    按多种编码尝试读取文本文件。

    Windows 上中文 txt 可能是 GBK，不是 UTF-8。
    直接按 UTF-8 读会抛 UnicodeDecodeError，所以要逐个试。

    返回 (内容字符串, 实际用的编码)
    """
    with open(path, "rb") as f:
        raw = f.read()

    # utf-8-sig 排在最前：它能顺带吃掉 BOM 头
    for enc in ("utf-8-sig", "utf-8", "gbk", "gb18030"):
        try:
            return raw.decode(enc), enc
        except UnicodeDecodeError:
            continue

    raise ValueError("无法识别文件编码: " + str(path))


def split_paragraphs(lines):
    """
    把按行组织的文本切成「段落块」。

    规则：连续的空白行作为分隔。非空的一段连续行算一个段落。

    返回 [(起始行号, 结束行号, 段落文字), ...]
    行号从 1 开始数（方便人对着原文找）。
    """
    paragraphs = []
    start = None
    buffer = []

    for idx, line in enumerate(lines, start=1):
        if line.strip() == "":
            # 遇到空行，把之前攒的段落收掉
            if buffer:
                paragraphs.append((start, idx - 1, "\n".join(buffer)))
                buffer = []
                start = None
        else:
            if start is None:
                start = idx
            buffer.append(line.rstrip("\n"))

    # 文件末尾没有空行时，也要收尾
    if buffer:
        paragraphs.append((start, len(lines), "\n".join(buffer)))

    return paragraphs
