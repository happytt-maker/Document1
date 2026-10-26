# -*- coding: utf-8 -*-
"""
readers/txt_reader.py
---------------------
纯文本读取器。

一个 txt 文件读成一个 JSON 对象：
{
  "file": "供应商联系方式.txt",
  "type": "txt",
  "encoding": "utf-8",
  "line_count": 26,
  "blocks": [
    {"line_start": 1, "line_end": 3, "text": "..."},
    ...
  ]
}
"""

from .base import make_block, read_text_file, split_paragraphs


def read_txt(path):
    """读一个 txt 文件，返回带位置信息的字典。"""
    text, encoding = read_text_file(path)
    lines = text.splitlines()

    blocks = []
    for start, end, para_text in split_paragraphs(lines):
        blocks.append(
            make_block(
                file_name=path.name,
                file_type="txt",
                text=para_text,
                locator={"line_start": start, "line_end": end},
            )
        )

    return {
        "file": path.name,
        "type": "txt",
        "encoding": encoding,
        "line_count": len(lines),
        "block_count": len(blocks),
        "blocks": blocks,
    }
