# -*- coding: utf-8 -*-
"""
readers/__init__.py
-------------------
读取器总入口。

按扩展名把文件分派给对应的读取器。
目前只支持 txt 和 md（第 1 步）。
"""

from pathlib import Path

from .txt_reader import read_txt
from .md_reader import read_md
from .docx_reader import read_docx
from .xlsx_reader import read_xlsx

# 扩展名 -> 读取函数
READERS = {
    ".txt": read_txt,
    ".md": read_md,
    ".markdown": read_md,
    ".docx": read_docx,
    ".xlsx": read_xlsx,
}

# 目前支持的扩展名（后面加上 docx / xlsx 时，只要往这里加）
SUPPORTED = sorted(READERS.keys())


def is_supported(path):
    """这个文件是不是我们现在能读的。"""
    return Path(path).suffix.lower() in READERS


def read_one(path):
    """
    读一个文件，返回字典。

    如果是暂不支持的格式，返回 None（调用方自己决定跳过还是记录）。
    """
    path = Path(path)
    suffix = path.suffix.lower()

    reader = READERS.get(suffix)
    if reader is None:
        return None

    return reader(path)
