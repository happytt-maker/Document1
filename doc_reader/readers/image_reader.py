# -*- coding: utf-8 -*-
"""
readers/image_reader.py
-----------------------
图片读取器（截图 / 拍照 / 扫描件 → OCR → 块）

支持的扩展名：.png .jpg .jpeg .bmp .webp .tif .tiff

定位方式：locator = {"kind": "image_para", "para": 第几段, "conf": 置信度}
——图片没有页码，所以定位落到「第几段」，段内顺序就是阅读顺序（从上到下）。

【和 image_ops.py 的区别，别搞混】
  image_ops.py    是「把图插进 Word/Excel」——图是素材，不认字；
  image_reader.py 是「把图里的字读出来当文档」——认字，进检索池。
  两条路都从 导入图片/ 之外来，互不影响。
"""

from pathlib import Path

from .base import make_block, load_ocr_module

IMAGE_EXTS = {".png", ".jpg", ".jpeg", ".bmp", ".webp", ".tif", ".tiff"}
# 认出来的有效字少于这个数 → 当"没认到东西"处理，不往池子里塞垃圾块
MIN_CHARS = 4


def _pil_size(path):
    try:
        from PIL import Image as PILImage
        with PILImage.open(str(path)) as im:
            return int(im.width), int(im.height)
    except Exception:
        return None, None


def read_image(path):
    """读一张图片：OCR → 段落块。"""
    path = Path(path)
    ocr = load_ocr_module()
    r = ocr.read_image(str(path))

    if not r.get("ok"):
        raise ValueError(r.get("error") or "OCR 失败")
    if r.get("char_count", 0) < MIN_CHARS:
        raise ValueError("这张图里没认到文字（认出的有效字不到 %d 个）——"
                         "可能是纯照片 / 图标，或者分辨率太低" % MIN_CHARS)

    blocks = []
    for i, p in enumerate(r.get("paragraphs") or []):
        text = (p.get("text") or "").strip()
        if not text:
            continue
        blocks.append(make_block(
            file_name=path.name,
            file_type="image",
            text=text,
            locator={"kind": "image_para", "para": i, "conf": p.get("conf")},
            source="ocr",
            ocr_conf=p.get("conf"),
        ))

    w, h = _pil_size(path)
    out = {
        "file": path.name,
        "type": "image",
        "ext": path.suffix.lower(),
        "width": w,
        "height": h,
        "block_count": len(blocks),
        "blocks": blocks,
        "ocr": True,
        "ocr_engine": r.get("engine"),
        "ocr_elapsed": r.get("elapsed"),
        "avg_conf": r.get("avg_conf"),
        "char_count": r.get("char_count"),
        "line_count": r.get("line_count"),
    }
    if r.get("low_conf_lines"):
        out["low_conf_lines"] = r["low_conf_lines"]
    if r.get("low_conf_tip"):
        out["notes"] = [r["low_conf_tip"]]
    return out
