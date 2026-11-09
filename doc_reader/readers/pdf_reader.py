# -*- coding: utf-8 -*-
"""
readers/pdf_reader.py
---------------------
PDF 读取器（带扫描页 OCR）。

一句话流程：
  打开 PDF → 一页一页看 → 这一页自带文字层吗？
      是 → 直接抽文字（快、准、0 误差）
      否（扫描页 / 拍照页）→ 把这一页渲染成图 → 交给 agent/ocr.py 认字
  → 两种结果都变成同一种「块」，带页码定位，后面的检索 / 抽取 / 填表照常用

为什么按页判断、而不是整份判断：
  真实 PDF 经常是混的——前面几页是电子版正文，中间夹几页扫描的签字页 / 盖章页。
  整份判一次就会把扫描页当成"空页"悄悄丢掉。

【坑 1】get_text("text") 不区分段落
  扫描件之外的 PDF，用 "text" 模式拿到的是一行一个 \\n，段落之间的空行
  不一定保留。所以这里用 get_text("dict")——它给的 blocks 就是版面块，
  一个 block 通常正好是一段（或一个表格单元格），比按空行切准得多。

【坑 2】扫描页的文字量不一定是 0
  有些扫描软件会加一层很淡的 OCR 文字层，也有些扫描页只带页码。
  所以判据是「去掉空白后还剩多少个字符」≥ MIN_TEXT_CHARS，而不是「是否为空」。

【坑 3】渲染 dpi 决定识别率
  150 dpi 小字会糊，300 dpi 慢一倍。200 是实测的平衡点：
  A4 单页 ≈ 1654×2339，中文小五号字能认准，单页约 0.5~1.2 秒。
"""

from pathlib import Path

from .base import make_block, load_ocr_module

# 一页去掉空白后少于这么多字符 → 认为是扫描页（没有可用文字层）
MIN_TEXT_CHARS = 20
# 扫描页渲染精度
DEFAULT_DPI = 200
# 单份 PDF 最多认这么多页（防止传错文件把服务卡死；超了会明确告知）
MAX_OCR_PAGES = 120


def _import_fitz():
    try:
        import fitz                                        # type: ignore
        return fitz, None
    except Exception as e:
        return None, ("读 PDF 需要 PyMuPDF，请先装：pip install pymupdf"
                      "（当前错误：%s）" % e)


def _clean(s: str) -> str:
    return (s or "").replace("\u3000", " ").strip()


def _text_chars(s: str) -> int:
    """数有效字符（不算空白）。"""
    return len("".join((s or "").split()))


def _blocks_from_text_layer(page, pno: int, file_name: str):
    """从文字层抽块：用 dict 模式的版面块划分，一块一段。"""
    blocks = []
    d = page.get_text("dict")
    for bi, b in enumerate(d.get("blocks", [])):
        if b.get("type") != 0:                     # type 1 = 图片，跳过
            continue
        lines = []
        sizes = []
        for ln in b.get("lines", []):
            txt = "".join(sp.get("text", "") for sp in ln.get("spans", []))
            txt = _clean(txt)
            if txt:
                lines.append(txt)
            for sp in ln.get("spans", []):
                if sp.get("text", "").strip() and sp.get("size"):
                    sizes.append(float(sp["size"]))
        text = "\n".join(lines).strip()
        if not text:
            continue
        blocks.append(make_block(
            file_name=file_name,
            file_type="pdf",
            text=text,
            locator={
                "kind": "pdf_block",
                "page": pno + 1,
                "block": bi,
                "bbox": [round(float(v), 1) for v in b.get("bbox", [])],
            },
            source="text_layer",
            font_size=(round(sum(sizes) / len(sizes), 1) if sizes else None),
        ))
    return blocks


def _blocks_from_ocr(ocr_res: dict, pno: int, file_name: str):
    """把一页 OCR 结果变成块：一个 OCR 段落 = 一块。"""
    blocks = []
    for i, p in enumerate(ocr_res.get("paragraphs") or []):
        text = (p.get("text") or "").strip()
        if not text:
            continue
        blocks.append(make_block(
            file_name=file_name,
            file_type="pdf",
            text=text,
            locator={
                "kind": "pdf_ocr",
                "page": pno + 1,
                "para": i,
                "conf": p.get("conf"),
            },
            source="ocr",
            ocr_conf=p.get("conf"),
        ))
    return blocks


def read_pdf(path, dpi: int = DEFAULT_DPI, force_ocr: bool = False,
             lang: str = "auto", vertical: str = "auto"):
    """读一个 PDF，返回和其它读取器同构的字典。

    force_ocr=True 时跳过文字层，整份都走 OCR（拿来核对文字层质量用）。
    lang / vertical：扫描页的 OCR 语言与竖排处理（见 agent/ocr.py，2026-11-07）。
    """
    path = Path(path)
    fitz, err = _import_fitz()
    if fitz is None:
        raise RuntimeError(err)

    doc = fitz.open(str(path))
    try:
        page_count = doc.page_count
        blocks = []
        text_pages, ocr_pages, ocr_failed = [], [], []
        ocr_elapsed = 0.0
        notes = []
        eff_dpi = int(dpi) if dpi else DEFAULT_DPI

        # OCR 引擎只在真的遇到扫描页时才加载（电子版 PDF 完全不受影响）
        ocr_mod = None

        for pno in range(page_count):
            page = doc[pno]
            raw = page.get_text("text") if not force_ocr else ""
            chars = _text_chars(raw)

            if chars >= MIN_TEXT_CHARS:
                blocks.extend(_blocks_from_text_layer(page, pno, path.name))
                text_pages.append(pno + 1)
                continue

            # ---- 扫描页：渲染成图再认 ----
            if len(ocr_pages) + len(ocr_failed) >= MAX_OCR_PAGES:
                notes.append("扫描页超过 %d 页，后面的页没认（防止一次请求跑太久）"
                             % MAX_OCR_PAGES)
                break
            if ocr_mod is None:
                ocr_mod = load_ocr_module()          # 延迟导入（只在遇到扫描页时加载）
                # 兜底引擎（Windows OCR）超过约 300 万像素会静默认不出，
                # 所以先把它钳到 160dpi（A4 约 240 万像素），别白渲染大图
                if ocr_mod.engine_now() == "winocr" and eff_dpi > 160:
                    notes.append(
                        "当前用的是 Windows 自带 OCR（像素上限约 300 万），本份按 160dpi 渲染；"
                        "想要 300dpi 的高精度请装 rapidocr_onnxruntime")
                    eff_dpi = 160

            try:
                pix = page.get_pixmap(dpi=eff_dpi)
                data = pix.tobytes("png")
            except Exception as e:
                ocr_failed.append({"page": pno + 1, "error": "渲染失败：%s" % e})
                continue

            r = ocr_mod.read_image_bytes(data, "p%d.png" % (pno + 1),
                                        lang=lang, vertical=vertical)
            if not r.get("ok"):
                ocr_failed.append({"page": pno + 1,
                                   "error": r.get("error") or "OCR 失败"})
                continue
            ocr_elapsed += float(r.get("elapsed") or 0)
            got = _blocks_from_ocr(r, pno, path.name)
            if not got:
                ocr_failed.append({"page": pno + 1,
                                   "error": "这一页没认到文字（可能是空白页或纯图表）"})
                continue
            blocks.extend(got)
            ocr_pages.append({"page": pno + 1, "paras": len(got),
                              "conf": r.get("avg_conf"), "elapsed": r.get("elapsed")})

        # 兜底：整份一个字都没读到
        if not blocks:
            raise ValueError(
                "这份 PDF 里没读到任何文字——%d 页，其中 %d 页走 OCR 也没认出来。"
                "常见原因：图片太糊 / 整页是手写字 / 文件加密。"
                % (page_count, len(ocr_failed)))

        out = {
            "file": path.name,
            "type": "pdf",
            "page_count": page_count,
            "has_text_layer": bool(text_pages),
            "text_pages": text_pages,
            "ocr_pages": ocr_pages,
            "ocr_page_count": len(ocr_pages),
            "ocr_failed": ocr_failed,
            "ocr_elapsed": round(ocr_elapsed, 2),
            "dpi": int(dpi),
            "block_count": len(blocks),
            "blocks": blocks,
        }
        if ocr_pages:
            out["scanned"] = True          # 前端/日志据此标「扫描件·已 OCR」
        if notes:
            out["notes"] = notes
        return out
    finally:
        doc.close()
