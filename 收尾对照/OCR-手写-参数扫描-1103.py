# -*- coding: utf-8 -*-
"""手写参数扫描：RapidOCR 的 text_score / box_thresh / unclip_ratio 怎么选。

本次真正的发现（先看这里再改代码）：
  RapidOCR 1.2.3 的 __call__ 默认 text_score=0.5，
  意思是"识别置信度低于 0.5 的框直接丢掉，连字都不返回"。
  而手写单字的置信度普遍在 0.20~0.55：
    验收单「数量」列的 3 / 2 / 5 → conf 0.50 / 0.44 / 0.20
    报销单「金额」列的 45.00     → conf 0.40 / 0.31
  结果是整格消失，人眼明明看得清。这不是模型不准，是**阈值把结果吞了**。
  所以调参不是玄学，是把"模型已经认出来、但被默认值扔掉"的字捞回来。

用法：python 收尾对照\OCR-手写-参数扫描-1103.py
"""

import itertools
import sys
from pathlib import Path

BASE = Path(__file__).resolve().parents[1]
IMG_DIR = BASE / "测试数据" / "OCR样本" / "手写"
sys.path.insert(0, str(BASE / "收尾对照"))

import ocrmetric as M                                     # noqa: E402
from rapidocr_onnxruntime import RapidOCR                 # noqa: E402

eng = RapidOCR()
rows = []
for ts, bt, ur in itertools.product((0.50, 0.35, 0.20, 0.10),
                                    (0.50, 0.30),
                                    (1.6, 2.0)):
    hit = key = 0
    cer_num = cer_den = 0.0
    boxes = chars = 0
    for name in M.GROUND_TRUTH:
        res = eng(str(IMG_DIR / name), text_score=ts, box_thresh=bt, unclip_ratio=ur)
        items = res[0] if isinstance(res, tuple) else res
        items = items or []
        boxes += len(items)
        text = "".join(str(it[1]) for it in items)
        s = M.score(name, text)
        hit += s["hit"]
        key += s["key"]
        chars += s["chars"]
        exp = len(M.norm(M.FULL_TEXT[name]))
        cer_num += M.lev(M.norm(text), M.norm(M.FULL_TEXT[name]))
        cer_den += exp
    cer = (1 - cer_num / cer_den) * 100
    rows.append("text_score=%.2f box=%.2f unclip=%.1f → 框%3d 字%4d 关键串 %2d/%d (%.0f%%) 字符 %.1f%%"
                % (ts, bt, ur, boxes, chars, hit, key, 100.0 * hit / key, cer))
    print(rows[-1])

out = BASE / "收尾对照" / "_手写参数扫描.txt"
out.write_text("\n".join(rows), encoding="utf-8")
print("\n已写入 %s" % out.name)
