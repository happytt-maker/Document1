# -*- coding: utf-8 -*-
"""手写漏检定位实验：问题是"没检到"还是"检到了认错"？

同一张图用不同放大倍数跑 RapidOCR，数"框数 / 字数"。
如果放大就大幅回血 → 问题在检测（小字没检到），不必换识别模型。
反过来如果框数不变、字错得多 → 才是识别模型的问题。

用法：python 收尾对照\OCR-手写-放大实验-1103.py
"""
import sys
from pathlib import Path

BASE = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(BASE / "agent"))

from PIL import Image
from rapidocr_onnxruntime import RapidOCR

IMG = BASE / "测试数据" / "OCR样本" / "手写"

eng = RapidOCR()
lines = []
for p in sorted(IMG.glob("*.png")):
    im = Image.open(p)
    w0, h0 = im.size
    for k in (1.0, 1.5, 2.0, 2.5, 3.0):
        if k == 1.0:
            q = p
        else:
            q = p.with_name("_tmp_x%s.png" % k)
            im.resize((int(w0 * k), int(h0 * k)), Image.LANCZOS).save(q)
        res = eng(str(q))
        items = res[0] if isinstance(res, tuple) else res
        n_box = len(items or [])
        n_ch = sum(len(str(it[1])) for it in (items or []))
        lines.append("%s x%.1f  框 %3d  字 %4d" % (p.stem, k, n_box, n_ch))
        if k != 1.0:
            q.unlink(missing_ok=True)
    lines.append("")

out = BASE / "收尾对照" / "_手写放大实验.txt"
out.write_text("\n".join(lines), encoding="utf-8")
print("\n".join(lines))
