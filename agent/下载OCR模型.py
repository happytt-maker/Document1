# -*- coding: utf-8 -*-
"""下载多语种 OCR 识别模型（2026-11-07）

为什么需要单独下：RapidOCR 自带的是 **简体** 识别模型（ch_PP-OCRv3），
「驗收單」会认成「收軍」、「監測」认成「盤測」—— 不是模糊，是字库里就没这些字。
换成 PP-OCRv5 的中文模型后简繁英一个模型全认，简体样本还不掉分。

下载源用 modelscope（国内直连，约 9MB/s；github 在这台机器上 SSL 不通）。

跑法（项目根目录）：
  python agent/下载OCR模型.py            下中文 + 英文两个
  python agent/下载OCR模型.py ch         只下中文（繁体/默认用它）
"""
import sys
import time
from pathlib import Path

import requests

OUT = Path(__file__).resolve().parents[1] / "models" / "ocr"
API = ("https://www.modelscope.cn/api/v1/models/RapidAI/RapidOCR/repo"
       "?Revision=master&FilePath=")

MODELS = {
    "ch": ("onnx/PP-OCRv5/rec/ch_PP-OCRv5_rec_mobile.onnx", "ch_PP-OCRv5_rec_mobile.onnx",
           "中文（简繁英混排）"),
    "en": ("onnx/PP-OCRv5/rec/en_PP-OCRv5_rec_mobile.onnx", "en_PP-OCRv5_rec_mobile.onnx",
           "英文（纯英文文档）"),
}


def fetch(name, remote, local, label):
    dst = OUT / local
    s = requests.Session()
    s.trust_env = False                       # 系统代理会把 127.0.0.1/国内站都劫持
    t = time.time()
    try:
        r = s.get(API + remote, stream=True, timeout=120)
    except Exception as e:
        print("✗ %s：连不上（%s）" % (label, e))
        return False
    if r.status_code != 200:
        print("✗ %s：HTTP %s" % (label, r.status_code))
        return False
    total = int(r.headers.get("content-length") or 0)
    tmp = dst.with_suffix(dst.suffix + ".part")
    n = 0
    with open(tmp, "wb") as f:
        for chunk in r.iter_content(1 << 16):
            if chunk:
                f.write(chunk)
                n += len(chunk)
    tmp.replace(dst)                          # 下完再改名，别留半截文件
    dt = max(time.time() - t, 0.1)
    print("✓ %s → %s（%.1f MB，%.1fs，%.0f kB/s）"
          % (label, dst, n / 1024 / 1024, dt, n / 1024 / dt))
    if total and n != total:
        print("  ⚠ 大小与服务器声明不一致（%d ≠ %d），建议重下" % (n, total))
    return True


def main():
    OUT.mkdir(parents=True, exist_ok=True)
    want = [a.strip() for a in sys.argv[1:] if a.strip()] or ["ch", "en"]
    ok = 0
    for k in want:
        if k not in MODELS:
            print("不认识的型号：%s（可选 %s）" % (k, " / ".join(MODELS)))
            continue
        remote, local, label = MODELS[k]
        if (OUT / local).is_file():
            print("· %s 已存在（%s），跳过" % (label, OUT / local))
            ok += 1
            continue
        if fetch(k, remote, local, label):
            ok += 1
    print("\n完成 %d/%d。识别时会自动启用，无需改代码。" % (ok, len(want)))
    return 0 if ok == len(want) else 1


if __name__ == "__main__":
    sys.exit(main())
