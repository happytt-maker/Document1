# -*- coding: utf-8 -*-
"""下载 bge-small-zh-v1.5（ONNX 版）到 models/bge-small-zh-v1.5/。

为什么用 Xenova 这个镜像仓库：官方 BAAI 仓库只有 pytorch 权重，
要转 ONNX 得装 torch（2GB+）。Xenova 仓库已经把 ONNX 导好了。

下哪几个文件：
    onnx/model.onnx          fp32，约 95MB（最终采用：精度稳）
    tokenizer.json           WordPiece 词表 + 规则（tokenizers 库直接加载）
    tokenizer_config.json    上下文长度等
    config.json              hidden size 等
支持断点续传（删了重跑即可）。
"""
import sys
import time
from pathlib import Path

import requests

REPO = "Xenova/bge-small-zh-v1.5"
DEST = Path(__file__).resolve().parents[1] / "models" / "bge-small-zh-v1.5"
FILES = [
    ("onnx/model.onnx", "model.onnx"),
    ("tokenizer.json", "tokenizer.json"),
    ("tokenizer_config.json", "tokenizer_config.json"),
    ("special_tokens_map.json", "special_tokens_map.json"),
    ("config.json", "config.json"),
    ("vocab.txt", "vocab.txt"),
]
BASES = ["https://hf-mirror.com", "https://huggingface.co"]

S = requests.Session()
S.trust_env = False      # 绕开环境里的坏代理
S.proxies = {}


def fetch(rel, local_name):
    dest = DEST / local_name
    if dest.exists() and dest.stat().st_size > 0:
        return "skip  %-22s %d bytes" % (local_name, dest.stat().st_size)

    last = ""
    for base in BASES:
        url = "%s/%s/resolve/main/%s" % (base, REPO, rel)
        try:
            r = S.get(url, timeout=60, stream=True)
            if r.status_code != 200:
                last = "HTTP %s @ %s" % (r.status_code, base)
                continue
            tmp = dest.with_suffix(dest.suffix + ".part")
            n = 0
            t0 = time.time()
            with open(tmp, "wb") as f:
                for chunk in r.iter_content(1024 * 512):
                    if chunk:
                        f.write(chunk)
                        n += len(chunk)
            tmp.replace(dest)
            return "ok    %-22s %d bytes  %.1fs" % (local_name, n, time.time() - t0)
        except Exception as e:
            last = "%s @ %s" % (type(e).__name__, base)
    return "FAIL  %-22s %s" % (local_name, last)


def main():
    DEST.mkdir(parents=True, exist_ok=True)
    lines = []
    for rel, name in FILES:
        msg = fetch(rel, name)
        print(msg)
        lines.append(msg)
    (DEST / "_download_log.txt").write_text("\n".join(lines), encoding="utf-8")
    total = sum(p.stat().st_size for p in DEST.iterdir() if p.is_file())
    print("total %.1f MB" % (total / 1024 / 1024))


if __name__ == "__main__":
    sys.exit(main())
