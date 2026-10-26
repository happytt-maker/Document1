# -*- coding: utf-8 -*-
"""下载「本机离线语音识别」的模型文件（faster-whisper-small）。

为什么单独写一个下载器：这台机器 pip 连 PyPI、huggingface 都会被掐（502 / 超时），
而 ModelScope 的国内站点是通的，所以模型从它的镜像仓库拉。

用法（在 agent 目录下）：
    python download_asr_model.py            # 缺什么补什么
    python download_asr_model.py --force    # 全部重下

产物：agent/models/faster-whisper-small/  （约 464MB）
"""
from __future__ import annotations

import json
import os
import ssl
import sys
import time
import urllib.request
from pathlib import Path

REPO = "Systran/faster-whisper-small"
DEST = Path(__file__).resolve().parent / "models" / "faster-whisper-small"
FILES = ["config.json", "tokenizer.json", "vocabulary.txt", "model.bin"]

# 本机所有出网都要绕开代理：代理会把镜像站掐成 502
for _k in ("HTTP_PROXY", "HTTPS_PROXY", "http_proxy", "https_proxy", "ALL_PROXY", "all_proxy"):
    os.environ.pop(_k, None)

_ctx = ssl.create_default_context()
_ctx.check_hostname = False
_ctx.verify_mode = ssl.CERT_NONE
_OP = urllib.request.build_opener(urllib.request.ProxyHandler({}),
                                  urllib.request.HTTPSHandler(context=_ctx))


def remote_sizes() -> dict:
    """向镜像要一份文件清单（名字 → 字节数），顺带确认仓库还在。"""
    url = ("https://www.modelscope.cn/api/v1/models/%s/repo/files?Revision=master&Root=" % REPO)
    j = json.loads(_OP.open(url, timeout=30).read().decode("utf-8"))
    return {f["Name"]: int(f.get("Size") or 0) for f in (j.get("Data", {}).get("Files") or [])}


def fetch(name: str, size: int) -> Path:
    dst = DEST / name
    url = "https://www.modelscope.cn/models/%s/resolve/master/%s" % (REPO, name)
    t0 = time.time()
    with _OP.open(url, timeout=180) as r, open(dst, "wb") as f:
        total = int(r.headers.get("Content-Length") or size or 0)
        done = 0
        while True:
            chunk = r.read(1 << 20)
            if not chunk:
                break
            f.write(chunk)
            done += len(chunk)
            if total:
                pct = done * 100 // total
                sys.stdout.write("\r  %-16s %3d%%  %6.1fMB / %.1fMB" %
                                 (name, pct, done / 1048576, total / 1048576))
                sys.stdout.flush()
    sys.stdout.write("  (%.1fs)\n" % (time.time() - t0))
    return dst


def main() -> int:
    force = "--force" in sys.argv
    DEST.mkdir(parents=True, exist_ok=True)
    print("目标目录：%s" % DEST)
    try:
        sizes = remote_sizes()
    except Exception as e:
        print("取不到文件清单：%s（检查网络）" % e)
        return 1

    ok = True
    for name in FILES:
        want = sizes.get(name, 0)
        p = DEST / name
        if not force and p.is_file() and want and abs(p.stat().st_size - want) < 1024:
            print("  %-16s 已就绪 %.1fMB" % (name, p.stat().st_size / 1048576))
            continue
        if not want:
            print("  %-16s 镜像上没有这个文件" % name)
            ok = False
            continue
        print("  %-16s 下载中 %.1fMB" % (name, want / 1048576))
        fetch(name, want)
        if abs(p.stat().st_size - want) > 1024:
            print("  !! %s 大小不对" % name)
            ok = False

    print("完成。" if ok else "有文件没拿全，重跑一次即可。")
    return 0 if ok else 2


if __name__ == "__main__":
    sys.exit(main())
