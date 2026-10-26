# -*- coding: utf-8 -*-
"""在没有外网（或 pip 被代理掐成 502）的机器上，装好离线语音识别的依赖。

背景：这台机器 pip 直连 PyPI 超时、走本地代理返回 502 Bad Gateway，但直连国内镜像是通的。
所以这里不走 pip 的索引，而是自己从镜像把 whl 抓下来，再让 pip 从本地目录装。

用法（在 agent 目录下）：
    python install_asr_deps.py            # 缺什么装什么
    python install_asr_deps.py --check    # 只检查依赖齐不齐

装完之后再跑 python download_asr_model.py 下模型（461MB）。
"""
from __future__ import annotations

import os
import re
import ssl
import subprocess
import sys
import time
import urllib.parse
import urllib.request
from pathlib import Path

MIRROR = "https://pypi.tuna.tsinghua.edu.cn/simple/%s/"
HERE = Path(__file__).resolve().parent
WHEELS = HERE / ".wheels"

PY_TAG = "cp%d%d" % sys.version_info[:2]
PLAT = "win_amd64"

# faster-whisper 及其依赖（间接依赖也列上，因为要一次装齐）
PKGS = ["faster-whisper", "ctranslate2", "huggingface_hub", "tokenizers", "onnxruntime",
        "av", "tqdm", "numpy", "filelock", "fsspec", "packaging", "pyyaml", "requests",
        "charset-normalizer", "idna", "urllib3", "certifi", "protobuf", "flatbuffers",
        "sympy", "mpmath", "typing-extensions", "coloredlogs", "humanfriendly",
        "setuptools", "colorama"]
# tokenizers 0.20 只认 huggingface-hub<1.0；不锁的话会抓到 2.x 然后依赖冲突
PIN = {"huggingface_hub": "1.0"}          # 上界（不含）

LINK = re.compile(r'href="([^"#]+)[^"]*"[^>]*>([^<]+)</a>')
PRE = re.compile(r"(rc|a|b|dev|post)\d*$", re.I)

for _k in ("HTTP_PROXY", "HTTPS_PROXY", "http_proxy", "https_proxy", "ALL_PROXY", "all_proxy"):
    os.environ.pop(_k, None)

_ctx = ssl.create_default_context()
_ctx.check_hostname = False
_ctx.verify_mode = ssl.CERT_NONE
_OP = urllib.request.build_opener(urllib.request.ProxyHandler({}),
                                  urllib.request.HTTPSHandler(context=_ctx))


def tag_ok(pylist: str, abi: str, plat: str) -> bool:
    """这个 wheel 能不能装在本机解释器上。"""
    if plat not in ("any", PLAT):
        return False
    if abi.startswith("cp") and abi.endswith("t"):     # cp313t = free-threaded，装不了
        return False
    tags = pylist.split(".")
    if "cp313t" in tags:
        return False
    if abi == "none":
        return plat == "any" or PY_TAG in tags
    if abi == "abi3":
        return any(t.startswith("cp3") and t[2:].isdigit() and int(t[2:]) <= 13 for t in tags)
    return abi == PY_TAG and PY_TAG in tags


def vkey(v: str) -> tuple:
    return tuple(int(x) for x in re.findall(r"\d+", v))[:6]


def pick(name: str):
    """挑该包最新、且本机能装的正式版 wheel。返回 (版本, 文件名, 下载地址)。"""
    base = MIRROR % name
    html = _OP.open(base, timeout=30).read().decode("utf-8", "ignore")
    cap = PIN.get(name)
    best = None
    for href, fn in LINK.findall(html):
        fn = fn.strip()
        if not fn.endswith(".whl") or "+" in fn:
            continue
        parts = fn[:-4].split("-")
        if len(parts) < 5 or PRE.search(parts[1]):
            continue
        if cap and vkey(parts[1]) >= vkey(cap):
            continue
        if not tag_ok(parts[-3], parts[-2], parts[-1]):
            continue
        if best is None or vkey(parts[1]) > best[0]:
            best = (vkey(parts[1]), fn, urllib.parse.urljoin(base, href))
    return best


def check() -> bool:
    try:
        import faster_whisper  # noqa: F401
        import ctranslate2     # noqa: F401
        import av              # noqa: F401
        print("依赖齐全（faster-whisper 可用）")
        return True
    except Exception as e:
        print("还缺：%s" % e)
        return False


def main() -> int:
    if "--check" in sys.argv:
        return 0 if check() else 1

    WHEELS.mkdir(exist_ok=True)
    print("解释器：%s" % sys.executable)
    print("缓存目录：%s" % WHEELS)

    for name in PKGS:
        try:
            hit = pick(name)
        except Exception as e:
            print("  !! %-20s 取不到索引：%s" % (name, e))
            continue
        if not hit:
            print("  !! %-20s 没有兼容 %s/%s 的 wheel" % (name, PY_TAG, PLAT))
            continue
        _k, fn, url = hit
        dst = WHEELS / fn
        if dst.exists() and dst.stat().st_size > 1000:
            print("  -- %-20s 已有 %s" % (name, fn))
            continue
        t0 = time.time()
        data = _OP.open(url, timeout=180).read()
        dst.write_bytes(data)
        print("  ++ %-20s %-52s %.1fMB %.1fs" % (name, fn, len(data) / 1048576, time.time() - t0))

    print("\n开始本地安装……")
    r = subprocess.run([sys.executable, "-m", "pip", "install", "--no-index",
                        "--find-links", str(WHEELS), "faster-whisper"],
                       text=True, encoding="utf-8", errors="replace")
    if r.returncode == 0:
        print("\n装好了。下一步：python download_asr_model.py")
        return 0
    print("\n安装失败：看上面 pip 的报错，多半是某个包还缺，重跑一次本脚本即可补齐。")
    return r.returncode


if __name__ == "__main__":
    sys.exit(main())
