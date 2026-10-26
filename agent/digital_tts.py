# -*- coding: utf-8 -*-
"""数字人「小橘」的声音：文字 → 语音（edge-tts），带本地缓存。

设计取舍：
- 只做「把一段文字读出来」，不做流式；客服一次回复就一两句，整段合成更简单。
- 同一句话第二次直接命中缓存文件，断网也能重放（演示现场最怕的就是网抖）。
- 合成失败不抛给前端 500，而是返回 None，由前端退回「只动嘴不出声」。
"""

from __future__ import annotations

import asyncio
import hashlib
import re
import threading
from pathlib import Path

try:
    import edge_tts
except Exception:                                    # 没装就整体降级
    edge_tts = None

VOICE = "zh-CN-XiaoyiNeural"                         # 卡通/活泼的年轻女声，配小橘猫
RATE = "+10%"                                        # 稍快一点 = 更有精神，不拖
PITCH = "+26Hz"                                      # 提调是"奶音"的关键，再高就假了
MAX_CHARS = 220
CACHE_DIR = Path(__file__).resolve().parent / "digital" / "cache"
_LOCK = threading.Lock()


def clean(text: str) -> str:
    """把气泡里的文字整理成「适合念出来」的句子。"""
    s = str(text or "")
    s = re.sub(r"```.*?```", "，", s, flags=re.S)              # 代码块直接跳过
    s = re.sub(r"[*_#>`~]+", "", s)                            # markdown 记号
    s = re.sub(r"https?://\S+", "链接", s)
    s = re.sub(r"[\U0001F300-\U0001FAFF\u2600-\u27BF]", "", s)  # emoji 不念
    s = re.sub(r"\s+", " ", s).strip()
    if len(s) > MAX_CHARS:
        cut = s[:MAX_CHARS]
        for sep in ("。", "！", "？", "；", "\n"):
            i = cut.rfind(sep)
            if i >= MAX_CHARS * 0.6:
                cut = cut[:i + 1]
                break
        s = cut
    return s


def _key(text: str) -> str:
    raw = "|".join([VOICE, RATE, PITCH, text])
    return hashlib.md5(raw.encode("utf-8")).hexdigest()[:20]


def speak(text: str) -> Path | None:
    """返回 mp3 路径；合成不了就返回 None（前端自会降级）。"""
    s = clean(text)
    if not s:
        return None
    CACHE_DIR.mkdir(parents=True, exist_ok=True)
    out = CACHE_DIR / (_key(s) + ".mp3")
    if out.exists() and out.stat().st_size > 900:
        return out
    if edge_tts is None:
        return None

    # 同一个 key 只让一个人去合成，其他人等结果（避免并发生成同一个文件）
    with _LOCK:
        if out.exists() and out.stat().st_size > 900:
            return out
        tmp = out.with_suffix(".part")
        try:
            async def _run():
                c = edge_tts.Communicate(s, VOICE, rate=RATE, pitch=PITCH)
                await c.save(str(tmp))
            asyncio.run(_run())
            if not tmp.exists() or tmp.stat().st_size < 900:
                tmp.unlink(missing_ok=True)
                return None
            tmp.replace(out)
            return out
        except Exception:
            tmp.unlink(missing_ok=True)
            return None


def stats() -> dict:
    n = len(list(CACHE_DIR.glob("*.mp3"))) if CACHE_DIR.exists() else 0
    size = sum(p.stat().st_size for p in CACHE_DIR.glob("*.mp3")) if CACHE_DIR.exists() else 0
    return {"voice": VOICE, "cached": n, "bytes": size, "ready": edge_tts is not None}
