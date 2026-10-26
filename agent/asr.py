# -*- coding: utf-8 -*-
"""本机离线语音识别（faster-whisper）。

为什么要有它：
    浏览器自带的语音识别（SpeechRecognition）要把麦克风音频送到谷歌 / 微软的服务器，
    国内网络经常连不上，前端只会拿到一句 network error（"语音服务连不上"）。
    这条通道把录音留在本机做识别，换掉那一次外网请求——断网也能用。

模型：models/faster-whisper-small/（由 download_asr_model.py 从 ModelScope 下载）
    461MB，int8 量化后跑 CPU，3~5 秒的短句一般 1~2 秒出结果。
"""
from __future__ import annotations

import os
import re
import threading
import time
from pathlib import Path

HERE = Path(__file__).resolve().parent
MODEL_DIR = HERE / "models" / "faster-whisper-small"

# 中文转写里不该留英文那种词间空格；这几个字符之间才允许保留空格
_CJK = r"\u4e00-\u9fa5\u3000-\u303f\uff00-\uffef"
_CJK_SPACE = re.compile(r"(?<=[%s])\s+(?=[%s])" % (_CJK, _CJK))
# whisper 爱输出半角标点（"金额,这两个"），中文里看着别扭，统一成全角
_HALF2FULL = str.maketrans({",": "，", "?": "？", "!": "！", ";": "；", ":": "："})
_PUNCT_SPACE = re.compile(r"\s*([，。？！；：、])\s*")

_lock = threading.Lock()
_model = None
_load_error = ""


def ready() -> bool:
    """模型文件是否齐（不加载模型，秒回）。"""
    return (MODEL_DIR / "model.bin").is_file() and (MODEL_DIR / "tokenizer.json").is_file()


def info() -> dict:
    """给 /api/asr/status 用的自检信息。"""
    files = {}
    for n in ("model.bin", "config.json", "tokenizer.json", "vocabulary.txt"):
        p = MODEL_DIR / n
        files[n] = round(p.stat().st_size / 1048576, 1) if p.is_file() else None
    return {
        "model_dir": str(MODEL_DIR),
        "file_mb": files,
        "ready": ready(),
        "loaded": _model is not None,
        "error": _load_error,
    }


def _get_model():
    """第一次调用时加载模型（约 2~3 秒），之后复用同一个实例。"""
    global _model, _load_error
    if _model is not None:
        return _model
    with _lock:
        if _model is not None:
            return _model
        if not ready():
            _load_error = "离线语音模型还没下载：在 agent 目录跑 python download_asr_model.py"
            raise RuntimeError(_load_error)
        try:
            from faster_whisper import WhisperModel
            threads = max(2, min(8, (os.cpu_count() or 4) // 2))
            _model = WhisperModel(str(MODEL_DIR), device="cpu",
                                  compute_type="int8", cpu_threads=threads)
            _load_error = ""
        except Exception as e:                       # 依赖缺失 / 模型文件损坏
            _load_error = "离线语音引擎起不来：%s" % e
            raise RuntimeError(_load_error)
        return _model


def warm_up() -> None:
    """后台预加载模型：用户第一次点麦克风时就不用再等那两三秒。失败不吵人。"""
    try:
        _get_model()
    except Exception:
        pass


def transcribe(audio_path: str | Path, language: str = "zh") -> dict:
    """音频文件 → 文字。返回 {text, seconds, duration}。"""
    model = _get_model()
    t0 = time.time()
    segments, meta = model.transcribe(
        str(audio_path),
        language=language,
        beam_size=5,
        vad_filter=True,                              # 掐掉首尾静音，短句更准
        vad_parameters={"min_silence_duration_ms": 300},
        condition_on_previous_text=False,             # 单句场景，别让它自己脑补上下文
        initial_prompt="以下是普通话口语，可能包含文档、表格、标题、格式等词。",
    )
    parts = []
    for s in segments:
        parts.append(s.text)
    text = "".join(parts).strip()
    text = _CJK_SPACE.sub("", text)                   # 去掉中文之间的多余空格
    if len(re.findall(r"[\u4e00-\u9fa5]", text)) >= max(2, len(text) // 3):
        text = text.translate(_HALF2FULL)             # 中文句子里的半角标点 → 全角
    text = _PUNCT_SPACE.sub(r"\1", text)              # 中文标点两边不留空格
    text = re.sub(r"\s{2,}", " ", text).strip()
    return {"text": text, "seconds": round(time.time() - t0, 2),
            "duration": round(float(getattr(meta, "duration", 0) or 0), 2)}
