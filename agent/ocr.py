# -*- coding: utf-8 -*-
"""
ocr.py —— 图片 / 扫描页的文字识别（OCR）

一句话流程：
  一张图（或 PDF 的某一页渲染成的图）→ 摆正（EXIF 方向）→
  找到文字框 → 按「视觉行」聚拢 → 同一行横向离得远的分成不同列 →
  再按行距切成段落 → 输出「全文 + 每段的来源定位（行号 / 坐标 / 置信度）」

为什么不用「整图丢给大模型」：
  · 大模型要联网、要 key、按量计费，比赛环境不一定有网；
  · OCR 是确定性计算，同样的图每次结果一样，可复现、可验收；
  · 赛题要「全量原文 + 来源定位」，OCR 能给出每个字在页面上的坐标。

两个引擎（自动选，也可用环境变量 A23_OCR_ENGINE 指定 rapidocr / winocr）：
  1) RapidOCR（PP-OCR + onnxruntime，纯 CPU 离线）——首选，中文准确率高
     依赖：pip install rapidocr_onnxruntime
  2) Windows 自带 OCR（Windows.Media.Ocr）——兜底，零下载
     只要系统装了简体中文语言包就能用；准确率略低，但不用连网装环境

手写（2026-11-03 补）：
  原来只会认印刷体，手写基本放弃。查下来**不是模型不行，是三件事**：
    ① RapidOCR 默认 text_score=0.5 —— 置信度低于 0.5 的结果直接丢，连字都不给。
       手写单字普遍 0.20~0.55（验收单「数量」列 3/2/5 = 0.50/0.44/0.20），
       于是整格凭空消失，人眼却看得清 → 改成"低分也收下，标成低置信"。
    ② 页边竖排 / 手机拍歪 —— 正向检测器认不出。→ 再跑 ±90° 两遍，按重叠去重合并。
    ③ 笔画细 —— 放大 1.5 倍再认，笔画分得开。
  这三件事只对"看起来像手写的页"做（低置信行占比 ≥30% 才触发），
  印刷体文档的耗时和行为一个字不变 —— 见 _escalate()。

繁体 / 英文 / 竖排（2026-11-07 补）：
  · 繁体：原来那个 ch_PP-OCRv3 是**简体字库**，「驗收單」会认成「收軍」、「監測」认成「盤測」。
    换成 PP-OCRv5 的中文模型（一套简繁英混合字表），同一个模型三种都认，简体还不掉分。
    模型放 models/ocr/，缺了就用自带的 v3（agent/下载OCR模型.py 可补）。
  · 英文：PP-OCRv5 也认英文；纯英文文档可以切 en 专用模型（更准，但**不认汉字**）。
  · 竖排：难点是顺序不是认字 —— 一列竖字常被检成一个细长框，版面还原是横的，
    两列会被拼成一行、"右起第一列"跑到后面。见 _vertical_pass()：
    同列多框拼整 + 整行竖排按从右到左。字本身是正立的，**别转 90°**（转了反而认不出）。

对外五个：
  status()              引擎状态（前端状态灯 / 启动日志）
  warmup()              预热（服务启动后调，第一次点按钮就不用等）
  read_image(path)      一张图 → 文字 + 行 + 段（lang=auto/chs/cht/en, vertical=auto/on/off）
  read_image_bytes(b)   内存里的图（上传链路用）
  engine_now()          当前实际在用的是哪个引擎
"""

import io
import json
import math
import os
import re
import shutil
import subprocess
import sys
import tempfile
import time
from pathlib import Path

# ==================== 配置 ====================

# 同一视觉行里，两个文字框的水平间隙超过「行高 × 这个倍数」→ 认为是不同列
COL_GAP_RATIO = 1.25
# 相邻两行的垂直间距超过「行高 × 这个倍数」→ 断段
PARA_GAP_RATIO = 1.75
# 置信度低于这个值 → 标记为「可能识别有误」，抽取时宁可不填
LOW_CONF = 0.60
# 图片长边超过这个像素数就先缩小（再大对准确率没帮助，只烧时间）
MAX_SIDE = 3200
# 纯像素上限（按引擎）：Windows 自带 OCR 超过约 300 万像素会「静默返回空」——
# 不报错、不给提示，就是认不出字。实测 3.87M 全空、3.13M 只认出一半、
# 2.48M 正常，所以这里压到 240 万（A4 ≈ 160dpi）保平安。
MAX_PIXELS = {"winocr": 2400000, "rapidocr": None}

# ---- 手写升级参数（2026-11-03）----
# 引擎默认值（= 11-03 之前的行为）。这几个数必须显式传：见 _run_rapidocr 的说明，
# 引擎自己会把阈值记在实例上不还原，不显式传就会互相污染。
DEFAULT_PARAMS = {"text_score": 0.50, "box_thresh": 0.50, "unclip_ratio": 1.6}
# 手写模式下放宽的三个阈值。数字都有实测依据，别随手改：
#   text_score  0.25：手写单字置信度 0.20~0.55，0.5 会把一半直接丢掉
#   box_thresh  0.35：连笔字的检测框得分偏低，0.5 会漏整行
#   unclip     1.80：手写笔画外扩得比印刷体多，否则裁掉笔画末端
HAND_PARAMS = {"text_score": 0.25, "box_thresh": 0.35, "unclip_ratio": 1.8}
# 放大倍数：手写笔画细，放大后检测框更稳（实测 1.5 最好，2.0 只会更慢）
HAND_SCALE = 1.5
# 长边已经超过这么多像素就别再放大 —— 原件本身够清晰，放大只烧时间不涨分。
# 实测：A4@200dpi（1654×2339）放大到 1.5 倍变成 870 万像素，
# 每页 2.4 秒涨到 5.5 秒，识别结果没变；而 1400×1100 这种"截图级"的小图放大后确实回血。
HAND_SCALE_MAX_SIDE = 1700
# 需要补跑 ±90° 的朝向
HAND_ANGLES = (90, -90)

# ---- 多语种 / 竖排（2026-11-07）----
# 为什么换识别模型就能认繁体：原来用的 ch_PP-OCRv3_rec 是**简体**字库，
# 碰到「驗收單」会认成「收軍」、「監測」认成「盤測」——不是模糊，是字库里就没这几个字。
# PP-OCRv5 的中文模型是一套「简繁英」混合字表，同一个模型三种都能认，
# 实测简体样本不掉分（置信度反而更高），所以默认用它。
_PROJ = Path(__file__).resolve().parents[1]
OCR_MODEL_DIR = Path(os.environ.get("A23_OCR_MODEL_DIR") or (_PROJ / "models" / "ocr"))

# 语言 → 识别模型文件。auto/chs/cht 都走中文模型（它本来就会认繁体和英文），
# 只有"明确只认英文"时才换成 en 模型：en 模型**不认汉字**，混排文档别用。
REC_MODELS = {
    "auto": ("ch_PP-OCRv5_rec_mobile.onnx", "PP-OCRv5 中文（简繁英混排，默认）"),
    "chs": ("ch_PP-OCRv5_rec_mobile.onnx", "PP-OCRv5 中文（简体优先）"),
    "cht": ("ch_PP-OCRv5_rec_mobile.onnx", "PP-OCRv5 中文（繁体识别）"),
    "en": ("en_PP-OCRv5_rec_mobile.onnx", "PP-OCRv5 英文（纯英文文档）"),
}
LANG_LABEL = {"auto": "自动", "chs": "简体", "cht": "繁体", "en": "英文"}

# 竖排判定：框的高 ÷ 宽 超过这个倍数，就当成"竖着排的一列字"
VERT_RATIO = 1.5
# 整页竖排的门槛：竖排框里的字数占全部字数的比例
VERT_FULL_SHARE = 0.55


def _norm_lang(lang) -> str:
    v = (lang or "").strip().lower()
    if v in ("", "auto", "自动"):
        return "auto"
    if v in ("chs", "zh", "zh-cn", "cn", "简体", "简"):
        return "chs"
    if v in ("cht", "zh-tw", "zh-hk", "tw", "繁体", "繁", "hk"):
        return "cht"
    if v in ("en", "english", "英", "英文", "latin"):
        return "en"
    return "auto"


def _rec_model(lang="auto"):
    """(模型路径 或 None, 说明)。文件不在就返回 None → 用 RapidOCR 自带模型。"""
    want = (os.environ.get("A23_OCR_REC") or "").strip().lower()
    if want in ("v3", "old", "0"):                      # 对照用：强制退回旧模型
        return None, "PP-OCRv3 中文（旧模型，A23_OCR_REC=v3）"
    name, label = REC_MODELS.get(_norm_lang(lang), REC_MODELS["auto"])
    p = OCR_MODEL_DIR / name
    if p.is_file():
        return p, label
    # 英文模型没下就用中文模型顶着（中文模型也认英文，只是不分大小写地略差）
    if name != REC_MODELS["auto"][0] and (OCR_MODEL_DIR / REC_MODELS["auto"][0]).is_file():
        return OCR_MODEL_DIR / REC_MODELS["auto"][0], REC_MODELS["auto"][1]
    return None, "PP-OCRv3 中文（多语种模型未下载，可用 agent/下载OCR模型.py 补齐）"


_PS1 = Path(__file__).resolve().parent / "win_ocr.ps1"

# 当前引擎："" 未定 / "rapidocr" / "winocr" / "none"
_ENGINE = ""
_ENGINE_OBJ = None
_ENGINE_ERR = ""
_ENGINE_TRIED = False
# 识别模型 → RapidOCR 实例（多语种：每个模型一台，第一次用到才加载）
_ENGINES = {}
_ENGINE_LOAD_ERR = {}


# ==================== 一、引擎选择 ====================

def _new_rapidocr(rec_path=None):
    """造一个 RapidOCR（可指定识别模型）。返回 (实例, 错误说明)。"""
    errs = []
    kw = {}
    if rec_path is not None:
        kw["rec_model_path"] = str(rec_path)
    for mod in ("rapidocr_onnxruntime", "rapidocr"):
        try:
            m = __import__(mod, fromlist=["RapidOCR"])
            return m.RapidOCR(**kw), ""
        except Exception as e:
            errs.append("%s: %s" % (mod, e))
            if kw:                                     # 指定模型失败时退回默认模型再试一次
                try:
                    m = __import__(mod, fromlist=["RapidOCR"])
                    return m.RapidOCR(), "指定模型加载失败（%s），已用默认模型" % e
                except Exception as e2:
                    errs.append("%s(默认): %s" % (mod, e2))
    return None, " ｜ ".join(errs)


def _engine_for(lang="auto"):
    """按语言取识别引擎（同模型复用同一个实例，别每次重建 —— 加载要 1s+）。"""
    key = _norm_lang(lang)
    path, label = _rec_model(key)
    cache_key = str(path) if path else ""
    if cache_key in _ENGINES:
        return _ENGINES[cache_key]
    obj, err = _new_rapidocr(path)
    if obj is None:
        raise RuntimeError("RapidOCR 加载失败：%s" % err)
    _ENGINES[cache_key] = obj
    if err:
        _ENGINE_LOAD_ERR[cache_key] = err
    return obj


def _try_rapidocr():
    """优先 RapidOCR 1.x（rapidocr_onnxruntime），再试 2.x（rapidocr）。"""
    obj, err = _new_rapidocr(_rec_model("auto")[0])
    if obj is not None:
        return obj, "rapidocr", err
    return None, "", err


def _winocr_ok():
    """Windows 自带 OCR 能不能用（只看这座机器上有没有 zh-Hans 语言包）。"""
    try:
        ps = ("$null=[Windows.Media.Ocr.OcrEngine,Windows.Foundation,ContentType=WindowsRuntime];"
              "$l=[Windows.Media.Ocr.OcrEngine]::AvailableRecognizerLanguages;"
              "if($l.Count -gt 0){'YES'}else{'NO'}")
        r = subprocess.run(["powershell", "-NoProfile", "-NonInteractive", "-Command", ps],
                           capture_output=True, timeout=30)
        return b"YES" in (r.stdout or b"")
    except Exception:
        return False


def _pick_engine():
    """定下用哪个引擎（只定一次）。"""
    global _ENGINE, _ENGINE_OBJ, _ENGINE_ERR, _ENGINE_TRIED
    if _ENGINE_TRIED:
        return _ENGINE
    _ENGINE_TRIED = True

    want = (os.environ.get("A23_OCR_ENGINE") or "auto").strip().lower()
    rapid_err = ""

    if want in ("auto", "rapidocr"):
        obj, name, err = _try_rapidocr()
        if obj is not None:
            _ENGINE, _ENGINE_OBJ = "rapidocr", obj
            return _ENGINE
        rapid_err = err
        if want == "rapidocr":
            _ENGINE = "none"
            _ENGINE_ERR = ("没装 RapidOCR。装一下：pip install rapidocr_onnxruntime"
                           "（详细原因：%s）" % err)
            return _ENGINE

    # 退到 Windows 自带 OCR
    if want in ("auto", "winocr"):
        if sys.platform.startswith("win") and _winocr_ok():
            _ENGINE = "winocr"
            return _ENGINE
        if want == "winocr":
            _ENGINE = "none"
            _ENGINE_ERR = "这台机器的 Windows OCR 不可用（需要装简体中文语言包）"
            return _ENGINE

    _ENGINE = "none"
    _ENGINE_ERR = (
        "没有可用的 OCR 引擎。两条路选一条：\n"
        "  ① 装 RapidOCR（推荐，准确率高）：pip install rapidocr_onnxruntime\n"
        "  ② 用 Windows 自带 OCR：在 设置 → 时间和语言 → 语言 里加「简体中文」\n"
        "（RapidOCR 失败原因：%s）" % (rapid_err or "未尝试"))
    return _ENGINE


def engine_now() -> str:
    return _pick_engine()


def status() -> dict:
    """引擎状态（给前端状态灯 / 启动日志 / 验收脚本用）。"""
    eng = _pick_engine()
    model_path, model_label = _rec_model("auto")
    out = {
        "ok": eng != "none",
        "engine": eng,
        "low_conf_threshold": LOW_CONF,
        # 手写：关掉开关时如实说"没开"，免得看界面的人以为手写能力没了
        "handwriting": (os.environ.get("A23_OCR_HAND") or "auto").strip().lower() != "0",
        # 多语种 / 竖排（2026-11-07）
        "rec_model": model_label,
        "langs": [{"key": k, "label": LANG_LABEL[k],
                   "model": _rec_model(k)[1],
                   "ready": _rec_model(k)[0] is not None or k in ("auto", "chs", "cht")}
                  for k in ("auto", "chs", "cht", "en")],
        "vertical": True,
        "vertical_tip": "竖排自动还原：同一列的多个框拼整，整页竖排按「从上到下、从右到左」排",
    }
    if out["handwriting"]:
        out["handwriting_tip"] = ("手写增强：低分字不被丢弃 + 小图放大 1.5 倍 + "
                                  "页边竖排补读（A23_OCR_HAND=0 可关）")
    if model_path is None:
        out["rec_tip"] = "多语种模型未下载 → 现在用 RapidOCR 自带简体模型（繁体易认错）"
    if eng == "rapidocr":
        out["tip"] = "RapidOCR（PP-OCR · 本机离线）"
    elif eng == "winocr":
        out["tip"] = "Windows 自带 OCR（兜底引擎；想更准可装 rapidocr_onnxruntime）"
    else:
        out["tip"] = _ENGINE_ERR
    try:
        import fitz                                        # noqa: F401
        out["pdf"] = True
    except Exception:
        out["pdf"] = False
        out["pdf_tip"] = "没装 PyMuPDF，PDF 拆页不可用：pip install pymupdf"
    return out


def warmup() -> bool:
    """预热：把模型读进内存（服务启动后调，第一次点按钮就不用等）。"""
    if _pick_engine() != "rapidocr":
        return True
    tmpdir = tempfile.mkdtemp(prefix="a23warm_")
    try:
        from PIL import Image
        p = Path(tmpdir) / "warm.png"
        Image.new("RGB", (320, 120), (255, 255, 255)).save(str(p), format="PNG")
        _run_rapidocr(p)
        return True
    except Exception:
        return False
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)


# ==================== 二、图片准备（摆正 / 转 PNG / 限尺寸） ====================

def _prepare(src_bytes=None, src_path=None):
    """把各种来源的图统一成「摆正的 PNG 临时文件」。

    做三件事：
      1) 按 EXIF 方向转正——手机拍的照片 90% 带旋转标记，cv2.imread 不认，
         不转正 OCR 出来就是躺着的（这是最常见的"认不出"原因）
      2) 统一成 RGB PNG——绕过 webp/16 位 PNG 这类引擎不支持的花样格式
      3) 长边超过 MAX_SIDE 就等比缩小

    返回 (png_path, width, height, tmpdir)，用完调 _cleanup()。
    """
    from PIL import Image, ImageOps

    if src_path is not None:
        im = Image.open(str(src_path))
    else:
        im = Image.open(io.BytesIO(src_bytes))
    im.load()

    try:
        im = ImageOps.exif_transpose(im)
    except Exception:
        pass
    if im.mode not in ("RGB", "L"):
        im = im.convert("RGB")

    w, h = im.size
    if max(w, h) > MAX_SIDE:
        k = MAX_SIDE / float(max(w, h))
        im = im.resize((max(1, int(w * k)), max(1, int(h * k))), Image.LANCZOS)

    # 引擎的像素上限（Windows OCR 超了会静默返回空）
    limit = MAX_PIXELS.get(_ENGINE)
    if limit:
        w, h = im.size
        if w * h > limit:
            k = (limit / float(w * h)) ** 0.5
            im = im.resize((max(1, int(w * k)), max(1, int(h * k))), Image.LANCZOS)

    tmpdir = tempfile.mkdtemp(prefix="a23ocr_")
    p = Path(tmpdir) / "page.png"
    im.save(str(p), format="PNG")
    return p, im.size[0], im.size[1], tmpdir


def _tmp_png(data: bytes, name: str = "x.png"):
    """把内存里的字节写成临时 PNG，返回 (path, tmpdir)。"""
    tmpdir = tempfile.mkdtemp(prefix="a23ocr_")
    p = Path(tmpdir) / name
    p.write_bytes(data)
    return p, tmpdir


# ==================== 三、跑引擎（统一成 [(box, text, score), ...]） ====================

def _run_rapidocr(png_path, params=None, lang="auto"):
    """跑一遍 RapidOCR。params 用来临时放宽 text_score 等阈值（手写模式）。

    踩过的坑（2026-11-03）：1.2.3 的 __call__ 收到 kwargs 后会把
    `self.text_score`、`self.text_detector.postprocess_op.box_thresh` **直接改在实例上，
    而且不还原**。于是"跑过一遍手写"之后，同一进程里后续所有页面都会继承放宽的阈值，
    印刷体悄悄多出一堆低分噪声 —— 表现为"关掉开关也不生效"，极难查。
    所以这里每次都把三个参数**显式传全**，让每次调用自成一体、结果可复现。

    2.x 的对象方法不一定收这些参数，抛 TypeError 就退回不带参调用。
    """
    eng = _engine_for(lang)
    if eng is None:
        raise RuntimeError("RapidOCR 没加载成功")
    p = dict(DEFAULT_PARAMS)
    p.update(params or {})
    try:
        res = eng(str(png_path), **p)
    except TypeError:
        res = eng(str(png_path))

    # 1.x：(结果列表, 耗时)；结果项 = [box, text, score]
    if isinstance(res, tuple):
        out = []
        for it in (res[0] or []):
            try:
                out.append((it[0], str(it[1] or ""), float(it[2] or 0)))
            except (IndexError, TypeError):
                continue
        return out

    # 2.x：对象带 boxes / txts / scores
    boxes = getattr(res, "boxes", None)
    txts = getattr(res, "txts", None)
    scores = getattr(res, "scores", None)
    if txts is None and isinstance(res, (list, tuple)):
        txts = res
    out = []
    for i in range(len(txts or [])):
        box = boxes[i] if boxes is not None and i < len(boxes) else None
        sc = scores[i] if scores is not None and i < len(scores) else 0.0
        out.append((box, str(txts[i] or ""), float(sc or 0)))
    return out


def _run_winocr(png_path):
    """子进程跑 PowerShell（结果写 JSON 再读，绕开中文编码问题）。"""
    if not _PS1.is_file():
        raise RuntimeError("缺 agent/win_ocr.ps1")
    tmpdir = tempfile.mkdtemp(prefix="a23win_")
    out_json = Path(tmpdir) / "r.json"
    try:
        cmd = ["powershell", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass",
               "-File", str(_PS1), "-Path", str(png_path), "-Out", str(out_json)]
        r = subprocess.run(cmd, capture_output=True, timeout=180)
        if not out_json.is_file():
            err = (r.stderr or b"").decode("utf-8", "replace")[:300]
            raise RuntimeError("Windows OCR 没返回结果：%s" % (err or "未知原因"))
        d = json.loads(out_json.read_text(encoding="utf-8-sig"))
        if not d.get("ok"):
            raise RuntimeError(d.get("error") or "Windows OCR 失败")
        items = []
        for ln in d.get("lines") or []:
            for wd in ln.get("words") or []:
                txt = str(wd.get("t") or "").strip()
                if not txt:
                    continue
                x = float(wd.get("x") or 0)
                y = float(wd.get("y") or 0)
                w = float(wd.get("w") or 0)
                h = float(wd.get("h") or 0)
                box = [[x, y], [x + w, y], [x + w, y + h], [x, y + h]]
                items.append((box, txt, None))     # 系统 OCR 不给置信度
        return items
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)


def _run_engine(png_path, lang="auto", params=None):
    eng = _pick_engine()
    if eng == "rapidocr":
        return _run_rapidocr(png_path, params=params, lang=lang)
    if eng == "winocr":
        return _run_winocr(png_path)
    raise RuntimeError(_ENGINE_ERR or "OCR 引擎不可用")


# ==================== 三·五、手写升级（2026-11-03） ====================
#
# 思路：不做"一上来就重活"，而是"先便宜地认一遍，看结果像不像手写再决定加码"。
# 印刷体文档永远走不到这里，所以原有耗时/准确率一个字不变。

def _rotate_png(png_path, angle, tmpdir):
    """把图转 angle 度存成新 PNG，返回 (新路径, 原尺寸, 新尺寸)。"""
    from PIL import Image

    im = Image.open(str(png_path))
    w0, h0 = im.size
    r = im.rotate(angle, resample=Image.BICUBIC, expand=True,
                  fillcolor=(255, 255, 255))
    out = Path(tmpdir) / ("rot%d.png" % angle)
    r.save(str(out), format="PNG")
    return out, (w0, h0), r.size


def _box_back(box, angle, size0, size1):
    """把「旋转图里的框」映射回「原图坐标」。

    PIL 的 rotate(正数) 是逆时针，expand=True 后会换画布。
    这里用同一个仿射的正向公式推逆变换，别凭直觉写 —— 写反了框会跑到对角去
    （验收脚本 OCR-手写-验收-1103.py 里有一条专门核对这个）。
    """
    if angle == 0:
        return box
    rad = math.radians(angle)
    cos_a, sin_a = math.cos(rad), math.sin(rad)
    w0, h0 = size0
    w1, h1 = size1
    out = []
    for x, y in box:
        dx, dy = float(x) - w1 / 2.0, float(y) - h1 / 2.0
        out.append([cos_a * dx - sin_a * dy + w0 / 2.0,
                    sin_a * dx + cos_a * dy + h0 / 2.0])
    return out


def _box_xyxy(box):
    try:
        xs = [float(p[0]) for p in box]
        ys = [float(p[1]) for p in box]
        return min(xs), min(ys), max(xs), max(ys)
    except Exception:
        return 0.0, 0.0, 0.0, 0.0


def _iou(a, b):
    ax0, ay0, ax1, ay1 = _box_xyxy(a)
    bx0, by0, bx1, by1 = _box_xyxy(b)
    ix = max(0.0, min(ax1, bx1) - max(ax0, bx0))
    iy = max(0.0, min(ay1, by1) - max(ay0, by0))
    inter = ix * iy
    if inter <= 0:
        return 0.0
    area = (ax1 - ax0) * (ay1 - ay0) + (bx1 - bx0) * (by1 - by0) - inter
    return inter / area if area > 0 else 0.0


def _merge_items(groups, iou_thr=0.45):
    """多遍（不同朝向 / 不同阈值）的结果合并。

    规则：
      · 两框重叠超过 iou_thr → 认为是同一处文字，留得分高的那条；
        得分一样时留字多的（"1280.00" 比 "11" 可信）。
      · 不重叠 → 都留下（这正是我们要的：正向漏掉的竖排便条会补进来）。
    为什么不用"按文字去重"：同一个人名在同一页出现两次很常见，按文字去重会误杀。
    """
    flat = []
    for items in groups:
        for box, text, score in items:
            flat.append((box, text, score if score is not None else 0.0))
    # 先按可信度从高到低，保证"留下的是好的那条"
    flat.sort(key=lambda it: (it[2], len(it[1] or "")), reverse=True)

    kept = []
    for box, text, score in flat:
        if not (text or "").strip():
            continue
        dup = False
        for kb, _, _ in kept:
            if _iou(box, kb) >= iou_thr:
                dup = True
                break
        if not dup:
            kept.append((box, text, None if score == 0.0 else score))
    return kept


def _overlap_min(a, b):
    """交集 ÷ 较小的那个框的面积。

    为什么不用 IoU：一个小框整个躺在另一个大框里时，IoU 只有 小/大（可能 0.1），
    会被当成"不重叠"而留下 —— 于是同一处文字出现两遍（一遍好的、一遍半截的）。
    用"占小框的比例"就没有这个洞：包含关系一定接近 1。
    """
    ax0, ay0, ax1, ay1 = _box_xyxy(a)
    bx0, by0, bx1, by1 = _box_xyxy(b)
    ix = max(0.0, min(ax1, bx1) - max(ax0, bx0))
    iy = max(0.0, min(ay1, by1) - max(ay0, by0))
    inter = ix * iy
    if inter <= 0:
        return 0.0
    small = min((ax1 - ax0) * (ay1 - ay0), (bx1 - bx0) * (by1 - by0))
    return inter / small if small > 0 else 0.0


def _merge_items(groups, thr=0.55):
    """多遍（不同阈值 / 不同朝向）的结果合并。

    规则：
      · 与已保留的框重叠（占小框比例）超过 thr → 同一处文字，留可信度高的那条；
        可信度相同留字多的（"1280.00" 比 "11" 可信）。
      · 不重叠 → 都留下（这正是要的：正向漏掉的竖排批注补进来）。
    为什么不按文字去重：同一个人名在同一页出现两次很常见，按文字去重会误杀。
    """
    flat = []
    for items in groups:
        for box, text, score in items:
            flat.append((box, text, score if score is not None else 0.0))
    flat.sort(key=lambda it: (it[2], len(it[1] or "")), reverse=True)

    kept = []
    for box, text, score in flat:
        if not (text or "").strip():
            continue
        if any(_overlap_min(box, kb) >= thr for kb, _, _ in kept):
            continue
        kept.append((box, text, None if score == 0.0 else score))
    return kept


def _uncovered_band(png_path, items):
    """页面里有没有"一窄条竖着的、没被认出来的墨迹" —— 即横躺的页边批注。

    为什么不用"检测结果里有没有竖框"来判断：竖排/横躺文字，正向检测器
    本来就检不到，等它给出竖框才补跑，那永远等不到。
    所以反过来做：**先看哪里没被认出来**，再看这些"漏网的墨"是不是聚成一条竖带。

    判据：漏网的文字级墨块，x 方向能聚成宽 < 12% 页宽的窄带，且纵向跨过 8% 页高，
    且带内至少 3 块 —— 印刷体页面的漏网墨是零散分布的，聚不成这种带。
    """
    try:
        import cv2
        import numpy as np
    except Exception:
        return False

    try:
        buf = np.frombuffer(Path(png_path).read_bytes(), np.uint8)
        img = cv2.imdecode(buf, cv2.IMREAD_GRAYSCALE)     # 中文路径必须走 imdecode
    except Exception:
        return False
    if img is None:
        return False

    H, W = img.shape
    _, bw = cv2.threshold(img, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)
    n, lab, stats, _ = cv2.connectedComponentsWithStats(bw, 8)

    mask = np.zeros((H, W), np.uint8)
    for box, _, _ in items:
        pts = np.array([[int(round(p[0])), int(round(p[1]))] for p in box], np.int32)
        cv2.fillPoly(mask, [pts], 255)
    mask = cv2.dilate(mask, np.ones((7, 7), np.uint8), iterations=1)

    # 「字大概多大」：拿已认出的框高度取中位数。用来筛掉表格线、下划线、印章
    # —— 它们也长得像墨块，但形状跟字差得远，不排掉的话带表格的印刷件全会被误判。
    hs = sorted(max(1.0, _box_xyxy(b)[3] - _box_xyxy(b)[1]) for b, _, _ in items)
    med_h = hs[len(hs) // 2] if hs else max(12.0, 0.02 * H)

    comps = []
    for i in range(1, n):
        x, y, w, h, area = stats[i]
        if area < 40 or h > 160 or w > 0.45 * W:
            continue
        if not (0.5 * med_h <= h <= 1.8 * med_h):     # 不是"一个字"那么高
            continue
        if not (0.6 <= w / float(h) <= 2.2):           # 太细：笔画碎片 / 表格线头
            continue
        sub = lab[y:y + h, x:x + w] == i
        inside = int(np.count_nonzero(sub & (mask[y:y + h, x:x + w] > 0)))
        if area - inside < area * 0.5:      # 这块墨基本被认出来了
            continue
        comps.append((x, y, x + w, y + h))

    if len(comps) < 3:
        return False
    # 按 x 聚成带（相邻 x 间隔小于 6% 页宽算同一带）
    comps.sort(key=lambda c: c[0])
    band = [comps[0]]
    for c in comps[1:]:
        if c[0] - band[-1][2] <= 0.06 * W:
            band.append(c)
        else:
            if _band_ok(band, W, H, med_h):
                return True
            band = [c]
    return _band_ok(band, W, H, med_h)


def _band_ok(band, W, H, med_h):
    """这一带像不像"一列竖着写的字"。

    ① 窄（< 12% 页宽），否则就是一个普通文字块，不是竖带；
    ② 竖着能连成串：把间距 > 3 个字高的地方断开，取最长的一段，
       要求这一段至少 3 个字、且纵向跨过 8% 页高。
       为什么要断开而不是"整条都必须紧"：一条边上可能有两段批注
       （实测：同一列的上下两条批注，中间空 169px，整条算下来必然不达标）。
    ③ 只统计"一个字那么大"的墨块（调用方已按 med_h 过滤），
       否则表格竖线、下划线、零碎标点会被当成竖排的字 —— 这是最典型的误判来源。
    """
    x0 = min(c[0] for c in band)
    x1 = max(c[2] for c in band)
    if (x1 - x0) >= 0.12 * W:
        return False

    band = sorted(band, key=lambda c: c[1])
    run = [band[0]]
    best = []
    for c in band[1:]:
        if c[1] - run[-1][3] <= 3.0 * med_h:
            run.append(c)
        else:
            best = run if len(run) > len(best) else best
            run = [c]
    best = run if len(run) > len(best) else best

    if len(best) < 3:
        return False
    return (best[-1][3] - best[0][1]) > 0.08 * H


def _escalate(png_path, size, items, tmpdir, lang="auto"):
    """手写升级：放宽阈值 + 放大（必做），±90° 补读（页面确实有竖着的东西才做）。

    为什么"放宽阈值"是必做项：RapidOCR 默认 text_score=0.5，手写单字分数普遍
    0.20~0.55（验收单「数量」列就是 0.50/0.44/0.20），会被整格丢掉。
    这跟"识别不准"是两件事 —— 模型认出来了，是阈值把结果吞了。

    为什么"放宽阈值"不能直接替换第一遍：印刷体页面里那些 0.4 分左右的框
    多半是真噪声（墨点、表格线头），直接放进来会污染印刷体的识别结果。
    两遍都跑再合并，就等于"印刷体保持原样，只在原来空着的地方补"。

    返回 (合并后的 items, 说明 dict)。任何一遍失败只是少一点结果，不让整页失败。
    """
    from PIL import Image

    info = {"text_score": HAND_PARAMS["text_score"],
            "box_thresh": HAND_PARAMS["box_thresh"]}
    groups = [items]
    n_pass = 1

    # ---------- ① 放宽阈值 + 放大（手写笔画细，放大后检测更稳） ----------
    w0, h0 = size
    scale = HAND_SCALE if max(size) < HAND_SCALE_MAX_SIDE else 1.0
    try:
        src = Path(png_path)
        if scale != 1.0:
            im = Image.open(str(src))
            src = Path(tmpdir) / "relaxed.png"
            im.resize((int(w0 * scale), int(h0 * scale)), Image.LANCZOS).save(
                str(src), format="PNG")
        got = _run_rapidocr(src, HAND_PARAMS, lang=lang)
        k = 1.0 / scale
        groups.append([([[p[0] * k, p[1] * k] for p in box], t, s)
                       for box, t, s in got])
        info["scale"] = scale
        n_pass += 1
    except Exception as e:
        info["relaxed_error"] = str(e)[:80]

    # ---------- ② ±90° 补读（横躺的页边批注 / 拍歪的纸） ----------
    need_angle = _uncovered_band(png_path, items)
    info["vertical_band"] = bool(need_angle)
    if need_angle:
        done = []
        for ang in HAND_ANGLES:
            try:
                rp, size0, size1 = _rotate_png(png_path, ang, tmpdir)
                got = _run_rapidocr(rp, HAND_PARAMS, lang=lang)
                groups.append([(_box_back(box, ang, size0, size1), t, s)
                               for box, t, s in got])
                done.append(ang)
                n_pass += 1
            except Exception as e:
                info.setdefault("angle_errors", {})[ang] = str(e)[:80]
        info["angles"] = done

    merged = _merge_items(groups)
    info["passes"] = n_pass
    info["before"] = len(items)
    info["after"] = len(merged)
    return merged, info


# ==================== 四、版面还原（行 / 列 / 段） ====================

def _box_geom(box):
    """四点框 → (x0, y0, x1, y1, 中心y, 高)。框缺失时给全 0。"""
    try:
        xs = [float(p[0]) for p in box]
        ys = [float(p[1]) for p in box]
        x0, x1 = min(xs), max(xs)
        y0, y1 = min(ys), max(ys)
        return x0, y0, x1, y1, (y0 + y1) / 2.0, max(1.0, y1 - y0)
    except Exception:
        return 0.0, 0.0, 0.0, 0.0, 0.0, 1.0


def _avg(vals):
    vals = [v for v in vals if v is not None]
    if not vals:
        return None
    return round(sum(vals) / len(vals), 3)


def _is_cjk(ch):
    """中文字符（含中文标点）。"""
    if not ch:
        return False
    o = ord(ch)
    return (0x4E00 <= o <= 0x9FFF or 0x3400 <= o <= 0x4DBF or
            0xF900 <= o <= 0xFAFF or ch in "，。、；：？！（）《》【】“”‘’…—～·")


def _join_parts(parts):
    """把同一「列」里的文字片段拼成一句（parts = [(text, x0, x1, h)]）。

    关键：不能无脑用空格连。
      · Windows 自带 OCR 会把中文按「一个字一个 word」返回，
        无脑连空格就变成「设 备 验 收 单」——检索和抽取全废；
      · 但「设备名称」和「数量」这两列如果挨得近，又确实需要分开。
    所以按视觉间隙判断：
      间隙 < 行高 × 0.5   → 贴在一起，直接粘（中文字与字之间就是这个距离）
      间隙更大的          → 加一个空格
    """
    out = ""
    prev_x1, prev_h = None, None
    for text, x0, x1, h in parts:
        text = (text or "").strip()
        if not text:
            continue
        if not out:
            out = text
        else:
            gap = x0 - prev_x1
            ref_h = prev_h or h or 1.0
            a, b = out[-1], text[0]
            if _is_cjk(a) and _is_cjk(b):
                thr = ref_h * 0.9        # 中文字贴字：松（顺带把中文括号连对）
            elif _is_cjk(a) or _is_cjk(b):
                thr = ref_h * 0.25       # 中英交界：紧一点，"1" 和 "0" 才不会被拆开
            else:
                thr = ref_h * 0.45       # 纯数字 / 字母内部
            if gap > max(2.0, thr):
                out += " "
            out += text
        prev_x1, prev_h = x1, h
    return out.strip()


def _group_lines(items, vert_rtl=True):
    """把文字框按「视觉行」聚拢。

    做法：先按中心 y 排序，再贪心往下吃——只要下一个框的中心 y 和
    当前行的基准 y 差得不超过「行高 × 0.6」，就算同一行。
    （用 0.6 而不是 0.5，是因为照片有轻微倾斜时同一行的 y 会飘。）
    """
    items = [it for it in items if (it[1] or "").strip()]
    items.sort(key=lambda it: _box_geom(it[0])[4])

    lines = []
    for box, text, score in items:
        x0, y0, x1, y1, cy, h = _box_geom(box)
        placed = False
        if lines:
            cur = lines[-1]
            base_cy = sum(_box_geom(b)[4] for b, _, _ in cur) / len(cur)
            base_h = max(_box_geom(b)[5] for b, _, _ in cur)
            if abs(cy - base_cy) <= max(6.0, base_h * 0.6):
                cur.append((box, text, score))
                placed = True
        if not placed:
            lines.append([(box, text, score)])

    out = []
    for ln in lines:
        # 整行都是竖排块时，按「从右到左」排（汉文竖排就是这么读的）；
        # 横排行里偶尔夹一个竖排表头，仍按从左到右，免得把正常版面打乱。
        vert_line = vert_rtl and len(ln) > 1 and all(_is_vert_box(b) for b, _t, _s in ln)
        ln.sort(key=lambda it: _box_geom(it[0])[0], reverse=vert_line)
        parts, confs, boxes = [], [], []
        for box, text, score in ln:
            x0, y0, x1, y1, cy, h = _box_geom(box)
            parts.append((text, x0, x1, h))
            confs.append(score)
            boxes.append(box)

        # 同一行里横向间隙大 → 分成不同"列"（表格行常见）
        if vert_line:
            # 竖排：一列就是一段，别把它们按"视觉行"硬拼成一串
            cols = [p[0] for p in parts if p[0]]
        else:
            cols, buf = [], [parts[0]]
            for i in range(1, len(parts)):
                gap = parts[i][1] - parts[i - 1][2]
                h = max(parts[i][3], parts[i - 1][3])
                if gap > h * COL_GAP_RATIO:
                    cols.append(_join_parts(buf))
                    buf = [parts[i]]
                else:
                    buf.append(parts[i])
            cols.append(_join_parts(buf))
        cols = [c for c in cols if c]

        ys = [_box_geom(b)[4] for b in boxes]
        heights = sorted(_box_geom(b)[5] for b in boxes)
        out.append({
            "text": "  ".join(cols).strip(),
            "conf": _avg(confs),
            "cells": cols,
            "box": boxes[0],
            "y": round(sum(ys) / len(ys), 1),
            "h": round(heights[len(heights) // 2], 1),
            "x0": round(_box_geom(boxes[0])[0], 1),
        })
    return out


# 一个段落最多几行：超过就切开。
# 为什么要有：OCR 看不出"这一段有多长"，行距均匀的公文会被并成一个几十行的巨块，
# 检索时块太大反而匹配不准；切成 12 行一块，来源定位也更细。
MAX_PARA_LINES = 12


def _group_paragraphs(lines):
    """行 → 段（行距明显变大 / 字号跳变就断段）。"""
    if not lines:
        return []
    paras, cur = [], [0]
    for i in range(1, len(lines)):
        prev, now = lines[i - 1], lines[i]
        gap = now["y"] - prev["y"]
        h = max(prev["h"], now["h"], 1.0)
        size_jump = now["h"] > prev["h"] * 1.35
        if gap > h * PARA_GAP_RATIO or size_jump:
            paras.append(cur)
            cur = [i]
        else:
            cur.append(i)
    paras.append(cur)

    out = []
    for idxs in paras:
        for k in range(0, len(idxs), MAX_PARA_LINES):
            chunk = idxs[k:k + MAX_PARA_LINES]
            texts = [lines[i]["text"] for i in chunk if lines[i]["text"]]
            if not texts:
                continue
            out.append((chunk[0], chunk[-1], "\n".join(texts),
                        _avg([lines[i]["conf"] for i in chunk])))
    return out


# ==================== 五、对外主函数 ====================

def _result_from_items(items, t0, extra=None, vert_rtl=True):
    """原始 OCR 结果 → 统一结构（行 / 段 / 全文 / 统计）。

    vert_rtl：竖排行是否按"从右到左"排（关掉竖排开关时为 False，完全按横排读）。
    """
    lines = _group_lines(items, vert_rtl=vert_rtl)
    paras = _group_paragraphs(lines)
    low = [i for i, ln in enumerate(lines)
           if ln["conf"] is not None and ln["conf"] < LOW_CONF]
    text = "\n\n".join(p[2] for p in paras)

    res = {
        "ok": True,
        "engine": _ENGINE,
        "elapsed": round(time.time() - t0, 2),
        "text": text,
        "lines": lines,
        "paragraphs": [
            {"text": t, "conf": c, "line_start": a, "line_end": b}
            for a, b, t, c in paras
        ],
        "box_count": len(items),
        "line_count": len(lines),
        "para_count": len(paras),
        "char_count": len(re.sub(r"\s", "", text)),
        "low_conf_lines": len(low),
        "avg_conf": _avg([ln["conf"] for ln in lines]),
    }
    if low:
        res["low_conf_tip"] = ("有 %d 行置信度低于 %.0f%%，八成是拍照歪了 / 字太小 / 表格线压字，"
                               "这些行建议人工复核" % (len(low), LOW_CONF * 100))
    if extra:
        res.update(extra)
    return res


# ==================== 三·六、竖排文字（2026-11-07） ====================
#
# 竖排的难点不在"认字"，在"顺序"。
# 实测（PP-OCRv5）：一列竖着的「環境空氣質量」整列被检成一个细长框，模型能认出来
# （它的训练数据里有竖排），但**框是竖的、版面还原是横的**——同一高度上的两列字
# 会被拼成一行，于是"右起第一列"跑到后面去了；字间空隙大时一列还会被拆成几个框，
# 读出来断断续续。
#
# 所以这里只做两件事，一个字都不重新认（把字转 90° 反而认不出，字本身是正立的）：
#   ① 同一列的多个竖排框 → 合成一格（bbox 取并集，文字按从上到下拼）；
#   ② 整页竖排时，行内排序改成"从右到左"（汉文竖排的阅读方向）——见 _group_lines。
# 判定依据：框的高 ÷ 宽 ≥ VERT_RATIO（横排单字框接近方形，竖排整列是细长条）。

def _box_wh(box):
    x0, y0, x1, y1 = _box_xyxy(box)
    return max(0.0, x1 - x0), max(0.0, y1 - y0)


def _is_vert_box(box, min_h=20):
    w, h = _box_wh(box)
    return w > 0 and h >= min_h and h / w >= VERT_RATIO


def _cjk_ratio(txt):
    s = re.sub(r"\s", "", txt or "")
    if not s:
        return 0.0
    return sum(1 for ch in s if 0x4E00 <= ord(ch) <= 0x9FFF) / float(len(s))


def _vert_share(items):
    """竖排框里的字数 ÷ 全部字数（0~1）。"""
    tot = vert = 0
    for box, txt, _sc in items:
        n = len(re.sub(r"\s", "", txt or ""))
        tot += n
        if box is not None and _is_vert_box(box):
            vert += n
    return vert / float(max(tot, 1))


def _cluster_cols(verts):
    """把竖排框按 x 中心聚成"列"。verts = [(idx, box, text, score)]。"""
    ws = sorted(_box_wh(v[1])[0] for v in verts)
    med_w = ws[len(ws) // 2] or 1.0
    span = max(med_w * 1.2, 12.0)                        # 同一列的左右容差
    cols = []
    for v in verts:
        x0, _y0, x1, _y1 = _box_xyxy(v[1])
        cx = (x0 + x1) / 2.0
        for c in cols:
            bc = sum((_box_xyxy(b)[0] + _box_xyxy(b)[2]) / 2.0 for _i, b, _t, _s in c) / len(c)
            if abs(cx - bc) <= span:
                c.append(v)
                break
        else:
            cols.append([v])
    return cols


def _merge_col(col):
    """一列竖排框 → 一个合成框（文字从上到下拼、bbox 取并集）。"""
    col = sorted(col, key=lambda v: _box_xyxy(v[1])[1])
    xs0 = [_box_xyxy(v[1])[0] for v in col]
    ys0 = [_box_xyxy(v[1])[1] for v in col]
    xs1 = [_box_xyxy(v[1])[2] for v in col]
    ys1 = [_box_xyxy(v[1])[3] for v in col]
    bbox = [[min(xs0), min(ys0)], [max(xs1), min(ys0)],
            [max(xs1), max(ys1)], [min(xs0), max(ys1)]]
    text = "".join((v[2] or "").strip() for v in col)
    conf = _avg([float(v[3]) for v in col if v[3] is not None])
    return bbox, text, conf, [v[0] for v in col]


def _stack_columns(items):
    """单字框竖向堆叠的情况：字距大时检测器会"一个字一个框"，
    这时框是方的（h/w≈1），_is_vert_box 抓不到，但排布仍是竖排。
    判据：一堆"单个汉字"的方框，按 x 聚成列后每列 ≥3 个，且至少 2 列。
    """
    sq = []
    for i, (box, txt, sc) in enumerate(items):
        if box is None:
            continue
        w, h = _box_wh(box)
        if w < 6 or h < 6 or h / w < 0.55 or h / w > 1.8:
            continue
        t = (txt or "").strip()
        if not t or len(t) > 2 or _cjk_ratio(t) < 0.5:
            continue
        sq.append((i, box, t, sc))
    if len(sq) < 6:
        return []
    ws = sorted(_box_wh(v[1])[0] for v in sq)
    span = max((ws[len(ws) // 2] or 10) * 0.9, 10.0)
    cols = []
    for v in sq:
        x0, _y0, x1, _y1 = _box_xyxy(v[1])
        cx = (x0 + x1) / 2.0
        for c in cols:
            bc = sum((_box_xyxy(b)[0] + _box_xyxy(b)[2]) / 2.0 for _i, b, _t, _s in c) / len(c)
            if abs(cx - bc) <= span:
                c.append(v)
                break
        else:
            cols.append([v])
    cols = [c for c in cols if len(c) >= 2]
    if len(cols) < 2 or sum(len(c) for c in cols) < 5:
        return []

    def _cx(c):
        return sum((_box_xyxy(v[1])[0] + _box_xyxy(v[1])[2]) / 2.0 for v in c) / len(c)

    # 零星漏网的单字框：x 落在某列附近就并进去（否则它会单独成行，把一列读断）
    loose = [v for v in sq if all(v not in c for c in cols)]
    for v in loose:
        x0, _y0, x1, _y1 = _box_xyxy(v[1])
        cx = (x0 + x1) / 2.0
        best = min(cols, key=lambda c: abs(cx - _cx(c)))
        if abs(cx - _cx(best)) <= 3 * span:
            best.append(v)
    return cols


def _vertical_pass(png, wh, items, tmpdir, lang="auto", vertical="auto"):
    """竖排处理（只重排、不重认）。返回 (items, 说明 or None)。

    vertical：auto（默认，只在需要时动手）/ on（强制所有竖排框按列拼整）/ off（关闭）
    """
    mode = (vertical or "").strip().lower()
    if mode in ("0", "off", "no", "关", "关闭"):
        return items, None
    if _ENGINE != "rapidocr" or not items:
        return items, None

    verts = []
    for i, (box, txt, sc) in enumerate(items):
        if box is None or not _is_vert_box(box):
            continue
        if len(re.sub(r"\s", "", txt or "")) < 2 or _cjk_ratio(txt) < 0.3:
            continue                                     # 单字框 / 英文不动，免得误伤表格窄列
        verts.append((i, box, txt, sc))

    cols = _cluster_cols(verts) if verts else []
    forced = mode in ("on", "1", "yes", "开", "强制")
    # 只有"被拆成好几个框"的列才需要合；强制模式下所有竖排框都合
    targets = cols if forced else [c for c in cols if len(c) >= 2]
    stacked = False
    if not targets:
        targets = _stack_columns(items)                   # 字距大 → 一个字一个框
        stacked = bool(targets)
    if not targets:
        return items, None

    drop, added = set(), []
    for col in targets:
        bbox, text, conf, idxs = _merge_col(col)
        if not text.strip():
            continue
        drop.update(idxs)
        added.append((bbox, text, conf if conf else None))
    if not added:
        return items, None
    merged = [it for i, it in enumerate(items) if i not in drop] + added
    if stacked:
        note = "竖排：%d 列（单字框竖向堆叠）已按「从上到下、从右到左」拼整" % len(added)
    else:
        note = "竖排：%d 列已按「从上到下」拼整" % len(added)
        if _vert_share(items) >= VERT_FULL_SHARE:
            note += "，整页竖排按「从右到左」排列"
    return merged, note


def _ocr_with_hand(png, wh, tmpdir, lang="auto", vertical="auto"):
    """跑一遍；再做竖排 / 手写升级。返回 (items, 手写说明 dict or None, 竖排说明 or None)。

    A23_OCR_HAND：
      未设 / auto → 升级（默认开。印刷体的识别结果不变，只是多跑一遍合并）
      0           → 关闭，只跑正向一遍（复现 11-03 之前的行为，做对照时用）

    为什么不拿"置信度低"当开关：默认那一遍 text_score=0.5，低分的已经被引擎
    丢掉了，外面根本看不到 —— 用置信度当判据等于用"已经被删掉的证据"判断。
    """
    items = _run_engine(png, lang=lang)
    hand_info = None
    vert_note = None
    if _ENGINE == "rapidocr":
        # 顺序有讲究：先"放宽阈值 + 补读"把框收全，再做竖排重排。
        # 反过来会漏——字距大的竖排单字在默认阈值下直接被丢掉（实测 11 个框剩 0 个），
        # 竖排逻辑看到的就是空结果，等于没做。
        want = (os.environ.get("A23_OCR_HAND") or "auto").strip().lower()
        if want != "0":
            items, hand_info = _escalate(png, wh, items, tmpdir, lang=lang)
        items, vert_note = _vertical_pass(png, wh, items, tmpdir, lang=lang, vertical=vertical)
    return items, hand_info, vert_note



def read_image(path, page=None, lang="auto", vertical="auto") -> dict:
    """认一张图（本地路径）。

    page      只用于标记来源（PDF 第几页）
    lang      auto / chs / cht / en —— 选识别模型（auto/chs/cht 同一个中文模型，它会认繁体）
    vertical  auto / on / off —— 竖排处理（on = 强制整页按竖排转正识别）
    """
    t0 = time.time()
    p = Path(path)
    if not p.is_file():
        return {"ok": False, "error": "图片不存在：%s" % p}
    if _pick_engine() == "none":
        return {"ok": False, "error": _ENGINE_ERR, "engine": "none"}

    lang = _norm_lang(lang)
    try:
        png, w, h, tmpdir = _prepare(src_path=p)
    except Exception as e:
        return {"ok": False, "error": "图片读不了（%s: %s）" % (type(e).__name__, e)}
    try:
        items, hand, vert = _ocr_with_hand(png, (w, h), tmpdir, lang=lang, vertical=vertical)
    except Exception as e:
        return {"ok": False, "error": "OCR 失败：%s" % e, "engine": _ENGINE}
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)

    extra = {"source": p.name, "width": w, "height": h, "lang": lang}
    if hand:
        extra["handwriting"] = hand
    if vert:
        extra["vertical"] = vert
    if page is not None:
        extra["page"] = int(page)
    res = _result_from_items(items, t0, extra,
                             vert_rtl=(str(vertical or "auto").strip().lower()
                                       not in ("0", "off", "no", "关", "关闭")))
    res["lang"] = lang
    res["lang_label"] = LANG_LABEL.get(lang, lang)
    res["rec_model"] = _rec_model(lang)[1]
    if not res["text"].strip():
        res["empty"] = True
        res["tip"] = "这张图里没认到文字（可能是纯图片 / 照片，或者分辨率太低）"
    return res


def read_image_bytes(data: bytes, name: str = "image", lang="auto", vertical="auto") -> dict:
    """认内存里的一张图（上传链路：文件还没落盘就能先看结果）。"""
    t0 = time.time()
    if _pick_engine() == "none":
        return {"ok": False, "error": _ENGINE_ERR, "engine": "none"}
    lang = _norm_lang(lang)
    try:
        png, w, h, tmpdir = _prepare(src_bytes=data)
    except Exception as e:
        return {"ok": False, "error": "图片读不了（%s: %s）" % (type(e).__name__, e)}
    try:
        items, hand, vert = _ocr_with_hand(png, (w, h), tmpdir, lang=lang, vertical=vertical)
    except Exception as e:
        return {"ok": False, "error": "OCR 失败：%s" % e, "engine": _ENGINE}
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)

    extra = {"source": name, "width": w, "height": h, "lang": lang}
    if hand:
        extra["handwriting"] = hand
    if vert:
        extra["vertical"] = vert
    res = _result_from_items(items, t0, extra,
                             vert_rtl=(str(vertical or "auto").strip().lower()
                                       not in ("0", "off", "no", "关", "关闭")))
    res["lang"] = lang
    res["lang_label"] = LANG_LABEL.get(lang, lang)
    res["rec_model"] = _rec_model(lang)[1]
    if not res["text"].strip():
        res["empty"] = True
        res["tip"] = "这张图里没认到文字（可能是纯图片 / 照片，或者分辨率太低）"
    return res


# ==================== 六、命令行（单张试跑，排查用） ====================

if __name__ == "__main__":
    st = status()
    print("引擎：%s ｜ PDF 拆页：%s" % (st["engine"], st.get("pdf")))
    print("说明：%s" % st.get("tip"))
    if not st["ok"]:
        sys.exit(1)

    if len(sys.argv) < 2:
        print("用法：python agent/ocr.py 某张图.png")
        sys.exit(0)

    r = read_image(sys.argv[1])
    if not r.get("ok"):
        print("失败：%s" % r.get("error"))
        sys.exit(1)
    print("耗时 %.2fs ｜ %d 个框 ｜ %d 行 ｜ %d 段 ｜ 平均置信度 %s"
          % (r["elapsed"], r["box_count"], r["line_count"],
             r["para_count"], r["avg_conf"]))
    print("-" * 60)
    print(r["text"])
