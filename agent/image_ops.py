# -*- coding: utf-8 -*-
"""
image_ops.py —— 图片插入（上传 → 自动定位置 → 插进文档 → 原格式不动）

一句话流程：
  用户上传一张图 → 系统先看指令里说的位置（"标题下面""第二部分""文末"）
  → 没说位置就看图片名和正文哪一段最相关（如"组织架构图" ↔ "组织架构"）
  → 都没对上就放文末
  → 在那一行/段之后插一个居中段落，宽度按页面可用宽度自适应
  → 另存新文件，原件不动

支持：
  .docx   python-docx（add_picture，段落里插图，版式不动）
  .xlsx   openpyxl（add_image，锚到单元格；需 Pillow）
  .md     Markdown 语法 ![图片](文件名)，锚到对应行
  .txt    不支持（纯文本没有图片表示法）——明确告知，不硬塞

位置判定的三级兜底（"自动"就体现在这里）：
  ① locate.hint（指令里的位置词 / 关键词）→ 命中段落
  ② 图片文件名与段落文字做片段匹配（≥2 字才算命中）
  ③ 文末
"""
import re
from pathlib import Path

BASE = Path(__file__).resolve().parents[1]
IMAGE_DIR = BASE / "导入图片"                       # 上传图片的存放目录
SUPPORTED_IMAGE_EXTS = {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".webp"}
MAX_IMAGE_BYTES = 20 * 1024 * 1024                  # 单张 20MB 上限

# 位置词（用户口语里常见的说法）
START_WORDS = ("开头", "文首", "最前", "顶部", "第一段前", "标题前", "扉页", "卷首")
END_WORDS = ("末尾", "文末", "最后", "结尾", "底部", "落款后", "文档最后")
TITLE_WORDS = ("标题", "题目", "大标题")


# ==================== 一、图片库（上传 / 列表 / 查找） ====================

def ensure_dir() -> Path:
    IMAGE_DIR.mkdir(parents=True, exist_ok=True)
    return IMAGE_DIR


def _safe_stem(name: str) -> str:
    """文件名去掉路径和非法字符（防路径注入、防 Windows 不接受的名字）。"""
    stem = Path(name or "image").stem.strip() or "image"
    stem = re.sub(r'[\\/:*?"<>|\r\n\t]', "_", stem)
    return stem[:80]


def image_size(path: Path):
    """读图片像素宽高。Pillow 不可用时返回 (None, None)——不影响插入。"""
    try:
        from PIL import Image as PILImage
    except Exception:
        return None, None
    try:
        with PILImage.open(str(path)) as im:
            return int(im.width), int(im.height)
    except Exception:
        return None, None


def _unique_name(stem: str, ext: str) -> str:
    cand = stem + ext
    i = 2
    while (IMAGE_DIR / cand).exists():
        cand = "%s-%d%s" % (stem, i, ext)
        i += 1
    return cand


def save_image(data: bytes, filename: str) -> dict:
    """保存上传的图片（重名自动加后缀，绝不覆盖），返回图片信息。"""
    ensure_dir()
    ext = Path(filename or "").suffix.lower()
    if ext not in SUPPORTED_IMAGE_EXTS:
        raise ValueError("不支持的图片格式「%s」——支持：%s"
                         % (ext or "无扩展名", " / ".join(sorted(SUPPORTED_IMAGE_EXTS))))
    if len(data) > MAX_IMAGE_BYTES:
        raise ValueError("图片太大（%.1f MB），上限 20 MB" % (len(data) / 1024 / 1024))

    stored = _unique_name(_safe_stem(filename), ext)
    path = IMAGE_DIR / stored
    path.write_bytes(data)
    w, h = image_size(path)
    return {
        "id": stored,
        "name": Path(filename or stored).name,
        "stored_name": stored,
        "width": w,
        "height": h,
        "size": len(data),
        "path": str(path),
        "preview": "/api/images/" + stored,
    }


def list_images(limit: int = 50) -> list:
    """最近的图片（新的在前）。"""
    ensure_dir()
    files = [p for p in IMAGE_DIR.glob("*") if p.is_file()
             and p.suffix.lower() in SUPPORTED_IMAGE_EXTS]
    files.sort(key=lambda p: p.stat().st_mtime, reverse=True)
    out = []
    for p in files[:max(1, int(limit))]:
        w, h = image_size(p)
        out.append({
            "id": p.name, "name": p.name, "width": w, "height": h,
            "size": p.stat().st_size,
            "preview": "/api/images/" + p.name,
            "mtime": p.stat().st_mtime,
        })
    return out


def latest_image():
    """最近上传的一张图（指令没说用哪张时的默认）。"""
    items = list_images(1)
    return Path(IMAGE_DIR / items[0]["id"]) if items else None


def resolve_image(ref: str):
    """把用户给的引用变成图片路径：支持存储名 / 原文件名 / 绝对路径 / 模糊包含。"""
    ref = (ref or "").strip().strip('"').strip("'")
    if not ref:
        return None
    p = Path(ref)
    if p.is_file():
        return p
    ensure_dir()
    cand = IMAGE_DIR / Path(ref).name
    if cand.is_file():
        return cand
    key = Path(ref).stem.lower()
    if not key:
        return None
    for f in IMAGE_DIR.glob("*"):
        if f.suffix.lower() in SUPPORTED_IMAGE_EXTS and key in f.stem.lower():
            return f
    return None


# ==================== 二、位置判定（三级兜底） ====================

def _clean(s: str) -> str:
    return re.sub(r"\s+", "", s or "")


def _clip(s: str, n: int = 24) -> str:
    s = (s or "").replace("\n", " ").strip()
    return s if len(s) <= n else s[:n] + "…"


def _frags(s: str, lo: int = 2, hi: int = 4) -> list:
    out = []
    for n in range(hi, lo - 1, -1):
        for i in range(len(s) - n + 1):
            f = s[i:i + n]
            if f not in out:
                out.append(f)
    return out


def _image_keywords(image_path) -> str:
    """图片名 → 关键词（去掉"图/截图/图片"这类无信息后缀）。"""
    stem = Path(image_path).stem
    stem = re.sub(r"(图片|插图|截图|照片|配图|image|img|photo|screenshot|pic)",
                  "", stem, flags=re.I)
    return _clean(stem.replace("_", "").replace("-", ""))


def pick_index(texts: list, hint: str, image_path) -> tuple:
    """在一串文本（段落 / 行）里挑插入位置。

    返回 (idx, where, reason)：
      idx   —— 目标下标（where=end 时无意义）
      where —— "start"（第 0 段之前）/ "after"（idx 之后）/ "end"（文末）
    """
    hint_c = _clean(hint)
    n = len(texts)

    # ① 指令里明确说了位置
    if hint_c:
        if any(w in hint_c for w in END_WORDS):
            return n, "end", "指令要求放文末"
        if any(w in hint_c for w in START_WORDS):
            return -1, "start", "指令要求放文首"
        if any(w in hint_c for w in TITLE_WORDS):
            first = next((i for i, t in enumerate(texts) if _clean(t)), None)
            if first is not None:
                return first, "after", "放到标题（第 %d 段）之后" % (first + 1)

        for i, t in enumerate(texts):
            if hint_c and hint_c in _clean(t):
                return i, "after", "定位词「%s」命中第 %d 段" % (hint.strip(), i + 1)

        # 位置词没命中原文时按片段打分（如"第二部分"可能原文是"二、实施计划"）
        best, best_score = None, 0
        for i, t in enumerate(texts):
            tc = _clean(t)
            score = 0
            for f in _frags(hint_c)[:60]:
                if f in tc:
                    score = max(score, len(f))
            if score > best_score:
                best, best_score = i, score
        if best is not None and best_score >= 2:
            return best, "after", ("定位词「%s」最接近第 %d 段「%s」"
                                   % (hint.strip(), best + 1, _clip(texts[best])))

    # ② 指令没说位置 → 图片名和正文哪段最相关
    key = _image_keywords(image_path)
    if key:
        best, best_score = None, 0
        for i, t in enumerate(texts):
            tc = _clean(t)
            score = 0
            for n2 in range(min(4, len(key)), 1, -1):
                for j in range(len(key) - n2 + 1):
                    if key[j:j + n2] in tc:
                        score = max(score, n2)
            if score > best_score:
                best, best_score = i, score
        if best is not None and best_score >= 2:
            return best, "after", ("图片名「%s」与第 %d 段「%s」相关"
                                   % (key, best + 1, _clip(texts[best])))

    # ③ 兜底：文末
    return n, "end", "没找到匹配的小节，放到文末"


# ==================== 三、各格式的落地 ====================

def _target_width(doc, image_path, ratio):
    """插入宽度：页面可用宽度 × 比例，且不超过原图像素宽（小图不放大）。"""
    from docx.shared import Emu
    try:
        r = float(ratio)
    except (TypeError, ValueError):
        r = 0.7
    r = min(max(r, 0.15), 1.0)
    sec = doc.sections[0]
    avail = int(sec.page_width - sec.left_margin - sec.right_margin)
    w = int(avail * r)
    try:
        from docx.image.image import Image as DocxImage
        di = DocxImage.from_file(str(image_path))
        if di.width and w > int(di.width):
            w = int(di.width)
    except Exception:
        pass
    return Emu(w)


def _para_after(doc, anchor):
    """在 anchor 段之后插一个空段落（python-docx 没有直接 API，先加再挪）。"""
    p = doc.add_paragraph()
    anchor._p.addnext(p._p)
    return p


def _para_before(doc, anchor):
    p = doc.add_paragraph()
    anchor._p.addprevious(p._p)
    return p


def _insert_docx(working: Path, out_path: Path, image_path: Path, intent: dict):
    import docx
    from docx.shared import Pt
    from docx.enum.text import WD_ALIGN_PARAGRAPH

    d = docx.Document(str(working))
    loc = intent.get("locate") or {}
    params = intent.get("params") or {}
    hint = (loc.get("hint") or "").strip()
    pos = (params.get("position") or "").strip().lower()
    if not hint:
        if pos == "start":
            hint = "开头"
        elif pos == "end":
            hint = "文末"

    paras = list(d.paragraphs)
    idx, where, reason = pick_index([p.text for p in paras], hint, image_path)

    width = _target_width(d, image_path, params.get("width_ratio"))

    if not paras:
        newp = d.add_paragraph()
        before = "（空文档）"
    elif where == "start":
        newp = _para_before(d, paras[0])
        before = paras[0].text
    elif where == "end":
        newp = d.add_paragraph()
        before = paras[-1].text
    else:
        newp = _para_after(d, paras[idx])
        before = paras[idx].text

    newp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    try:
        newp.paragraph_format.space_before = Pt(6)
        newp.paragraph_format.space_after = Pt(6)
    except Exception:
        pass
    pic = newp.add_run().add_picture(str(image_path), width=width)
    d.save(str(out_path))

    if where == "start":
        pos_desc = "文档开头（第 1 段之前）"
    elif where == "end" or idx >= len(paras):
        pos_desc = "文档末尾"
    else:
        pos_desc = "第 %d 段「%s」之后" % (idx + 1, _clip(paras[idx].text))

    changes = [{
        "para": idx,
        "before": before or "（空段落）",
        "after": "［图片 %s，宽约 %.1f cm］" % (image_path.name, pic.width.cm),
        "how": "插入图片到%s（%s）" % (pos_desc, reason),
    }]
    return changes, True, "%s ｜ %s" % (pos_desc, reason)


def _insert_xlsx(working: Path, out_path: Path, image_path: Path, intent: dict):
    try:
        from openpyxl import load_workbook
        from openpyxl.drawing.image import Image as XLImage
    except Exception as e:
        return [], False, "Excel 插图需要 Pillow 支持（%s）" % e

    loc = intent.get("locate") or {}
    params = intent.get("params") or {}
    hint = (loc.get("hint") or "").strip()

    wb = load_workbook(str(working))
    ws = wb.active
    m = re.search(r"([A-Za-z]{1,2}\d{1,4})", hint)
    if m:
        anchor = m.group(1).upper()
        why = "锚到单元格 %s" % anchor
    elif any(w in _clean(hint) for w in END_WORDS):
        anchor = "A%d" % (ws.max_row + 2)
        why = "锚到表尾 %s" % anchor
    else:
        anchor = "A1"
        why = "锚到工作表左上角 A1"

    img = XLImage(str(image_path))
    ratio = params.get("width_ratio")
    try:
        ratio = float(ratio)
    except (TypeError, ValueError):
        ratio = 0.5
    max_w = int(max(120, min(900, 900 * min(max(ratio, 0.15), 1.0))))
    if img.width and img.width > max_w:
        k = max_w / img.width
        img.width, img.height = int(img.width * k), int(img.height * k)

    ws.add_image(img, anchor)
    wb.save(str(out_path))
    changes = [{
        "para": 0, "before": "工作表「%s」" % ws.title,
        "after": "［图片 %s，%d×%d px］" % (image_path.name, img.width, img.height),
        "how": "插入图片到 %s（%s）" % (anchor, why),
    }]
    return changes, True, why


def _insert_markdown(working: Path, out_path: Path, image_path: Path, intent: dict):
    loc = intent.get("locate") or {}
    hint = (loc.get("hint") or "").strip()
    text = working.read_text(encoding="utf-8")
    lines = text.splitlines()

    idx, where, reason = pick_index(lines, hint, image_path)
    md = "![%s](%s)" % (image_path.stem, image_path.name)

    if where == "start":
        new_lines = [md, ""] + lines
        pos_desc = "文档开头"
    elif where == "end":
        new_lines = lines + ([""] if lines and lines[-1].strip() else []) + [md]
        pos_desc = "文档末尾"
    else:
        new_lines = lines[:idx + 1] + ["", md] + lines[idx + 1:]
        pos_desc = "第 %d 行「%s」之后" % (idx + 1, _clip(lines[idx]))

    out_path.write_text("\n".join(new_lines) + ("\n" if text.endswith("\n") else ""),
                        encoding="utf-8")
    changes = [{
        "para": idx, "before": lines[idx] if 0 <= idx < len(lines) else "（文末）",
        "after": md, "how": "插入 Markdown 图片到%s（%s）" % (pos_desc, reason),
    }]
    return changes, True, "%s ｜ %s" % (pos_desc, reason)


def insert_image(working: Path, out_path: Path, image_path: Path, intent: dict):
    """按文档类型分发。返回 (changes, ok, msg)，与 edit.py 的约定一致。"""
    suffix = working.suffix.lower()
    if suffix == ".docx":
        return _insert_docx(working, out_path, image_path, intent)
    if suffix == ".xlsx":
        return _insert_xlsx(working, out_path, image_path, intent)
    if suffix in (".md", ".markdown"):
        return _insert_markdown(working, out_path, image_path, intent)
    return [], False, ("暂不支持往 %s 里插图——Word(.docx) / Excel(.xlsx) / Markdown(.md) "
                       "可以，纯文本没有图片表示法" % (suffix or "该类型"))
