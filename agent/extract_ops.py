# -*- coding: utf-8 -*-
"""
extract_ops.py —— 模块（1）· 第 4 步：内容提取接入（说一句话 → 信息整理出来）

链路：一句话 → ①解析（parse_intent）→ ②字段规范化（本文件）
      → ③逐字段复用抽取线（matcher：筛→判→定）→ ④整理输出（配对表 / 分列列表）

为什么能复用抽取线？"内容提取"与"填表抽取"本质是同一件事：
  目标（字段名）→ 找证据（召回）→ 让模型摘值（判）→ 标出处（定）。
  区别只是：填表＝模板给列名、全文档里找；这里＝用户给字段、指定文档里找。
  所以直接调 matcher/extract.py 的 run_column（用 only_file 限定在目标文档内）。

三层字段规范化（本步新增，解决"口语说法对不上文档"）：
  1. 合称拆分："甲乙双方名称" → 甲方名称 + 乙方名称          （内置规则）
  2. 别名表：查 字段别名表.json（如"工资"→"基本工资"）        （手工，退成兜底）
  3. 语义对齐：用本地向量把字段对到文档里真实出现的说法
     （如"待办事项"→"任务内容"、"合同总价"→"合同金额合计"）——2026-10-26 新增，
     不需要配任何东西，换行业自动生效；模型缺失时自动降级为原词直抽

三层是"从便宜到贵"排的：合称拆分和语义对齐都不花钱、不联网；
只有前两层都没对上，才轮到第四条"大模型翻译"（一次几秒、有成本）。

跑法（在项目根目录）：
  python agent/extract_ops.py "把员工信息表里所有员工的姓名和工号整理出来"

输出：
  屏幕：字段 → 条目（值 + 出处）的报告
  文件：agent/输出/<文档名>_提取结果.md / .json（中间结果在 输出/中间抽取/）
"""
import json
import re
import sys
import time
from collections import Counter
from pathlib import Path

import requests

# 输出兜底：遇到终端表示不了的字符就替换，不让程序崩
try:
    sys.stdout.reconfigure(errors="replace")
except Exception:
    pass

BASE = Path(__file__).resolve().parents[1]                  # = 研一项目/
AGENT_DIR = Path(__file__).resolve().parent                 # = 研一项目/agent/
sys.path.insert(0, str(AGENT_DIR))                          # 复用同目录的解析器
sys.path.insert(0, str(BASE / "matcher"))                   # 复用抽取线

from parse_intent import API_KEY, API_URL, MODEL            # noqa: E402
from llm_config import chat as _llm_chat                    # noqa: E402
import extract as matcher_extract                           # noqa: E402
import semantic                                             # noqa: E402  语义字段匹配

OUT_DIR = BASE / "agent" / "输出"
MID_DIR = OUT_DIR / "中间抽取"          # run_column 的落盘目录（每字段一份 json）
OUTPUT_JSON = BASE / "doc_reader" / "output.json"
ALIAS_FILE = AGENT_DIR / "字段别名表.json"

# 合称 → 逐个展开（用于"甲乙双方名称"这类一个词顶两个字段的说法）
SPLIT_RULES = {
    "甲乙双方": ["甲方", "乙方"],
    "甲乙两方": ["甲方", "乙方"],
    "供需双方": ["供方", "需方"],
}

TRANSLATE_SYSTEM = ("你是文档信息提取助手。用户给一个想提取的字段名，"
                    "你要在给定的一份文档要素清单里，找出含义最接近的说法。"
                    "严格只输出 JSON，不要任何解释。")


# ==================== 一、字段规范化 ====================

def load_aliases() -> dict:
    """读"口语 → 正式"别名表（不存在或坏了就当空表，不影响流程）。

    2026-10-26：这张表从"主力"退成"兜底/覆盖"——
    常见说法由语义匹配自动对上，这里只留两种东西：
      · 语义容易搞错、需要人工拍板的（如"总价"在有的文档里指不含税价）
      · 客户固定叫法（必须按他们的说法走）
    """
    if ALIAS_FILE.exists():
        try:
            data = json.loads(ALIAS_FILE.read_text(encoding="utf-8"))
            return {k: v for k, v in data.items()
                    if not k.startswith("_") and isinstance(v, str)}
        except Exception:
            pass
    return {}


# ---------- 文档里的"说法"清单（语义匹配的候选空间） ----------

_DOC_CACHE = {"key": None, "docs": None}


def _load_docs():
    """读 output.json 的 documents（按 mtime 缓存，避免每个字段重读一遍）。"""
    if not OUTPUT_JSON.exists():
        return []
    st = OUTPUT_JSON.stat()
    key = (str(OUTPUT_JSON), st.st_mtime_ns, st.st_size)
    if _DOC_CACHE["key"] != key:
        try:
            data = json.loads(OUTPUT_JSON.read_text(encoding="utf-8"))
        except Exception:
            return []
        _DOC_CACHE["key"] = key
        _DOC_CACHE["docs"] = data.get("documents", [])
    return _DOC_CACHE["docs"] or []


_BAD_TERM = re.compile(r"^[\d\s.,%％+\-/：:（）()年月日]+$")
_LABEL_RE = re.compile(r"^([\u4e00-\u9fa5A-Za-z（）()·]{2,12})\s*[：:]")
# 编码型文本（E1001 / M-1001 / SO-2601-001 / VMC-850B）不是字段名。
# 不滤掉的话它们会占候选名额，还会把语义排名带偏。
_CODE_TERM = re.compile(r"^[A-Za-z]{1,6}[-_]?\d|^\d+[A-Za-z]|^=")
# 条款标题（"第三条 交付与验收"）不是字段名；地名/人名那类值也不是。
# 2026-10-26 补：候选池里混进这些，语义排名会把用户带偏（"合同总价"→"合同编号"）。
_CLAUSE_TERM = re.compile(r"^第\s*[一二三四五六七八九十百千\d]+\s*[条章节款项部分]")


def doc_candidates(file_name: str, limit: int = 400) -> list:
    """把一份文档里"可能当字段名"的说法捞出来，作为语义匹配的候选空间。

    为什么不用"全文档短文本"（doc_element_list）：
      那样 339 个候选里大半是数据值（"920"、"E1001"、"制造中心"），噪声大、还慢。
      字段名只可能出现在两个地方 —— ①表格的列头 ②"XXX："这种标签。
      只捞这两类，候选能压到几十个，匹配既准又快。

    列头取每列**前 5 个非空**而不是最上面那一个：
      Excel 表头上面常压着一行大标题（"设备采购台账"），最上面那个不是列头。
    """
    docs = _load_docs()
    doc = next((d for d in docs if d.get("file") == file_name), None)
    if not doc:
        return []

    cols = {}          # (区域, 列号) -> [(行号, 文本)]
    labels = []
    for b in doc.get("blocks", []):
        loc = b.get("locator") or {}
        raw = b.get("text") or ""
        if not raw.strip():
            continue

        kind = loc.get("kind")
        if kind in ("cell", "table_cell"):
            zone = (kind, loc.get("sheet") or loc.get("table_index"), loc.get("col"))
            cols.setdefault(zone, []).append((loc.get("row") or 0,
                                              re.sub(r"\s+", " ", raw).strip()))
        else:
            # 按行看：一个块里可能压着好几条「甲方：… 乙方：…」
            for line in raw.split("\n"):
                line = re.sub(r"\s+", " ", line).strip()
                if not line:
                    continue
                m = _LABEL_RE.match(line)
                if m:
                    labels.append(m.group(1).strip())

    cands, seen = [], set()
    title = semantic.norm(Path(file_name).stem)     # 文档大标题不是字段名

    def add(t):
        t = str(t).strip()
        if not t or len(t) > 12 or len(t) < 2:
            return
        if _BAD_TERM.match(t) or _CODE_TERM.match(t) or _CLAUSE_TERM.match(t):
            return
        k = semantic.norm(t)
        if not k or k == title:
            return
        if k not in seen:
            seen.add(k)
            cands.append(t)

    # 表头到底在第几行：按「同一张表」统计各列第一个非空行的众数。
    # 取到众数这一行为止（含它上方的行）：两层表头时众数落在第二层，父列名在上面也不会漏；
    # 单层表头时众数就是表头行，它下面那行是数据 —— 不能要。
    # 必须按表分组算：一份文档里可能好几张表（项目进度表=总览+工作明细），
    # 各表的表头行不一样，混在一起算会把整张表的列头漏掉。
    # 2026-10-26 改：原先取"每列前 5 个非空"，把数据值也收进来了 ——
    # 合同表"设备名称"列的 数控加工中心/激光切割机/工业机器人 全成了候选，
    # 语义排名被这些值挤歪（"合同总价" 排到"合同编号"上）。
    groups = {}
    for zone, items in cols.items():
        groups.setdefault((zone[0], zone[1]), {})[zone[2]] = items
    body = set()                      # 表格正文里的值，补候选时要排除
    for g in groups.values():
        firsts = [min(r for r, _t in it) for it in g.values() if it]
        hrow = Counter(firsts).most_common(1)[0][0] if firsts else 0
        for it in g.values():
            for r, t in it:
                if r > hrow:
                    body.add(semantic.norm(t))
        for it in g.values():
            for r, t in sorted(it)[:5]:    # 每列前 5 个非空，但只认表头行及其上方
                if r <= hrow:
                    add(t)

    for t in labels:
        add(t)

    # 不足时用短段落补齐（有些文档没有表格，只有"甲方：XX"这种行）
    if len(cands) < 20:
        for t in doc_element_list(file_name, limit=120):
            if semantic.norm(t) in body:      # 表格正文的数据值不是字段名
                continue
            add(t)

    return cands[:limit]


def normalize_fields(fields: list, candidates: list = None, aliases: dict = None) -> list:
    """字段规范化：合称拆分 + 别名表 + 语义对齐。

    返回 [{"orig", "field", "note", "semantic": [(说法, 分数), ...]}]

    semantic 是"这个字段在文档里还可能叫什么"的候选，留给后面抽空时按顺序重试
    ——都是本地向量算的，不花钱，所以可以一次算好备用。

    顺序有讲究（2026-10-26 定）：
      ① 合称拆分 → ② 别名表 → ③ 语义对齐
      前两步是"确定的规则"，语义只在前两步没结论时改写。
      而且语义改写只在**原词根本不在文档里**时才做：
      原词在文档里出现过，就按原词抽（用户说的就是文档里的词，别自作聪明）。
    """
    if aliases is None:
        aliases = load_aliases()
    cands = [str(c) for c in (candidates or []) if str(c).strip()]

    out = []
    for f in fields:
        f = str(f).strip()
        if not f:
            continue

        # 1) 合称拆分
        done = False
        for k, reps in SPLIT_RULES.items():
            if k in f:
                for r in reps:
                    out.append({"orig": f, "field": f.replace(k, r),
                                "note": "「%s」拆分 → 「%s」" % (f, f.replace(k, r)),
                                "semantic": []})
                done = True
                break
        if done:
            continue

        field, note = f, ""

        # 2) 别名表（人工确认过的，优先级最高）
        if f in aliases:
            field = aliases[f]
            note = "别名表：「%s」→「%s」" % (f, field)

        # 3) 语义对齐
        sem_list = []
        if cands:
            exact = next((c for c in cands if semantic.lexical(f, c) >= 0.999), None)
            if exact:
                if not note:                    # 文档里就有这个说法，直接用
                    field = exact
            else:
                res = semantic.match(f, cands, topk=3, aliases=aliases,
                                     min_score=semantic.SEM_WEAK, tag="doc-terms")
                sem_list = [(r["term"], r["score"]) for r in res]
                # 改写（把用户说法直接换成文档说法）要很克制：
                #   ① 相似度 ≥ 0.85（"入职时间→入职日期" 0.92、"设备名字→设备名称" 0.98）
                #   ② 与第 2 名拉开 ≥ 0.15（唯一性：不是"两个都像"）
                # 只满足 ① 不满足 ② 的典型反例：
                #   "库存数量"→期初库存(0.75)/期末库存(0.72)、"完成度"→计划完成(0.71)/进度(0.71)
                #   —— 这种改写会抽错列，比不改还糟。此时只把它放进候选，
                #   等直抽为空时再逐个试，让模型按原文判。
                # 【0.15 这条线是实测调出来的】0.10 时"岗位工资→岗位"(领先 0.12)会漏过去，
                #   而"岗位工资"在薪酬表里指的是"岗位津贴"，改了就抽错列。
                if not note and len(res) >= 2:
                    margin = res[0]["score"] - res[1]["score"]
                    if res[0]["score"] >= semantic.SEM_REWRITE and margin >= 0.15 \
                            and res[0]["via"] in ("semantic", "lexical"):
                        field = res[0]["term"]
                        note = "语义对齐：「%s」→ 文档里的「%s」（相似度 %.2f，领先第 2 名 %.2f）" % (
                            f, field, res[0]["score"], margin)
                elif not note and len(res) == 1:
                    if res[0]["score"] >= semantic.SEM_REWRITE \
                            and res[0]["via"] in ("semantic", "lexical"):
                        field = res[0]["term"]
                        note = "语义对齐：「%s」→ 文档里的「%s」（相似度 %.2f）" % (
                            f, field, res[0]["score"])

        out.append({"orig": f, "field": field, "note": note, "semantic": sem_list})
    return out


# ==================== 二、语义翻译（直抽为空时的兜底） ====================

def ask_model(system: str, user: str):
    """调大模型一次，返回 (回答文本, 耗时秒)。失败返回 (None, 耗时)。"""
    content, _usage, cost, _err = _llm_chat(
        [{"role": "system", "content": system},
         {"role": "user", "content": user}],
        timeout=60)
    return content, cost


def doc_element_list(file_name: str, limit: int = 60) -> list:
    """从 output.json 拿该文档的"要素清单"（短的文本块＝标题/表头/标签）。"""
    if not OUTPUT_JSON.exists():
        return []
    try:
        data = json.loads(OUTPUT_JSON.read_text(encoding="utf-8"))
    except Exception:
        return []
    for d in data.get("documents", []):
        if d.get("file") != file_name:
            continue
        terms, seen = [], set()
        for b in d.get("blocks", []):
            t = re.sub(r"\s+", " ", (b.get("text") or "")).strip()
            if not t or len(t) > 20 or t in seen:
                continue
            seen.add(t)
            terms.append(t)
            if len(terms) >= limit:
                break
        return terms
    return []


def translate_field(field: str, doc_file: str, elements: list, exclude=None) -> list:
    """把字段翻译成文档里实际出现的说法。返回候选清单（照抄要素清单）。

    exclude：已试过但未命中的说法——追问轮里排除掉，避免重复试同一个词。
    """
    if not elements:
        return []
    ex = [str(x).strip() for x in (exclude or []) if str(x).strip()]
    ex_line = ""
    if ex:
        ex_line = ("已经试过、但都没提取到信息的说法（请不要再选）：%s\n"
                   % "、".join("「%s」" % x for x in ex))
    prompt = (
        "用户想从文档《%s》里提取【%s】的信息。\n\n"
        "这份文档里出现过的要素清单（标题、表头、关键短语）：\n%s\n\n"
        "请判断：【%s】在这份文档里最可能以什么说法出现？\n"
        "如果要提取的是逐条的数据（名单、清单、多条记录），优先挑数据所在表格的"
        "列头/字段名，而不是章节标题。\n"
        "%s"
        "从清单里挑出 1~3 个含义最接近的说法（按接近程度排序，照抄清单里的原文，"
        "不许改字）；若清单里确实没有含义接近的说法，terms 返回空数组。\n"
        '只输出 JSON：{"terms": ["说法1", "说法2"]}'
        % (doc_file, field, "\n".join("- " + e for e in elements), field, ex_line))
    content, _cost = ask_model(TRANSLATE_SYSTEM, prompt)
    if not content:
        return []
    t = content.strip()
    i, j = t.find("{"), t.rfind("}")
    if i == -1 or j == -1:
        return []
    try:
        terms = json.loads(t[i:j + 1]).get("terms") or []
    except Exception:
        return []
    return [str(x).strip() for x in terms if str(x).strip()]


# ==================== 三、抽一个字段（直抽 + 翻译兜底） ====================

def extract_field(field: str, doc_file: str, notes: list, sem_candidates=None) -> dict:
    """抽一个字段。直抽为空时，先用**语义候选**重试（本地、免费、快），
    再不行才让大模型翻译（远程、有成本、慢）。

    为什么把顺序倒过来（2026-10-26）：
      原来直抽一空就去调模型翻译，一次 3~8 秒、还要花钱，而且同一句话
      两次可能给出不同说法（抖动）。语义候选是本地向量算的，0.3 秒出 Top3，
      稳定、可复现、断网也能用 —— 大模型只留作最后的"兜底"。

    返回 {"field", "items", "mapped_to", "elapsed", "via"}
        via:  ""（原词直接抽到）｜"semantic"（语义候选命中）｜"model"（模型翻译命中）
    """
    t0 = time.time()
    res = matcher_extract.run_column(field, quiet=True, only_file=doc_file,
                                     out_dir=str(MID_DIR))
    items = list((res or {}).get("items", []))
    mapped_to, via = "", ""

    if not items:
        # ---------- ① 语义候选（本地向量，不花钱） ----------
        tried = []
        for term, sc in (sem_candidates or [])[:3]:
            term = str(term).strip()
            if not term or term == field or term in tried:
                continue
            tried.append(term)
            res2 = matcher_extract.run_column(term, quiet=True, only_file=doc_file,
                                              out_dir=str(MID_DIR))
            items2 = list((res2 or {}).get("items", []))
            if items2:
                res, items, mapped_to, via = res2, items2, term, "semantic"
                notes.append("「%s」语义候选命中：「%s」（相似度 %.2f）"
                             % (field, term, sc))
                break
        if not mapped_to and tried:
            notes.append("「%s」的语义候选 %s 都没抽到值 → 交给模型翻译"
                         % (field, "、".join("「%s」" % x for x in tried)))

    if not items:
        # ---------- ② 大模型翻译（最后一道兜底） ----------
        elements = doc_element_list(doc_file)
        tried = []
        if elements:
            # 第一轮：翻译候选（最多 3 个，逐个试）
            terms = [x for x in translate_field(field, doc_file, elements)
                     if x and x != field][:3]
            if terms:
                notes.append("「%s」字面在文档里没找到 → 对照文档要素翻译，试：%s"
                             % (field, "、".join("「%s」" % x for x in terms)))
            for term in terms:
                res2 = matcher_extract.run_column(term, quiet=True, only_file=doc_file,
                                                  out_dir=str(MID_DIR))
                items2 = list((res2 or {}).get("items", []))
                tried.append(term)
                if items2:
                    res, items, mapped_to, via = res2, items2, term, "model"
                    break
            # 第二轮：首轮挑偏未命中 → 请模型换说法再试（最多 2 个，排除已试词）
            if not mapped_to and tried:
                terms2 = [x for x in translate_field(field, doc_file, elements, exclude=tried)
                          if x and x != field and x not in tried][:2]
                if terms2:
                    notes.append("首轮未命中 → 再换说法试：%s"
                                 % "、".join("「%s」" % x for x in terms2))
                for term in terms2:
                    res2 = matcher_extract.run_column(term, quiet=True, only_file=doc_file,
                                                      out_dir=str(MID_DIR))
                    items2 = list((res2 or {}).get("items", []))
                    if items2:
                        res, items, mapped_to, via = res2, items2, term, "model"
                        break
        if mapped_to:
            notes.append("「%s」已按文档里的说法「%s」提取成功（模型翻译）"
                         % (field, mapped_to))
        else:
            notes.append("「%s」：文档里没有找到对应信息（诚实留空，不编造）" % field)

    return {"field": field, "items": items, "mapped_to": mapped_to, "via": via,
            "elapsed": round((res or {}).get("elapsed_sec", time.time() - t0), 2)}


# ==================== 四、整理输出（配对表 / 分列列表） ====================

def row_key(item: dict):
    """从出处解析"行锚"：用于把多个字段按行对齐成表。解析不出返回 None。"""
    loc = item.get("loc_text") or ""
    m = re.search(r"工作表「([^」]+)」\s*[A-Z]+(\d+)", loc)      # xlsx 单元格
    if m:
        return ("xlsx:" + m.group(1), int(m.group(2)))
    m = re.search(r"表格(\d+)\s*·\s*行(\d+)", loc)               # docx 表格
    if m:
        return ("docx表" + m.group(1), int(m.group(2)))
    m = re.search(r"第\s*(\d+)\s*行", loc)                       # md/txt 行号
    if m:
        return ("行", int(m.group(1)))
    m = re.search(r"第\s*(\d+)\s*段", loc)                       # docx 段落
    if m:
        return ("段", int(m.group(1)))
    if item.get("line"):
        return ("行", int(item["line"]))
    return None


def build_pairs(fields_kept: list):
    """能按行配对时返回配对表（[行值列表, ...]），否则 None。

    条件：≥2 个字段、每字段 ≥2 条、每条都有行锚、且行锚属于同一"区域"（同表/同行体系）。
    """
    if len(fields_kept) < 2:
        return None
    if any(len(f["items"]) < 2 for f in fields_kept):
        return None
    zone = None
    cols = []
    for f in fields_kept:
        keys = []
        for it in f["items"]:
            k = row_key(it)
            if k is None:
                return None
            if zone is None:
                zone = k[0]
            elif k[0] != zone:
                return None
            keys.append(k[1])
        cols.append(keys)
    rows = sorted(set(r for ks in cols for r in ks))
    table = []
    for r in rows:
        row = []
        for f, ks in zip(fields_kept, cols):
            val = "—"
            for it, rr in zip(f["items"], ks):
                if rr == r:
                    val = str(it["value"])
                    break
            row.append(val)
        table.append(row)
    return table


def _loc_of(item: dict) -> str:
    """条目出处的可读展示。"""
    return item.get("loc_text") or (("第 %d 行/段" % item["line"]) if item.get("line") else "未定位")


def format_screen(results: list, notes: list, cmd: str, doc: str, total: float) -> str:
    """屏幕报告。"""
    lines = []
    lines.append("=" * 62)
    lines.append("【提取】一句话 → 信息整理出来")
    lines.append("=" * 62)
    lines.append("指令：%s" % cmd)
    lines.append("文档：%s" % doc)
    lines.append("")
    n_all = sum(len(f["items"]) for f in results)
    lines.append("【结果】%d 个字段，共 %d 条信息（用时 %.1f 秒）"
                 % (len(results), n_all, total))
    lines.append("")

    pairs = build_pairs(results)
    if pairs:
        lines.append("----- 配对整理（按源文档行对齐） -----")
        lines.append("  " + " ｜ ".join(f["field"] for f in results))
        for row in pairs:
            lines.append("  " + " ｜ ".join(row))
        lines.append("")

    lines.append("----- 字段明细 -----")
    for f in results:
        title = f["field"]
        if f.get("mapped_to"):
            title += "（映射为文档中的「%s」）" % f["mapped_to"]
        lines.append("◆ %s ｜ %d 条" % (title, len(f["items"])))
        if not f["items"]:
            lines.append("   （文档里没有找到对应信息，诚实留空）")
        for i, it in enumerate(f["items"], start=1):
            lines.append("   %d. %s ｜ 出处：%s" % (i, it["value"], _loc_of(it)))
        lines.append("")
    if notes:
        lines.append("----- 说明 -----")
        for n in notes:
            lines.append("  · " + n)
    lines.append("=" * 62)
    return "\n".join(lines)


def format_md(results: list, notes: list, cmd: str, doc: str,
              total: float, tstamp: str) -> str:
    """Markdown 版结果（给文件）。"""
    L = []
    L.append("# 内容提取结果")
    L.append("")
    L.append("- **指令**：%s" % cmd)
    L.append("- **文档**：%s" % doc)
    L.append("- **时间**：%s ｜ 用时 %.1f 秒" % (tstamp, total))
    L.append("- **字段**：%s" % "、".join(f["field"] for f in results))
    L.append("")

    pairs = build_pairs(results)
    if pairs:
        L.append("## 配对整理（按源文档行对齐）")
        L.append("")
        L.append("| " + " | ".join(f["field"] for f in results) + " |")
        L.append("|" + "---|" * len(results))
        for row in pairs:
            L.append("| " + " | ".join(row) + " |")
        L.append("")

    L.append("## 字段明细")
    L.append("")
    for f in results:
        title = f["field"]
        if f.get("mapped_to"):
            title += "（映射为文档中的「%s」）" % f["mapped_to"]
        L.append("### %s（%d 条）" % (title, len(f["items"])))
        L.append("")
        if not f["items"]:
            L.append("_文档里没有找到对应信息，诚实留空。_")
            L.append("")
            continue
        L.append("| # | 值 | 出处 |")
        L.append("|---|---|---|")
        for i, it in enumerate(f["items"], start=1):
            val = str(it["value"]).replace("|", "\\|")
            L.append("| %d | %s | %s |" % (i, val, _loc_of(it)))
        L.append("")
    if notes:
        L.append("## 说明")
        L.append("")
        for n in notes:
            L.append("- " + n)
        L.append("")
    return "\n".join(L)


# ==================== 五、主流程 ====================

def run_extract(text: str, intent: dict, src: Path, out_dir: Path = None):
    """执行一条"内容提取"指令。返回 (是否成功, 报告文本, 产出文件列表)。"""
    out_dir = Path(out_dir) if out_dir else OUT_DIR
    out_dir.mkdir(parents=True, exist_ok=True)
    notes = []
    t0 = time.time()

    raw_fields = (intent.get("params") or {}).get("fields") or []
    # 语义匹配的候选空间 = 这份文档里真实出现的"说法"（列头 + 标签）
    cands = doc_candidates(src.name)
    norm = normalize_fields(raw_fields, candidates=cands)
    if not norm:
        return False, "【失败】指令里没有说明要提取哪些字段，无法执行。", []

    # 字段规范化的说明（拆分/别名/语义对齐才有）
    for f in norm:
        if f["note"]:
            notes.append(f["note"])

    results = []
    for f in norm:
        r = extract_field(f["field"], src.name, notes,
                          sem_candidates=f.get("semantic"))
        r["orig"] = f["orig"]
        results.append(r)

    total = time.time() - t0
    tstamp = time.strftime("%Y-%m-%d %H:%M")

    screen = format_screen(results, notes, text, src.name, total)
    md = format_md(results, notes, text, src.name, total, tstamp)

    md_path = out_dir / (src.stem + "_提取结果.md")
    json_path = out_dir / (src.stem + "_提取结果.json")
    md_path.write_text(md, encoding="utf-8")
    json_path.write_text(json.dumps({
        "cmd": text, "doc": src.name, "time": tstamp,
        "elapsed_sec": round(total, 2),
        "fields": results, "notes": notes,
    }, ensure_ascii=False, indent=2), encoding="utf-8")

    screen += "\n【输出】%s\n【输出】%s\n【原件】未改动（提取是只读操作）" % (md_path, json_path)
    return True, screen, [str(md_path), str(json_path)]


def main():
    """命令行直跑：解析 + 提取一条指令（不经 edit.py 时的入口）。"""
    import argparse
    from parse_intent import scan_docs, parse_one

    ap = argparse.ArgumentParser(description="模块（1）· 内容提取执行器")
    ap.add_argument("text", help="一条内容提取指令")
    ap.add_argument("--docs", default=str(BASE / "测试数据" / "源文档"), help="文档库目录")
    args = ap.parse_args()

    docs_dir = Path(args.docs)
    doc_names = scan_docs(docs_dir)
    intent, cost, ok = parse_one(args.text, doc_names)
    if not ok:
        print("解析失败（网络或格式问题），请重跑一次。")
        sys.exit(1)
    print("解析（%.1fs）：%s" % (cost, json.dumps(intent, ensure_ascii=False)))
    if intent.get("action") != "内容提取":
        print("这不是一条「内容提取」指令（action=%s）。" % intent.get("action"))
        sys.exit(1)

    # 找文档（和 edit.py 一致的宽松匹配）
    doc_name = intent.get("doc") or intent.get("doc_hint") or ""
    src = None
    for p in docs_dir.rglob("*"):
        if p.is_file() and (p.name == doc_name or p.stem == Path(doc_name).stem):
            src = p
            break
    if src is None:
        print("文档库里找不到：%s" % doc_name)
        sys.exit(1)

    ok, report, files = run_extract(args.text, intent, src)
    print(report)
    sys.exit(0 if ok else 1)


if __name__ == "__main__":
    main()
