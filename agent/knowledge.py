# -*- coding: utf-8 -*-
"""
knowledge.py —— 模块（2）· 统一知识抽取（2026-11-05）

为什么要重做
------------
老的三条线（画像 / 实体 / 字段）各自背着一套**写死的分类**：
  · 画像：主题 + 要点 + 要素（固定三个筐）
  · 实体：人名 / 机构 / 日期 / 金额 / 编号 / 联系方式 / 产品设备（固定 7 类）
  · 字段：用户点名要哪几个字段
换一个行业，固定分类就装不下：检验报告要"检验项目 / 结果 / 参考区间"，
判决书要"案号 / 原告 / 被告 / 诉讼请求"，运单要"运单号 / 寄件人 / 收件人 / 计费重量"。
这些既不是"人名机构"，也不是"关键要素"，用户还得先想清楚该点哪条线 —— 认知成本高、召回还低。

这一版改成「先读懂文档自己的结构，再按它的结构抽」：
  ① 结构扫描（零模型）：标题 / 章节 / 表格列头 / "XX：" 标签 / 样例行
  ② 结构归纳（1 次模型调用）：这份文档属于什么领域、在讲哪些**业务对象**、
     每个对象有哪些**属性**、文档目的与关键结论。属性名必须能在骨架里找到，
     找不到的一律丢弃 —— 不套行业模板，也不许模型拿常识补。
  ③ 抽值：
       · 表格 → 程序直接读「列头 × 单元格」，不调模型（这一类天然零误差）
       · 叙述 → 模型按归纳出的属性抽，每条必须带原文摘抄
  ④ 程序复核：值回原文定位（原样 / 归一 / 数字边界），定位不上的降级或丢弃
  ⑤ 统一入库：一条条 knowledge_items（对象 / 属性 / 关系 / 事件 / 要点），
     外加每份文档自己的 doc_schemas（这就是它长出来的"行业词表"）

对用户只有一个概念：**这份文档里有哪些对象、各自什么属性、值是多少、出自哪里**。
实体、画像、字段不再作为三个入口；它们是同一个知识视图里的不同 kind。

跑法（项目根目录）：
  python agent/knowledge.py "自动读懂设备采购合同并提取关键信息入库"
  python agent/knowledge.py "..." --requirement "我最关心合同金额和交付期"
"""
import json
import math
import os
import re
import sys
import time
from collections import Counter
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

try:
    sys.stdout.reconfigure(errors="replace")
except Exception:
    pass

BASE = Path(__file__).resolve().parents[1]
AGENT_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(AGENT_DIR))

from parse_intent import ask_model, parse_json_reply                      # noqa: E402
from summary_ops import (load_doc, pick_for_budget, build_digest,          # noqa: E402
                         locate_quote, loc_desc, _norm, _split_parts, _find_part)

try:
    import semantic                                                       # noqa: E402
except Exception:                                                         # 语义模型缺失不致命
    semantic = None

OUT_DIR = BASE / "agent" / "输出"

SCHEMA_BUDGET = 4200        # 结构骨架送模型的字符预算（骨架不是全文，够判断结构即可）
TEXT_BUDGET = 12000         # 叙述体抽值的字符预算
MAX_TABLE_ROWS = int(os.environ.get("A23_KNOWLEDGE_MAX_ROWS") or 2000)
MAX_TEXT_ITEMS = 120
# 叙述体抽值分批（2026-11-05 加）。一次喂 30 多段让它"一省一条"地抽，
# 模型会偷懒只写几句总结（实测：31 个省只抽到 1 条属性 + 4 条要点）。
# 切成小批之后它才肯逐条照抄。
#
# 2026-11-05 实测（同一份 31 段的新冠疫情记叙文，每省 7~8 个属性）：
#   8 块/批 × 4 批 = 45 秒
#   16 块/批 × 2 批 = 36 秒   ← 甜点，用它
#   31 块/批 × 1 批 = 51 秒（一次吐 170 条，单次生成太久，反而更慢）
#
# 关键认知：当前平台（智谱）是**单并发**，批再多也是排队串行，
# 每批还有十来秒固定开销 —— 所以"批数少"比"批小"更重要。
# 但批太大模型会偷懒（早期 31 段一次喂进去只肯写几句总结），
# 所以取 16 块为上限，段数再多就靠 TEXT_MAX_CHUNKS 兜住批数。
TEXT_CHUNK_BLOCKS = 16      # 每批最多几块（模型不偷懒的上限）
TEXT_MAX_CHUNKS = 4         # 批数上限：段数多时按此数平均分，别串太久
TEXT_WORKERS = 4            # 并发批数（换到支持并发的平台时才有用）

KIND_ORDER = ["object", "attribute", "relation", "event", "summary"]
KIND_CN = {"object": "对象", "attribute": "属性", "relation": "关系",
           "event": "事件", "summary": "要点"}

# 结构列：能区分记录，但不是"信息属性"
_ID_HEADERS = ["编号", "工号", "学号", "单号", "案号", "序号", "id", "code",
               "名称", "姓名", "项目", "商品", "物料", "产品", "运单号"]
_STRUCT_HEADERS = {"序号", "编号", "工号", "学号", "id", "no", "no."}
# 表格里这些列只是"第几条"，不是信息属性（编号/工号可能是身份列，另有处理，不在此列）
_SKIP_ATTR = {"序号", "no", "no.", "#", "行号"}

# 标题式的行：不是表头（"员工信息表（2026年第一季度）"）
_TITLEISH = re.compile(r"(表|清单|台账|明细|报表|记录|统计|名册|汇总|一览)[）)]?$")

SCHEMA_SYSTEM = (
    "你是文档结构分析器。你的任务不是套用任何行业模板，而是先读懂这份文档，"
    "再把它自身的信息结构说出来：文档在讲哪些业务对象、每个对象有哪些属性。"
    "所有属性名和判断依据必须来自给定的文档骨架，不许用行业常识补、不许编造。"
    "严格只输出 JSON，不要任何解释。"
)

EXTRACT_SYSTEM = (
    "你是文档信息抽取器。你已经拿到这份文档归纳出的对象与属性清单，"
    "现在只负责把对应的值从原文里找出来。值必须逐字照抄原文，"
    "文档里没有的绝不推测、绝不补全。严格只输出 JSON，不要任何解释，"
    "也不要输出 JSON 之外的任何文字。"
)


# ====================================================================
#  一、结构扫描（零模型）
# ====================================================================

def _cell_text(v) -> str:
    return re.sub(r"\s+", " ", str(v or "")).strip()


def _is_titleish(t: str) -> bool:
    t = (t or "").strip()
    return len(t) >= 8 and bool(_TITLEISH.search(t))


def _row_score(row: dict) -> int:
    """一行像不像表头：短、非数字、非标题句。"""
    n = 0
    for c, v in row.items():
        v = _cell_text(v)
        if not v or len(v) > 14:
            continue
        if re.fullmatch(r"[\d.,%％+\-—/年月日:：()（）]+", v):
            continue
        if _is_titleish(v):
            continue
        n += 1
    return n


def group_tables(blocks: list) -> list:
    """把块按「同一张表」分组，并认出表头行。

    返回 [{"zone":(kind, sheet/table_index), "header_row", "header_rows",
           "headers":{col:name}, "rows":{row:{col:text}}, "blocks":{(r,c):block},
           "n_data":int}]
    """
    groups = {}
    for b in blocks:
        loc = b.get("locator") or {}
        kind = loc.get("kind")
        if kind not in ("cell", "table_cell"):
            continue
        r, c = loc.get("row"), loc.get("col")
        if r is None or c is None:
            continue
        zone = (kind, loc.get("sheet") if kind == "cell" else loc.get("table_index"))
        g = groups.setdefault(zone, {"cells": {}, "blocks": {}})
        g["cells"].setdefault(int(r), {})[int(c)] = _cell_text(b.get("text"))
        g["blocks"][(int(r), int(c))] = b

    out = []
    for zone, g in groups.items():
        cells = g["cells"]
        if not cells:
            continue
        max_row = max(cells)
        # 候选表头行：前 5 行里"最像表头"的那一行（不是简单取第一行 ——
        # 表格上面常压着一行大标题，取第一行会把标题当列头、把真列头当数据）
        best_r, best_s = None, -1
        for r in range(1, min(5, max_row) + 1):
            s = _row_score(cells.get(r) or {})
            if s > best_s:
                best_r, best_s = r, s
        if not best_r or best_s <= 0:
            continue
        # 多层表头：表头行上方那些"也像表头"的行一起算进来
        header_rows = [r for r in range(1, best_r + 1)
                       if _row_score(cells.get(r) or {}) >= max(2, best_s // 2)]
        if best_r not in header_rows:
            header_rows.append(best_r)

        cols = set()
        for r in range(1, max_row + 1):
            cols |= set((cells.get(r) or {}).keys())

        headers = {}
        for c in sorted(cols):
            parts = []
            for r in sorted(header_rows):
                v = _cell_text((cells.get(r) or {}).get(c))
                if not v or _is_titleish(v):
                    continue
                if parts and (v == parts[-1] or v in parts[-1] or parts[-1] in v):
                    continue
                parts.append(v)
            if parts:
                headers[c] = "·".join(parts[:2])

        rows = {r: d for r, d in cells.items() if r > best_r and any(v for v in d.values())}
        if not headers or not rows:
            continue
        # 版面式表格：一个单元格里塞了多条数据（政府公报/PDF 转 Word 常见）。
        # 这类表"列头 × 单元格"根本不是一一对应，硬读会得到
        # "数量：34 4 23 5 2 2843 …" 这种把整列糊在一起的值。
        # 宁可明确不抽（并说明原因），也不往库里写垃圾。
        layout = False
        for d in rows.values():
            for v in d.values():
                toks = [x for x in v.split() if 1 <= len(x) <= 12]
                if len(toks) >= 4:
                    layout = True
                    break
            if layout:
                break
        out.append({"zone": zone, "header_row": best_r, "header_rows": header_rows,
                    "headers": headers, "rows": rows, "blocks": g["blocks"],
                    "n_data": len(rows), "layout": layout})
    return out


def _table_title(blocks: list, table: dict) -> str:
    """给表起个名字：取表前面最近的一块短文本（常是"表1 设备清单"）。"""
    first = table["blocks"].get((table["header_row"], min(table["headers"])))
    if first is None:
        return ""
    try:
        i = blocks.index(first)
    except ValueError:
        return ""
    for j in range(i - 1, max(-1, i - 4), -1):
        t = _cell_text(blocks[j].get("text"))
        if 2 <= len(t) <= 24:
            return t
    return ""


def build_skeleton(blocks: list, doc_file: str, doc_type: str = "") -> dict:
    """文档骨架：章节、表格列头与样例、标签行、开头。全零模型。"""
    headings, labels, plain = [], [], []
    for b in blocks:
        t = _cell_text(b.get("text"))
        if not t:
            continue
        loc = b.get("locator") or {}
        if loc.get("kind") in ("cell", "table_cell"):
            continue
        if b.get("is_heading") or b.get("style") in ("Title", "Heading1", "Heading2", "Heading3"):
            headings.append(t[:60])
            continue
        if re.match(r"^第\s*[一二三四五六七八九十百千\d]+\s*[条章节款项]", t):
            headings.append(t[:60])
            continue
        for line in t.split("\n"):
            line = line.strip()
            if not line:
                continue
            m = re.match(r"^([\u4e00-\u9fa5A-Za-z（）()·]{2,12})\s*[：:]", line)
            if m:
                labels.append(m.group(1).strip())
            if len(plain) < 8 and len(line) >= 10:
                plain.append(line[:160])

    tables = group_tables(blocks)
    for t in tables:
        t["title"] = _table_title(blocks, t)
    terms, seen = [], set()

    def add_term(t):
        t = _cell_text(t)
        if not t or len(t) < 2 or len(t) > 16:
            return
        k = _norm(t)
        if k and k not in seen:
            seen.add(k)
            terms.append(t)

    for h in headings:
        add_term(re.sub(r"^第[一二三四五六七八九十百千\d]+[条章节款项部分]\s*", "", h)[:16])
    for l in labels:
        add_term(l)
    for t in tables:
        for c, name in t["headers"].items():
            add_term(name)

    # 骨架文本（给模型看，控制在预算内）
    L = ["【基本信息】文件：%s ｜ 格式：%s ｜ 块数：%d" % (doc_file, doc_type or "?", len(blocks))]
    if headings:
        L.append("【标题/章节】" + " ｜ ".join(headings[:18]))
    for i, t in enumerate(tables[:6], start=1):
        title = _table_title(blocks, t)
        names = [t["headers"][c] for c in sorted(t["headers"])]
        L.append("【表格%d】%s列头：%s" % (i, ("「%s」" % title) if title else "",
                                        " | ".join(names[:16])))
        if t.get("layout"):
            L.append("    ※ 版面式：一个单元格里有多条数据，不要按列头×单元格一一对应去理解")
        sample = list(sorted(t["rows"]))[:3]
        for r in sample:
            row = t["rows"][r]
            vals = [_cell_text(row.get(c))[:18] for c in sorted(t["headers"])]
            if any(vals):
                L.append("    样例：%s" % " | ".join(vals))
        L.append("    数据行数：%d" % t["n_data"])
    if labels:
        L.append("【标签】" + " ｜ ".join(dict.fromkeys(labels))[:600])
    if plain:
        L.append("【开头原文】")
        for p in plain[:6]:
            L.append("  " + p)
    # 抽样正文：叙述体文档光看章节名不够，给模型一段正文才能判断"有哪些属性"
    # （表格类文档到上面就够了，这里只为补齐叙述体的结构线索）
    nar = [b for b in blocks if (b.get("locator") or {}).get("kind") not in ("cell", "table_cell")
           and (b.get("text") or "").strip()]
    if nar and not tables:
        idx, _note = pick_for_budget(nar, budget=1500)
        L.append("【正文抽样】")
        for i, j in enumerate(idx[:24], start=1):
            t = _cell_text(nar[j].get("text"))[:180]
            if t:
                L.append("  [%d] %s" % (i, t))
    text = "\n".join(L)
    if len(text) > SCHEMA_BUDGET:
        text = text[:SCHEMA_BUDGET] + "\n…（骨架过长，已截断）"

    return {"text": text, "terms": terms, "headings": headings,
            "labels": labels, "tables": tables, "n_chars": sum(len(b.get("text") or "") for b in blocks)}


# ====================================================================
#  二、结构归纳（1 次模型调用）
# ====================================================================

def build_schema_prompt(doc_file: str, sk: dict, requirement: str) -> str:
    req = requirement.strip() or "（用户没有额外说明，按文档本身最重要、最能代表它的信息来）"
    return f"""请先读懂下面这份文档的骨架，再归纳它自身的信息结构。

【文档骨架】
{sk['text']}

【用户关注点】
{req}

【输出 JSON】（严格按这个结构）
{{
  "domain": "这份文档属于哪个业务领域（从内容判断，例如：设备采购/民事诉讼/医学检验/公路货运/中学教务；不要写'通用'或'其他'）",
  "doc_type": "文档类型（例如：采购合同/判决书/检验报告/运单/成绩单/统计公报）",
  "purpose": "这份文档用来做什么（一句话，20~50 字）",
  "objects": [
    {{
      "type": "业务对象类型（用文档里的说法，如 合同、设备、检验项目、被告、运单、课程）",
      "why": "判断依据（引用骨架里的原文）",
      "records": "single=整份文档只有一个（如一份合同）；many=表格里有很多条（如每行一个员工）",
      "identity": "用来区分每条记录的列名/字段名（records=many 时必填，必须是列头原文）",
      "attrs": [
        {{"name": "属性名（必须来自骨架：列头/标签/章节名）", "evidence": "骨架里支持它的原文", "unit": "单位，没有就空字符串"}}
      ]
    }}
  ],
  "key_points": [
    {{"text": "这份文档最值得记住的结论/要点（一句话）", "quote": "支撑它的原文摘抄（必须逐字来自骨架）"}}
  ],
  "focus": ["把用户关注点对应到本文档的属性名；用户没提就空数组"]
}}

规则（很重要，请逐条遵守）：
- **不要套行业模板**。只看这份文档本身有什么，不要因为它是"合同"就硬塞甲乙方，也不要因为它是"报告"就硬塞结论。
- **属性名要写成"能装值的名词短语"**（如 判决结果、合同金额、签发日期），不要把动词性行文
  （"判决如下"、"经审理查明"、"详见附件"、"特此公告"）直接当属性名 —— 这类措辞只作为 evidence。
- 属性名优先用骨架里的说法；骨架里只有"判决如下""甲方（采购方）"这类**行文措辞**时，
  可以用更规范的业务叫法（如 判决结果、甲方），但必须把那一行原文写进 evidence。
- 文档头部**没有 "XX：" 标签**的信息（案号、报告编号、文号、签发日期等）也要抽成属性，
  evidence 就填那一行原文（例如"（2026）苏01民初1234号"）。
- 属性名和 evidence 至少要有一个能在骨架里逐字找到；两个都找不到的宁可不写。
- 表格的列头就是属性，尽量列全（"序号"这类纯结构列可以不要）。
- 一份文档可能有多个对象（如"合同"和它附的"设备清单"），都要列。
- **文档里用显式角色标签分块的，一个角色一个对象**（原告：/被告：/甲方：/乙方：/买方：/卖方：/
  患者：/委托方：等）。不要合并成"当事人""双方""甲乙双方"这类概括对象 ——
  一合并就把"是哪一方、叫什么名字"丢了，后面既查不出原告是谁，也没法按角色取值。
- records=many 时，identity 必须是列头里真实存在的那个（如 工号 / 员工姓名 / 运单号）。
- key_points 给 2~5 条；每条 quote 必须逐字来自骨架，不许改写。
- 文档里没有的信息绝对不要虚构。"""


def _match_term(name: str, terms: list) -> str:
    """属性名是否真的出现在骨架里。返回命中的骨架说法，没命中返回空串。

    先精确（归一后相等/互相包含），再语义近似（≥0.88，要求很高）——
    只放"合同总价 ↔ 合同总金额"这种同义改写，不放"金额 ↔ 单价"这种近邻。
    """
    n = _norm(name)
    if not n:
        return ""
    for t in terms:
        tn = _norm(t)
        if not tn:
            continue
        if n == tn or (len(n) >= 2 and n in tn) or (len(tn) >= 2 and tn in n):
            return t
    if semantic is not None:
        best, best_s = "", 0.0
        for t in terms:
            s = semantic.lexical(name, t)
            if s > best_s:
                best, best_s = t, s
        if best_s >= 0.88:
            return best
    return ""


def _evidence_ok(evidence: str, skeleton_text: str) -> bool:
    """模型给的支持依据，是否真的在骨架里出现过（归一后子串）。

    叙述体文档没有"列头"这种东西，属性名往往是概括出来的（如"负责人"），
    但它的依据必须能在骨架里找到（如"**汇报人**：吴晓敏"）。
    这一条是"允许概括、禁止编造"的分界线。

    允许"抄长/带省略号"：模型经常把整段判词当 evidence（骨架只喂了 4200 字，
    抄全了反而对不上），所以除了整串比，还退到几个递减前缀和按标点切开的最长片段。
    底线不变 —— 至少要**连续 10 个字**在原文里逐字出现，否则算编造。
    """
    ev = _norm(evidence)
    if len(ev) < 2:
        return False
    sks = _norm(skeleton_text)
    if ev in sks:
        return True
    for n in (80, 60, 40, 24, 16, 10):
        if len(ev) >= n and ev[:n] in sks:
            return True
    for part in re.split(r"[，,。；;：:、\s（）()\[\]【】]+", ev):
        if len(part) >= 10 and part in sks:
            return True
    return False


def clean_schema(raw: dict, sk: dict, doc_file: str) -> dict:
    """清洗模型给的结构：属性名回骨架核对，核对不上的丢掉并计数。"""
    objects, dropped = [], 0
    for o in (raw.get("objects") or []):
        if not isinstance(o, dict):
            continue
        typ = _cell_text(o.get("type"))
        if not typ:
            continue
        attrs, seen = [], set()
        for a in (o.get("attrs") or []):
            if not isinstance(a, dict):
                continue
            name = _cell_text(a.get("name"))
            if not name or len(name) > 20:
                continue
            hit = _match_term(name, sk["terms"])
            ev = _cell_text(a.get("evidence"))
            via = "name"
            if not hit:
                # 叙述体：属性名可以是概括，但依据必须在骨架里
                if not _evidence_ok(ev, sk["text"]):
                    dropped += 1
                    continue
                via = "evidence"
            k = _norm(name)
            if k in seen:
                continue
            seen.add(k)
            attrs.append({"name": name, "doc_term": hit or ev[:20],
                          "via": via,
                          "unit": _cell_text(a.get("unit"))[:12],
                          "evidence": ev[:120]})
        objects.append({"type": typ[:24], "why": _cell_text(o.get("why"))[:160],
                        "records": "many" if str(o.get("records")).lower() == "many" else "single",
                        "identity": _cell_text(o.get("identity"))[:24],
                        "attrs": attrs[:24]})

    # 一个属性都没有的对象没有信息量：只在"所有对象都是空"时保留一个占位
    if any(o["attrs"] for o in objects):
        objects = [o for o in objects if o["attrs"]]

    focus = []
    for f in (raw.get("focus") or []):
        f = _cell_text(f)
        if f and len(f) <= 20:
            focus.append(f)

    points = []
    for p in (raw.get("key_points") or []):
        if isinstance(p, dict) and _cell_text(p.get("text")):
            points.append({"text": _cell_text(p["text"])[:200],
                           "quote": _cell_text(p.get("quote"))[:300]})

    return {
        "domain": _cell_text(raw.get("domain"))[:40] or "未识别",
        "doc_type": _cell_text(raw.get("doc_type"))[:40] or "未识别",
        "purpose": _cell_text(raw.get("purpose"))[:200],
        "objects": objects[:8],
        "key_points": points[:5],
        "focus": focus[:8],
        "attr_dropped": dropped,
        "source": "model",
    }


def fallback_schema(sk: dict, doc_file: str, requirement: str = "") -> dict:
    """模型不可用时的兜底：只用骨架里的表格列头和标签，照样能抽出东西。

    这不是"降级到更差的模型"，而是承认一件事实：表格类文档的结构本来就在文档里，
    不需要模型也能读出来。叙述体的标签（"甲方：XX"）同理。
    """
    objects = []
    for i, t in enumerate(sk["tables"][:4], start=1):
        title = (t.get("title") or "").strip()
        typ = title if 2 <= len(title) <= 20 else ("记录%d" % i)
        attrs = []
        for c in sorted(t["headers"]):
            name = t["headers"][c]
            if _norm(name) in {_norm(x) for x in _STRUCT_HEADERS}:
                continue
            attrs.append({"name": name, "doc_term": name, "unit": "", "evidence": "表格列头"})
        objects.append({"type": typ, "why": "表格「%s」的列头" % title,
                        "records": "many",
                        "identity": _guess_identity(t["headers"]), "attrs": attrs[:24]})
    if not objects and sk["labels"]:
        attrs = [{"name": l, "doc_term": l, "unit": "", "evidence": "标签行"}
                 for l in dict.fromkeys(sk["labels"])][:24]
        objects.append({"type": Path(doc_file).stem[:20], "why": "文档里的 XXX： 标签",
                        "records": "single", "identity": "", "attrs": attrs})
    return {"domain": "未识别", "doc_type": Path(doc_file).suffix.lstrip(".") or "未知",
            "purpose": "", "objects": objects, "key_points": [], "focus": [],
            "attr_dropped": 0, "source": "structure"}


def induce_schema(doc_file: str, sk: dict, requirement: str = ""):
    """调模型归纳结构。返回 (schema, 耗时, 错误说明)。失败返回兜底结构。

    失败会**重试一次**，且第二次只喂截短后的骨架 —— 实测长文档（判决书这种）
    第一次容易撞模型超时（zhipu 408，95 秒才回来），喂短一点反而过得去。
    注意：校验（clean_schema）始终拿**完整骨架**去核 evidence，
    截短只影响"让模型看到多少"，不能因此把有据可依的属性判成没依据。
    """
    # 注意：没有列头/标签 ≠ 没结构。通报、纪实、新闻稿这类叙述体天生没有「XX：」标签，
    # 它的结构线索就在正文里（"湖北省常住人口约 5775 万人，人均 GDP 约 7.3 万元…"
    # 这一段说的就是一个省的几个指标）。所以只有"既无标签、又几乎没有正文"才认输；
    # 否则一律让模型读正文去归纳 —— 早期这里一见 terms 为空就 return，
    # 导致所有纯叙述体文档 100% 归零。
    if not sk["terms"] and sk.get("n_chars", 0) < 200:
        return fallback_schema(sk, doc_file, requirement), 0.0, "骨架里没有可用的列头/标签"
    prompt = build_schema_prompt(doc_file, sk, requirement)
    t0 = time.time()
    err = ""
    for attempt in (1, 2):
        content, _usage, cost = ask_model(prompt, system=SCHEMA_SYSTEM)
        if content:
            data = parse_json_reply(content)
            if isinstance(data, dict) and data.get("objects"):
                sch = clean_schema(data, sk, doc_file)
                if sch["objects"]:
                    if attempt > 1:
                        sch["_retry"] = True
                    return sch, time.time() - t0, ""
                err = "模型没归纳出对象，已按文档骨架直接抽"
            else:
                err = "模型返回不是有效结构，已按骨架直接抽"
        else:
            err = "模型调用失败（超时或返回空），已按文档骨架直接抽"
        if attempt == 1:
            slim = dict(sk)
            slim["text"] = (sk.get("text") or "")[:2200] + \
                "\n（骨架过长已截断，请只依据上面已给出的部分归纳）"
            prompt = build_schema_prompt(doc_file, slim, requirement)
    return fallback_schema(sk, doc_file, requirement), time.time() - t0, err


# ====================================================================
#  三、抽值 A：表格直读（零模型，按列头 × 单元格）
# ====================================================================

def _guess_identity(headers: dict) -> str:
    names = [headers[c] for c in sorted(headers)]
    for key in _ID_HEADERS:
        for n in names:
            if _norm(n) == _norm(key) or key in _norm(n):
                return n
    return names[0] if names else ""


def _time_col(headers: dict, skip=None):
    """找"时间/日期"列（明细表的天然区分列）。找不到返回 None。"""
    for c in sorted(headers):
        if c == skip:
            continue
        n = _norm(headers[c])
        if any(k in n for k in ("时间", "日期", "时刻", "年月", "date", "time")):
            return c
    return None


def _name_col(headers: dict, skip=None):
    """找"名称"列（站点名称/姓名/商品名称）——比编码好认，拿它当对象名。"""
    for c in sorted(headers):
        if c == skip:
            continue
        n = _norm(headers[c])
        if any(k in n for k in ("名称", "姓名", "品名")) and not any(k in n for k in ("编码", "编号", "代码")):
            return c
    return None


def _pick_object_for_table(table: dict, schema: dict) -> dict:
    """表格该挂到哪个对象上：按属性名重合度选，选不上就用"记录"。"""
    hdr = {_norm(v) for v in table["headers"].values()}
    best, best_n = None, 0
    for o in schema.get("objects") or []:
        if o.get("records") != "many":
            continue
        n = sum(1 for a in o["attrs"] if _norm(a["name"]) in hdr)
        if n > best_n:
            best, best_n = o, n
    if best is not None:
        return best
    return {"type": "记录", "identity": _guess_identity(table["headers"]), "attrs": []}


def table_units(blocks: list, schema: dict, notes: list) -> list:
    """表格 → 对象 + 属性条目。纯程序，值 100% 来自单元格。"""
    units = []
    total = 0
    for ti, t in enumerate(schema.get("_tables") or [], start=1):
        if t.get("layout"):
            notes.append("表格「%s」是版面式排列（一个单元格里有多条数据），"
                         "不按列头×单元格直读；这部分改由正文抽取，或请提供 Excel 版"
                         % (t.get("title") or ("表%d" % ti)))
            continue
        obj = _pick_object_for_table(t, schema)
        otype = obj.get("type") or "记录"
        ident_name = obj.get("identity") or _guess_identity(t["headers"])
        # identity 列：先按名字精确找，再退回第一列
        id_col = None
        for c, name in t["headers"].items():
            if _norm(name) == _norm(ident_name) or (_norm(ident_name) and _norm(ident_name) in _norm(name)):
                id_col = c
                break
        if id_col is None and t["headers"]:
            id_col = sorted(t["headers"])[0]

        rows = sorted(t["rows"])
        truncated = 0
        if len(rows) > MAX_TABLE_ROWS:
            truncated = len(rows) - MAX_TABLE_ROWS
            rows = rows[:MAX_TABLE_ROWS]

        # 明细表（空气质量逐时数据、每班巡检记录这类"一次测量一行"）的对象名：
        #   · 身份列只是流水号（序号/编号/工号）时，"第 2000 行"这种名字没有信息量；
        #   · 身份列本身重复（同一站点几十条）时，拿它当对象名会把记录合并，
        #     时间与数值的对应关系就散了。
        # 两种情况都改用「名称｜时间」作对象名（找不到名称/时间就补行号）。
        tail_col = None
        head_col = None
        if len(rows) >= 20:
            ident_norm = _norm(ident_name)
            weak_id = ident_norm in {_norm(x) for x in ("序号", "编号", "id", "no", "no.", "#", "工号", "学号", "code")}
            ids = [_cell_text(t["rows"][r].get(id_col)) for r in rows] if id_col is not None else []
            dup = bool(ids) and len({v for v in ids if v}) < len(rows) * 0.9
            if weak_id or dup:
                tail_col = _time_col(t["headers"], id_col)
                head_col = _name_col(t["headers"], id_col)
                if tail_col is None and head_col is None and dup:
                    tail_col = -1                      # -1 表示用行号

        for r in rows:
            row = t["rows"][r]
            oname = _cell_text(row.get(id_col)) if id_col is not None else ""
            if head_col is not None:
                hn = _cell_text(row.get(head_col))
                if hn:
                    oname = hn
            if not oname:
                nonempty = [_cell_text(row.get(c)) for c in sorted(row)]
                oname = next((v for v in nonempty if v), "")
            if not oname:
                continue
            if tail_col == -1:
                oname = "%s｜第%d行" % (oname, r)
            elif tail_col is not None:
                tv = _cell_text(row.get(tail_col))
                oname = "%s｜%s" % (oname, tv) if tv else "%s｜第%d行" % (oname, r)
            b0 = t["blocks"].get((r, id_col)) if id_col is not None else None
            units.append({
                "kind": "object", "obj_type": otype, "obj_name": oname[:200],
                "attr": "", "value": oname[:500], "unit": "",
                "quote": _cell_text(row.get(id_col))[:300] if id_col is not None else "",
                "loc": loc_desc(b0) if b0 else "", "verified": True,
                "focus": False, "source": "table",
            })
            total += 1
            for c in sorted(t["headers"]):
                if c == id_col:
                    continue
                name = t["headers"][c]
                if _norm(name) in {_norm(x) for x in _SKIP_ATTR}:
                    continue
                val = _cell_text(row.get(c))
                if not val:
                    continue
                b = t["blocks"].get((r, c))
                unit, val_clean = _split_unit(val)
                units.append({
                    "kind": "attribute", "obj_type": otype, "obj_name": oname[:200],
                    "attr": name[:60], "value": val_clean[:800], "unit": unit,
                    "quote": val[:300], "loc": loc_desc(b) if b else "",
                    "verified": True, "focus": False, "source": "table",
                })
        if truncated:
            notes.append("表格「%s」共 %d 行，按上限只入库前 %d 行（%d 行未入库，可调 A23_KNOWLEDGE_MAX_ROWS）"
                         % (_table_title(blocks, t) or ("表%d" % ti),
                            truncated + MAX_TABLE_ROWS, MAX_TABLE_ROWS, truncated))
    notes.append("表格直读：%d 个对象 / %d 条属性" % (total, len(units) - total)) if units else None
    return units


def _split_unit(v: str):
    """从值里拆单位："1,800,000 元" → (元, 1,800,000)；"92.5%" → (%, 92.5)。"""
    v = (v or "").strip()
    m = re.match(r"^([+-]?[\d,，.]+)\s*(万元|亿元|千元|元|万份|万人|吨|kg|千克|米|m|㎡|%|％)$", v)
    if m:
        return (m.group(2).replace("％", "%"), m.group(1))
    return ("", v)


# ====================================================================
#  四、抽值 B：叙述体（1 次模型调用，逐条回原文核验）
# ====================================================================

def _narrative_blocks(blocks: list) -> list:
    return [b for b in blocks if (b.get("locator") or {}).get("kind") not in ("cell", "table_cell")
            and (b.get("text") or "").strip()]


def _attr_codes(schema: dict):
    """给结构里的对象/属性编短号：O1/O2…、A1/A2…

    2026-11-05 加，纯为提速：叙述体抽值时模型每条都要重复写一遍
    「省级行政区」「常住人口」这种长名字，输出量的大头全耗在这上面，
    而输出量就是耗时本身。编号之后每条能省一半 token。
    """
    omap, amap, n = {}, {}, 0
    for oi, o in enumerate(schema.get("objects") or [], start=1):
        code = "O%d" % oi
        omap[code] = o["type"]
        for a in o["attrs"]:
            n += 1
            amap["A%d" % n] = a["name"]
    return omap, amap


def build_extract_prompt(doc_file: str, schema: dict, digest: str,
                         requirement: str, n_show: int) -> str:
    omap, amap = _attr_codes(schema)
    rev_a, rev_o = {v: k for k, v in amap.items()}, {v: k for k, v in omap.items()}
    obj_lines = []
    for o in schema.get("objects") or []:
        codes = "、".join("%s=%s" % (rev_a.get(a["name"], "?"), a["name"]) for a in o["attrs"])
        obj_lines.append("- %s = %s（%s）：%s"
                         % (rev_o.get(o["type"], "?"), o["type"],
                            "多条记录" if o["records"] == "many" else "单条",
                            codes or "（无固定属性，可自行归纳）"))
    focus = "、".join(schema.get("focus") or []) or "（无特别关注）"
    return f"""请按下面这份结构，从文档《{doc_file}》的原文里把值找出来。

【文档已归纳出的结构】（左边是编号，右边是名字；输出时写编号）
{chr(10).join(obj_lines)}

【用户关注点】{focus}
【原始要求】{requirement.strip() or "（无）"}

注意：关注点只决定"哪些要重点标出来"，**绝不代表只抽这几项**。
上面清单里的每个属性都要抽，关注点之外的同样要抽、抽不全算漏。
（早期这里让模型只看关注点，结果用户一填关注点，其他字段全丢了。）

【原文】（共 {n_show} 块）
{digest}

【输出 JSON】
{{
  "items": [
    {{"o": "对象编号（O1/O2…）", "n": "这条记录叫什么（如 湖北省 / 张三）",
      "p": {{"A1": "值", "A2": "值", "A5": "值"}}}}
  ],
  "relations": [
    {{"subject": "主体（对象名）", "predicate": "关系（如 供应/签约/负责/隶属于）",
      "object": "客体（对象名）", "quote": "能同时看到两者的原文片段"}}
  ],
  "events": [
    {{"date": "日期（逐字照抄）", "what": "发生了什么", "quote": "原文片段"}}
  ],
  "points": [
    {{"text": "值得记住的结论（一句话）", "quote": "支撑它的原文摘抄"}}
  ]
}}

规则：
- o / p 的键必须填上面的**编号**（O1、A2 这种），不要填名字。
- **一条记录只写一个 item**：把这个记录的所有属性都塞进它的 p 里，
  不要把同一条记录拆成好几个 item。没有的属性不要写 "未提及"。
- items 只写 o / n / p 三个字段，**不要写 quote、不要写 unit**。
  程序会拿值自己回原文定位；单位也从值里拆（写成"约 5775 万人"就够）。
  只有一种例外：值经过换算或概括、原文不是这么写的，才在 p 里写成
  {{"A3": ["值", "原文片段"]}} 这种两元组。
- 值必须逐字出现在原文里，不许换算、不许总结、不许补全；原文怎么写就怎么抄。
- 标着「多条记录」的对象：上面原文里有几条就列几条，一条都不能少。
  一段说一个的（如"湖北省…""广东省…"），每段都要各出一条；
  不要合并成"各省情况"一条，也不要用几句总结代替明细。
- 文档里没有的就不报；拿不准的不报。宁少不编。
- relations 的 subject 必须是**动作的发出者**（人 / 组织 / 主体），object 是动作的承受者。
  不要把方向搞反（"吴晓敏负责 xlsx 解析器"应写成 subject=吴晓敏、predicate=负责、object=xlsx 解析器）。
- relations / events / points 没有就返回空数组。"""


def _verify_value(value: str, quote: str, blocks: list):
    """值回原文：返回 (loc 描述, 命中方式, 是否值本身命中)。

    ① 值原样在块里 → verified
    ② 值归一（去空白/标点/千分位）后命中 → verified
    ③ 值不在但摘抄在 → 保留但标未核验（可能经过换算）
    """
    v = (value or "").strip()
    if v:
        for b in blocks:
            if v in (b.get("text") or ""):
                return loc_desc(b), "exact", True
        vn = _norm(v)
        if len(vn) >= 2:
            for b in blocks:
                if vn in _norm(b.get("text") or ""):
                    return loc_desc(b), "ws", True
        if re.fullmatch(r"[\d,，.]+", v):
            v2 = v.replace(",", "").replace("，", "")
            rx = re.compile(r"(?<!\d)" + re.escape(v2) + r"(?!\d)")
            for b in blocks:
                if rx.search((b.get("text") or "").replace(",", "").replace("，", "")):
                    return loc_desc(b), "num", True
    loc, how = locate_quote(quote, blocks)
    if loc:
        return loc, "quote-only", False
    return "", "miss", False


def _snippet_for(value: str, blocks: list, pad: int = 18) -> str:
    """值在原文里的上下文片段（一定包含 value），用来代替模型自己抄的 quote。

    2026-11-05 加：叙述体抽值里让模型逐条再抄一遍 quote 是纯浪费 ——
    quote 的契约本来就只是"必须包含 value"，那么拿 value 反查一次就能得到。
    实测这能砍掉约四成输出量（输出量就是耗时本身）。
    """
    v = (value or "").strip()
    if not v:
        return ""
    for b in blocks:
        t = b.get("text") or ""
        i = t.find(v)
        if i >= 0:
            a, z = max(0, i - pad), min(len(t), i + len(v) + pad)
            return t[a:z].strip()
    vn = _norm(v)
    if len(vn) >= 2:
        for b in blocks:
            t = b.get("text") or ""
            if vn in _norm(t):
                return t[:120].strip()          # 归一坐标对不上原文，给整块开头
    return ""


def clean_items(raw, schema: dict, blocks: list, notes: list) -> list:
    """清洗模型抽的值：属性名核对 + 回原文核验 + 去重。

    兼容两种回包：编号式（o/n/a/v，2026-11-05 起的默认）和全名式（obj_type/attr/value）。
    编号先翻译回名字，再走同一套校验，所以两者的验收标准完全一致。
    """
    omap, amap = _attr_codes(schema)
    valid_attrs = {}
    for o in schema.get("objects") or []:
        for a in o["attrs"]:
            valid_attrs[_norm(a["name"])] = a["name"]
    focus = {_norm(f) for f in (schema.get("focus") or [])}

    out, seen = [], set()
    n_unverified = n_dropped = 0

    # 两种回包形状都吃：
    #   紧凑式 {"o":"O1","n":"湖北省","p":{"A1":"湖北省","A2":"约 5775 万人"}}
    #   扁平式 {"o":"O1","n":"湖北省","a":"A2","v":"约 5775 万人"}
    # 紧凑式是 2026-11-05 提速用的：一条记录只写一次编号和名字，
    # 几十条属性不再各自重复一遍 "o"/"n"/"a" 的键名。
    rows = []
    for it in (raw.get("items") or []):
        if not isinstance(it, dict):
            continue
        o = _cell_text(it.get("o") or it.get("obj_type"))
        n = _cell_text(it.get("n") or it.get("obj_name"))
        q = _cell_text(it.get("q") or it.get("quote"))
        pmap = it.get("p") or it.get("pairs")
        if isinstance(pmap, dict) and pmap:
            for ak, av in pmap.items():
                if isinstance(av, (list, tuple)):        # {"A3": ["值", "原文片段"]}
                    pair = list(av) + [None]
                    rows.append((o, n, _cell_text(ak), pair[0], _cell_text(pair[1]) or q))
                else:
                    rows.append((o, n, _cell_text(ak), av, q))
            continue
        if isinstance(pmap, list):
            for pair in pmap:
                if isinstance(pair, dict):
                    rows.append((o, n, _cell_text(pair.get("a")), pair.get("v"), q))
            continue
        rows.append((o, n, _cell_text(it.get("a") or it.get("attr")),
                     it.get("v") if it.get("v") is not None else it.get("value"), q))

    for o_raw, n_raw, attr_raw, value_raw, q_raw in rows[:MAX_TEXT_ITEMS * 4]:
        attr = amap.get(attr_raw.upper(), attr_raw) if attr_raw else ""
        value = _cell_text(value_raw)
        if not attr or not value:
            continue
        hit = valid_attrs.get(_norm(attr))
        if hit is None:
            n_dropped += 1
            continue
        loc, how, ok = _verify_value(value, q_raw, blocks)
        if not loc:
            n_dropped += 1
            continue
        if not ok:
            n_unverified += 1
        quote = q_raw or _snippet_for(value, blocks)
        unit = _split_unit(value)[0]
        otype = omap.get(o_raw.upper(), o_raw) if o_raw else ""
        if not otype:
            otype = schema["objects"][0]["type"] if schema["objects"] else "记录"
        oname = n_raw
        key = (_norm(otype), _norm(oname), _norm(hit), _norm(value))
        if key in seen:
            continue
        seen.add(key)
        out.append({"kind": "attribute", "obj_type": otype[:40], "obj_name": oname[:200],
                    "attr": hit[:60], "value": value[:800],
                    "unit": unit[:12], "quote": quote[:400],
                    "loc": loc, "verified": ok, "focus": _norm(hit) in focus,
                    "source": "text:%s" % how})
    for r in (raw.get("relations") or [])[:60]:
        if not isinstance(r, dict):
            continue
        s, p, o = (_cell_text(r.get(k)) for k in ("subject", "predicate", "object"))
        if not (s and p and o):
            continue
        quote = _cell_text(r.get("quote"))
        loc, how, _ok = _verify_value(s, quote, blocks)
        if not loc:
            n_dropped += 1
            continue
        out.append({"kind": "relation", "obj_type": "关系", "obj_name": s[:200],
                    "attr": p[:40], "value": "%s → %s" % (s, o), "unit": "",
                    "quote": quote[:400], "loc": loc, "verified": True,
                    "focus": False, "source": "text:%s" % how})
    for e in (raw.get("events") or [])[:60]:
        if not isinstance(e, dict):
            continue
        date, what = _cell_text(e.get("date")), _cell_text(e.get("what"))
        if not (date and what):
            continue
        loc, how, ok = _verify_value(date, _cell_text(e.get("quote")), blocks)
        if not loc:
            n_dropped += 1
            continue
        out.append({"kind": "event", "obj_type": "事件", "obj_name": date[:100],
                    "attr": what[:100], "value": what[:500], "unit": "",
                    "quote": _cell_text(e.get("quote"))[:400], "loc": loc,
                    "verified": ok, "focus": False, "source": "text:%s" % how})
    if n_dropped:
        notes.append("叙述抽值：丢弃 %d 条（属性名不在结构里，或值/摘抄都回不到原文）" % n_dropped)
    if n_unverified:
        notes.append("叙述抽值：%d 条值本身未逐字命中（摘抄可溯，可能经换算/规整），已标记 verified=0" % n_unverified)
    return out


def extract_text(doc_file: str, schema: dict, blocks: list, requirement: str):
    """叙述体抽值。返回 (units, 耗时, 错误)。

    分批调模型：每批 ≤ TEXT_CHUNK_BLOCKS 块。原因见 TEXT_CHUNK_BLOCKS 处的注释。
    """
    nar = _narrative_blocks(blocks)
    if not nar:
        return [], 0.0, ""
    idx, note = pick_for_budget(nar, budget=TEXT_BUDGET)
    if not idx:
        return [], 0.0, ""

    # 切批：批数少 → 固定开销少；批太大 → 模型偷懒。取两者折中（见上方常量注释）。
    n_blk = len(idx)
    step = int(math.ceil(n_blk / float(max(1, TEXT_MAX_CHUNKS))))
    size = max(TEXT_CHUNK_BLOCKS, step)
    chunks = [idx[i:i + size] for i in range(0, n_blk, size)]

    units, total_cost = [], 0.0
    errs, notes = [], []
    if note:
        notes.append(note)

    def _run_chunk(ci, ch):
        """跑一批。返回 (units, points, cost, err)。"""
        digest = build_digest(nar, ch)
        if not digest.strip():
            return [], [], 0.0, ""
        prompt = build_extract_prompt(doc_file, schema, digest, requirement, len(ch))
        content, _usage, cost = ask_model(prompt, system=EXTRACT_SYSTEM)
        if not content:
            return [], [], cost, "第%d/%d批模型无返回" % (ci, len(chunks))
        data = parse_json_reply(content)
        if not isinstance(data, dict):
            return [], [], cost, "第%d/%d批返回不是 JSON" % (ci, len(chunks))
        local_notes = []
        us = clean_items(data, schema, blocks, local_notes)
        pts = []
        for p in (data.get("points") or [])[:3]:
            if isinstance(p, dict) and _cell_text(p.get("text")):
                loc, _how = locate_quote(_cell_text(p.get("quote")), blocks)
                pts.append({"kind": "summary", "obj_type": "要点", "obj_name": "",
                            "attr": "", "value": _cell_text(p["text"])[:300], "unit": "",
                            "quote": _cell_text(p.get("quote"))[:400], "loc": loc or "",
                            "verified": bool(loc), "focus": False, "source": "text"})
        return us, pts, cost, "；".join(local_notes)

    # 各批互不依赖（最终靠 quote 回原文定位、靠 (对象,属性,值) 去重），
    # 所以并发跑：串行 6 批要 120 秒，超出「每文档 ≤90 秒」的红线。
    with ThreadPoolExecutor(max_workers=max(1, min(TEXT_WORKERS, len(chunks)))) as ex:
        futs = [ex.submit(_run_chunk, ci, ch) for ci, ch in enumerate(chunks, start=1)]
        for f in futs:
            try:
                us, pts, cost, e = f.result()
            except Exception as exc:            # 单批炸了不拖累其他批
                errs.append("某批异常：%s" % exc)
                continue
            units += us
            units += pts
            total_cost += cost
            if e:
                errs.append(e)

    err = "；".join(dict.fromkeys(errs + notes))
    return units, total_cost, err


# ====================================================================
#  五、整理 / 报告 / 落盘
# ====================================================================

def _dedupe(units: list) -> list:
    """同一 (kind, 对象, 属性, 值) 只留一条；表格优先于叙述（表格是原文坐标）。"""
    order = {"table": 0, "text:exact": 1, "text:ws": 2, "text:num": 3, "text": 4}
    units = sorted(units, key=lambda u: order.get(u.get("source", ""), 9))
    out, seen = [], set()
    for u in units:
        k = (u["kind"], _norm(u["obj_type"]), _norm(u["obj_name"]),
             _norm(u["attr"]), _norm(u["value"]))
        if k in seen:
            continue
        seen.add(k)
        out.append(u)
    return out


def _mark_focus(units: list, schema: dict) -> int:
    """按用户的关注点，给已抽到的条目补打 focus 标记；返回命中条数。

    为什么放在最后统一做：表格是「直读」的（table_units 完全不经过模型），
    自然拿不到 schema 里的 focus 名单，早先版本于是出现"关注点填了、一条都没命中"。
    这里对全部条目做一次统一认领：对象名 / 属性名里出现关注点就算命中；
    值里也认，但只认 ≥3 个字的词，免得"数""量"这种单字到处乱命中。
    """
    terms = [t for t in (_norm(f) for f in (schema.get("focus") or [])) if len(t) >= 2]
    if not terms:
        return 0
    n = 0
    for u in units:
        if u.get("focus"):
            n += 1
            continue
        name = _norm("%s %s" % (u.get("obj_type", ""), u.get("attr", "")))
        hit = bool(name) and any(t in name or name in t for t in terms)
        if not hit:
            val = _norm(u.get("value", ""))
            hit = any(len(t) >= 3 and t in val for t in terms)
        if hit:
            u["focus"] = True
            n += 1
    return n


def summarize_schema(schema: dict) -> str:
    if not schema.get("objects"):
        return "（没有归纳出对象）"
    bits = []
    for o in schema["objects"]:
        bits.append("%s×%d属性" % (o["type"], len(o["attrs"])))
    return " ｜ ".join(bits)


def format_screen(doc_file: str, schema: dict, units: list, notes: list,
                  cmd: str, total: float, requirement: str) -> str:
    L = []
    L.append("=" * 62)
    L.append("【知识抽取】读懂文档 → 对象 / 属性 / 关系 / 事件（模块2 · 统一抽取）")
    L.append("=" * 62)
    L.append("指令：%s" % cmd)
    if requirement.strip():
        L.append("关注点：%s" % requirement.strip())
    L.append("文档：%s" % doc_file)
    L.append("")
    L.append("◆ 领域：%s ｜ 类型：%s" % (schema.get("domain"), schema.get("doc_type")))
    if schema.get("purpose"):
        L.append("◆ 用途：%s" % schema["purpose"])
    if schema.get("source") == "structure":
        L.append("◆ 结构来源：文档自身（表格列头/标签）——模型本次未参与结构归纳")
    L.append("◆ 对象：%s" % summarize_schema(schema))
    if schema.get("attr_dropped"):
        L.append("◆ 结构清洗：丢弃 %d 个「骨架里找不到」的属性名（防模型套模板）"
                 % schema["attr_dropped"])
    L.append("")
    cnt = Counter(u["kind"] for u in units)
    L.append("◆ 入库条目：%d 条 ｜ %s" % (
        len(units), " · ".join("%s %d" % (KIND_CN[k], cnt.get(k, 0)) for k in KIND_ORDER)))
    attrs = [u for u in units if u["kind"] == "attribute"]
    if attrs:
        ok = sum(1 for u in attrs if u["verified"])
        L.append("◆ 值可核：%d/%d = %.1f%%（逐条回原文/单元格）"
                 % (ok, len(attrs), 100.0 * ok / len(attrs)))
    if schema.get("focus"):
        hit = sum(1 for u in units if u.get("focus"))
        L.append("◆ 关注点命中：%s ｜ %d 条" % ("、".join(schema["focus"][:6]), hit))
    L.append("")

    # 每个对象给一块示例
    for o in (schema.get("objects") or [])[:4]:
        rows = [u for u in units if u["kind"] == "attribute" and _norm(u["obj_type"]) == _norm(o["type"])]
        if not rows:
            continue
        L.append("◆ 【%s】%d 个对象 · %d 条属性" % (
            o["type"], len({_norm(u["obj_name"]) for u in rows}), len(rows)))
        for u in rows[:6]:
            nm = ("（%s）" % u["obj_name"]) if u["obj_name"] and u["obj_name"] != u["obj_type"] else ""
            L.append("  · %s%s ｜ %s：%s%s ｜ %s" % (
                o["type"], nm, u["attr"], u["value"],
                (" " + u["unit"]) if u["unit"] else "", u["loc"] or "未定位"))
        if len(rows) > 6:
            L.append("  … 其余 %d 条在产物文件 / 数据库里" % (len(rows) - 6))
        L.append("")
    rels = [u for u in units if u["kind"] == "relation"]
    if rels:
        L.append("◆ 关系（%d 条）" % len(rels))
        for u in rels[:5]:
            L.append("  · %s —%s→ %s" % (u["obj_name"], u["attr"], u["value"].split("→")[-1].strip()))
        L.append("")
    pts = [u for u in units if u["kind"] == "summary"]
    if pts:
        L.append("◆ 关键结论（%d 条）" % len(pts))
        for u in pts[:5]:
            L.append("  · %s（%s）" % (u["value"], u["loc"] or "未定位"))
        L.append("")
    for n in notes:
        L.append("◆ 说明：%s" % n)
    L.append("")
    L.append("用时 %.1f 秒" % total)
    return "\n".join(L)


def format_md(doc_file: str, schema: dict, units: list, notes: list,
              cmd: str, total: float, tstamp: str, requirement: str) -> str:
    L = ["# 文档知识抽取", "",
         "- **文档**：%s" % doc_file,
         "- **指令**：%s" % cmd]
    if requirement.strip():
        L.append("- **关注点**：%s" % requirement.strip())
    L.append("- **时间**：%s ｜ 用时 %.1f 秒" % (tstamp, total))
    L.append("")
    L.append("> **领域**：%s ｜ **类型**：%s" % (schema.get("domain"), schema.get("doc_type")))
    if schema.get("purpose"):
        L.append(">")
        L.append("> **用途**：%s" % schema["purpose"])
    L.append("")
    L.append("## 归纳出的结构")
    L.append("")
    L.append("| 对象 | 记录形态 | 标识列 | 属性 |")
    L.append("|---|---|---|---|")
    for o in schema.get("objects") or []:
        L.append("| %s | %s | %s | %s |" % (
            o["type"], "多条" if o["records"] == "many" else "单条",
            o.get("identity") or "—", "、".join(a["name"] for a in o["attrs"]) or "—"))
    L.append("")
    for o in schema.get("objects") or []:
        rows = [u for u in units if u["kind"] == "attribute" and _norm(u["obj_type"]) == _norm(o["type"])]
        if not rows:
            continue
        L.append("## %s" % o["type"])
        L.append("")
        L.append("| 对象 | 属性 | 值 | 单位 | 出处 | 核验 |")
        L.append("|---|---|---|---|---|---|")
        for u in rows[:200]:
            L.append("| %s | %s | %s | %s | %s | %s |" % (
                (u["obj_name"] or "—").replace("|", "\\|"),
                u["attr"].replace("|", "\\|"),
                u["value"].replace("|", "\\|")[:80],
                u["unit"] or "—", u["loc"] or "—",
                "✓" if u["verified"] else "⚠ 摘抄可溯"))
        if len(rows) > 200:
            L.append("| … | 其余 %d 条见 JSON / 数据库 | | | | |" % (len(rows) - 200))
        L.append("")
    rels = [u for u in units if u["kind"] == "relation"]
    if rels:
        L.append("## 关系")
        L.append("")
        for u in rels:
            L.append("- %s —**%s**→ %s（%s）" % (u["obj_name"], u["attr"],
                                                u["value"].split("→")[-1].strip(), u["loc"]))
        L.append("")
    pts = [u for u in units if u["kind"] == "summary"]
    if pts:
        L.append("## 关键结论")
        L.append("")
        for u in pts:
            L.append("- %s（_%s_）" % (u["value"], u["loc"] or "未定位"))
        L.append("")
    if notes:
        L.append("## 说明")
        L.append("")
        for n in notes:
            L.append("- %s" % n)
        L.append("")
    return "\n".join(L)


# ====================================================================
#  六、主流程
# ====================================================================

def run_knowledge(text: str, intent: dict, src: Path, out_dir: Path = None,
                  requirement: str = ""):
    """执行一次统一知识抽取。返回 (是否成功, 报告, 产出文件列表)。

    与 run_profile / run_entities / run_extract 同签名 —— edit.py 按 op 分流。
    """
    out_dir = Path(out_dir) if out_dir else OUT_DIR
    out_dir.mkdir(parents=True, exist_ok=True)
    t0 = time.time()

    requirement = (requirement or (intent.get("params") or {}).get("requirement")
                   or (intent.get("params") or {}).get("require") or "").strip()

    doc, err = load_doc(src.name, src)
    if doc is None:
        return False, "【失败】%s" % err, []
    blocks = doc.get("blocks") or []
    if not blocks:
        return False, "【失败】《%s》读不出内容块。" % src.name, []

    notes = []
    sk = build_skeleton(blocks, src.name, doc.get("type", ""))
    schema, cost_s, err_s = induce_schema(src.name, sk, requirement)
    if err_s:
        notes.append("结构归纳：%s" % err_s)
    schema["_tables"] = sk["tables"]
    # 表格里认出的对象如果模型没提到，补进去（保证对象清单完整）
    for t in sk["tables"]:
        if t.get("layout"):
            continue
        o = _pick_object_for_table(t, schema)
        if o.get("type") == "记录" and not any(x["type"] == "记录" for x in schema["objects"]):
            schema["objects"].append({"type": "记录", "why": "表格列头", "records": "many",
                                      "identity": o.get("identity") or "", "attrs": []})

    units = []
    units += table_units(blocks, schema, notes)

    t1 = time.time()
    text_units, _cost_t, err_t = extract_text(src.name, schema, blocks, requirement)
    if err_t:
        notes.append("叙述体：%s" % err_t)
    units += text_units

    # 结构里的关键结论（模型在归纳阶段给的）
    for p in schema.get("key_points") or []:
        loc, _how = locate_quote(p.get("quote", ""), blocks)
        units.append({"kind": "summary", "obj_type": "要点", "obj_name": "", "attr": "",
                      "value": p["text"], "unit": "", "quote": p.get("quote", ""),
                      "loc": loc or "", "verified": bool(loc), "focus": False,
                      "source": "schema"})

    units = _dedupe(units)
    units.sort(key=lambda u: KIND_ORDER.index(u["kind"]))
    n_focus = _mark_focus(units, schema)
    if requirement and schema.get("focus") and not n_focus:
        notes.append("关注点（%s）在本文档的属性里没有直接对应的项，已按常规全量抽取"
                     % "、".join(schema["focus"][:4]))

    if not units:
        return False, ("【失败】没有抽到任何可入库的内容。这份文档可能是空的，"
                       "或结构过于特殊 —— 报告里会写明卡在哪一步。"), []

    total = time.time() - t0
    tstamp = time.strftime("%Y-%m-%d %H:%M")
    schema_view = {k: v for k, v in schema.items() if not k.startswith("_")}

    screen = format_screen(src.name, schema_view, units, notes, text, total, requirement)
    md = format_md(src.name, schema_view, units, notes, text, total, tstamp, requirement)

    md_path = out_dir / (src.stem + "_知识.md")
    json_path = out_dir / (src.stem + "_知识.json")
    payload = {
        "cmd": text, "doc": src.name, "time": tstamp,
        "elapsed_sec": round(total, 2),
        "sec_schema": round(cost_s, 2), "sec_extract": round(time.time() - t1, 2),
        "requirement": requirement,
        "schema": schema_view,
        "counts": {k: sum(1 for u in units if u["kind"] == k) for k in KIND_ORDER},
        "units": units,
        "notes": notes,
    }
    md_path.write_text(md, encoding="utf-8")
    json_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")

    # 入库（失败不拦主流程，报告里说明）
    store_line = ""
    try:
        import store as db_store
        res = db_store.save_knowledge(src.name, schema_view, units,
                                      run_id=time.strftime("%Y%m%d%H%M%S"),
                                      requirement=requirement)
        store_line = "【入库】%s：%s" % (res.get("db"), res.get("line"))
    except Exception as e:
        store_line = "【入库】失败（不影响文件产物）：%s" % e

    screen += ("\n【输出】%s\n【输出】%s\n%s\n【原件】未改动（抽取是只读操作）"
               % (md_path, json_path, store_line))
    return True, screen, [str(md_path), str(json_path)]


def main():
    import argparse
    from parse_intent import scan_docs, parse_one

    ap = argparse.ArgumentParser(description="模块（2）· 统一知识抽取")
    ap.add_argument("text", help="一条抽取指令")
    ap.add_argument("--docs", default=str(BASE / "测试数据" / "源文档"))
    ap.add_argument("--requirement", default="", help="用户关注点（可选）")
    args = ap.parse_args()

    docs_dir = Path(args.docs)
    names = scan_docs(docs_dir)
    intent, cost, ok = parse_one(args.text, names)
    if not ok:
        print("解析失败，请重跑一次。")
        sys.exit(1)
    print("解析（%.1fs）：%s" % (cost, json.dumps(intent, ensure_ascii=False)))
    name = intent.get("doc") or intent.get("doc_hint") or ""
    src = next((p for p in docs_dir.rglob("*")
                if p.is_file() and (p.name == name or p.stem == Path(name).stem)), None)
    if src is None:
        print("文档库里找不到：%s" % name)
        sys.exit(1)
    ok, report, _files = run_knowledge(args.text, intent, src, requirement=args.requirement)
    print(report)
    sys.exit(0 if ok else 1)


if __name__ == "__main__":
    main()
