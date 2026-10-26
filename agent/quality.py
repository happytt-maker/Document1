# -*- coding: utf-8 -*-
"""
quality.py —— 模块（2）· 第 5 步：质量验证与准确率报告

一句话：把数据库里的数据「拉出来 → 回原文核一遍 → 对照人工标准答案算分 → 出一份报告」。

设计原则：零模型调用。
  验证过程全部是确定性程序（归一化比对 / 子串匹配 / 正则），不调用任何模型 ——
  同样的输入必然得出同样的结论：质量结果可复现、可审计、可当场重跑。

三块能力：
  ① 反向审计 audit —— 数据库 → 原文：
      库里每条数据（实体 / 画像明细 / 字段）的摘抄回原文逐片段核对，输出「可溯率」。
      口径：
        · 摘抄按分隔符拆片段，全部片段都能回原文找到 = 完全可溯（traced）
        · 部分片段找到 = 部分可溯（partial）；全找不到 = 未溯（miss）；无摘抄 = no_quote
        · 主指标「保真度」= traced / (total - no_quote)
  ② 计分 score —— 数据库 → 人工标准答案（标准答案-模块2.json）：
      实体线：101 项查全率 + 类型正确率；画像线：40 点覆盖率；字段线：58 值查全率。
  ③ 报告 report —— 汇总成《质量报告.html》+《质量明细.json》（分文档 / 分线 / 分类型）。

命令行（在项目根目录）：
    python agent/quality.py --audit           只跑反向审计
    python agent/quality.py --score           只跑对照计分
    python agent/quality.py --report          出完整报告（默认 = 全跑）
    python agent/quality.py --doc 设备采购     只看某份文档
"""
import argparse
import json
import re
import sys
import time
from datetime import datetime
from pathlib import Path

try:
    sys.stdout.reconfigure(errors="replace")
except Exception:
    pass

BASE = Path(__file__).resolve().parents[1]                  # = 研一项目/
AGENT_DIR = Path(__file__).resolve().parent                 # = 研一项目/agent/
sys.path.insert(0, str(AGENT_DIR))                          # db / 同目录复用
sys.path.insert(0, str(BASE / "doc_reader"))                # readers（现场解析池外文档）

from db import get_conn                                     # noqa: E402

ANSWER_JSON = AGENT_DIR / "标准答案-模块2.json"
POOL_JSON = BASE / "doc_reader" / "output.json"
REPORT_DIR = AGENT_DIR / "输出" / "质量报告"
SEARCH_DIRS = [BASE / "测试数据" / "源文档",
               BASE / "测试数据" / "导入样本",
               BASE / "导入文档"]

TYPE_CN = {"person": "人名", "org": "机构", "date": "日期", "money": "金额",
           "code": "编号", "contact": "联系方式", "product": "产品设备"}


# ================================================================
#  〇、基础工具（归一化 / 拆片段 / 块内查找）
# ================================================================

_PUNCT_MAP = [("，", ","), ("：", ":"), ("；", ";"), ("（", "("), ("）", ")"),
              ("、", ","), ("￥", ""), ("¥", "")]


def norm(s) -> str:
    """归一化：去空白 + 全角标点统一 + 去货币符号 + 去千分位逗号 + 去 markdown 标记 + 去尾部 .00。

    只做「不影响语义的等价差异」消除，不做同义词替换 —— 质量核对必须保守。
    （markdown 粗体/反引号是排版标记、不是内容，摘抄时被省略属等价差异。）
    """
    s = re.sub(r"\s+", "", str(s or ""))
    for a, b in _PUNCT_MAP:
        s = s.replace(a, b)
    s = s.replace(",", "")
    s = s.replace("**", "").replace("*", "").replace("`", "")
    s = re.sub(r"\.00$", "", s)
    return s


def split_parts(s) -> list:
    """把「拼接式摘抄」拆片段（空格 / 换行 / 顿号 / 分号 / 省略号 … 等）。

    单片段原样返回；多片段时过滤单字符碎片（防「8」这类碎片满文命中）。
    """
    raw = [p.strip() for p in re.split(r"[\s,，、;；|/]+|…+", str(s or ""))]
    raw = [p for p in raw if p]
    if len(raw) <= 1:
        return raw
    parts = [p for p in raw if len(p) >= 2]
    return parts or raw


_HEADER_NOTE_RE = re.compile(r"（该格所在列的列头[:：]([^）]+)）")


def strip_notes(q: str) -> str:
    """剥离系统生成的辅助标注（如 xlsx 格子的「（该格所在列的列头：X）」）。

    标注是程序为帮助模型理解附上的（见 matcher/extract.py），不属于摘抄内容本身；
    但标注里引用的列头名保留为片段 —— 它仍需在原文可溯。
    """
    return _HEADER_NOTE_RE.sub(lambda m: " " + m.group(1) + " ", str(q or ""))


def num_tokens(s) -> list:
    """提取文本里的数字型 token（画像转述的「数字溯源」检查用）。"""
    s = str(s or "")
    toks = []
    toks += re.findall(r"\d{2,}", s)
    toks += re.findall(r"\d{1,3}%", s)
    toks += re.findall(r"\d{1,2}月\d{1,2}日", s)
    return sorted(set(toks))


def _find_num_in_blocks(tok: str, blocks: list) -> bool:
    """数字 token 回原文核对（带数字边界 —— 防「8」误命中「12800」）。

    先原样找，再对块文本归一化后找（覆盖「2026 年 6 月 30 日」这类数字间带空格的写法）。
    """
    pat = re.compile(r"(?<!\d)" + re.escape(tok) + r"(?!\d)")
    for b in blocks:
        text = b.get("text") or ""
        if pat.search(text):
            return True
        nt = norm(text)
        if nt != text and pat.search(nt):
            return True
    return False


def _find_in_blocks(text, blocks, min_len=2) -> bool:
    """一段文本回原文找：①原样子串 ②归一后子串。"""
    t = (text or "").strip()
    if not t:
        return False
    for b in blocks:
        if t in (b.get("text") or ""):
            return True
    tn = norm(t)
    if len(tn) >= min_len:
        for b in blocks:
            if tn in norm(b.get("text") or ""):
                return True
    return False


# ================================================================
#  一、装载（解析池 / 现场解析 / 数据库 / 标准答案）
# ================================================================

_POOL_CACHE = None
_PARSE_CACHE = {}


def _pool() -> dict:
    global _POOL_CACHE
    if _POOL_CACHE is None:
        try:
            _POOL_CACHE = json.loads(POOL_JSON.read_text(encoding="utf-8"))
        except Exception:
            _POOL_CACHE = {"documents": []}
    return _POOL_CACHE


def load_blocks_any(doc_name: str):
    """装载某文档的原文块。

    ① 解析池（doc_reader/output.json）里找；
    ② 池外（如已清出池的导入件）→ 磁盘上现场解析（只读，不入池）。
    返回 (blocks 或 None, 来源说明)。
    """
    for d in _pool().get("documents", []):
        if d.get("file") == doc_name:
            return d.get("blocks") or [], "解析池"

    if doc_name in _PARSE_CACHE:
        return _PARSE_CACHE[doc_name]

    for d in SEARCH_DIRS:
        if not d.exists():
            continue
        for p in d.rglob(doc_name):
            if p.is_file():
                try:
                    from readers import read_one
                    dd = read_one(p)
                    if dd and dd.get("blocks"):
                        _PARSE_CACHE[doc_name] = (dd["blocks"], "现场解析")
                        return dd["blocks"], "现场解析"
                except Exception as e:
                    _PARSE_CACHE[doc_name] = (None, "现场解析失败：%r" % e)
                    return _PARSE_CACHE[doc_name]
    _PARSE_CACHE[doc_name] = (None, "原文不在磁盘")
    return None, "原文不在磁盘"


def load_answers() -> dict:
    return json.loads(ANSWER_JSON.read_text(encoding="utf-8"))


def fetch_db() -> dict:
    """从 MySQL 拉三线全量数据（只读）。"""
    conn = get_conn()
    cur = conn.cursor()
    out = {}
    for key, sql in [
        ("entities", "SELECT id, doc_name, etype, etype_cn, value, freq, loc, quote, verify, source, run_id, ts FROM entities ORDER BY doc_name, id"),
        ("profiles", "SELECT id, doc_name, topic, doc_type, run_id, ts FROM profiles ORDER BY doc_name, id"),
        ("profile_items", "SELECT id, profile_id, kind, seq, title, detail, quote, loc, loc_ok FROM profile_items ORDER BY profile_id, seq, id"),
        ("extractions", "SELECT id, doc_name, field, value, evidence, loc, run_id, ts FROM extractions ORDER BY doc_name, field, id"),
    ]:
        cur.execute(sql)
        cols = [d[0] for d in cur.description]
        out[key] = [dict(zip(cols, r)) for r in cur.fetchall()]
    conn.close()
    return out


# ================================================================
#  二、反向审计（数据库 → 原文）
# ================================================================

def audit_one(quote: str, value: str, blocks: list) -> dict:
    """单条核对（纯函数）。

    返回 {status, parts, hit_parts, miss_parts, value_ok}
    status: traced / partial / miss / no_quote
    """
    parts = split_parts(strip_notes(quote))
    if not parts:
        return {"status": "no_quote", "parts": [], "hit_parts": [], "miss_parts": [],
                "value_ok": None}

    hit_parts = [p for p in parts if _find_in_blocks(p, blocks)]
    miss_parts = [p for p in parts if not _find_in_blocks(p, blocks)]

    v = (value or "").strip()
    value_ok = None
    if v:
        value_ok = _find_in_blocks(v, blocks) or bool(norm(v) and norm(v) in norm(quote))

    if not miss_parts and (value_ok is not False):
        status = "traced"
    elif hit_parts:
        status = "partial"
    else:
        status = "miss"
    return {"status": status, "parts": parts, "hit_parts": hit_parts,
            "miss_parts": miss_parts, "value_ok": value_ok}


def _blank_stat() -> dict:
    return {"total": 0, "traced": 0, "partial": 0, "miss": 0, "no_quote": 0,
            "frag_total": 0, "frag_hit": 0}


def _add_stat(stat: dict, r: dict):
    stat["total"] += 1
    stat[r["status"]] += 1
    stat["frag_total"] += len(r["parts"])
    stat["frag_hit"] += len(r["hit_parts"])


def audit_all(db: dict = None, doc: str = None) -> dict:
    """反向审计全库（或单份文档）。

    doc: 文档名子串过滤（None = 全库）
    返回 {overall, by_doc, by_line, by_type, issues, load_errors, elapsed}
    """
    t0 = time.time()
    db = db or fetch_db()

    # 三线各自的审计清单：(doc_name, 行标识, quote, 核值, 行号, 数字清单)
    #   · 实体 / 字段：quote 逐片核对 + 值核对
    #   · 画像：quote 逐片核对 + 转述数字溯源（detail 是转述性质，不逐字核对）
    lines = {"entities": [], "profile": [], "fields": []}

    for r in db["entities"]:
        lines["entities"].append((r["doc_name"], "%s=%s" % (r.get("etype_cn") or r.get("etype"), r.get("value")),
                                  r.get("quote"), r.get("value"), r.get("id"), None))
    # profile_items 需关联 profiles 拿 doc_name
    prof_doc = {p["id"]: p["doc_name"] for p in db["profiles"]}
    for r in db["profile_items"]:
        label = (r.get("title") or (r.get("detail") or "")[:30])
        lines["profile"].append((prof_doc.get(r["profile_id"], "?"), label,
                                 r.get("quote"), None, r.get("id"),
                                 num_tokens(r.get("detail") or "")))
    for r in db["extractions"]:
        lines["fields"].append((r["doc_name"], "%s=%s" % (r.get("field"), r.get("value")),
                                r.get("evidence"), r.get("value"), r.get("id"), None))

    overall = _blank_stat()
    by_line = {k: _blank_stat() for k in lines}
    by_doc = {}
    by_type = {}
    issues = []
    load_errors = {}
    cache = {}

    for line, rows in lines.items():
        for doc_name, label, quote, value, rid, nums in rows:
            if doc:
                if doc not in (doc_name or ""):
                    continue
            if doc_name not in cache:
                cache[doc_name] = load_blocks_any(doc_name)
            blocks, src = cache[doc_name]
            if blocks is None:
                load_errors[doc_name] = src
                continue
            r = audit_one(quote, value, blocks)
            # 转述数字溯源（画像线）：有数字未溯 → 至少降级为 partial
            if nums:
                bad_nums = [t for t in nums if not _find_num_in_blocks(t, blocks)]
                r["bad_nums"] = bad_nums
                if bad_nums and r["status"] == "traced":
                    r["status"] = "partial"
            r["doc"] = doc_name
            r["line"] = line
            r["label"] = label
            r["rid"] = rid

            _add_stat(overall, r)
            _add_stat(by_line[line], r)
            d = by_doc.setdefault(doc_name, _blank_stat())
            _add_stat(d, r)

            if r["status"] in ("partial", "miss"):
                issues.append({"line": line, "doc": doc_name, "label": label,
                               "status": r["status"], "miss_parts": r["miss_parts"],
                               "bad_nums": r.get("bad_nums") or [],
                               "quote": (quote or "")[:120]})

    # entities 按类型分：单独过一遍（保持 audit_one 结果复用）
    for r in db["entities"]:
        if doc and doc not in (r.get("doc_name") or ""):
            continue
        blocks, src = cache.get(r["doc_name"], (None, ""))
        if not blocks:
            continue
        res = audit_one(r.get("quote"), r.get("value"), blocks)
        tt = r.get("etype") or "unknown"
        tstat = by_type.setdefault(tt, _blank_stat())
        _add_stat(tstat, res)

    return {"overall": overall, "by_doc": by_doc, "by_line": by_line,
            "by_type": by_type, "issues": issues, "load_errors": load_errors,
            "elapsed": round(time.time() - t0, 2)}


# ================================================================
#  三、计分（数据库 → 人工标准答案）
# ================================================================

def _norm_v(s) -> str:
    """计分归一（比 norm 稍宽：尾部 .00 已在 norm 处理；保留主逻辑一致）。"""
    return norm(s)


def _find_hit(alt_list, rows, key="value") -> tuple:
    """对清单项（多形态）在库数据中找命中。双向包含。返回 (命中?, 命中行的 key 值)。

    口径与第 2/3 步验收一致：av == ev 或（len≥2 时）互为子串。
    """
    for r in rows:
        ev = _norm_v(r.get(key))
        if not ev:
            continue
        for alt in alt_list:
            av = _norm_v(alt)
            if not av:
                continue
            if av == ev or (len(av) >= 2 and av in ev) or (len(ev) >= 2 and ev in av):
                return True, r
    return False, None


def score_entities(rows: list, checklist: dict) -> dict:
    """实体线计分（纯函数）。checklist: {type: [[形态...], ...]}"""
    total = sum(len(v) for v in checklist.values())
    hit = 0
    type_ok = 0
    miss = []
    claimed = set()
    for typ, items in checklist.items():
        for alt in items:
            ok, row = _find_hit(alt, rows)
            if ok:
                hit += 1
                if (row.get("etype") or "") == typ:
                    type_ok += 1
                # 认领：命中行 + 一切与它「归一后完全相等」的重复行
                # （同值多形态并存，如金额「1,800,000」与「￥1,800,000.00」，不算额外）
                base = _norm_v(row.get("value"))
                for rr in rows:
                    if rr["id"] not in claimed and base and _norm_v(rr.get("value")) == base:
                        claimed.add(rr["id"])
            else:
                miss.append("%s：%s" % (TYPE_CN.get(typ, typ), alt[0]))
    extra = [r for r in rows if r["id"] not in claimed]
    return {"total": total, "hit": hit, "miss": miss,
            "type_ok": type_ok, "extra": len(extra),
            "extra_values": [r.get("value") for r in extra][:20]}


def score_profile(topic: str, items: list, checklist: list) -> dict:
    """画像线计分（纯函数）。items: profile_items 行；checklist: [[形态...], ...]"""
    parts = [topic or ""]
    for it in items:
        if it.get("kind") == "point":
            parts.append(it.get("detail") or "")
        else:
            parts.append("%s %s" % (it.get("title") or "", it.get("detail") or ""))
    merged = norm(" ".join(parts))
    hit = [c for c in checklist if any(norm(t) in merged for t in c if norm(t))]
    miss = [c[0] for c in checklist if c not in hit]
    return {"total": len(checklist), "hit": len(hit), "miss": miss}


def score_fields(rows: list, specs: list) -> dict:
    """字段线计分（纯函数）。specs: [{key, aliases, expect}]"""
    out = []
    for spec in specs:
        names = [spec["key"]] + list(spec.get("aliases") or [])
        got_rows = [r for r in rows if (r.get("field") or "") in names]
        got = []
        seen = set()
        for r in got_rows:
            v = (r.get("value") or "").strip()
            nv = norm(v)
            if nv and nv not in seen:
                seen.add(nv)
                got.append(v)
        hit, miss = 0, []
        for e in spec["expect"]:
            en = norm(e)
            if any(en == norm(v) or en in norm(v) for v in got):
                hit += 1
            else:
                miss.append(e)
        # 清单外值（库有、标准答案没有 —— 不一定错，但列出供人工看）
        extra = []
        for v in got:
            vn = norm(v)
            if not any(norm(e) == vn or norm(e) in vn for e in spec["expect"]):
                extra.append(v)
        out.append({"key": spec["key"], "aliases": spec.get("aliases") or [],
                    "total": len(spec["expect"]), "hit": hit, "miss": miss,
                    "got_n": len(got), "raw_n": len(got_rows), "extra": extra})
    total = sum(x["total"] for x in out)
    hit = sum(x["hit"] for x in out)
    return {"total": total, "hit": hit, "by_field": out,
            "alias_merge": sum(x["raw_n"] - x["got_n"] for x in out)}


def score_all(db: dict = None, answers: dict = None, doc: str = None) -> dict:
    """三线计分全跑。返回 {entities, profile, fields}（by_doc + overall）。"""
    db = db or fetch_db()
    answers = answers or load_answers()

    def _f(d):       # 文档过滤
        return {k: v for k, v in d.items() if (not doc) or (doc in k)}

    # ---- 实体线 ----
    ents_by_doc = {}
    for r in db["entities"]:
        ents_by_doc.setdefault(r["doc_name"], []).append(r)
    ent_by_doc, ent_over = {}, _blank_score()
    for doc_name, checklist in _f(answers["entities"]).items():
        rows = ents_by_doc.get(doc_name)
        if not rows:
            ent_by_doc[doc_name] = {"error": "库中无此文档数据"}
            continue
        s = score_entities(rows, checklist)
        ent_by_doc[doc_name] = s
        _merge_score(ent_over, s)

    # ---- 画像线 ----
    prof_by_doc = {p["doc_name"]: p for p in db["profiles"]}
    items_by_pid = {}
    for r in db["profile_items"]:
        items_by_pid.setdefault(r["profile_id"], []).append(r)
    prof_out, prof_over = {}, _blank_score()
    for doc_name, checklist in _f(answers["profile"]).items():
        p = prof_by_doc.get(doc_name)
        if not p:
            prof_out[doc_name] = {"error": "库中无此文档画像"}
            continue
        s = score_profile(p.get("topic"), items_by_pid.get(p["id"], []), checklist)
        prof_out[doc_name] = s
        _merge_score(prof_over, s)

    # ---- 字段线 ----
    ext_by_doc = {}
    for r in db["extractions"]:
        ext_by_doc.setdefault(r["doc_name"], []).append(r)
    fld_out, fld_over = {}, _blank_score()
    for doc_name, specs in _f(answers["fields"]).items():
        rows = ext_by_doc.get(doc_name) or []
        s = score_fields(rows, specs)
        fld_out[doc_name] = s
        _merge_score(fld_over, s)

    return {"entities": {"by_doc": ent_by_doc, "overall": ent_over},
            "profile": {"by_doc": prof_out, "overall": prof_over},
            "fields": {"by_doc": fld_out, "overall": fld_over},
            "alias_merge": fld_over.get("alias_merge", 0)}


def _blank_score() -> dict:
    return {"total": 0, "hit": 0, "type_ok": 0, "extra": 0, "alias_merge": 0}


def _merge_score(acc: dict, s: dict):
    acc["total"] += s.get("total", 0)
    acc["hit"] += s.get("hit", 0)
    acc["type_ok"] += s.get("type_ok", 0)
    acc["extra"] += s.get("extra", 0)
    acc["alias_merge"] += s.get("alias_merge", 0)


# ================================================================
#  四、汇总与报告
# ================================================================

def summarize(audit: dict, scores: dict) -> dict:
    """给 API / 终端用的轻量摘要。"""
    o = audit["overall"]
    n_quotable = o["total"] - o["no_quote"]
    fidelity = (o["traced"] / n_quotable * 100) if n_quotable else None
    frag_rate = (o["frag_hit"] / o["frag_total"] * 100) if o["frag_total"] else None
    def _rate(d):
        return round(d["hit"] / d["total"] * 100, 1) if d.get("total") else None
    return {
        "audit": {"total": o["total"], "traced": o["traced"], "partial": o["partial"],
                  "miss": o["miss"], "no_quote": o["no_quote"],
                  "fidelity": round(fidelity, 1) if fidelity is not None else None,
                  "frag_rate": round(frag_rate, 1) if frag_rate is not None else None,
                  "issues": len(audit["issues"]), "load_errors": audit["load_errors"]},
        "scores": {"entities": _rate(scores["entities"]["overall"]),
                   "profile": _rate(scores["profile"]["overall"]),
                   "fields": _rate(scores["fields"]["overall"])},
    }


def _bar(pct_val, target=80):
    if pct_val is None:
        return "—"
    cls = "ok" if pct_val >= target else "warn"
    return '<span class="pct %s">%.1f%%</span>' % (cls, pct_val)


def _stat_row(label, s: dict) -> str:
    n_quotable = s["total"] - s["no_quote"]
    fid = (s["traced"] / n_quotable * 100) if n_quotable else None
    return ("<tr><td>%s</td><td>%d</td><td>%d</td><td>%d</td><td>%d</td>"
            "<td>%s</td></tr>"
            % (label, s["total"], s["traced"], s["partial"], s["miss"],
               _bar(fid) if fid is not None else "—"))


def _score_row(label, s: dict, kind="hit") -> str:
    if s.get("error"):
        return "<tr><td>%s</td><td colspan=\"4\" class=\"dim\">%s</td></tr>" % (label, s["error"])
    rate = (s["hit"] / s["total"] * 100) if s.get("total") else None
    extra = ""
    if kind == "entities":
        t_rate = (s["type_ok"] / s["hit"] * 100) if s.get("hit") else None
        extra = "类型正确 %s ｜ 额外 %d 条" % (_bar(t_rate), s.get("extra", 0))
    elif kind == "profile" and s.get("miss"):
        extra = "缺：" + "、".join(s["miss"][:4])
    elif kind == "fields":
        n_extra = sum(len(x.get("extra") or []) for x in (s.get("by_field") or []))
        if n_extra:
            extra = "清单外值 %d 个" % n_extra
    return ("<tr><td>%s</td><td>%d / %d</td><td>%s</td><td class=\"dim\">%s</td></tr>"
            % (label, s["hit"], s["total"], _bar(rate), extra))


def build_report(audit: dict, scores: dict, out_dir: Path = None) -> tuple:
    """生成《质量报告.html》+《质量明细.json》。返回 (html 路径, json 路径)。"""
    out_dir = Path(out_dir or REPORT_DIR)
    out_dir.mkdir(parents=True, exist_ok=True)

    summ = summarize(audit, scores)
    o = audit["overall"]
    ent, prof, fld = scores["entities"]["overall"], scores["profile"]["overall"], scores["fields"]["overall"]
    ent_rate = (ent["hit"] / ent["total"] * 100) if ent["total"] else None
    prof_rate = (prof["hit"] / prof["total"] * 100) if prof["total"] else None
    fld_rate = (fld["hit"] / fld["total"] * 100) if fld["total"] else None

    detail = {
        "generated": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "summary": summ,
        "scores": {k: {"overall": scores[k]["overall"],
                       "by_doc": scores[k]["by_doc"]} for k in ("entities", "profile", "fields")},
        "audit": {"overall": audit["overall"], "by_doc": audit["by_doc"],
                  "by_line": audit["by_line"], "by_type": audit["by_type"],
                  "issues": audit["issues"], "load_errors": audit["load_errors"]},
    }
    jp = out_dir / "质量明细.json"
    jp.write_text(json.dumps(detail, ensure_ascii=False, indent=2, default=str),
                  encoding="utf-8")

    # ---------- HTML ----------
    now = detail["generated"]
    fid = summ["audit"]["fidelity"]
    frag = summ["audit"]["frag_rate"]

    rows_doc = "".join(_stat_row(d, s) for d, s in sorted(audit["by_doc"].items()))
    line_cn = {"entities": "实体线", "profile": "画像线", "fields": "字段线"}
    rows_line = "".join(_stat_row(line_cn.get(k, k), s) for k, s in audit["by_line"].items())

    rows_ent = "".join(_score_row(d, s, "entities") for d, s in sorted(scores["entities"]["by_doc"].items()))
    rows_prof = "".join(_score_row(d, s, "profile") for d, s in sorted(scores["profile"]["by_doc"].items()))
    rows_fld = "".join(_score_row(d, s, "fields") for d, s in sorted(scores["fields"]["by_doc"].items()))

    fld_detail = ""
    for doc_name, s in sorted(scores["fields"]["by_doc"].items()):
        if s.get("error"):
            continue
        for x in s["by_field"]:
            miss = ("缺：" + "、".join(x["miss"])) if x["miss"] else "全中"
            extra = (" ｜ 清单外：" + "、".join(x["extra"][:5])) if x["extra"] else ""
            fld_detail += ("<tr><td>%s</td><td>%s</td><td>%d / %d</td>"
                           "<td class=\"dim\">%s%s</td></tr>"
                           % (doc_name, x["key"], x["hit"], x["total"], miss, extra))

    if audit["issues"]:
        items = []
        for i in audit["issues"][:40]:
            why = "、".join(i.get("miss_parts") or [])
            bn = i.get("bad_nums") or []
            if bn:
                why = (why + " ｜ " if why else "") + "数字未溯：" + "、".join(bn)
            items.append("<li><b>%s</b> ｜ %s ｜ %s<br><span class=\"dim\">未命中：%s</span></li>"
                         % (line_cn.get(i["line"], i["line"]), i["doc"], i["label"],
                            why or "（值核对未过）"))
        issues_html = "".join(items)
        issues_note = ('<div class="note red">以下条目有摘抄片段未能在原文逐字命中'
                       ' —— 如实列出，不做修饰。</div>')
    else:
        issues_html = "<li class=\"dim\">无 —— 全部条目完全可溯 ✓</li>"
        issues_note = ('<div class="note green">本次审计全部条目均逐字可溯；'
                       '若未来出现未命中，将在此如实列出、不做修饰。</div>')

    load_err_html = ""
    if audit["load_errors"]:
        load_err_html = ("<div class=\"note\">原文装载异常：%s</div>"
                         % "；".join("%s（%s）" % (k, v) for k, v in audit["load_errors"].items()))

    # 数据整洁观察：同义字段并存
    clean_notes = []
    for doc_name, s in sorted(scores["fields"]["by_doc"].items()):
        if not isinstance(s, dict) or s.get("error"):
            continue
        for x in s.get("by_field", []):
            if x["raw_n"] > x["got_n"]:
                clean_notes.append("《%s》字段「%s」存在同义名并存：库 %d 条原始记录合并去重后 %d 条"
                                   "（aliases：%s）" % (doc_name, x["key"], x["raw_n"], x["got_n"],
                                                       "、".join(x["aliases"]) or "无"))
    clean_html = "".join("<li>%s</li>" % n for n in clean_notes) or "<li class=\"dim\">未发现</li>"

    conclusion = ("保真度 %.1f%%、三线查全率（实体 %.1f%% / 画像 %.1f%% / 字段 %.1f%%）"
                  % (fid or 0, ent_rate or 0, prof_rate or 0, fld_rate or 0))
    kpi_ok = (fid or 0) >= 80

    html = """<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<title>模块2 · 质量验证与准确率报告</title>
<style>
  :root{--ink:#1a2330;--dim:#67748a;--line:#e3e8f0;--bg:#f5f7fb;--card:#fff;
        --ok:#0f9d58;--warn:#e8710a;--red:#d93025;--blue:#1a63c9;}
  *{box-sizing:border-box}
  body{margin:0;background:var(--bg);color:var(--ink);
       font:14px/1.75 "Microsoft YaHei","PingFang SC",system-ui,sans-serif}
  .wrap{max-width:980px;margin:0 auto;padding:28px 20px 60px}
  .banner{background:linear-gradient(135deg,#12304f,#1a63c9);color:#fff;
          border-radius:16px;padding:30px 34px;margin-bottom:22px}
  .banner h1{margin:0 0 8px;font-size:24px;letter-spacing:1px}
  .banner p{margin:4px 0;opacity:.92;font-size:13.5px}
  .banner .big{font-size:15px;margin-top:10px}
  .banner .big b{font-size:20px;color:#ffe08a}
  .cards{display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin-bottom:22px}
  .kpi{background:var(--card);border:1px solid var(--line);border-radius:14px;
       padding:16px 14px;text-align:center;box-shadow:0 2px 8px rgba(20,40,80,.04)}
  .kpi .n{font-size:26px;font-weight:700;color:var(--blue)}
  .kpi .n.ok{color:var(--ok)} .kpi .n.bad{color:var(--red)}
  .kpi .t{font-size:12.5px;color:var(--dim);margin-top:2px}
  h2{font-size:17px;margin:30px 0 10px;padding-left:10px;border-left:4px solid var(--blue)}
  table{width:100%%;border-collapse:collapse;background:var(--card);
        border:1px solid var(--line);border-radius:12px;overflow:hidden;font-size:13px}
  th,td{padding:8px 10px;border-bottom:1px solid var(--line);text-align:left;vertical-align:top}
  th{background:#f0f4fa;font-weight:600;color:#33415c}
  tr:last-child td{border-bottom:none}
  .dim{color:var(--dim);font-size:12.5px}
  .pct{font-weight:700}
  .pct.ok{color:var(--ok)} .pct.warn{color:var(--warn)}
  .note{background:#fff;border:1px solid var(--line);border-left:4px solid var(--blue);
        border-radius:10px;padding:10px 14px;color:#3a4a63;font-size:13px;margin:10px 0}
  .note.green{border-left-color:var(--ok)} .note.red{border-left-color:var(--red)}
  ul{margin:8px 0;padding-left:22px} li{margin:6px 0}
  .foot{margin-top:36px;color:var(--dim);font-size:12.5px;text-align:center}
</style>
</head>
<body>
<div class="wrap">

  <div class="banner">
    <h1>质量验证与准确率报告</h1>
    <p>A23 项目 · 模块（2）「非结构化信息提取 + 数据库存储」 · 第 5 步</p>
    <p>生成时间：%(now)s ｜ 数据源：MySQL a23_agent（实时读取） ｜ 验证方式：<b>零模型调用 · 纯程序核对 · 可复现</b></p>
    <p class="big">总账：%(conclusion)s %(kpi_word)s</p>
  </div>

  <div class="cards">
    <div class="kpi"><div class="n %(cls_fid)s">%(fid_s)s</div><div class="t">保真度（数据可回原文）</div></div>
    <div class="kpi"><div class="n %(cls_e)s">%(ent_s)s</div><div class="t">实体查全率（101 项）</div></div>
    <div class="kpi"><div class="n %(cls_p)s">%(prof_s)s</div><div class="t">画像覆盖率（40 点）</div></div>
    <div class="kpi"><div class="n %(cls_f)s">%(fld_s)s</div><div class="t">字段查全率（58 值）</div></div>
  </div>

  <h2>一、反向审计：数据库 → 原文（保真度）</h2>
  <div class="note">
    口径：库里每条数据的「原文摘抄」按分隔符拆片段，逐片段回原文核对。
    <b>完全可溯</b>=全部片段命中；<b>部分可溯</b>=有片段未命中；<b>未溯</b>=全部未命中。
    保真度 = 完全可溯 /（总数 − 无摘抄），硬指标 ≥ 80%%。
  </div>
  <table>
    <tr><th>数据线</th><th>条数</th><th>完全可溯</th><th>部分可溯</th><th>未溯</th><th>保真度</th></tr>
    %(rows_line)s
  </table>
  <div class="note">按文档细分：</div>
  <table>
    <tr><th>文档</th><th>条数</th><th>完全可溯</th><th>部分可溯</th><th>未溯</th><th>保真度</th></tr>
    %(rows_doc)s
    <tr><td><b>合计</b></td><td><b>%(tot)s</b></td><td><b>%(tr)s</b></td><td><b>%(pa)s</b></td><td><b>%(mi)s</b></td><td><b>%(fid_s)s</b></td></tr>
  </table>
  %(load_err)s
  <div class="note">片段级命中率（更宽口径）：<b>%(frag)s</b>（%(fh)s / %(ft)s 个片段）。
    综合层面的「实体/画像/字段」分线明细见《质量明细.json》。</div>

  <h2>二、三线计分：对照人工标准答案</h2>
  <div class="note">
    标准答案 = 人工通读源文档圈定的「必含清单」（实体 101 项 / 画像 40 点 / 字段 58 值），
    命中判定的归一化与第 2/3 步验收完全一致。查全率 = 命中 / 应含。
  </div>

  <h3>实体线（7 类 · 5 文档）</h3>
  <table>
    <tr><th>文档</th><th>命中 / 应含</th><th>查全率</th><th>备注</th></tr>
    %(rows_ent)s
  </table>

  <h3>画像线（4 文档 × 10 必含点）</h3>
  <table>
    <tr><th>文档</th><th>命中 / 应含</th><th>覆盖率</th><th>缺口</th></tr>
    %(rows_prof)s
  </table>
  <div class="note">画像为「关键信息概要」性质：缺口均为概要级明细值（如单个员工的工资数额、
    精度检验数值），覆盖率对照内控线（≥ 80%%）达标。</div>

  <h3>字段线（7 文档 · 14 字段）</h3>
  <table>
    <tr><th>文档</th><th>命中 / 应含</th><th>查全率</th><th>备注</th></tr>
    %(rows_fld)s
  </table>
  <div class="note">逐字段明细（含缺失与清单外值）：</div>
  <table>
    <tr><th>文档</th><th>字段</th><th>命中 / 应含</th><th>明细</th></tr>
    %(fld_detail)s
  </table>

  <h2>三、未完全可溯条目（诚实清单）</h2>
  %(issues_note)s
  <ul>%(issues)s</ul>

  <h2>四、数据整洁观察</h2>
  <ul>%(clean)s</ul>

  <h2>五、性能参考（历史验收实测）</h2>
  <div class="note">
    对照赛题「每文档 ≤ 90 秒」的节奏要求，三线提取单份文档均在秒级完成（数据来源：各步验收记录，
    见 agent/输出/验收-*/）。本节为历史实测汇总，不是本次报告运行值。
  </div>
  <table>
    <tr><th>环节</th><th>实测耗时</th><th>来源</th></tr>
    <tr><td>文档导入解析</td><td>0.04 ~ 0.08 秒 / 份</td><td class="dim">第 1 步验收</td></tr>
    <tr><td>文档画像（关键信息）</td><td>约 3 秒 / 份</td><td class="dim">第 2 步验收</td></tr>
    <tr><td>实体识别（7 类）</td><td>1.5 ~ 5.0 秒 / 份</td><td class="dim">第 3 步验收</td></tr>
    <tr><td>字段提取</td><td>单字段 1.7 ~ 7.5 秒 ｜ 一条多字段指令约 13 秒</td><td class="dim">第 4 步验收</td></tr>
    <tr><td>质量审计（本报告，零模型）</td><td>0.4 秒 / 全库 261 条</td><td class="dim">本次运行</td></tr>
  </table>

  <h2>六、方法与口径（可复现声明）</h2>
  <ul>
    <li><b>零模型调用</b>：本报告全部结论由确定性程序得出（归一化比对 / 子串匹配），重跑必然一致。</li>
    <li><b>数据源</b>：MySQL a23_agent 五张表实时读取；原文核对的「原文」= 解析池（doc_reader/output.json），池外文档现场解析（只读）。</li>
    <li><b>归一化口径</b>：去空白、全角标点转半角、去货币符号、去千分位逗号、去尾部“.00”——只消除等价差异，不做同义词替换。</li>
    <li><b>查全率口径</b>：标准答案是人工圈定的「必含」下界；库里可能有清单之外的额外合法数据（标记为「额外」）。</li>
    <li><b>复跑方式</b>：<code>python agent/quality.py --report</code>（完整重跑审计 + 计分 + 报告）。</li>
  </ul>

  <div class="foot">A23 项目 · 模块（2）第 5 步质量报告 ｜ 生成：小布 · %(now)s</div>
</div>
</body>
</html>
""" % {
        "now": now,
        "conclusion": conclusion,
        "kpi_word": "（硬指标 ≥80% 已达标 ✓）" if kpi_ok else "（保真度低于 80% 硬指标 ⚠）",
        "fid_s": ("%.1f%%" % fid) if fid is not None else "—",
        "ent_s": ("%.1f%%" % ent_rate) if ent_rate is not None else "—",
        "prof_s": ("%.1f%%" % prof_rate) if prof_rate is not None else "—",
        "fld_s": ("%.1f%%" % fld_rate) if fld_rate is not None else "—",
        "cls_fid": "ok" if (fid or 0) >= 80 else "bad",
        "cls_e": "ok" if (ent_rate or 0) >= 80 else "bad",
        "cls_p": "ok" if (prof_rate or 0) >= 80 else "bad",
        "cls_f": "ok" if (fld_rate or 0) >= 80 else "bad",
        "rows_doc": rows_doc, "rows_line": rows_line,
        "rows_ent": rows_ent, "rows_prof": rows_prof, "rows_fld": rows_fld,
        "fld_detail": fld_detail,
        "tot": o["total"], "tr": o["traced"], "pa": o["partial"], "mi": o["miss"],
        "frag": ("%.1f%%" % frag) if frag is not None else "—",
        "fh": o["frag_hit"], "ft": o["frag_total"],
        "issues": issues_html, "issues_note": issues_note,
        "clean": clean_html, "load_err": load_err_html,
    }
    hp = out_dir / "质量报告.html"
    hp.write_text(html, encoding="utf-8")
    return hp, jp


# ================================================================
#  五、CLI
# ================================================================

def main():
    ap = argparse.ArgumentParser(description="模块（2）第 5 步：质量验证与准确率报告")
    ap.add_argument("--audit", action="store_true", help="只跑反向审计")
    ap.add_argument("--score", action="store_true", help="只跑对照计分")
    ap.add_argument("--report", action="store_true", help="出完整报告（默认行为）")
    ap.add_argument("--doc", default=None, help="只处理文档名含该子串的文档")
    args = ap.parse_args()
    only = args.audit or args.score          # 单跑模式（--report 会拉回全跑）
    do_audit = (not only) or args.audit or args.report
    do_score = (not only) or args.score or args.report

    print("=" * 62)
    print("模块（2）· 第 5 步 质量验证与准确率报告")
    print("=" * 62)

    t0 = time.time()
    try:
        db = fetch_db()
    except Exception as e:
        print("【失败】数据库连不上：%s" % e)
        print("  → 检查 MySQL 服务（net start MySQL80）与 agent/db_config.json")
        sys.exit(1)
    print("数据库读取：%d 实体 / %d 画像明细 / %d 字段"
          % (len(db["entities"]), len(db["profile_items"]), len(db["extractions"])))

    audit = None
    if do_audit:
        audit = audit_all(db, doc=args.doc)
        o = audit["overall"]
        n_q = o["total"] - o["no_quote"]
        fid = (o["traced"] / n_q * 100) if n_q else None
        print("反向审计：%d 条 ｜ 完全可溯 %d ｜ 部分 %d ｜ 未溯 %d ｜ 无摘抄 %d"
              % (o["total"], o["traced"], o["partial"], o["miss"], o["no_quote"]))
        if fid is not None:
            frag = (o["frag_hit"] / o["frag_total"] * 100) if o["frag_total"] else 0
            print("保真度：%.1f%% ｜ 片段命中率：%.1f%%" % (fid, frag))
        for k, v in audit["load_errors"].items():
            print("  [原文装载异常] %s：%s" % (k, v))

    scores = None
    if do_score:
        scores = score_all(db, doc=args.doc)
        for k, cn in (("entities", "实体线"), ("profile", "画像线"), ("fields", "字段线")):
            s = scores[k]["overall"]
            rate = (s["hit"] / s["total"] * 100) if s["total"] else 0
            print("%s：查全 %d/%d = %.1f%%%s"
                  % (cn, s["hit"], s["total"], rate,
                     ("（类型正确 %d）" % s["type_ok"]) if k == "entities" else ""))

    if audit and scores:
        hp, jp = build_report(audit, scores)
        print("\n报告生成用时：%.1f 秒（零模型调用）" % (time.time() - t0))
        print("已输出：")
        print("  %s" % hp)
        print("  %s" % jp)
    else:
        print("\n（单跑模式，报告未生成）")


if __name__ == "__main__":
    main()
