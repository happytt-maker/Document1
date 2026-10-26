# -*- coding: utf-8 -*-
"""
entity_ops.py —— 模块（2）· 第 3 步：实体数据提取（命名实体 · 7 类）

链路：一句话「识别这份文档里的实体数据」→ ①解析（parse_intent，op=entities）
      → ②装载文档块（doc_reader/output.json，解析池里的最新内容）
      → ③预算选块（实体比梗概更需要覆盖 → 预算放宽到 8000 字符）
      → ④模型识别 7 类实体（人名 / 机构 / 日期 / 金额 / 编号 / 联系方式 / 产品设备）
      → ⑤程序复核（模型负责认，程序负责查）：
           · 出处定位 —— 每个实体回原文定位（复用第 2 步的跨块区域定位链）
           · 格式核验 —— 日期 / 金额 / 编号 / 联系方式 用正则二次校验（对不上标"存疑"）
           · 规则补漏 —— 正则全文档扫描，补齐模型漏报的 日期 / 金额 / 联系方式
           · 去重合并 —— 同值合并并统计出现次数
      → ⑥输出：分类实体表（屏幕报告 + <文档名>_实体.md + .json）

不编造原则：模型给的每个值都要回原文定位；定位不上的标「未定位」，绝不圆谎。

跑法（在项目根目录）：
  python agent/entity_ops.py "识别设备采购合同里的实体数据"
"""
import json
import re
import sys
import time
from pathlib import Path

# 输出兜底：遇到终端表示不了的字符就替换，不让程序崩
try:
    sys.stdout.reconfigure(errors="replace")
except Exception:
    pass

BASE = Path(__file__).resolve().parents[1]                  # = 研一项目/
AGENT_DIR = Path(__file__).resolve().parent                 # = 研一项目/agent/
sys.path.insert(0, str(AGENT_DIR))                          # 复用同目录的解析器与画像链

from parse_intent import ask_model, parse_json_reply        # noqa: E402
from summary_ops import (load_doc, pick_for_budget, build_digest,   # noqa: E402
                         locate_quote, loc_desc, _norm)

OUT_DIR = BASE / "agent" / "输出"

ENTITY_BUDGET = 8000   # 实体提取预算（字符）：比画像(6000)更宽 —— 实体散布全文，覆盖率优先

TYPE_ORDER = ["person", "org", "date", "money", "code", "contact", "product"]
TYPE_CN = {"person": "人名", "org": "机构", "date": "日期", "money": "金额",
           "code": "编号", "contact": "联系方式", "product": "产品设备"}

ENTITY_SYSTEM = (
    "你是文档实体识别助手。用户给出一份文档的内容，你要识别里面所有"
    "「实体数据」，归入 7 类：人名、机构、日期、金额、编号、联系方式、产品设备。"
    "所有实体值必须逐字来自原文，绝对不许改写、拼凑或编造；"
    "文档里没有出现的一律不报。严格只输出 JSON，不要任何解释。"
)

# ==================== 一、强模式正则（核验 + 补漏共用） ====================
# 日期：2026年3月12日 / 2026-03-16 / 2026 年 6 月 30 日 / 04-15（月-日，须两位月）
RE_DATE = re.compile(
    r"(?:19|20)\d{2}\s*[-/.年]\s*\d{1,2}(?:\s*[-/.月]\s*\d{1,2}\s*日?)?"
    r"|(?<!\d)(?:0[1-9]|1[0-2])[-/.](?:0[1-9]|[12]\d|3[01])(?!\d)"
)
# 金额：￥1,800,000.00 / 1.5万元 / 5000元 / 1,800,000（千分位）/ 壹佰捌拾万元整（大写）
RE_MONEY = re.compile(
    r"[¥￥]\s*\d[\d,，]*(?:\.\d{1,2})?"
    r"|\d{1,3}(?:[,，]\d{3})+(?:\.\d{1,2})?"
    r"|\d[\d,，]*\.\d{1,2}\s*(?:万元|亿元|千元|元)"
    r"|\d[\d,，]*\s*(?:万元|亿元|千元|元)"
    r"|[零壹贰叁肆伍陆柒捌玖拾佰仟]{2,}[万亿]?[元整角分]{0,3}"
)
# 联系方式：手机 / 座机 / 邮箱
RE_CONTACT = re.compile(
    r"1[3-9]\d{9}"
    r"|0\d{2,3}[-—\s]?\d{7,8}"
    r"|[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}"
)
# 编号：字母 + 数字混合（HT-2026-0417 / E1001 / TL-850-20260512 / YS-2026-0089）
RE_CODE = re.compile(r"[A-Za-z]{1,10}\s*[-—–]?\s*\d{2,}(?:[-—–][A-Za-z0-9]+)*")


def _looks_address(v) -> bool:
    """地址的宽松判断：含 ≥2 种行政区划/街道用字。"""
    return sum(1 for c in "省市区县镇乡村路街道号" if c in (v or "")) >= 2


def verify_entity(etype: str, value: str) -> str:
    """正则二次校验。返回 ok（通过）/ suspect（格式存疑）/ na（该类无强模式）。"""
    v = value or ""
    if etype == "date":
        return "ok" if RE_DATE.search(v) else "suspect"
    if etype == "money":
        return "ok" if RE_MONEY.search(v) else "suspect"
    if etype == "contact":
        return "ok" if (RE_CONTACT.search(v) or _looks_address(v)) else "suspect"
    if etype == "code":
        return "ok" if RE_CODE.search(v) else "suspect"
    return "na"   # person / org / product：形态无强模式，跳过（对齐人工复核）


# ==================== 二、值归一（去重 / 比对用） ====================
_FULL2HALF = str.maketrans(
    "０１２３４５６７８９ＡＢＣＤＥＦＧＨＩＪＫＬＭＮＯＰＱＲＳＴＵＶＷＸＹＺ"
    "ａｂｃｄｅｆｇｈｉｊｋｌｍｎｏｐｑｒｓｔｕｖｗｘｙｚ",
    "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz")


def norm_value(v: str) -> str:
    """实体值的归一形态（只用于"是不是同一个值"的比对，不改展示原文）。"""
    v = re.sub(r"\s+", "", v or "")
    v = v.translate(_FULL2HALF)
    for a, b in (("，", ","), ("：", ":"), ("（", "("), ("）", ")")):
        v = v.replace(a, b)
    v = v.replace("￥", "").replace("¥", "")     # 币种符号不参与比对
    v = v.replace(",", "")                       # 千分位逗号不参与比对
    v = re.sub(r"\.00$", "", v)                  # 尾巴的 .00 不参与比对
    return v


def _seen(have: set, nv: str) -> bool:
    """nv 是否与已收集合里的某个值"同一个"（相等或互相包含，含长度防误判）。"""
    for h in have:
        if nv == h:
            return True
        if len(nv) >= 3 and nv in h:
            return True
        if len(h) >= 3 and h in nv:
            return True
    return False


# ==================== 三、模型识别 → 实体清单 ====================
def build_entity_prompt(doc_file: str, digest: str, n_show: int, note: str) -> str:
    return f"""请阅读文档《{doc_file}》的内容，识别里面的实体数据。

【文档内容】（共 {n_show} 块，编号仅供定位参考）
{digest}

【7 类实体】（每个实体归入最贴切的一类）
- person 人名：自然人的姓名（员工、联系人、签字人、汇报人、整理人…）
- org 机构：公司、单位、部门、车间等组织名
- date 日期：具体的年-月-日（照原文写法："2026年3月12日"、"2026-03-16"、"04-15" 都算）
- money 金额：带币种 / 单位 / 千分位 / 大写写法的钱数（如"￥1,800,000.00"、"5000元"、"壹佰捌拾万元"）
- code 编号：合同号、工号、设备编号、验收编号等编码（如"HT-2026-0417"、"E1001"）
- contact 联系方式：电话、邮箱、地址
- product 产品设备：设备名称、规格型号、产品名（如"数控加工中心"、"VMC-850B"）

【输出 JSON】（严格按这个结构）
{{
  "entities": [
    {{"type": "person", "value": "吴晓敏", "quote": "**汇报人**：吴晓敏"}}
  ]
}}

规则：
- value 必须逐字照抄文档中的原文（保留原有写法、大小写、标点）
- quote：实体值所在处的原文片段（5~40 字，逐字照抄原文），供程序定位出处
- 同一个值只报一次（同一实体不要重复出现）
- 文档里出现的都要报（尽量报全），但绝不编造、不推测；拿不准的不报
- 表格类内容：同一列 / 同一行的同类实体（如整列姓名、整列工号、整行供应商）要列全，不要只举一两个例子
- 注意：设备型号（如 VMC-850B、IRB-2600）属于「产品设备」，不要归为「编号」
- 按 person → org → date → money → code → contact → product 的顺序输出"""


def ask_entities(doc_file: str, digest: str, n_show: int, note: str):
    """调模型识别实体。返回 (dict 或 None, 耗时秒, 说明)。"""
    prompt = build_entity_prompt(doc_file, digest, n_show, note)
    t0 = time.time()
    content, _usage, cost = ask_model(prompt, system=ENTITY_SYSTEM)
    if not content:
        return None, cost, "模型调用失败（网络或服务问题），请重跑一次。"
    data = parse_json_reply(content)
    if not isinstance(data, dict):
        return None, cost, "模型返回不是 JSON，无法解析。"
    return data, cost, ""


_TYPE_ALIAS = {
    "人名": "person", "人物": "person", "员工": "person", "姓名": "person",
    "机构": "org", "组织": "org", "公司": "org", "部门": "org", "单位": "org",
    "日期": "date", "时间": "date",
    "金额": "money", "价格": "money", "钱": "money",
    "编号": "code", "编码": "code", "号码": "code",
    "联系方式": "contact", "电话": "contact", "邮箱": "contact", "地址": "contact",
    "产品设备": "product", "产品": "product", "设备": "product", "型号": "product",
}


def _fix_type(t) -> str:
    """类型归一：英文字段 / 中文名 / 常见别名 → 标准英文键。识别不了返回 None。"""
    s = str(t or "").strip()
    if s.lower() in TYPE_ORDER:
        return s.lower()
    return _TYPE_ALIAS.get(s)


def clean_entities(raw) -> list:
    """清理模型输出：类型归一 + 去重 + 字段整理。"""
    out, seen = [], set()
    if not isinstance(raw, list):
        return out
    for e in raw:
        if not isinstance(e, dict):
            continue
        t = _fix_type(e.get("type"))
        v = str(e.get("value") or "").strip()
        if not t or not v:
            continue
        key = (t, norm_value(v))
        if key in seen:
            continue
        seen.add(key)
        out.append({"type": t, "value": v,
                    "quote": str(e.get("quote") or "").strip(),
                    "source": "model"})
    return out


# ==================== 四、出处定位 / 出现次数 ====================
def _is_numeric(v: str) -> bool:
    """是不是「纯数字值」（可能带千分位逗号 / 小数点）。"""
    return bool(re.fullmatch(r"[\d,，.]+", (v or "").strip()))


def _num_pattern(v: str):
    """数字值的匹配式：要求数字边界（防 "2800" 撞进 "12800" 的子串）。"""
    return re.compile(r"(?<!\d)" + re.escape(v.strip()) + r"(?!\d)")


def _find_value_block(value: str, blocks: list):
    """按实体值找块：先原样子串，后归一子串。返回 (块序号 或 None, 方式)。

    纯数字值要求数字边界（"2800" 不算命中 "12800"）。
    """
    v = (value or "").strip()
    if not v:
        return None, "empty"
    numeric = _is_numeric(v)
    rx = _num_pattern(v) if numeric else None
    for i, b in enumerate(blocks):
        t = b.get("text") or ""
        if v in t and (rx is None or rx.search(t)):
            return i, "exact"
    if numeric:
        # 数字的归一兜底："2,800" ↔ "2800"（去千分位后仍要数字边界）
        vn = v.replace(",", "").replace("，", "")
        rxn = _num_pattern(vn)
        for i, b in enumerate(blocks):
            tn = (b.get("text") or "").replace(",", "").replace("，", "")
            if rxn.search(tn):
                return i, "ws-num"
        return None, "miss"
    vn = _norm(v)
    if len(vn) >= 2:
        for i, b in enumerate(blocks):
            if vn in _norm(b.get("text") or ""):
                return i, "ws"
    return None, "miss"


def locate_entity(value: str, quote: str, blocks: list):
    """实体出处：①按值定位（最直接）②失败用摘抄走完整定位链（含跨块区域）。"""
    i, how = _find_value_block(value, blocks)
    if i is not None:
        return loc_desc(blocks[i]), True
    d, _how2 = locate_quote(quote, blocks)
    if d:
        return d, True
    return "未定位", False


def count_in_blocks(blocks: list, value: str) -> int:
    """值在全文的出现次数。纯数字要求数字边界（"2800" 不计 "12800" 里的子串）；
    计数为 0 时用归一形态兜底。至少记 1 次。"""
    v = (value or "").strip()
    if not v:
        return 0
    if _is_numeric(v):
        rx = _num_pattern(v)
        total = sum(len(rx.findall(b.get("text") or "")) for b in blocks)
        if not total:
            vn = v.replace(",", "").replace("，", "")
            rxn = _num_pattern(vn)
            total = sum(len(rxn.findall((b.get("text") or "").replace(",", "").replace("，", "")))
                        for b in blocks)
        return total or 1
    total = sum((b.get("text") or "").count(v) for b in blocks)
    if total:
        return total
    vn = _norm(v)
    if len(vn) >= 2:
        total = sum(_norm(b.get("text") or "").count(vn) for b in blocks)
    return total or 1


# ==================== 五、规则补漏（正则全文档扫描） ====================
def _ctx_quote(txt: str, s: int, e: int, pad: int = 14) -> str:
    """截取值所在处的上下文片段（供展示 / 定位）。"""
    a, b = max(0, s - pad), min(len(txt), e + pad)
    return txt[a:b].replace("\n", " ").strip()


def scan_missing(blocks: list, ents: list) -> list:
    """正则扫描全文档，补模型漏报的 日期 / 金额 / 联系方式（形态最特异的三类）。

    只补"模型完全没报到"的值；与已报值（归一）相等或互相包含的都跳过。
    """
    have = set(norm_value(e["value"]) for e in ents)
    found = []
    for bi, b in enumerate(blocks):
        txt = b.get("text") or ""
        if not txt:
            continue
        for etype, rx in (("date", RE_DATE), ("money", RE_MONEY), ("contact", RE_CONTACT)):
            for m in rx.finditer(txt):
                v = m.group(0).strip()
                if not v or len(v) < 4:
                    continue
                nv = norm_value(v)
                if _seen(have, nv):
                    continue
                have.add(nv)
                found.append({"type": etype, "value": v,
                              "quote": _ctx_quote(txt, m.start(), m.end()),
                              "source": "rule"})
    return found


# ==================== 六、整理输出 ====================
def format_screen(doc_file: str, doc_type: str, blocks: list, ents: list,
                  cmd: str, total: float, note: str) -> str:
    L = []
    L.append("=" * 62)
    L.append("【实体】一句话 → 实体数据识别（模块2 · 第3步）")
    L.append("=" * 62)
    L.append("指令：%s" % cmd)
    L.append("文档：%s（%s ｜ %d 块）" % (doc_file, doc_type or "?", len(blocks)))
    if note:
        L.append("说明：%s" % note)
    L.append("")

    counts = {}
    for e in ents:
        counts[e["type"]] = counts.get(e["type"], 0) + 1
    overview = " · ".join("%s %d" % (TYPE_CN[t], counts.get(t, 0)) for t in TYPE_ORDER)
    L.append("◆ 总览：共 %d 个实体 ｜ %s" % (len(ents), overview))

    n_hit = sum(1 for e in ents if e["loc_ok"])
    L.append("◆ 出处定位：%d/%d = %.0f%%" % (n_hit, len(ents), 100.0 * n_hit / max(len(ents), 1)))

    vstats = []
    for t in ("date", "money", "code", "contact"):
        g = [e for e in ents if e["type"] == t]
        if not g:
            continue
        ok = sum(1 for e in g if e["verify"] == "ok")
        vstats.append("%s %d/%d" % (TYPE_CN[t], ok, len(g)))
    L.append("◆ 格式核验（正则二次校验）：%s" % (" ｜ ".join(vstats) if vstats else "本次无强模式实体"))

    n_rule = sum(1 for e in ents if e["source"] == "rule")
    L.append("◆ 规则补漏：正则全扫描补充 %d 条" % n_rule)
    L.append("")

    for t in TYPE_ORDER:
        group = [e for e in ents if e["type"] == t]
        if not group:
            continue
        L.append("◆ 【%s】%d 条" % (TYPE_CN[t], len(group)))
        for e in group:
            tag = " ⚠格式存疑" if e["verify"] == "suspect" else ""
            src = "（规则补漏）" if e["source"] == "rule" else ""
            loc = e["loc"] if e["loc_ok"] else "未定位（摘抄与原文有出入）"
            L.append("  · %s ×%d — %s%s%s" % (e["value"], e["count"], loc, tag, src))
        L.append("")
    L.append("用时 %.1f 秒" % total)
    return "\n".join(L)


def format_md(doc_file: str, blocks: list, ents: list, cmd: str,
              total: float, tstamp: str, note: str) -> str:
    L = []
    L.append("# 实体数据表")
    L.append("")
    L.append("- **文档**：%s" % doc_file)
    L.append("- **指令**：%s" % cmd)
    L.append("- **时间**：%s ｜ 用时 %.1f 秒" % (tstamp, total))
    if note:
        L.append("- **说明**：%s" % note)
    counts = {}
    for e in ents:
        counts[e["type"]] = counts.get(e["type"], 0) + 1
    overview = " · ".join("%s %d" % (TYPE_CN[t], counts.get(t, 0)) for t in TYPE_ORDER)
    n_hit = sum(1 for e in ents if e["loc_ok"])
    L.append("- **总览**：共 %d 个实体 ｜ %s" % (len(ents), overview))
    L.append("- **出处定位**：%d/%d" % (n_hit, len(ents)))
    L.append("")
    for t in TYPE_ORDER:
        group = [e for e in ents if e["type"] == t]
        if not group:
            continue
        L.append("## %s（%d 条）" % (TYPE_CN[t], len(group)))
        L.append("")
        L.append("| 实体 | 次数 | 出处 | 核验 |")
        L.append("|---|---|---|---|")
        for e in group:
            v = e["value"].replace("|", "\\|")
            loc = e["loc"] if e["loc_ok"] else "_未定位_"
            if e["source"] == "rule":
                loc += "（规则补漏）"
            vf = {"ok": "✓", "suspect": "⚠ 存疑", "na": "—"}[e["verify"]]
            L.append("| %s | ×%d | %s | %s |" % (v, e["count"], loc, vf))
        L.append("")
    return "\n".join(L)


# ==================== 七、主流程 ====================
def run_entities(text: str, intent: dict, src: Path, out_dir: Path = None):
    """执行一条「实体数据提取」指令。返回 (是否成功, 报告文本, 产出文件列表)。

    与 summary_ops.run_profile / extract_ops.run_extract 同签名 —— edit.py 按 op 分流。
    """
    out_dir = Path(out_dir) if out_dir else OUT_DIR
    out_dir.mkdir(parents=True, exist_ok=True)
    t0 = time.time()

    doc, err = load_doc(src.name)
    if doc is None:
        return False, "【失败】%s" % err, []

    blocks = doc.get("blocks") or []
    if not blocks:
        return False, "【失败】《%s》在解析池里没有内容块。" % src.name, []

    idx_list, note = pick_for_budget(blocks, budget=ENTITY_BUDGET)
    digest = build_digest(blocks, idx_list)
    if not digest.strip():
        return False, "【失败】《%s》的内容块都是空的，无法识别实体。" % src.name, []

    data, cost, err2 = ask_entities(src.name, digest, len(idx_list), note)
    if data is None:
        return False, "【失败】%s（用时 %.1f 秒）" % (err2, cost), []

    ents = clean_entities(data.get("entities"))
    if not ents:
        return False, ("【失败】模型没有识别出实体（可能文档内容为空或调用异常），"
                       "请重跑一次。"), []

    # 定位 + 计数 + 核验
    for e in ents:
        e["loc"], e["loc_ok"] = locate_entity(e["value"], e.get("quote", ""), blocks)
        e["count"] = count_in_blocks(blocks, e["value"])
        e["verify"] = verify_entity(e["type"], e["value"])

    # 规则补漏（日期 / 金额 / 联系方式）
    for e in scan_missing(blocks, ents):
        e["loc"], e["loc_ok"] = locate_entity(e["value"], e.get("quote", ""), blocks)
        e["count"] = count_in_blocks(blocks, e["value"])
        e["verify"] = verify_entity(e["type"], e["value"])
        ents.append(e)

    # 排序：按 7 类固定顺序（同类保持发现顺序）
    ents.sort(key=lambda e: TYPE_ORDER.index(e["type"]))

    total = time.time() - t0
    tstamp = time.strftime("%Y-%m-%d %H:%M")

    screen = format_screen(src.name, doc.get("type", ""), blocks, ents,
                           text, total, note)
    md = format_md(src.name, blocks, ents, text, total, tstamp, note)

    md_path = out_dir / (src.stem + "_实体.md")
    json_path = out_dir / (src.stem + "_实体.json")
    md_path.write_text(md, encoding="utf-8")
    json_path.write_text(json.dumps({
        "cmd": text, "doc": src.name, "time": tstamp,
        "elapsed_sec": round(total, 2),
        "counts": {TYPE_CN[t]: sum(1 for e in ents if e["type"] == t) for t in TYPE_ORDER},
        "entities": [{"type": e["type"], "type_cn": TYPE_CN[e["type"]],
                      "value": e["value"], "quote": e["quote"],
                      "count": e["count"], "loc": e["loc"], "loc_ok": e["loc_ok"],
                      "verify": e["verify"], "source": e["source"]}
                     for e in ents],
        "note": note,
    }, ensure_ascii=False, indent=2), encoding="utf-8")

    screen += "\n【输出】%s\n【输出】%s\n【原件】未改动（提取是只读操作）" % (md_path, json_path)
    return True, screen, [str(md_path), str(json_path)]


def main():
    """命令行直跑：解析 + 实体提取一条指令（不经 edit.py 时的入口）。"""
    import argparse
    from parse_intent import scan_docs, parse_one

    ap = argparse.ArgumentParser(description="模块（2）· 实体数据提取执行器")
    ap.add_argument("text", help="一条实体数据识别指令")
    ap.add_argument("--docs", default=str(BASE / "测试数据" / "源文档"), help="文档库目录")
    args = ap.parse_args()

    docs_dir = Path(args.docs)
    doc_names = scan_docs(docs_dir)
    intent, cost, ok = parse_one(args.text, doc_names)
    if not ok:
        print("解析失败（网络或格式问题），请重跑一次。")
        sys.exit(1)
    print("解析（%.1fs）：%s" % (cost, json.dumps(intent, ensure_ascii=False)))
    if intent.get("op") != "entities":
        print("这不是一条「实体数据识别」指令（op=%s）。" % intent.get("op"))
        sys.exit(1)

    doc_name = intent.get("doc") or intent.get("doc_hint") or ""
    src = None
    for p in docs_dir.rglob("*"):
        if p.is_file() and (p.name == doc_name or p.stem == Path(doc_name).stem):
            src = p
            break
    if src is None:
        print("文档库里找不到：%s" % doc_name)
        sys.exit(1)

    ok, report, files = run_entities(args.text, intent, src)
    print(report)
    sys.exit(0 if ok else 1)


if __name__ == "__main__":
    main()
