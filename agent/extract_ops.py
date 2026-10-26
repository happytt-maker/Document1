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
  2. 口语别名：查 字段别名表.json（如"工资"→"基本工资"）      （可配置）
  3. 语义翻译：直抽为空时，让模型看"文档要素清单"，把字段翻译成
     文档里真实存在的说法（如"待办事项"→"任务内容"）——再抽一遍（模型兜底）

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
    """读"口语 → 正式"别名表（不存在或坏了就当空表，不影响流程）。"""
    if ALIAS_FILE.exists():
        try:
            data = json.loads(ALIAS_FILE.read_text(encoding="utf-8"))
            return {k: v for k, v in data.items()
                    if not k.startswith("_") and isinstance(v, str)}
        except Exception:
            pass
    return {}


def normalize_fields(fields: list) -> list:
    """字段规范化：合称拆分 + 别名映射。

    返回 [{"orig": 用户说法, "field": 实际去抽的字段, "note": 说明（没处理则空）}]
    """
    aliases = load_aliases()
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
                                "note": "「%s」拆分 → 「%s」" % (f, f.replace(k, r))})
                done = True
                break
        if done:
            continue
        # 2) 别名映射
        if f in aliases:
            out.append({"orig": f, "field": aliases[f],
                        "note": "别名表：「%s」→「%s」" % (f, aliases[f])})
        else:
            out.append({"orig": f, "field": f, "note": ""})
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

def extract_field(field: str, doc_file: str, notes: list) -> dict:
    """抽一个字段。直抽为空时翻译后重抽——多候选 + 失败追问一轮（防翻译抖动）。

    返回 {"field", "items", "mapped_to", "elapsed"}
    """
    t0 = time.time()
    res = matcher_extract.run_column(field, quiet=True, only_file=doc_file,
                                     out_dir=str(MID_DIR))
    items = list((res or {}).get("items", []))
    mapped_to = ""

    if not items:
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
                    res, items, mapped_to = res2, items2, term
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
                        res, items, mapped_to = res2, items2, term
                        break
        if mapped_to:
            notes.append("「%s」已按文档里的说法「%s」提取成功" % (field, mapped_to))
        else:
            notes.append("「%s」：文档里没有找到对应信息（诚实留空，不编造）" % field)

    return {"field": field, "items": items, "mapped_to": mapped_to,
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
    norm = normalize_fields(raw_fields)
    if not norm:
        return False, "【失败】指令里没有说明要提取哪些字段，无法执行。", []

    # 字段规范化的说明（拆分/别名才有）
    for f in norm:
        if f["note"]:
            notes.append(f["note"])

    results = []
    for f in norm:
        r = extract_field(f["field"], src.name, notes)
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
