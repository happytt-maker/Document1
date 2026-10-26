# -*- coding: utf-8 -*-
"""
summary_ops.py —— 模块（2）· 第 2 步：关键信息提取（文档画像）

链路：一句话「提取这份文档的关键信息」→ ①解析（parse_intent，op=profile）
      → ②装载文档块（doc_reader/output.json，解析池里的最新内容）
      → ③预算选块（大文档：开头 + 结尾 + 中间抽样）
      → ④模型通读 → 输出画像（主题 / 类型 / 核心要点 / 关键要素）
      → ⑤程序拿"原文摘抄"回文档定位出处（不让模型数编号）
      → ⑥输出：屏幕报告 + <文档名>_画像.md + .json

画像三件套：
  topic     一句话主题（这份文档是什么、关于什么）
  points    3~7 条核心要点（每条带原文摘抄 → 程序定位出处）
  elements  ≤8 条关键要素（时间 / 主体 / 金额 / 编号……最值得记住的具体信息）

不编造原则：要求模型所有摘抄逐字照抄原文；程序把摘抄回原文核对，
核对不上的标「未定位」，绝不替模型圆谎——出处可核是这条链的底线。

跑法（在项目根目录）：
  python agent/summary_ops.py "提取设备采购合同的关键信息"
"""
import difflib
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
sys.path.insert(0, str(AGENT_DIR))                          # 复用同目录的解析器

from parse_intent import ask_model, parse_json_reply        # noqa: E402

OUT_DIR = BASE / "agent" / "输出"
OUTPUT_JSON = BASE / "doc_reader" / "output.json"

BUDGET = 6000          # 送入模型的文档文本预算（字符）；超出走"头+尾+中间抽样"
HEAD_RATIO = 0.45      # 开头的字符占比
TAIL_RATIO = 0.15      # 结尾的字符占比（中间为剩余）

PROFILE_SYSTEM = (
    "你是文档关键信息提取助手（文档画像）。用户给出一份文档的内容，"
    "你要通读全文，输出它的关键信息画像：主题、类型、核心要点、关键要素。"
    "所有摘抄必须逐字照抄原文，绝对不许改写、拼凑或编造；"
    "文档里没有出现的信息一律不写。严格只输出 JSON，不要任何解释。"
)


# ==================== 一、装载文档块 ====================

def load_doc(doc_name: str):
    """从 doc_reader/output.json 装载目标文档。

    返回 (doc_dict 或 None, 错误说明)。
    """
    if not OUTPUT_JSON.exists():
        return None, "解析池不存在（doc_reader/output.json 没找到）——先跑一次文档读取。"
    try:
        data = json.loads(OUTPUT_JSON.read_text(encoding="utf-8"))
    except Exception as e:
        return None, "解析池读取失败：%s" % e
    for d in data.get("documents", []):
        if d.get("file") == doc_name:
            return d, ""
    return None, ("文档《%s》不在解析池里——如果它是刚导入的，"
                  "可跑 agent/pool.py --rebuild 重新并入。" % doc_name)


def loc_desc(block: dict) -> str:
    """把块的 locator 变成人话出处描述（docx 段 / docx 表格 / xlsx 格子 / 行号）。"""
    loc = block.get("locator") or {}
    if "para_index" in loc:
        return "第 %s 段" % loc["para_index"]
    if "table_index" in loc:
        return "表格%s · 行%s" % (loc.get("table_index"), loc.get("row"))
    if "cell" in loc:
        sheet = loc.get("sheet") or ""
        return "工作表「%s」%s" % (sheet, loc["cell"])
    if "line_start" in loc:
        a, b = loc.get("line_start"), loc.get("line_end")
        return ("第 %s 行" % a) if a == b else ("第 %s~%s 行" % (a, b))
    if "path" in loc:      # md 标题路径（如果读取器以后补上）
        return str(loc["path"])
    return "（定位信息缺失）"


# ==================== 二、预算选块 ====================

def pick_for_budget(blocks: list, budget: int = BUDGET):
    """把要展示的块序号挑出来。返回 (序号列表, 说明)。

    小文档：全量。大文档：开头 45% + 结尾 15% + 中间均匀抽样（剩余预算）。
    """
    texts = [(b.get("text") or "") for b in blocks]
    total = sum(len(t) for t in texts)
    if total <= budget or len(blocks) <= 3:
        return list(range(len(blocks))), ""

    head_b, tail_b = int(budget * HEAD_RATIO), int(budget * TAIL_RATIO)
    mid_b = budget - head_b - tail_b

    picked, acc, i = [], 0, 0
    while i < len(blocks) and acc < head_b:
        picked.append(i)
        acc += len(texts[i])
        i += 1
    head_end = i

    tail_pick, acc2, j = [], 0, len(blocks) - 1
    while j >= head_end and acc2 < tail_b:
        tail_pick.append(j)
        acc2 += len(texts[j])
        j -= 1
    tail_start = j + 1

    mid_ids = list(range(head_end, tail_start))
    sel = []
    if mid_ids and mid_b > 0:
        avg = max(1, sum(len(texts[k]) for k in mid_ids) // len(mid_ids))
        want = max(1, min(len(mid_ids), mid_b // avg))
        if want >= len(mid_ids):
            sel = mid_ids
        else:
            step = len(mid_ids) / want
            sel = [mid_ids[int(x * step)] for x in range(want)]

    all_picked = picked + sel + sorted(tail_pick)
    skipped = len(blocks) - len(all_picked)
    note = ("原文共 %d 块（约 %d 字），超出单次阅读预算；"
            "已按「开头 + 结尾 + 中间抽样」展示 %d 块，中间略过 %d 块。"
            % (len(blocks), total, len(all_picked), skipped))
    return all_picked, note


def build_digest(blocks: list, idx_list: list) -> str:
    """把挑出的块拼成给模型的文本（每块一行，带块内序号）。"""
    lines = []
    for n, i in enumerate(idx_list, start=1):
        t = (blocks[i].get("text") or "").strip()
        if not t:
            continue
        lines.append("[%d] %s" % (n, t))
    return "\n".join(lines)


# ==================== 三、模型读文档 → 画像 ====================

def build_profile_prompt(doc_file: str, digest: str, n_show: int, note: str) -> str:
    return f"""请阅读文档《{doc_file}》的内容，输出它的关键信息画像。

【文档内容】（共 {n_show} 块，编号仅供定位参考）
{digest}

【输出 JSON】（严格按这个结构）
{{
  "topic": "一句话说清：这份文档是什么、关于什么（30~60 字）",
  "doc_type": "合同 | 报告 | 会议纪要 | 台账/清单 | 手册/制度 | 说明 | 通知 | 其他（选最贴切的）",
  "points": [
    {{"text": "核心要点（一句话）", "quote": "支撑它的原文摘抄（20~60 字，必须逐字照抄）"}}
  ],
  "elements": [
    {{"name": "要素名", "value": "要素值", "quote": "该值所在的原文摘抄（短，5~30 字）"}}
  ]
}}
规则：
- points 给 3~7 条，按重要性排序，覆盖文档的主要脉络（是什么 / 涉及谁 / 关键内容 / 结论或要求）。
- elements 给 0~8 条，只挑最值得记住的具体信息（时间、主体、金额、编号、数量、结论等）。
- 所有 quote 必须逐字照抄上面的文档内容，不许改写、不许拼凑、不许自己总结成别的字。
- 每一条要点都必须能在文里找到出处；找不到出处的不要写。
- 文档里没有的信息绝对不要虚构。"""


def ask_profile(doc_file: str, digest: str, n_show: int, note: str):
    """调模型出画像。返回 (dict 或 None, 耗时秒, 说明)。"""
    prompt = build_profile_prompt(doc_file, digest, n_show, note)
    t0 = time.time()
    content, _usage, cost = ask_model(prompt, system=PROFILE_SYSTEM)
    if not content:
        return None, cost, "模型调用失败（网络或服务问题），请重跑一次。"
    data = parse_json_reply(content)
    if not isinstance(data, dict):
        return None, cost, "模型返回不是 JSON，无法解析。"
    return data, cost, ""


# ==================== 四、出处定位（程序核对，不让模型数编号） ====================

_PUNCT_MAP = [("，", ","), ("。", "."), ("：", ":"), ("；", ";"),
              ("（", "("), ("）", ")"), ("、", ""), ("　", "")]


def _norm(s: str) -> str:
    s = re.sub(r"\s+", "", s or "")
    for a, b in _PUNCT_MAP:
        s = s.replace(a, b)
    return s


def _locate_single(quote: str, blocks: list):
    """单块定位：①原样子串 ②去空白/标点归一后子串 ③最相似块（≥0.5）。"""
    q = (quote or "").strip()
    if not q:
        return None, "empty"
    for i, b in enumerate(blocks):
        if q in (b.get("text") or ""):
            return i, "exact"
    qn = _norm(q)
    if len(qn) >= 4:
        for i, b in enumerate(blocks):
            if qn in _norm(b.get("text") or ""):
                return i, "ws"
    best, best_r = None, 0.0
    if len(qn) >= 6:
        for i, b in enumerate(blocks):
            bt = _norm(b.get("text") or "")
            if not bt:
                continue
            r = difflib.SequenceMatcher(None, qn, bt).ratio()
            if r > best_r:
                best, best_r = i, r
    if best is not None and best_r >= 0.5:
        return best, "near(%.2f)" % best_r
    return None, "miss"


def _split_parts(quote: str) -> list:
    """把"多个片段拼起来的摘抄"拆成片段（表格类常见：整行/整列引用）。"""
    parts = [p.strip() for p in re.split(r"[\s,，、;；|/]+", quote or "")]
    return [p for p in parts if len(p) >= 2]


def _find_part(p: str, blocks: list):
    """找包含该片段的块（先原样、后归一）。找不到返回 None。"""
    for i, b in enumerate(blocks):
        if p in (b.get("text") or ""):
            return i
    pn = _norm(p)
    if len(pn) >= 3:
        for i, b in enumerate(blocks):
            if pn in _norm(b.get("text") or ""):
                return i
    return None


def _col_letter(n) -> str:
    """列号 → 字母（1→A，27→AA）。"""
    try:
        n = int(n)
    except Exception:
        return str(n)
    s = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def _span_desc(blocks: list, idxs: list):
    """跨块引用的区域描述。聚不到一起（散布全文档）时返回 None。"""
    bs = [blocks[i] for i in idxs]
    locs = [b.get("locator") or {} for b in bs]

    # xlsx：同一工作表里 → 同一行 / 同一列
    if all(("cell" in l) for l in locs):
        sheets = {l.get("sheet") for l in locs}
        if len(sheets) == 1:
            sheet = list(sheets)[0]
            rows = [l.get("row") for l in locs]
            cols = [l.get("col") for l in locs]
            if len(set(rows)) == 1 and len(set(cols)) > 1:
                cs = sorted(int(c) for c in cols)
                return ("工作表「%s」第 %s 行（%s~%s）"
                        % (sheet, rows[0], "%s%s" % (_col_letter(cs[0]), rows[0]),
                           "%s%s" % (_col_letter(cs[-1]), rows[0])))
            if len(set(cols)) == 1 and len(set(rows)) > 1:
                rs = sorted(int(r) for r in rows)
                col = cols[0]
                return ("工作表「%s」第 %s 列（%s%s~%s%s）"
                        % (sheet, _col_letter(col), _col_letter(col), rs[0],
                           _col_letter(col), rs[-1]))
        return None

    # docx 表格：同一表格里 → 同一行
    if all(("table_index" in l) for l in locs):
        tabs = {l.get("table_index") for l in locs}
        if len(tabs) == 1 and len({l.get("row") for l in locs}) == 1:
            return "表格%s · 第 %s 行" % (locs[0].get("table_index"), locs[0].get("row"))
        return None

    # 行号类（md/txt）：距离 ≤15 行算一段范围
    if all(("line_start" in l) for l in locs):
        a = min(int(l["line_start"]) for l in locs)
        b = max(int(l.get("line_end") or l["line_start"]) for l in locs)
        if b - a <= 15:
            return "第 %s 行" % a if a == b else ("第 %s~%s 行" % (a, b))
        return None

    # 段落类（docx）
    if all(("para_index" in l) for l in locs):
        a = min(int(l["para_index"]) for l in locs)
        b = max(int(l["para_index"]) for l in locs)
        if b - a <= 15:
            return "第 %s 段" % a if a == b else ("第 %s~%s 段" % (a, b))
        return None

    return None


def locate_quote(quote: str, blocks: list):
    """拿摘抄回原文定位（尽量宽厚）。

    ① 单块定位（原样 / 归一 / 近似）；
    ② 单块不中 → 拆分定位：把摘抄按分隔符拆片段，若片段命中同一行 /
       同一列 / 相邻行段等"连续区域"，给出区域描述（表格整行/整列引用常见）。
    返回 (出处描述 或 None, 命中方式)。
    """
    i, how = _locate_single(quote, blocks)
    if i is not None:
        return loc_desc(blocks[i]), how

    parts = _split_parts(quote)
    if len(parts) >= 2:
        hits, seen = [], set()
        for p in parts:
            j = _find_part(p, blocks)
            if j is not None and j not in seen:
                seen.add(j)
                hits.append(j)
        if len(hits) >= 2:
            desc = _span_desc(blocks, hits)
            if desc:
                return desc, "span(%d块)" % len(hits)
    return None, "miss"


# ==================== 五、整理输出 ====================

def _clean_points(raw) -> list:
    out = []
    if isinstance(raw, list):
        for p in raw:
            if isinstance(p, dict) and str(p.get("text") or "").strip():
                out.append({"text": str(p["text"]).strip(),
                            "quote": str(p.get("quote") or "").strip()})
            elif isinstance(p, str) and p.strip():
                out.append({"text": p.strip(), "quote": ""})
    return out[:7]


def _clean_elements(raw) -> list:
    out = []
    if isinstance(raw, list):
        for e in raw:
            if isinstance(e, dict) and str(e.get("value") or "").strip():
                out.append({"name": str(e.get("name") or "").strip() or "要素",
                            "value": str(e["value"]).strip(),
                            "quote": str(e.get("quote") or "").strip()})
    return out[:8]


def format_screen(doc_file: str, doc_type: str, blocks: list, data: dict,
                  cmd: str, total: float, note: str) -> str:
    """屏幕报告。"""
    L = []
    L.append("=" * 62)
    L.append("【画像】一句话 → 文档关键信息整理出来（模块2 · 第2步）")
    L.append("=" * 62)
    L.append("指令：%s" % cmd)
    L.append("文档：%s（%s ｜ %d 块）" % (doc_file, doc_type or "?", len(blocks)))
    if note:
        L.append("说明：%s" % note)
    L.append("")
    L.append("◆ 主题：%s" % (data.get("topic") or "（未给出）"))
    L.append("◆ 类型：%s" % (data.get("doc_type") or "（未给出）"))
    L.append("")

    pts = data["_points"]
    hit = sum(1 for p in pts if p["_loc_text"])
    L.append("◆ 核心要点（%d 条 ｜ 出处定位 %d/%d）" % (len(pts), hit, len(pts)))
    for i, p in enumerate(pts, start=1):
        L.append("  %d. %s" % (i, p["text"]))
        if p["_loc_text"]:
            L.append("     出处：%s" % p["_loc_text"])
        else:
            L.append("     出处：未定位（摘抄与原文有出入）")
    L.append("")

    els = data["_elements"]
    if els:
        hit2 = sum(1 for e in els if e["_loc_text"])
        L.append("◆ 关键要素（%d 条 ｜ 出处定位 %d/%d）" % (len(els), hit2, len(els)))
        for e in els:
            L.append("  · %s：%s（%s）" % (e["name"], e["value"], e["_loc_text"] or "未定位"))
        L.append("")
    L.append("用时 %.1f 秒" % total)
    return "\n".join(L)


def format_md(doc_file: str, blocks: list, data: dict, cmd: str,
              total: float, tstamp: str, note: str) -> str:
    """Markdown 版画像（给文件）。"""
    L = []
    L.append("# 关键信息画像")
    L.append("")
    L.append("- **文档**：%s" % doc_file)
    L.append("- **指令**：%s" % cmd)
    L.append("- **时间**：%s ｜ 用时 %.1f 秒" % (tstamp, total))
    if note:
        L.append("- **说明**：%s" % note)
    L.append("")
    L.append("> **主题**：%s" % (data.get("topic") or "（未给出）"))
    L.append(">")
    L.append("> **类型**：%s" % (data.get("doc_type") or "（未给出）"))
    L.append("")

    L.append("## 核心要点")
    L.append("")
    for i, p in enumerate(data["_points"], start=1):
        L.append("### %d. %s" % (i, p["text"]))
        L.append("")
        if p["_loc_text"]:
            L.append("- 出处：**%s**" % p["_loc_text"])
        else:
            L.append("- 出处：_未定位（摘抄与原文有出入）_")
        if p["quote"]:
            L.append("- 摘抄：「%s」" % p["quote"].replace("|", "\\|"))
        L.append("")

    els = data["_elements"]
    if els:
        L.append("## 关键要素")
        L.append("")
        L.append("| 要素 | 值 | 出处 |")
        L.append("|---|---|---|")
        for e in els:
            L.append("| %s | %s | %s |" % (e["name"],
                                           e["value"].replace("|", "\\|"),
                                           e["_loc_text"] or "_未定位_"))
        L.append("")
    return "\n".join(L)


# ==================== 六、主流程 ====================

def run_profile(text: str, intent: dict, src: Path, out_dir: Path = None):
    """执行一条「文档画像」指令。返回 (是否成功, 报告文本, 产出文件列表)。

    与 extract_ops.run_extract 同签名——edit.py 按 op 分流即可。
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

    idx_list, note = pick_for_budget(blocks)
    digest = build_digest(blocks, idx_list)
    if not digest.strip():
        return False, "【失败】《%s》的内容块都是空的，无法生成画像。" % src.name, []

    data, cost, err2 = ask_profile(src.name, digest, len(idx_list), note)
    if data is None:
        return False, "【失败】%s（用时 %.1f 秒）" % (err2, cost), []

    data["_points"] = _clean_points(data.get("points"))
    data["_elements"] = _clean_elements(data.get("elements"))
    if not data["_points"]:
        return False, ("【失败】模型没有给出核心要点（可能文档内容为空或调用异常），"
                       "请重跑一次。"), []

    # 出处定位（程序核对）
    for p in data["_points"]:
        p["_loc_text"], p["_how"] = locate_quote(p.get("quote", ""), blocks)
    for e in data["_elements"]:
        e["_loc_text"], e["_how"] = locate_quote(e.get("quote", ""), blocks)

    total = time.time() - t0
    tstamp = time.strftime("%Y-%m-%d %H:%M")

    screen = format_screen(src.name, doc.get("type", ""), blocks, data,
                           text, total, note)
    md = format_md(src.name, blocks, data, text, total, tstamp, note)

    md_path = out_dir / (src.stem + "_画像.md")
    json_path = out_dir / (src.stem + "_画像.json")
    md_path.write_text(md, encoding="utf-8")
    json_path.write_text(json.dumps({
        "cmd": text, "doc": src.name, "time": tstamp,
        "elapsed_sec": round(total, 2),
        "topic": data.get("topic", ""), "doc_type": data.get("doc_type", ""),
        "points": [{"text": p["text"], "quote": p["quote"],
                    "loc": p["_loc_text"] or "", "loc_ok": bool(p["_loc_text"])}
                   for p in data["_points"]],
        "elements": [{"name": e["name"], "value": e["value"], "quote": e["quote"],
                      "loc": e["_loc_text"] or "", "loc_ok": bool(e["_loc_text"])}
                     for e in data["_elements"]],
        "note": note,
    }, ensure_ascii=False, indent=2), encoding="utf-8")

    screen += "\n【输出】%s\n【输出】%s\n【原件】未改动（提取是只读操作）" % (md_path, json_path)
    return True, screen, [str(md_path), str(json_path)]


def main():
    """命令行直跑：解析 + 画像一条指令（不经 edit.py 时的入口）。"""
    import argparse
    from parse_intent import scan_docs, parse_one

    ap = argparse.ArgumentParser(description="模块（2）· 文档画像执行器")
    ap.add_argument("text", help="一条关键信息提取指令")
    ap.add_argument("--docs", default=str(BASE / "测试数据" / "源文档"), help="文档库目录")
    args = ap.parse_args()

    docs_dir = Path(args.docs)
    doc_names = scan_docs(docs_dir)
    intent, cost, ok = parse_one(args.text, doc_names)
    if not ok:
        print("解析失败（网络或格式问题），请重跑一次。")
        sys.exit(1)
    print("解析（%.1fs）：%s" % (cost, json.dumps(intent, ensure_ascii=False)))
    if intent.get("op") != "profile":
        print("这不是一条「文档画像」指令（op=%s）。" % intent.get("op"))
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

    ok, report, files = run_profile(args.text, intent, src)
    print(report)
    sys.exit(0 if ok else 1)


if __name__ == "__main__":
    main()
