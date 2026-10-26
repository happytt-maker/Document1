# -*- coding: utf-8 -*-
"""
extract.py —— 第 4 步 · 4b：抽取器（先筛，后判）

任务：给一个列名（比如"责任判定"），输出这一列的「值 + 出处（行号）」。

流程三段：
  ① 筛（不花钱）—— 用 retrieve.py 的打分函数，从 701 个内容块里挑候选块
  ② 判（花钱）  —— 把候选块 + 列名一起交给大模型，让它抽「值 + 一句原文摘抄」
  ③ 定（不花钱）—— 程序拿"原文摘抄"回原文里找行号

为什么不让模型直接报行号？
  模型不擅长"数行号"（容易数错），但擅长"挑出正确的句子"。
  让每个角色干它擅长的事 —— 模型负责判，程序负责算。

跑法（VSCode 打开 matcher 文件夹）：
  方式一（推荐）：跑 run_all.py —— 全部列一次跑完（并行），不用动这里
  方式二：改下面配置区的 COL_NAME，按 F5（只想跑一列时用）
  方式三：终端里  python extract.py 责任判定

输出：
  屏幕：逐条打印「值 ｜ 出处」
  文件：抽取结果-<列名>.json（给 evaluate.py 评测用）
"""
import json
import os
import re
import sys
import threading
import time
from pathlib import Path

import requests

from retrieve import build_column_contexts, locate_text, score_blocks, strip_unit

# 输出兜底：遇到终端表示不了的字符就替换，不让程序崩
try:
    sys.stdout.reconfigure(errors="replace")
except Exception:
    pass

# ================== 配置区（只改这里） ==================
# 2026-10-25：换成「阿里云百炼 MaaS」的 OpenAI 兼容接口。
# 密钥 / 地址 / 模型名统一在 agent/llm_config.py 里配（只填那一处），这里只引用。
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "agent"))
from llm_config import API_KEY, API_URL, MODEL, chat as _llm_chat  # noqa: E402

COL_NAME = "责任判定"   # 想抽哪一列，改这里
TOP_K = 15              # 最多挑几块给模型看（2026-10-11 由 8 提到 15：
                        #   整列格子数常到 10~12 个（表头+数据行），原上限 8 会把
                        #   整列截掉尾巴 → 模型只能抽到前几条，"列条数不齐"的元凶之一）
MIN_SCORE = 20          # 筛选阈值（和 retrieve.py 保持一致）
SPLIT_TOP_K = 36        # 拆词兜底的候选上限（要能装下 2~3 个整列；实测「库存数量」
                        #   重审时盘点表的 物料编码/期初库存/期末库存 三列共 33 个格子）
# ========================================================

BASE = Path(__file__).resolve().parents[1]           # = 研一项目/
# 读取器的产物。默认 doc_reader/output.json；
# 环境变量 A23_OUTPUT_JSON 可指定另一份池文件（压测用的"大池"与日常小池隔离，
# 互不污染；不设变量时行为完全不变）——2026-10-14 压测前加。
OUTPUT_JSON = Path(os.environ.get("A23_OUTPUT_JSON") or (BASE / "doc_reader" / "output.json"))

# 池缓存（2026-10-14 压测后优化）：
#   大池 194MB / 50 万块，json.loads + 展平 + 列上下文预建 每列都要几秒~十几秒；
#   8 列并行 + 一致性重审会把同一份池重复解析十几次，纯浪费。
#   按 (路径, mtime, 大小) 缓存；解析后立即预建列上下文（_ctx / _colheads_norm），
#   让之后所有全池/子集打分都免建。run_all 的同进程多线程共享这一份。
_POOL_CACHE = {"key": None, "blocks": None}
_POOL_LOCK = threading.Lock()


def load_pool():
    """读出全部内容块（带缓存与上下文预建）。"""
    st = OUTPUT_JSON.stat()
    key = (str(OUTPUT_JSON), st.st_mtime_ns, st.st_size)
    with _POOL_LOCK:
        if _POOL_CACHE["key"] != key:
            data = json.loads(OUTPUT_JSON.read_text(encoding="utf-8"))
            blocks = [b for d in data.get("documents", []) for b in d.get("blocks", [])]
            build_column_contexts(blocks)          # 预热：挂 _ctx / _colheads_norm
            _POOL_CACHE["key"] = key
            _POOL_CACHE["blocks"] = blocks
        return _POOL_CACHE["blocks"]

SYSTEM_PROMPT = "你是一个严谨的信息抽取助手，只依据给定的原文抽取信息，绝不编造。只输出 JSON。"


def build_user_prompt(col_name: str, blocks: list, split_subs: list = None) -> str:
    """把候选块和任务拼成给模型的一段话。

    split_subs：本列走了拆词兜底时传入子词（如 ["库存","数量"]）。
    2026-10-11 实测：兜底场景（目标是"库存数量"，候选是"期初库存/期末库存"这类
    近似列头）模型会"纠结到放弃"，同题两次跑出 0 条 / 10 条。所以专门加一条
    规则：近似也算，选最接近的整列，别输出空。
    """
    core = strip_unit(col_name)

    parts = []
    parts.append("下面是从文档中检索到的 %d 个内容块（已编号，每个带『相关度』分数，分数越高越可能含要找的值）：\n" % len(blocks))
    for i, b in enumerate(blocks, start=1):
        parts.append("[块%d] 相关度 %d ｜ %s ｜ %s" % (i, b.get("_score", 0), b.get("file"), locate_text(b)))
        # 表格格子：把"这格所在列的列头"一起展示 —— 否则模型看到裸值（如"VMC-850B"）
        # 不知道它属于哪一列，只能去抄别处带标签的孤条（2026-10-11 规格型号案例实测）
        head = (b.get("_head") or "").strip()
        if head and head != (b.get("text") or "").strip():
            parts.append("（该格所在列的列头：%s）" % head)
        parts.append(b.get("text", ""))
        parts.append("")

    # 规则逐条攒起来（有特殊情况就多加一条，编号自动顺延）
    rules = [
        "每个值必须来自上面的原文；evidence 字段摘抄包含该值的原文原句（一个字都不要改）",
        "原文里没有的，不要编造；找不到的记录就不输出",
        "以「=」开头的单元格文本是还没求值的公式（如 =E3+F3-G3），不是可用的值 —— 不要抽取它们",
        "有几条记录就输出几条，按在原文中出现的顺序排",
        "优先级（同一列有多处来源时，按完整性选）："
        "①表格的整列数据（同一份文档、同一列连续多行）＝完整记录集，整列取全；"
        "②多条同样格式的记录（如多条「客户名称：…」）也＝完整记录集，全部取；"
        "③单独出现一次的零散值信息不全，不能替代①②；"
        "④①②同时存在时取相关度更高的一组；"
        "⑤多组相关度并列（分数相同）时不要反复比较：取候选里格子数更多、更连续的一组；"
        "格子数也一样时取先出现的一组，选定后立即整列照抄；"
        "⑥候选里能看到几格就照抄几格（候选就是全部可用证据），"
        "不要因为一组看起来「没列全」而犹豫、改选或混合别的组；"
        "同一列的所有记录必须取自同一份文档，不要跨文档混着来",
    ]
    if core != col_name:
        unit_hint = col_name.replace(core, "").strip("（）() ")
        if unit_hint:
            rules.append("列名里的「%s」是单位说明：value 只填数字部分（比如 2），不要带单位" % unit_hint)
    if split_subs:
        rules.append("本列在原文里可能以近似说法出现（检索时是按「%s」「%s」两个字分别召回候选的）。"
                     "若候选里有一组数据（同一表格、同一列连续多行）与目标列名部分匹配，就取这组、整列取全；"
                     "多组部分匹配时，取与目标列名共同字更多、或语义更接近的一组；"
                     "不要因为「不完全等于」就输出空" % (split_subs[0], split_subs[1]))
    if any(k in col_name for k in ("摘要", "概要", "简述", "概述", "说明", "描述",
                                   "内容", "事项", "情况", "要点")):
        rules.append("这一列是「摘要/提炼」类：value 由原文直接删减得到 —— 只删，不写新词、不做同义替换；"
                     "删法：看句首，如果是时间或来源的词（如“2026年1月8日交付的”“从供应商处转来的”），"
                     "把它删掉；句首不是这类词就整句照抄、别动；删完这一处就停；"
                     "其余部分原样保留，数字和数量词（如“其中一台”“第 6 天”“2 至 3 次”）一个都不能删")
    rules.append('输出格式（JSON）：{"items": [{"value": "值", "evidence": "原文原句", "block": 块编号}, ...]}')
    rules.append("只输出这个 JSON，不要任何解释文字")

    parts.append("=" * 40)
    parts.append("任务：抽取【%s】这一列的值。" % col_name)
    parts.append("")
    parts.append("规则（照做，不要发挥）：")
    for i, r in enumerate(rules, start=1):
        parts.append("%d. %s" % (i, r))
    return "\n".join(parts)


def ask_model(user_prompt: str):
    """调用大模型一次，返回 (回答文本, 用量, 耗时秒数)。"""
    # 2026-10-11 试验过 temperature=0：反而更糟 —— 模板5「库存数量」重审时
    # 模型"死磕"了 79 秒 / 16k reasoning tokens 后输出空（默认温度时 4.4 秒出 10 条）。
    # 推理型模型固定 0 会退化，保持默认温度（即不传 temperature）。
    content, usage, cost, err = _llm_chat(
        [{"role": "system", "content": SYSTEM_PROMPT},
         {"role": "user", "content": user_prompt}],
        timeout=90)
    if err:
        print("请求失败！", err)
        return None, None, cost
    return content, usage, cost


def parse_json_reply(text: str):
    """宽容地解析模型的回复：允许 ```json 围栏、前后多余文字。"""
    t = (text or "").strip()
    if t.startswith("```"):
        t = re.sub(r"^```[A-Za-z]*\s*", "", t)
        t = re.sub(r"```\s*$", "", t)
    i, j = t.find("{"), t.rfind("}")
    if i == -1 or j == -1:
        return None
    try:
        return json.loads(t[i:j + 1])
    except Exception:
        return None


PUNCT = r"[\s，。、；：！？,.;:!?（）()【】\[\]「」『』\"'']+"


def norm_text(s: str) -> str:
    """归一化：去掉空白与标点（定位用，不影响展示）。"""
    return re.sub(PUNCT, "", s or "")


def locate_evidence(evidence: str, block: dict):
    """在块里定位 evidence。返回 (行号, 位置说明) —— 都找不到时是 (None, None)。

    位置体系按块的类型自动切换（这是"跨文档"的关键）：
      txt / md     → 行号（locator.line_start + 偏移）   ← 行号可给评测用
      docx 段落    → 段号（locator.para_index + 偏移）
      docx 表格格  → 格子坐标（如"表格1 · 行2·列2"）      ← 格子本身即出处
      xlsx 单元格  → 工作表 + 单元格坐标（如"工作表「员工信息」 B3"）

    分两轮，从严到宽：
      第一轮：整句互相包含（正常情况一句话总能对上）
      第二轮：拿前 10 个字当锚点（模型偶尔会多抄/少抄几个字）
    """
    ev = norm_text(evidence)
    if not ev:
        return None, None

    loc = block.get("locator") or {}
    lines = (block.get("text") or "").split("\n")

    # —— 连续文本块（txt / md / docx 段落）：逐行找 ——
    base = loc.get("line_start") or loc.get("para_index")
    if base:
        is_line = "line_start" in loc          # True=行号体系，False=段落体系
        for k, line in enumerate(lines):
            nl = norm_text(line)
            if nl and (ev in nl or nl in ev):
                pos = base + k
                return (pos, "第 %d 行" % pos) if is_line else (None, "第 %d 段" % pos)
        anchor = ev[:10]
        for k, line in enumerate(lines):
            if anchor and anchor in norm_text(line):
                pos = base + k
                return (pos, "第 %d 行" % pos) if is_line else (None, "第 %d 段" % pos)
        return None, None

    # —— 单元格块（docx 表格 / xlsx）：格子本身即出处 ——
    if loc.get("kind") in ("table_cell", "cell"):
        nl = norm_text(block.get("text") or "")
        if nl and (ev in nl or nl in ev or ev[:10] in nl):
            return None, locate_text(block)

    return None, None


def run_column(col_name: str, quiet: bool = False, only_file: str = None,
               allow_split: bool = True, out_dir: str = None):
    """完整跑一列：筛 → 判 → 定。返回结果 dict（失败返回 None）。

    quiet=True 时不在屏幕唠叨（给 run_all.py 并行调用用，避免多列打印打架）。
    only_file="某文件名" 时只在那一份文档里检索（给一致性重审用，见 run_all.py）。
    allow_split=False 时不做拆词兜底（2026-10-11 全量实测：拆词兜底在"全文档池"
      里干扰太大，模型会纠结几十秒还只抽一半；改由一致性重审阶段在"主源"内做，
      环境干净、快且全。单列调试时保持默认 True）。
    out_dir=结果文件写到哪个目录（默认写在本目录 matcher/；模块（1）的内容提取
      接入时传自己的输出目录，避免污染本目录）。
    """
    core = strip_unit(col_name)
    say = (lambda *a, **k: None) if quiet else print
    safe_col = re.sub(r'[\\/:*?"<>|]', "_", col_name)     # 文件名字符兜底
    out_base = Path(out_dir) if out_dir else Path(__file__).resolve().parent
    out_base.mkdir(parents=True, exist_ok=True)

    say("=" * 60)
    say("抽取器 · 列名：%s" % col_name)
    say("=" * 60)

    if not OUTPUT_JSON.exists():
        say("找不到 output.json：%s" % OUTPUT_JSON)
        say("先去 doc_reader 里跑一次 run_all.py，把它生成出来。")
        return None

    # ---------- ① 筛 ----------
    all_blocks = load_pool()
    pool = all_blocks
    contexts = None
    if only_file:
        pool = [b for b in all_blocks if b.get("file") == only_file]
        contexts = [b.get("_ctx") or "" for b in pool]   # 预建缓存，直接复用
        say("① 筛：只在《%s》内检索（一致性重审）" % only_file)
    total = len(pool)
    scored = score_blocks(pool, core, min_score=MIN_SCORE, contexts=contexts)
    split_used = False
    subs = []
    if allow_split and len(scored) < 2 and len(core) >= 4:
        # —— 拆词兜底（2026-10-11 B1 修复）——
        # 场景：列名「库存数量」对源表头「期初库存 / 期末库存」——字面一个候选都进不来。
        # 做法：把列名拆成前 2 字、后 2 字，分别召回再合并（同一块取最高分）。
        # 只在前面的正常召回"几乎为空"（<2 块）时触发，正常列不受影响。
        # 注意："是否紧缺"这类源文档里真没有的列，拆词后依然没候选 —— 保持为空，这是对的。
        for w in (core[:2], core[-2:]):
            if w and w != core and w not in subs:
                subs.append(w)
        merged = {id(b): (s, b) for s, b in scored}     # 先装正常候选
        for sub in subs:
            for s, b in score_blocks(pool, sub, min_score=MIN_SCORE, contexts=contexts):
                k = id(b)
                if k not in merged or s > merged[k][0]:
                    merged[k] = (s, b)
        scored = sorted(merged.values(), key=lambda t: -t[0])
        if scored:
            split_used = True
            say("① 筛：正常召回不足 → 拆词兜底（%s）→ %d 块" % ("＋".join(subs), len(scored)))

    limit = SPLIT_TOP_K if split_used else TOP_K
    cands = [dict(b, _score=s) for s, b in scored[:limit]]

    say("① 筛：从 %d 个块中挑出 %d 个候选块" % (total, len(cands)))
    for i, b in enumerate(cands, start=1):
        say("   [块%d] %d分 ｜ %s ｜ %s" % (i, b.get("_score", 0), b.get("file"), locate_text(b)))

    if not cands:
        # "一个候选都筛不出来" ≠ "程序坏了" —— 多数是源文档里就没有这一列的信息
        # （如模板5「是否紧缺」）。返回一个空结果（items=[]）而不是 None：
        #   ① 下游能区分"源里没有"（诚实留空）和"真失败"（调用错误）；
        #   ② 一致性重审时，"主源里也没有" 才能被采纳为空
        #      （2026-10-11 实测教训：备注列重审返回 None 被当成"失败"，
        #       结果把库存盘点表的错误数据保了回来）。
        say("没筛出候选块 —— 源文档里大概率没有这一列的信息（诚实留空）。")
        empty = {
            "col": col_name,
            "time": time.strftime("%Y-%m-%d %H:%M"),
            "elapsed_sec": 0,
            "usage": None,
            "candidate_count": 0,
            "items": [],
            "no_candidate": True,
        }
        if only_file:
            empty["only_file"] = only_file
        out_path = out_base / ("抽取结果-%s.json" % safe_col)
        out_path.write_text(json.dumps(empty, ensure_ascii=False, indent=2), encoding="utf-8")
        say("（结果已存：%s）" % out_path.name)
        return empty

    # ---------- ② 判 ----------
    prompt = build_user_prompt(col_name, cands, split_subs=(subs if split_used else None))
    say("")
    say("② 判：调用大模型（%s）……" % MODEL)
    content, usage, cost = ask_model(prompt)
    if content is None:
        return None

    parsed = parse_json_reply(content)
    if not parsed or "items" not in parsed:
        say("模型输出不是合法 JSON，原样打印：")
        say(content)
        return None

    raw_items = parsed["items"]

    # ---------- ③ 定 ----------
    items = []
    for it in raw_items:
        value = str(it.get("value", "")).strip()
        evidence = str(it.get("evidence", "")).strip()
        bno = it.get("block")

        line = None
        loc_text = None
        src_block = None
        if isinstance(bno, int) and 1 <= bno <= len(cands):
            src_block = cands[bno - 1]
            line, loc_text = locate_evidence(evidence, src_block)
        if line is None and loc_text is None:
            for b in cands:                      # 兜底：在所有候选块里找
                line, loc_text = locate_evidence(evidence, b)
                if line is not None or loc_text is not None:
                    src_block = b
                    break

        items.append({
            "value": value,
            "evidence": evidence,
            "block": bno,
            "file": (src_block or {}).get("file"),
            "line": line,
            "loc_text": loc_text,
        })

    # ---------- 打印 ----------
    say("")
    say("③ 定：拿摘抄回原文定位出处")
    say("-" * 60)
    if not items:
        say("模型没有抽到任何值。")
    for i, it in enumerate(items, start=1):
        loc = it.get("loc_text") or "未定位"
        say("[%d] 值 = %s" % (i, it["value"]))
        say("    出处 = %s ｜ %s" % (it.get("file") or "?", loc))
        say("    原文 = %s" % it["evidence"])
        say("")

    say("-" * 60)
    say("本次耗时：%.2f 秒（赛题：每份文档 ≤ 90 秒 —— 7 列相加它）" % cost)
    say("本次用量：%s" % usage)
    say("")
    say("下一步：跑另外几列，然后运行 evaluate.py 对标准答案算准确率。")

    # ---------- 存文件 ----------
    result = {
        "col": col_name,
        "time": time.strftime("%Y-%m-%d %H:%M"),
        "elapsed_sec": round(cost, 2),
        "usage": usage,
        "candidate_count": len(cands),
        "items": items,
    }
    if only_file:
        result["only_file"] = only_file        # 一致性重审：本列被限定在这份文档里抽
    if split_used:
        result["split_fallback"] = True         # 本列走了拆词兜底召回
    out_path = out_base / ("抽取结果-%s.json" % safe_col)
    out_path.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    say("（结果已存：%s）" % out_path.name)
    return result


def main():
    col_name = sys.argv[1] if len(sys.argv) > 1 else COL_NAME
    run_column(col_name)


if __name__ == "__main__":
    main()
