# -*- coding: utf-8 -*-
"""
retrieve.py —— 第 4 步 · 4a：筛选器（先筛，后判）

作用：给一个列名（比如"投诉编号"），从 output.json 的 701 个内容块里，
      挑出"最可能含这个字段"的候选块 —— 只筛选、不判断、不花钱。

跑法（VSCode 打开 matcher 文件夹）：
    方式一：改下面配置区的 COL_NAME，按 F5
    方式二：终端里  python retrieve.py 责任判定

输出：按分数排序的候选块清单（分数 / 文件 / 位置 / 文本片段）。
看一眼、回答两个问题：
    1. 该进来的块都进来了吗？（拿标准答案里的「来源」行号对一对）
    2. 有没有不相关的块混进来？（它们会干扰后面模型判断）
"""
import datetime
import json
import os
import re
import sys
import threading
from pathlib import Path

# ================== 配置区（只改这里） ==================
COL_NAME = "投诉编号"   # 想筛哪一列，改这里
TOP_N = 15              # 显示前几名
MIN_SCORE = 20          # 低于这个分数不显示（滤掉噪音）
# ========================================================

BASE = Path(__file__).resolve().parents[1]            # = 研一项目/
# 读取器的产物。默认 doc_reader/output.json；
# 环境变量 A23_OUTPUT_JSON 可指定另一份池（2026-11-07：表格填写只勾了几份源文档时，
# 流水线会把共享池裁成"这一单专用"的小池并通过它传进来 —— 这里必须跟着走，
# 否则语义召回还在全库里捞，等于范围没限住）。
OUTPUT_JSON = Path(os.environ.get("A23_OUTPUT_JSON") or (BASE / "doc_reader" / "output.json"))


def strip_unit(name: str) -> str:
    """去掉列名里的括号说明：'处理时长（天）' -> '处理时长'。"""
    return re.sub(r"[（(].*?[)）]", "", name).strip()


def bigrams(text: str):
    """切成连续两字片段：'投诉编号' -> ['投诉','诉编','编号']。"""
    return [text[i:i + 2] for i in range(len(text) - 1)]


def score_block(text: str, core: str) -> int:
    """给一个内容块打分：它多大程度上像'含这个字段'的块。

    三层信号，从强到弱：
      A 标签命中：文本里出现"某前缀："（如 '投诉内容：' 对 '投诉内容摘要' 也算）
      B 完整命中：文本里出现整个列名（不带冒号也算，比如表头行、说明句）
      C 局部命中：列名的两字片段命中比例（兜底，覆盖不完全一致的说法）
    """
    score = 0

    # A) 标签命中 —— 最强信号：列名的任一中长前缀后面紧跟冒号
    prefixes = [core[:i] for i in range(3, len(core) + 1)]
    if any((p + "：" in text) or (p + ":" in text) for p in prefixes):
        score += 90

    # B) 完整命中：整个列名出现在文本里
    if core in text:
        score += 40

    # C) 局部命中：两字片段命中比例
    parts = bigrams(core)
    if parts:
        hit = sum(1 for p in parts if p in text)
        score += int(40 * hit / len(parts))

    return score


def build_column_contexts(blocks: list) -> list:
    """预计算每个块的"同列上下文"：同列最上面 3 个非空文本（去掉它自己）。

    为什么需要它：
      表格里"列名在表头、值在数据格里"，而值本身往往不含列名文字 ——
      比如"任务内容"这一列，值是"完成数据接入模块开发"，一个字都不沾"任务内容"。
      光拿列名去匹配，数据格一个都召不回来（2026-10-11 实测过）。
      把同列最上面的文本（通常含表头）拼进搜索文本，字面命中才能生效。

    支持两种表格块：
      docx 表格单元格：locator 里有 table_index / row / col
      xlsx 单元格：    locator 里有 sheet / row / col

    2026-10-14 压测后优化：
      ① 结果按"同一份 blocks"缓存 —— 大池 50 万块构建一次要好几秒，
         多列（8 列）与一致性重审会反复调用同一份池，之前每次都全量重建。
      ② 构建时给所有块统一挂 _ctx（含非表格块，"值为空串"），
         并预挂 _colheads_norm（列头去括号版）—— 打分时不再跑正则、不再现场拼。
      ③ only_file 子集（同一批块对象）打分时，可直接从块上取 _ctx 复用。
    """
    key = id(blocks)
    with _CTX_LOCK:
        hit = _CTX_CACHE.get(key)
        if hit is not None and hit[0] is blocks:
            return hit[1]

    cols = {}   # key -> [(row, 块下标, 文本), ...]
    for i, b in enumerate(blocks):
        loc = b.get("locator") or {}
        text = (b.get("text") or "").strip()
        if not text:
            continue
        kind = loc.get("kind")
        if kind == "table_cell":
            key_c = ("d", b.get("file"), loc.get("table_index"), loc.get("col"))
        elif kind == "cell":
            key_c = ("x", b.get("file"), loc.get("sheet"), loc.get("col"))
        else:
            continue        # 段落 / 行文本块本身就是连续文本，不需要增强
        cols.setdefault(key_c, []).append((loc.get("row") or 0, i, text))

    contexts = [""] * len(blocks)
    for items in cols.values():
        items.sort()
        top = items[:3]                        # 该列最上面 3 个非空
        head = top[0][2] if top else ""        # 最上面那个，一般就是表头
        heads5 = [t for (_, _, t) in items[:5]]  # 前 5 个非空（列头匹配检查用，见 score_blocks）
        for (_, j, _t) in items:
            contexts[j] = " ".join(t for (_, jj, t) in top if jj != j)
            # 把上下文挂到块上（打分、展示共用）：
            #   _ctx  = 同列最上面 3 个非空拼起来（打分用）
            #   _head = 同列最上面那一个文字，通常是表头（展示给模型看用 —— 2026-10-11 补：
            #           规格型号案例实测，模型看裸值"VMC-850B"不知道属于哪列，只敢抄
            #           "规格型号：VMC-850B"那条标签；把列头给它看，它才能整列取）
            #   _colheads = 同列前 5 个非空文本（2026-10-11 A 类修复：
            #           列头不一定在最顶上——员工信息表 A 列实测，最上面是标题行
            #           "某某科技有限公司员工信息表"，表头"工号"在第 2 个。
            #           所以匹配时要在前几个里"找得到就行"，不能只认最顶上那个）
            blocks[j]["_ctx"] = contexts[j]
            blocks[j]["_head"] = head
            blocks[j]["_colheads"] = heads5

    # 统一挂 _ctx（非表格块 = ""）＋ 预归一化列头（打分热路径不做正则）
    for j, b in enumerate(blocks):
        b["_ctx"] = contexts[j]
        ch = b.get("_colheads")
        if ch:
            b["_colheads_norm"] = [strip_unit(h) for h in ch]

    with _CTX_LOCK:
        if len(_CTX_CACHE) >= _CTX_CACHE_KEEP:
            _CTX_CACHE.pop(next(iter(_CTX_CACHE)))      # 简单 FIFO 淘汰
        _CTX_CACHE[key] = (blocks, contexts)
    return contexts


def colhead_hits(b: dict, core: str) -> bool:
    """该块所在列的"列头"里，有没有一个与目标列名完全相等。

    为什么需要（2026-10-11 A 类修复之一——列头精确匹配加权）：
      目标列「规格」：盘点表列头「规格」（完全相等） vs 合同列头「规格型号」（不相等）
      目标列「工号」：员工信息表列头「工号」（相等） vs 别的表就算同名也各凭证据
      完全相等是比"字面沾边"更强的信号，给它加分，让正确的列排到候选前面。

    2026-10-14 优化：优先用构建时预归一化的 _colheads_norm（打分热路径免正则）；
    老数据（未预建）时回退到现场 strip_unit —— 行为不变。
    """
    norms = b.get("_colheads_norm")
    if norms is None:
        norms = [strip_unit(h) for h in (b.get("_colheads") or [])]
    return core in norms


HEAD_EXACT_BONUS = 60      # 列头与目标列名完全相等时的加分

# ===== 键值段落兜底（2026-11-02 新增）=====
# 语义召回的盲区：它只拿"表格列头"算向量（_colheads 只给表格块挂），
# 而大量字段标签其实写在**段落**里 —— 合同「乙方（供货方）：泰隆机电设备有限公司」、
# 台账「投诉编号：TS-2026-011」…… 这些段落没有列头，语义分永远是 0，
# 于是一整类"换说法"的列都召不回来（模板1「供应商」实测：0 候选，模型只能交白卷）。
#
# 兜底做法：把源里形如「标签：值」的段落整条挑出来，交给模型判断。
# 为什么不自己判"是不是目标字段"：标签叫法可以千变万化（乙方/供货方/供方/厂商…），
# 而"是不是同一个字段"正是大模型最擅长、也最便宜的一步判断。
# 触发范围刻意收窄：只给 anchor 判为 weak（源内字面 <20，字面完全对不上）的列用，
# 正常列一根手指都不碰 —— 70 列零退化靠这条守住。
_KV_RE = re.compile(r"^\s*([^：:\n]{2,16})\s*[：:]\s*(\S.*)$")


def paragraph_kv_blocks(blocks, limit=8):
    """挑出「标签：值」形式的段落块。返回 [(标签, block), ...]。

    过滤掉大标题/纯数字标签（复用 _HEAD_NOISE / _HEAD_NUM 的清洗规则），
    以及值太长（>60 字，多半是正文句子而不是字段值）的段落。
    """
    out = []
    for b in blocks:
        loc = b.get("locator") or {}
        if loc.get("kind") != "paragraph":
            continue
        t = (b.get("text") or "").strip()
        m = _KV_RE.match(t)
        if not m:
            continue
        label, val = m.group(1).strip(), m.group(2).strip()
        if len(val) > 60:
            continue
        if _HEAD_NOISE.search(label) or _HEAD_NUM.match(label):
            continue
        out.append((label, b))
        if len(out) >= limit:
            break
    return out


# ===== 语义召回（2026-10-26 新增）=====
# 字面打分的死穴：「合同总价」和源表头「合同总金额」只有「价/金额」沾边，
# 两字片段命中率 1/4，分数 10 分 —— 连 MIN_SCORE=20 都过不去，整列召不回来，
# 只能靠人往别名表里补一条映射。换行业（医疗/物流/教育）又得重配。
#
# 这里加一路语义分：拿目标列名和「每个块的列头」算向量相似度。
#
# 【为什么按排名给分，而不是按绝对相似度】
#   短字段名（4 个字）的余弦分布很挤：「员工手机号」对「工号」0.73、
#   对「联系电话」0.61 —— 绝对阈值切不出对错，但**排序**是对的（该进前几名的都在）。
#   所以取语义 Top-8 的列头，按名次给 50/45/40…15 的阶梯分。
#   实测（真实 42 个列头 × 30 组字段名）：Top3 命中 28/30，按名次给分
#   足以把正确的整列捞进候选，又不至于把无关列顶上来。
SEM_TOPK = 8               # 语义排名前几名给分
SEM_TOP_SCORE = 50         # 第 1 名给多少分
SEM_RANK_STEP = 5          # 每退一名少多少分（第 8 名 = 15 分）
SEM_HINT_MIN = 0.65        # 只有这么像的才配"告诉模型原文里可能叫这个"
                           # 定在 0.65 而不是更高：pipeline 里的真实换说法
                           # （合同总价→合同总金额 0.88、月薪→基本工资 0.66）
                           # 要能进来；同时 prompt 里写了"看着都不对就输出空"兜住误提示。

# 列头候选池的清洗规则：混进"大标题/数据值"会让语义排名失准。
#   「某某科技有限公司员工信息表（2026年第一季度）」和第 1 名的相似度 0.58，
#   比「联系电话」的 0.61 只低一点点，但它是标题不是列头。
_HEAD_NOISE = re.compile(
    r"有限公司|股份公司|集团|台账|汇总表|信息表|登记表|跟踪表|明细表|一览表|统计表"
    r"|[（(]?20\d\d\s*年|[（(]截至")
_HEAD_NUM = re.compile(r"^[\d\s.,%+\-/：:年月日（）()]+$")

_SEM_CACHE = {}            # {(id(blocks), core): (raw列表, 加分列表)} —— 大池多列复用
_SEM_KEEP = 8


def _clean_head(text):
    """这一格像不像"列头"（而不是大标题/数据值）。不像就返回 None。"""
    t = str(text or "").strip()
    if not (2 <= len(t) <= 14):
        return None
    if _HEAD_NUM.match(t):          # 920 / 100% / 2026-01-08
        return None
    if _HEAD_NOISE.search(t):       # 大标题、跨列表名
        return None
    return t


def _semantic_column_scores(blocks, core):
    """给每个块算语义加分。返回 (raw, bonus, ranked) 三个结果。

    raw    —— 与 blocks 等长：该块所在列与目标列名的最高余弦（展示/排查用）
    bonus  —— 与 blocks 等长：按**列头排名**折算的阶梯分（真正参与打分的）
    ranked —— [(列头, 余弦), ...]，按相似度降序取前 SEM_TOPK 个
              （给 extract.py 用：让模型知道"这个字段在原文里可能叫这个"）

    为什么挂在"列"上而不是"块"上：同一列的格子共享同一批列头，
    算一次就能给整列用 —— 这也正是我们要的效果（整列一起被召回）。
    """
    key = (id(blocks), core)
    hit = _SEM_CACHE.get(key)
    if hit is not None and len(hit[0]) == len(blocks):
        return hit

    zero = ([0.0] * len(blocks), [0] * len(blocks), [])
    try:
        # semantic.py 在 agent/ 下；matcher 跑起来时 sys.path 里通常没有它。
        # 【踩过的坑】这里不补路径的话 import 会抛 ModuleNotFoundError，
        # 被 except 吞掉 → 语义分全 0，表现为"改了代码但召回一点没变"。
        import sys as _sys
        _agent = str(Path(__file__).resolve().parents[1] / "agent")
        if _agent not in _sys.path:
            _sys.path.insert(0, _agent)
        import semantic
    except Exception:
        return zero

    # 收集唯一的列头文本（清洗后约 150 个，编码一次 ~0.2s）
    terms, seen = [], set()
    for b in blocks:
        for h in (b.get("_colheads") or [])[:5]:
            t = _clean_head(h)
            if t and t not in seen:
                seen.add(t)
                terms.append(t)
    if not terms:
        return zero

    try:
        idx = semantic.build_index(terms, tag="colheads")
        emb, _err = semantic.get_embedder()
        if idx.vectors is None or emb is None:
            return zero
        qv = emb.encode([core])[0]
        sims = idx.score(qv)
        if sims is None:
            return zero
        sims = [float(x) for x in sims]
    except Exception:
        return zero

    # 排名 → 阶梯分
    order = sorted(range(len(terms)), key=lambda i: -sims[i])[:SEM_TOPK]
    rank_bonus = {}
    for rank, i in enumerate(order):
        rank_bonus[terms[i]] = max(0, SEM_TOP_SCORE - rank * SEM_RANK_STEP)
    ranked = [(terms[i], round(sims[i], 4)) for i in order]

    raw_of = {t: sims[i] for i, t in enumerate(terms)}
    raw, bonus = [], []
    for b in blocks:
        heads = [_clean_head(h) for h in (b.get("_colheads") or [])[:5]]
        heads = [h for h in heads if h]
        raw.append(max((raw_of.get(h, 0.0) for h in heads), default=0.0))
        bonus.append(max((rank_bonus.get(h, 0) for h in heads), default=0))

    res = (raw, bonus, ranked)
    if len(_SEM_CACHE) >= _SEM_KEEP:
        _SEM_CACHE.pop(next(iter(_SEM_CACHE)))
    _SEM_CACHE[key] = res
    return res

# 列上下文缓存（2026-10-14 压测后优化）：
#   {id(blocks): (blocks 引用, contexts)} —— 用引用校验防 id 复用误命中；
#   大池 50 万块的构建是秒级开销，多列/重审反复调用时只构建一次。
_CTX_CACHE = {}
_CTX_CACHE_KEEP = 2        # 最多保留最近 2 份（日常小池 + 压测大池这类场景）
_CTX_LOCK = threading.Lock()


def score_blocks(blocks: list, core: str, min_score: int = None, contexts: list = None,
                 use_semantic: bool = True, hints_out: list = None):
    """对全部块打分（含表格列上下文增强），从高到低返回 [(score, block)]。

    统一入口：retrieve.py 和 extract.py 都用它，让"筛选"逻辑只有一份。

    contexts（2026-10-14 优化）：可选，与 blocks 对齐的"同列上下文"列表。
      大池场景下由调用方预先构建/复用（避免每次全量重建）；only_file 子集可直接
      从块上取 _ctx 拼出来传进来。不传时行为与旧版完全一致（现场构建）。
    打分主体（2026-10-14 优化）：prefixes/bigrams/冒号标签提前算好，循环内不再
      每块重复构造——50 万块 × 8 列下这是实打实的热路径。

    use_semantic（2026-10-26 新增）：是否加语义召回分。
      默认开。置 False 可退回"纯字面"，用于回归对比（证明没退化）与
      模型缺失环境（那时 _semantic_column_scores 本身会返回全 0）。
    """
    if min_score is None:
        min_score = MIN_SCORE
    if contexts is None:
        contexts = build_column_contexts(blocks)

    prefixes = [core[:i] for i in range(3, len(core) + 1)]
    labels = [p + "：" for p in prefixes] + [p + ":" for p in prefixes]
    parts = bigrams(core)
    n_parts = len(parts)

    # ---------- 第一遍：只算字面分 ----------
    # 为什么先算一遍：语义是"字面召不回时的兜底"，不是常驻加分项。
    # 只有当**全池字面最高分都不到 40**（= 没有任何一处完整出现这个词）时，
    # 才请语义介入。这样"供应商""设备名称"这类本来就能字面对上的列，
    # 候选和以前一模一样 —— 回归 70 列零变化靠的就是这条。
    literal = []
    for b, ctx in zip(blocks, contexts):
        text = b.get("text") or ""
        probe = (ctx + " " + text).strip() if ctx else text
        s = 0
        if any(l in probe for l in labels):
            s += 90
        if core in probe:
            s += 40
        if n_parts:
            hit = sum(1 for p in parts if p in probe)
            s += int(40 * hit / n_parts)
        if colhead_hits(b, core):
            s += HEAD_EXACT_BONUS
        literal.append(s)

    best_literal = max(literal) if literal else 0
    need_sem = use_semantic and best_literal < 40
    if need_sem:
        sem_raw, sem_bonus, sem_ranked = _semantic_column_scores(blocks, core)
    else:
        sem_raw, sem_bonus, sem_ranked = [0.0] * len(blocks), [0] * len(blocks), []
    if hints_out is not None and sem_ranked:
        # 给调用方（extract.py）一句提示：目标字段在原文里最可能叫这几个名字。
        # 【必须带门槛】实测「合同总价」在只有设备/员工数据的池子里，语义前三名是
        # 「月度合计」「单价（元）」「预算（万元）」—— 全是金额类但都不是它。
        # 把这种提示喂给模型，等于诱导它去抽单价。所以只在 sim ≥ 0.72 时才说，
        # 宁可不说，也不能给错方向。
        hints_out.extend([t for t, s in sem_ranked[:3] if s >= SEM_HINT_MIN])

    scored = []
    for i, b in enumerate(blocks):
        s = literal[i]

        # D) 语义召回 —— 只补"字面没对上"的块
        #    字面已经完整命中（≥40）的块不加：它已经排得很前了，语义再抬它
        #    只会把原本正确的顺序搅乱（回归时 10 张模板的候选顺序必须不变）。
        b["_sem"] = round(sem_raw[i], 4) if i < len(sem_raw) else 0.0
        via_sem = bool(i < len(sem_bonus) and sem_bonus[i] and s < 40)
        b["_via_sem"] = via_sem
        if via_sem:
            s += sem_bonus[i]
        # 记住这一块属于哪一列 —— 下游要按"列"分配候选名额
        loc = b.get("locator") or {}
        if loc.get("kind") in ("cell", "table_cell"):
            b["_colkey"] = (b.get("file"), loc.get("sheet") or loc.get("table_index"),
                            loc.get("col"))
        else:
            b["_colkey"] = (b.get("file"), "text")

        if s >= min_score:
            scored.append((s, b))
    scored.sort(key=lambda t: -t[0])
    return scored


def locate_text(block: dict) -> str:
    """把块的 locator 转成人看得懂的位置说明。"""
    loc = block.get("locator", {})

    # txt / md：行号区间
    if "line_start" in loc:
        if loc["line_start"] == loc.get("line_end"):
            return "第 %d 行" % loc["line_start"]
        return "第 %d–%d 行" % (loc["line_start"], loc["line_end"])

    kind = loc.get("kind")
    if kind == "cell":      # xlsx：单元格
        return "工作表「%s」 %s" % (loc.get("sheet"), loc.get("cell"))
    if kind == "paragraph":  # docx：段落
        return "第 %d 段" % loc.get("para_index")
    if kind == "table_cell":  # docx 表格：单元格
        return "表格%d · 行%d·列%d" % (loc.get("table_index"), loc.get("row"), loc.get("col"))

    # 兜底：没见过的新格式，原样显示
    return " / ".join("%s=%s" % (k, v) for k, v in loc.items())


def shorten(text: str, limit: int = 72) -> str:
    """把多行长文本压成一行预览。"""
    one = text.replace("\n", " ⏎ ")
    if len(one) > limit:
        return one[:limit] + "……"
    return one


def main():
    col_name = sys.argv[1] if len(sys.argv) > 1 else COL_NAME
    core = strip_unit(col_name)

    # 既打印到屏幕，也收集起来，最后统一存成文件
    lines = []

    def say(text=""):
        lines.append(text)
        print(text)

    say("=" * 60)
    say("筛选器 · 列名：%s" % col_name)
    if core != col_name:
        say("（已去掉括号说明，按「%s」匹配）" % core)
    say("=" * 60)

    if not OUTPUT_JSON.exists():
        say("找不到 output.json：%s" % OUTPUT_JSON)
        say("先去 doc_reader 里跑一次 run_all.py，把它生成出来。")
    else:
        data = json.loads(OUTPUT_JSON.read_text(encoding="utf-8"))
        docs = data.get("documents", [])
        # 与 extract.load_pool 同一套范围过滤（用户只勾了几份源文档时）
        _raw = (os.environ.get("A23_DOC_SCOPE") or "").replace(";", "\n")
        _scope = {x.strip() for x in _raw.split("\n") if x.strip()}
        if _scope:
            docs = [d for d in docs if d.get("file") in _scope]
        total = sum(len(d.get("blocks", [])) for d in docs)
        say("扫描范围：%d 个文件 / %d 个内容块" % (len(docs), total))
        say("-" * 60)

        all_blocks = [b for d in docs for b in d.get("blocks", [])]
        scored = score_blocks(all_blocks, core)

        if not scored:
            say("没有找到任何候选块（分数都低于 %d）。" % MIN_SCORE)
        else:
            say("候选块 Top %d（分数越高，越可能含这一列的值）：" % min(TOP_N, len(scored)))
            for i, (s, b) in enumerate(scored[:TOP_N], start=1):
                say()
                say("[%2d] %3d 分 ｜ %s ｜ %s" % (i, s, b.get("file"), locate_text(b)))
                say("      %s" % shorten(b.get("text", "")))

            say()
            say("-" * 60)
            say("人工看一眼：")
            say("  1. 该进的块进来了吗？（拿标准答案的「来源」行号对一对）")
            say("  2. 不相关的块混进来了吗？（它们会干扰后面的模型判断）")

    # 结果留档成文件（终端里万一中文乱码，就打开这个文件看）
    out_path = Path(__file__).resolve().parent / ("筛选结果-%s.txt" % col_name)
    stamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
    out_path.write_text("生成时间：%s\n\n%s\n" % (stamp, "\n".join(lines)), encoding="utf-8")
    print()
    print("（结果已存：%s）" % out_path.name)


if __name__ == "__main__":
    main()
