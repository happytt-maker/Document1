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

from retrieve import (build_column_contexts, locate_text, paragraph_kv_blocks,
                      score_blocks, strip_unit)

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
SEM_TOP_K = 45          # 语义召回场景的候选上限（2026-10-26 新增）
                        #   为什么必须放大：一列常 10~12 个格子，15 个名额只装得下
                        #   1 列多；而"换说法"时正确的列往往排在语义第 3~6 名，
                        #   名额不够它就永远进不来。45 ≈ 能装 3~4 个整列。
                        #   只在"有块是靠语义才进候选"时才放大 —— 字面能对上的
                        #   常规场景仍是 15，行为与以前完全一致（回归靠这条守住）。
KV_SCORE = 45           # 「标签：值」段落兜底块的分数（2026-11-02）
                        #   排在语义 Top1（50）之后、其余语义名次之前：
                        #   weak 列的语义候选本来就是"猜的"，让模型优先看原文标签。
HIST_SCORE = 95         # 历史参照块的分数（2026-11-02）
                        #   只比列头精确命中（100）低一点：这是唯一"真填对过一次"的
                        #   映射，比字面猜和语义猜都可靠，所以整体排到最前面。
HIST_MAX = 30           # 历史锚定最多带几块（一列数据一般 5~15 块；30 是防"备注"这类
                        #   宽列头一次匹配上百块，把模型淹掉）
NARRATIVE_CHARS = 30000 # 叙述体文档"整篇进候选"的字符预算（2026-11-02）
                        #   判据：全池没有一个块带列头 = 这份文档没有"表格列"这回事
                        #   （公报/纪要/报告）。此时"按列名打分挑 15 块"是错的：
                        #   叙述体里一段就是一条记录，挑 15 块 = 只取前 15 条。
                        #   实测（真实数据·城市百强报告 113 段）：「城市名」字面最高分
                        #   只有 20，100 个城市里 85 个进不了候选 → 整表只有 15 行。
                        #   预算 3 万字 ≈ 2 万 token，一列一次调用可以接受；
                        #   超过就退回按分截断（保住大文档不炸），并在报告里说明。
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


def _pool_scope():
    """本次只看哪几份文档（换行分隔）；空 = 全池。

    A23_DOC_SCOPE（2026-11-07）：用户勾了源文档时，只要在**读池子的时候**
    把其余的滤掉就行 —— 不用把上百 MB 的池子拷贝一份小池再让子进程去读。
    以前那套拷贝（api_server._scoped_pool）在 87MB 的池上要 9 秒，
    比整条填表流水线还长，等于"限个范围"要付两次读盘的钱。
    """
    raw = os.environ.get("A23_DOC_SCOPE") or ""
    return {x.strip() for x in raw.replace(";", "\n").split("\n") if x.strip()}


def load_pool():
    """读出全部内容块（带缓存与上下文预建）。"""
    st = OUTPUT_JSON.stat()
    scope = _pool_scope()
    key = (str(OUTPUT_JSON), st.st_mtime_ns, st.st_size, tuple(sorted(scope)))
    with _POOL_LOCK:
        if _POOL_CACHE["key"] != key:
            data = json.loads(OUTPUT_JSON.read_text(encoding="utf-8"))
            docs = data.get("documents", [])
            if scope:
                docs = [d for d in docs if d.get("file") in scope]
            blocks = [b for d in docs for b in d.get("blocks", [])]
            build_column_contexts(blocks)          # 预热：挂 _ctx / _colheads_norm
            _POOL_CACHE["key"] = key
            _POOL_CACHE["blocks"] = blocks
        return _POOL_CACHE["blocks"]

SYSTEM_PROMPT = "你是一个严谨的信息抽取助手，只依据给定的原文抽取信息，绝不编造。只输出 JSON。"


def build_user_prompt(col_name: str, blocks: list, split_subs: list = None,
                      sem_hints: list = None, kv_hint: bool = False,
                      hist_hint: dict = None, narrative: bool = False) -> str:
    """把候选块和任务拼成给模型的一段话。

    split_subs：本列走了拆词兜底时传入子词（如 ["库存","数量"]）。
    2026-10-11 实测：兜底场景（目标是"库存数量"，候选是"期初库存/期末库存"这类
    近似列头）模型会"纠结到放弃"，同题两次跑出 0 条 / 10 条。所以专门加一条
    规则：近似也算，选最接近的整列，别输出空。

    sem_hints（2026-10-26）：语义召回时"最像的几个列头"。
    为什么需要告诉模型：语义把「合同总金额」那一列捞进来了，但任务是"抽取【合同总价】"，
    模型看到列头不是同一个词，可能不敢抽、直接返回空 —— 前面白忙一场。
    把"原文里最可能叫这些"明说出来，模型才敢按语义取整列。

    hist_hint（2026-11-02）：历史填写参照。同一张模板上次是从哪份文档的哪个列头
    抽的（如「库存数量」← 「期末库存」）。这是唯一被真实填写验证过一次的线索，
    比字面和语义都可靠，所以只在候选里确实找到那一列时才传进来。
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
        kvl = (b.get("_kv_label") or "").strip()
        if kvl:
            parts.append("（「标签：值」段落，标签 = %s）" % kvl)
        parts.append(b.get("text", ""))
        parts.append("")

    # 规则逐条攒起来（有特殊情况就多加一条，编号自动顺延）
    rules = [
        "每个值必须来自上面的原文；evidence 字段摘抄包含该值的原文原句（一个字都不要改）",
        "原文里没有的，不要编造；找不到的记录就不输出",
        "以「=」开头的单元格文本是还没求值的公式（如 =E3+F3-G3），不是可用的值 —— 不要抽取它们",
        "有几条记录就输出几条，按在原文中出现的顺序排",
        "一行就是一条记录：同一个值在原文里占了多行（比如同一个人负责两项任务、同一家供应商供多台设备），"
        "每行都要单独输出一次，不要合并、不要去重 —— 合并会让后面所有行整体错位",
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
    # 时间修饰歧义（2026-11-02）：源里常有成对的时间列，目标列名却不带修饰。
    # 模板5「库存数量」：源有「期初库存」「期末库存」两列，语义得分几乎一样，
    # 模型选了期初 → 10 格全错。业务上"库存数量"指当前值 = 期末。
    # 只在**目标列名本身不带**这类修饰时生效，所以「期初库存」这类列不受影响。
    if not any(k in core for k in ("期初", "期末", "上期", "本期", "年初", "年末",
                                   "上月", "本月")):
        rules.append("如果候选里有成对的时间修饰列（如「期初库存」和「期末库存」、"
                     "「上期余额」和「本期余额」、「年初」和「年末」），而目标列名本身不带这类修饰，"
                     "就取时间更晚的那一列（期末／本期／年末／本月）—— 目标指的通常是当前/最终的值")
    if narrative:
        rules.append("本文档是**叙述体**（没有表格）：每个自然段描述一条记录（例如每个城市一段）。"
                     "请逐段抽取，一段一条记录，按段落出现顺序**全部取全** —— "
                     "不要只取前几条、不要把格式相同的段落合并、不要因为段落里没有出现列名就跳过它。"
                     "候选块的相关度分数在这里没有意义（都是并列记录），一律视为 0，不要按分数挑")
        rules.append("数字写法照原值转成规范形式：千分位逗号去掉（56,708.71 → 56708.71），"
                     "小数位保持原样，不要四舍五入、不要改精度")
    if split_subs:
        rules.append("本列在原文里可能以近似说法出现（检索时是按「%s」「%s」两个字分别召回候选的）。"
                     "若候选里有一组数据（同一表格、同一列连续多行）与目标列名部分匹配，就取这组、整列取全；"
                     "多组部分匹配时，取与目标列名共同字更多、或语义更接近的一组；"
                     "不要因为「不完全等于」就输出空" % (split_subs[0], split_subs[1]))
    if sem_hints:
        rules.append("语义检索发现：原文里与【%s】意思最接近的写法是 %s。"
                     "若候选里有一组数据（同一表格、同一列连续多行）的列头正是其中之一，"
                     "就把它当作目标列、整列取全 —— 叫法不同不代表不是这一列；"
                     "但若这几个说法看着都不对、或原文里确实没有这一列，"
                     "就输出空（宁可留空，不要硬凑到金额/数量这类近邻列上）。"
                     % (col_name, "、".join("「%s」" % h for h in sem_hints)))
    if kv_hint:
        rules.append("本列在源文档里可能根本不叫这个名字，所以候选里额外放了几条"
                     "「标签：值」形式的段落。若某个标签的意思与【%s】相同、只是叫法不同"
                     "（例如「乙方（供货方）」就是「供应商」、「投诉单号」就是「投诉编号」），"
                     "就把冒号后面的那一段当作值抽出来，这类段落有几条就输出几条；"
                     "标签意思无关的（如「签订地点」「盘点日期」）一条都不要抽 —— "
                     "看不对就输出空，不要硬凑" % col_name)
    if hist_hint and hist_hint.get("head"):
        _f = hist_hint.get("file")
        _scope = ("候选里只放了这一列（共 %d 块）" % hist_hint.get("n", 0)
                  if hist_hint.get("whole") else "这一列的候选已排在最前面")
        rules.append("历史填写记录：这一列上一次是从%s「%s」那一整列抽出来的，%s。"
                     "若这组候选确实是同一列连续多行，就整列取全 —— 有几块数据就输出几条，不要只抽一部分；"
                     "但原文是唯一依据：这一列在这批文档里没有、或内容明显不是【%s】，"
                     "就输出空，不要为了对上历史去抄别的列"
                     % (("《%s》的" % _f) if _f else "", hist_hint["head"], _scope, col_name))
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


def _unwrap_items(d):
    """剥掉模型自作主张加的外壳。

    有些模型在 response_format=json_object 下会把结果包一层：
        {"answer": {"items": [...]}}    ← 智谱 GLM 就常这样
    也可能包成 result / data / output。顶层已经有 items 的不动。
    2026-11-02 加：换到智谱后实测到这一层，不剥就会"解析成功但 0 条"，
    看起来像模型没抽到东西，极难排查。
    """
    if not isinstance(d, dict) or "items" in d:
        return d
    for k in ("answer", "result", "data", "output", "response", "json"):
        v = d.get(k)
        if isinstance(v, dict) and "items" in v:
            return v
    if len(d) == 1:                      # 只有一个键、值还是字典：多半就是壳
        v = next(iter(d.values()))
        if isinstance(v, dict) and "items" in v:
            return v
    return d


def parse_json_reply(text: str):
    """宽容地解析模型的回复：允许 ```json 围栏、前后多余文字、多余外壳。"""
    t = (text or "").strip()
    if t.startswith("```"):
        t = re.sub(r"^```[A-Za-z]*\s*", "", t)
        t = re.sub(r"```\s*$", "", t)
    i, j = t.find("{"), t.rfind("}")
    if i == -1 or j == -1:
        return None
    try:
        return _unwrap_items(json.loads(t[i:j + 1]))
    except Exception:
        return None


PUNCT = r"[\s，。、；：！？,.;:!?（）()【】\[\]「」『』\"'']+"


def norm_text(s: str) -> str:
    """归一化：去掉空白与标点（定位用，不影响展示）。"""
    return re.sub(PUNCT, "", s or "")


def _hnorm(s: str) -> str:
    """列头归一（2026-11-02）：只统一空白和括号/冒号，**不删标点**。

    为什么不直接用 norm_text：列头要跟历史记录里的字面比对，
    "期初库存（件）" 和 "期初库存(件)" 应当相等，但 "库存/数量" 不该被
    压成 "库存数量"（那是两回事）。所以这里只做最保守的等价替换。
    """
    s = re.sub(r"[\s\u3000]+", "", str(s or ""))
    for a, b in (("（", "("), ("）", ")"), ("：", ":"), ("，", ","), ("、", ",")):
        s = s.replace(a, b)
    return s.strip().lower()


def _cell_match(nl: str, ev: str) -> bool:
    """格子的文本和 evidence 算不算对上了（只用于表格单元格定位）。

    2026-11-02 修：不能简单用"互相包含"。模板5「库存数量」实测：
    evidence = "95"，而 H3=95、H5=1950 —— 归一化后 "95" in "1950" 成立，
    于是 1950 被定位到 H3，同一批里的 95 也定位到 H3，行号撞车、整列跟着错位
    一格（评分 100% → 97.6%，漏 2 格）。
    纯数字的格子必须"独立出现"（前后不是数字或小数点）才算对上；
    文字仍走包含判断（模型抄的常是含该值的整句）。
    """
    if not nl or not ev:
        return False
    if nl == ev:
        return True
    if re.fullmatch(r"[\d.]+", nl):
        return re.search(r"(?<![\d.])" + re.escape(nl) + r"(?![\d.])", ev) is not None
    return ev in nl or nl in ev or ev[:10] in nl


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
    raw = evidence or ""
    # 剥掉 prompt 里我们自己加的说明（模型常把整块内容连说明一起抄回来）：
    #   「（该格所在列的列头：X）」「（「标签：值」段落，标签 = X）」
    # 不剥的后果（2026-11-02 模板5 实测）：归一化会去掉标点，于是 evidence 里
    # 含"期初库存"，**表头格 E2 反查也能命中**（nl in ev），10 条值全被定位到 E2 →
    # 写回时基准行号取成 2，整表下移一格，模板5 从 88.1% 掉到 15.5%。
    raw = re.sub(r"（该格所在列的列头[^）]*）", "", raw)
    raw = re.sub(r"（「标签：值」段落[^）]*）", "", raw)
    ev = norm_text(raw)
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
            if _cell_match(nl, ev):
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
        if _cell_match(nl, ev):
            return None, locate_text(block)

    return None, None


def _assign_by_value(items, blocks):
    """把记录分配到格子上，规则：**同一个格子只能对应一条记录**。

    2026-11-02 修（模板5「单位」实测）：一列里"个"出现 3 次、"台"出现 2 次，
    而定位是按"值反查格子"，于是 3 条"个"全挤在同一格、另外 2 格空着
    → 漏填 2 格。这不是模型抽错，是定位分配错了。
    同值的格子按原文顺序各用一次；格子用完了才允许重复（那是模型多输出了）。

    返回重新分配了几条（0 = 原定位本来就都对）。
    """
    cells = [b for b in blocks
             if ((b.get("locator") or {}).get("kind") in ("table_cell", "cell"))]
    if not cells:
        return 0
    by_val = {}
    for k, b in enumerate(cells):
        by_val.setdefault(norm_text(b.get("text") or ""), []).append(k)

    used, changed = set(), 0
    for it in items:
        if it.get("from_paragraph"):        # 段落来源是"文档级"值，本来就该广播
            continue
        v = norm_text(it.get("value") or "")
        if not v:
            continue
        idxs = by_val.get(v) or []
        cur = it.get("loc_text")
        pick = None
        for k in idxs:                      # ① 优先保住当前定位的那一格
            if k not in used and cur and locate_text(cells[k]) == cur:
                pick = k
                break
        if pick is None:                    # ② 换一个同值、还没被占用的格子
            for k in idxs:
                if k not in used:
                    pick = k
                    break
        if pick is None:
            continue
        used.add(pick)
        b = cells[pick]
        new_loc = locate_text(b)
        if new_loc != cur:
            it["loc_text"] = new_loc
            it["line"] = None
            it["file"] = b.get("file")
            it["head"] = b.get("_head")
            it["aligned"] = "same_value_once"
            changed += 1
    return changed


# =================================================================
# 叙述体 · 整表一次抽取（2026-11-02）
# =================================================================
# 为什么不能沿用 run_column 一列一列抽：
#   真实数据实测（《2025 年中国城市经济百强全景报告》→ 5 列 × 100 行）：
#   单列「城市名」56.9 秒，5 列 = 285 秒，赛题红线是 90 秒/文档 —— 差 3 倍多。
#   根因不是模型慢，是**同一份 1.1 万字原文被完整喂了 5 遍**，每一遍都在
#   "重新读一遍文章，只回答一个问题"。
#   叙述体的结构决定了正确答案：一段 = 一行记录。所以正确切法是一次把
#   "每一段对应哪一行"抽出来 —— 原文只读一次，答案按行给全。
#
# 与 run_column 的差别（为什么不能简单合并）：
#   * 表格文档：列头是导航，一列一抽才能"整列取全、不串列"；
#   * 叙述体：没有列头，段落是记录单位，逐列抽等于把同一段读 5 遍。
#   两条路各有各的适用面，所以并列实现，由 anchor/pipeline 按"有无列头"选路。

TABLE_MAX_ROWS = 400        # 一次最多输出多少行（防模型吐太长直接超时）


def _num_norm(s: str) -> str:
    """数字归一：去千分位逗号与空白（56,708.71 → 56708.71）。"""
    return re.sub(r"[,，\s\u3000]", "", str(s or ""))


def value_in_text(value, text) -> bool:
    """这个值是不是真在这段原文里。

    它挡的是"编造/算错"（模型把 5.49% 算成数字、把不存在的值填进来），
    挡不住"取错字段"（收入取成支出 —— 那个数字确实在原文里）。
    后者归提示词规则和人工抽检管，不能指望这个校验。
    """
    v = str(value or "").strip()
    if not v:
        return False
    t = str(text or "")
    if v in t:
        return True
    nv = _num_norm(v)
    return bool(nv) and nv in _num_norm(t)


def snippet_around(text, value, width: int = 28) -> str:
    """从段落里截一段含该值的原文当 evidence。

    叙述体一段可能有 170 字，整段抄回结果文件会让 100 行 × 5 列的结果膨胀到
    几百 KB，而且人看报告时也抓不到重点。截「值前后各 28 字」足够自证出处。
    """
    t = str(text or "")
    v = str(value or "").strip()
    if not t:
        return ""
    idx = t.find(v) if v else -1
    if idx < 0:
        # 千分位差异（原文 56,708.71 / 值 56708.71）：去逗号后再找，
        # 用位置映射还原回原串下标 —— 否则截出来的片段对不上这个值。
        keep = [i for i, ch in enumerate(t) if ch not in ",，"]
        clean = "".join(t[i] for i in keep)
        j = clean.find(_num_norm(v)) if v else -1
        idx = keep[j] if 0 <= j < len(keep) else -1
    if idx < 0:
        return t[:width * 2] + ("…" if len(t) > width * 2 else "")
    a = max(0, idx - width)
    b = min(len(t), idx + len(v) + width)
    return ("…" if a else "") + t[a:b] + ("…" if b < len(t) else "")


def build_table_prompt(col_names, blocks) -> str:
    """整表抽取的提示词：一次读原文，按「段 → 行」输出全部列。"""
    parts = []
    parts.append("下面是源文档的全部 %d 个段落（按原文顺序编号，没有表格）：" % len(blocks))
    parts.append("")
    for i, b in enumerate(blocks, start=1):
        parts.append("[块%d] %s" % (i, (b.get("text") or "").replace("\n", " ")))
    parts.append("")
    parts.append("=" * 40)
    parts.append("任务：把这张表**一次填完**。要抽的列（顺序固定，共 %d 列）：" % len(col_names))
    for i, c in enumerate(col_names, start=1):
        core = strip_unit(c)
        unit = c.replace(core, "").strip("（）() ")
        parts.append("  [%d] %s%s" % (i, c, ("（只填数字，不带单位「%s」）" % unit) if unit else ""))
    parts.append("")

    rules = [
        "一段 = 一条记录。这一行的所有字段都从**同一个段落**里取，不要跨段拼凑",
        "只输出真正在描述一条记录的段落。章节标题（如「一、双极领航……」）、"
        "综述/概括性段落（如「2025 年，中国城市经济……」这类总述）不要输出成记录",
        "v 数组的长度必须等于列数 %d，严格按上面给的列顺序。"
        "某一段确实没有这个字段就写 null —— 不要猜、不要拿别的数字顶替、不要留空字符串" % len(col_names),
        "数字照原文取：去掉千分位逗号（56,708.71 → 56708.71），小数位原样保留，不要四舍五入、不加单位",
        "同一个段落里往往有好几个数字，**按含义匹配，不要按出现顺序硬套**："
        "叙述顺序会变（有的段先讲人口再讲 GDP），顺序一套错就整行皆错",
        "记录顺序 = 段落顺序（叙述体的段落顺序常带排名/时间含义，不要重排、不要按数值大小排）",
        "找不到的段落不要硬凑；宁可行数少，不要编造",
    ]
    # 百分比/增速：只有当目标列本身就是"率/比/增速"时才可能是值
    if not any(k in "".join(col_names) for k in ("率", "占比", "比重", "增速", "增幅")):
        rules.append("「同比增长 X%」「占比 X%」这类百分比是增速，不是这几列的值 —— 不要抽它")
    # 成对反义字段（收入/支出、期初/期末）：叙述体里最爱串的一处
    pairs = [("收入", "支出"), ("期初", "期末"), ("进口", "出口"), ("应收", "应付"),
             ("上年", "本年"), ("流入", "流出"), ("买入", "卖出")]
    for a, b in pairs:
        hit = [c for c in col_names if a in c]
        if hit and any(b in c for c in col_names):
            rules.append("注意区分「%s」和「%s」：这两列在这一段里都出现了，各取各的，不要互相顶替"
                         % (a, b))
            break

    parts.append("规则（照做，不要发挥）：")
    for i, r in enumerate(rules, start=1):
        parts.append("%d. %s" % (i, r))
    parts.append("")
    parts.append('输出格式（JSON）：{"records": [{"block": 段落编号, "v": [值1, 值2, ...]}, ...]}')
    parts.append("block 填上面方括号里的块编号（不是行序号）。只输出这个 JSON，不要任何解释文字。")
    return "\n".join(parts)


def run_table(col_names, out_dir=None, quiet=False, only_file=None,
              max_rows: int = TABLE_MAX_ROWS):
    """叙述体整表抽取：一次调用，填满整张表。

    返回 {"cols": {列名: 条数}, "elapsed_sec": 秒, "usage": {...}, "records": 行数}；
    失败返回 None。

    写出的文件与 run_column 完全一致（抽取结果-<列名>.json），所以 writer/fill.py
    不需要区分这次是"整表抽的"还是"逐列抽的"。
    """
    say = (lambda *a, **k: None) if quiet else print
    out_base = Path(out_dir) if out_dir else Path(__file__).resolve().parent
    out_base.mkdir(parents=True, exist_ok=True)

    # 结构列（序号）不来自源文档，跳过；否则模型会去原文里找"1、2、3"
    cols = [c for c in col_names if c and c != "序号"]
    if not cols:
        say("没有要抽的列。")
        return None

    say("=" * 60)
    say("整表抽取 · %d 列：%s" % (len(cols), "、".join(cols)))
    say("=" * 60)

    if not OUTPUT_JSON.exists():
        say("找不到 output.json：%s（先跑 doc_reader）" % OUTPUT_JSON)
        return None

    all_blocks = load_pool()
    pool = [b for b in all_blocks if b.get("file") == only_file] if only_file else all_blocks
    # 保原文顺序：记录顺序有业务含义（排名/时间），按相关度排序会把它打乱
    cands = list(pool)
    chars = sum(len(b.get("text") or "") for b in cands)
    say("① 筛：%d 个块全量进候选（%d 字，叙述体整篇）" % (len(cands), chars))
    if not cands:
        say("池子是空的。")
        return None
    if chars > NARRATIVE_CHARS:
        say("   注意：%d 字已超预算 %d，模型可能漏记录 —— 建议先分段" % (chars, NARRATIVE_CHARS))

    # ---------- ② 判（一次） ----------
    prompt = build_table_prompt(cols, cands)
    say("")
    say("② 判：调用大模型一次（%s）……" % MODEL)
    content, usage, cost = ask_model(prompt)
    if content is None:
        return None
    parsed = parse_json_reply(content)
    if not parsed or "records" not in parsed:
        say("模型输出不是合法 JSON，原样打印：")
        say(content)
        return None

    records = parsed["records"]
    if not isinstance(records, list):
        say("records 不是数组，原样打印：")
        say(content)
        return None
    if len(records) > max_rows:
        say("   模型输出了 %d 行，超过上限 %d —— 只取前 %d 行" % (len(records), max_rows, max_rows))
        records = records[:max_rows]

    # ---------- ③ 定：按列拆开 + 逐值回原文校验 ----------
    per_col = {c: [] for c in cols}
    no_src, unverified = [], 0
    for rec in records:
        if not isinstance(rec, dict):
            continue
        bno = rec.get("block")
        vals = rec.get("v")
        if not isinstance(vals, list):
            continue
        if not (isinstance(bno, int) and 1 <= bno <= len(cands)):
            no_src.append(rec)
            continue
        blk = cands[bno - 1]
        text = blk.get("text") or ""
        for k, col in enumerate(cols):
            if k >= len(vals):
                break
            raw = vals[k]
            if raw is None:
                continue
            value = str(raw).strip()
            if not value:
                continue
            ok = value_in_text(value, text)
            if not ok:
                # 原文里找不到 → 不写进表（宁可留空，不留错值），但要记下来给人看
                unverified += 1
                continue
            per_col[col].append({
                "value": value,
                "evidence": snippet_around(text, value),
                "block": bno,
                "file": blk.get("file"),
                "head": None,
                "colheads": None,
                "line": None,
                "loc_text": "第 %d 段" % ((blk.get("locator") or {}).get("para_index") or bno),
                "from_table_once": True,
            })

    say("")
    say("③ 定：%d 行记录 → 逐值回原文校验" % len(records))
    if unverified:
        say("   %d 个值在对应段落里找不到（已丢弃，不写进表）—— 多半是模型抄错或取错字段" % unverified)
    if no_src:
        say("   %d 行没给出有效块编号（已跳过）" % len(no_src))

    result_cols = {}
    for col in cols:
        items = per_col[col]
        safe_col = re.sub(r'[\\/:*?"<>|]', "_", col)
        res = {
            "col": col,
            "time": time.strftime("%Y-%m-%d %H:%M"),
            "elapsed_sec": round(cost, 2),
            "usage": usage,
            "candidate_count": len(cands),
            "items": items,
            "table_once": True,          # 标记：这份结果是"整表一次抽"出的
            "shared_cost": True,         # 耗时是全表共用的，不是这一列自己的
        }
        if only_file:
            res["only_file"] = only_file
        (out_base / ("抽取结果-%s.json" % safe_col)).write_text(
            json.dumps(res, ensure_ascii=False, indent=2), encoding="utf-8")
        result_cols[col] = len(items)
        say("   %s：%d 条" % (col, len(items)))

    say("-" * 60)
    say("本次耗时：%.2f 秒（整表 %d 列共用这一次调用）" % (cost, len(cols)))
    say("本次用量：%s" % usage)
    say("")

    return {
        "cols": result_cols,
        "elapsed_sec": round(cost, 2),
        "usage": usage,
        "records": len(records),
        "unverified": unverified,
        "no_src": len(no_src),
    }


def run_column(col_name: str, quiet: bool = False, only_file: str = None,
               allow_split: bool = True, out_dir: str = None, weak: bool = False,
               hist: dict = None):
    """完整跑一列：筛 → 判 → 定。返回结果 dict（失败返回 None）。

    quiet=True 时不在屏幕唠叨（给 run_all.py 并行调用用，避免多列打印打架）。
    only_file="某文件名" 时只在那一份文档里检索（给一致性重审用，见 run_all.py）。
    allow_split=False 时不做拆词兜底（2026-10-11 全量实测：拆词兜底在"全文档池"
      里干扰太大，模型会纠结几十秒还只抽一半；改由一致性重审阶段在"主源"内做，
      环境干净、快且全。单列调试时保持默认 True）。
    out_dir=结果文件写到哪个目录（默认写在本目录 matcher/；模块（1）的内容提取
      接入时传自己的输出目录，避免污染本目录）。
    weak=True（2026-11-02）：定源判定"这一列在源文档里字面完全对不上"。
      此时追加「标签：值」段落兜底 —— 换说法的字段（供应商→乙方（供货方））
      只可能藏在段落里，光靠表格列头永远召不回。正常列不传，行为不变。
    hist（2026-11-02）：历史填写参照 {"file": 上次的源文档, "head": 上次的列头}。
      这是唯一"被真实填写验证过"的映射（模板5「库存数量」← 《库存盘点表》「期末库存」），
      优先于字面和语义：把该列头的整列块插到候选最前面，并在提示词里点名。
      找不到那一列时这条线索自然失效（不注入、不提示词），退化到今天的行为。
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
    sem_hints = []
    scored = score_blocks(pool, core, min_score=MIN_SCORE, contexts=contexts,
                          hints_out=sem_hints)
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

    # ---------- 叙述体整篇召回（2026-11-02）----------
    # 判据不是"文档名/扩展名"，而是结构：全池没有一个块带列头。
    # 有列头 = 表格文档（列头本身就是导航，按列名挑块是对的）；
    # 没列头 = 叙述体（一段一条记录，块之间是"并列记录"关系，没有"列"可锚）。
    # 这时必须整篇进候选，且**保文档原序** —— 记录顺序有业务含义（排名/时间/条目号），
    # 按相关度排序会把 100 条记录打乱成无意义的顺序。
    narrative_used = False
    narrative_note = ""
    legacy_trunc = os.environ.get("A23_LEGACY_TRUNC") == "1"
    if legacy_trunc:
        narrative_note = "对照实验：关闭叙述体整篇召回，回到按相关度取候选"
        say("① 筛：" + narrative_note)
    elif pool and not any((b.get("_head") or "").strip() for b in pool):
        chars = sum(len(b.get("text") or "") for b in pool)
        if chars <= NARRATIVE_CHARS:
            smap = {id(b): s for s, b in scored}
            cands = [dict(b, _score=smap.get(id(b), 0)) for b in pool]
            narrative_used = True
            narrative_note = "叙述体文档（%d 段，无一处带列头），整篇 %d 字进候选" % (len(pool), chars)
            say("① 筛：%s → 保原序，不按相关度截断" % narrative_note)
        else:
            narrative_note = ("叙述体文档但 %d 字 > 预算 %d，仍按相关度取候选"
                              "（长报告会漏记录，需先分段）" % (chars, NARRATIVE_CHARS))
            say("① 筛：" + narrative_note)

    # 候选名额：常规 15；拆词兜底 36；靠语义才召回的场景 45（装得下 3~4 个整列）
    if not narrative_used:
        if split_used:
            limit = SPLIT_TOP_K
        elif any(b.get("_via_sem") for _s, b in scored[:SEM_TOP_K]):
            limit = SEM_TOP_K
        else:
            limit = TOP_K
        cands = [dict(b, _score=s) for s, b in scored[:limit]]

    # ---------- 键值段落兜底（2026-11-02）----------
    # 只对 weak 列启用（anchor 判定"源内字面分 <20"= 字面完全对不上）。
    # 典型：模板1「供应商」，源里写的是「乙方（供货方）：泰隆机电设备有限公司」——
    # 段落块没有列头，语义分恒为 0，永远进不了候选，模型只能交白卷。
    # 把这类「标签：值」段落整条送进去，由模型判断"标签是不是同一个字段的另一种叫法"。
    kv_used = False
    if weak:
        kv = paragraph_kv_blocks(pool)
        have = {id(b) for b in cands}
        for label, b in kv:
            if id(b) in have:          # 已经靠字面/语义进来的，不重复加
                continue
            cands.append(dict(b, _score=KV_SCORE, _kv_label=label))
            kv_used = True
        if kv_used:
            say("① 筛：本列字面对不上 → 追加 %d 条「标签：值」段落（%s）"
                % (len([c for c in cands if c.get("_kv_label")]),
                   "、".join(c["_kv_label"] for c in cands if c.get("_kv_label"))))

    # ---------- 历史填写参照（2026-11-02）----------
    # 放在字面/语义/段落兜底都做完之后：那些是"猜这一列可能在哪"，
    # 这一步是"这一列上次就在这儿、而且填对了"。可信度天然高一档，
    # 所以把那整列提到候选最前面，并在提示词里点名。
    hist_used = {}
    if hist and hist.get("head"):
        hh = _hnorm(hist["head"])
        hf = hist.get("file")
        same_file, other_file = [], []
        for b in pool:
            heads = b.get("_colheads") or ([b.get("_head")] if b.get("_head") else [])
            if not any(_hnorm(h) == hh for h in heads):
                continue
            if hf and b.get("file") != hf:
                other_file.append(b)
            else:
                same_file.append(b)
        anchors = same_file or other_file
        if anchors:
            anchors = anchors[:HIST_MAX]
            # "锚到了整列"还是"只锚到一格"：≥2 块才认整列（列头 + 至少一条数据）。
            # 认整列时**收缩**候选 —— 历史已经确定这一列在哪，再把期初库存、
            # 存放库位这些干扰块摆给模型看，只会让它重新纠结一遍：
            # 实测叠加式把 33.4s/13065token 拖到 36.2s/17055token（2026-11-02）。
            whole = len(anchors) >= 2
            if whole:
                cands = [dict(b, _score=HIST_SCORE, _hist=True) for b in anchors]
            else:
                a_ids = {id(b) for b in anchors}
                rest = [c for c in cands if id(c) not in a_ids]
                cands = [dict(b, _score=HIST_SCORE, _hist=True) for b in anchors] + rest
            hist_used = {
                "head": hist["head"],
                "file": anchors[0].get("file"),
                "n": len(anchors),
                "whole": whole,                   # True = 候选只剩这一列
                "cross_file": not same_file,      # 换了一批数据文档，但列头对上了
            }
            say("① 筛：历史参照 → 《%s》「%s」%d 块%s%s"
                % (anchors[0].get("file"), hist["head"], len(anchors),
                   "，候选收缩为这一列" if whole else "，排到最前",
                   "（历史源文件不在本次池里，按同名列头跨文件取）" if not same_file else ""))

    if narrative_used:
        say("① 筛：从 %d 个块中取 %d 个候选（叙述体整篇）" % (total, len(cands)))
    else:
        say("① 筛：从 %d 个块中挑出 %d 个候选块%s" % (
            total, len(cands),
            "（含语义召回，故放宽到 %d）" % limit if limit != TOP_K else ""))
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
    # 叙述体不走语义：候选是整篇段落，语义名次没有意义（也省一次编码）
    sem_used = False if narrative_used else \
        any(b.get("_via_sem") for _s, b in scored[:limit])
    if hist_used.get("whole"):
        # 候选已经收缩成历史锚定的那一列，语义候选全被剔掉了 ——
        # 这时再报"靠语义找回"是假的（会给评委一个不存在的证据）。
        sem_used = False
    if sem_used and sem_hints:
        say("① 筛：字面没对上 → 语义召回（原文里最像的写法：%s）"
            % "、".join("「%s」" % h for h in sem_hints[:3]))
    prompt = build_user_prompt(col_name, cands, split_subs=(subs if split_used else None),
                               sem_hints=(sem_hints if sem_used else None),
                               kv_hint=kv_used, hist_hint=(hist_used or None),
                               narrative=narrative_used)
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
    head_skipped = []
    for it in raw_items:
        value = str(it.get("value", "")).strip()
        evidence = str(it.get("evidence", "")).strip()
        bno = it.get("block")

        # 表头自指护栏（2026-11-02）：模型偶尔把"列头那一格"也当成一条记录。
        # 模板2「部门」实测抽到 9 条，第 1 条 value="部门"、出处 D2（表头行）——
        # 于是整列下移一格、8 条数据全部错位，模板2 从 100% 掉到 26%。
        # 判据取"值 == 列名"（去单位后），这是所有表格里最干净的一条：
        # 表头那格的文本就是列名，抄下来必然相等；正常数据格不会恰好等于列名。
        if value and strip_unit(value) == core:
            head_skipped.append(value)
            continue

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

        item = {
            "value": value,
            "evidence": evidence,
            "block": bno,
            "file": (src_block or {}).get("file"),
            # 来源块所在列的列头（2026-11-02）：历史填写参照靠它落地 ——
            # 记一次「模板列 ← 源列头」，下次同名列头直接当锚点，不用再猜。
            # colheads 多存一份"同列前 5 个非空文本"：_head 取的是最上面那条，
            # 碰上 A 列顶上压着文档大标题（"××盘点表（截至…）"）就会记错，
            # 由 history 从这 5 个里挑真正像列头的那个。
            "head": (src_block or {}).get("_head"),
            "colheads": (src_block or {}).get("_colheads"),
            "line": line,
            "loc_text": loc_text,
        }
        # 来源是"表外段落"的标记（2026-11-02）：合同「乙方（供货方）：泰隆机电」
        # 这种值是**文档级**的 —— 一份合同只有一个乙方，却适用于表里每一行。
        # 写回时靠这个标记把它广播到每条记录（fill.build_slot_map）。
        # 判据必须是"段落来源"而不是"只有一条"：后者可能是模型漏抽，广播=把漏抽变错填。
        # 例外（2026-11-02）：叙述体文档里段落就是记录本身（一段一个城市），
        # 广播会把 100 个城市全写成同一个 —— 这条标记必须关掉。
        if (not narrative_used) and \
                ((src_block or {}).get("locator") or {}).get("kind") == "paragraph":
            item["from_paragraph"] = True
        items.append(item)

    # ---------- ③.5 格子不重复占用（2026-11-02）----------
    # 对所有列都做：这一条只在"两条记录指向同一个格子"时才动手，没冲突就一条不改
    # （所以已经验过的表不会被它带偏）。有冲突本身就是 bug —— 同一个格子不可能
    # 同时是两行的值；换成"另一个同值、还没被占用的格子"几乎必然更对。
    realigned = 0
    if items:
        realigned = _assign_by_value(items, cands)

    # ---------- 打印 ----------
    say("")
    say("③ 定：拿摘抄回原文定位出处")
    if realigned:
        say("（%d 条重新分配了格子：同值只占用一个格子 —— 修的是「个」「台」这类"
            "重复值挤在同一行）" % realigned)
    if head_skipped:
        say("（跳过 %d 条「表头自指」：值是列名本身，不是数据 —— 模型把表头那格当记录了）"
            % len(head_skipped))
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
    # 语义召回的依据（2026-10-26）：把"这个字段在原文里最可能叫什么"记下来，
    # 供上游（run_all / pipeline / 界面）展示"模板列名 → 源文档说法"的对照。
    if sem_used:
        result["semantic_hints"] = sem_hints[:3]
    if only_file:
        result["only_file"] = only_file        # 一致性重审：本列被限定在这份文档里抽
    if split_used:
        result["split_fallback"] = True         # 本列走了拆词兜底召回
    if head_skipped:
        result["head_skipped"] = head_skipped   # 被"表头自指"护栏丢掉的条目（排查用）
    if hist_used:
        result["history"] = hist_used           # 本列用了历史参照（锚定在哪一列）
    if narrative_note:
        result["narrative"] = {"used": narrative_used, "note": narrative_note}
    out_path = out_base / ("抽取结果-%s.json" % safe_col)
    out_path.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    say("（结果已存：%s）" % out_path.name)
    return result


def main():
    args = sys.argv[1:]
    plain = [a for a in args if not a.startswith("--")]
    # 整表模式（叙述体）：python extract.py --table "城市名" "GDP总量（亿元）" ...
    if "--table" in args:
        run_table(plain)
        return
    col_name = plain[0] if plain else COL_NAME
    run_column(col_name)


if __name__ == "__main__":
    main()
