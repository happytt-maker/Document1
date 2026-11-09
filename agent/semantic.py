# -*- coding: utf-8 -*-
"""
semantic.py —— 字段语义匹配：不靠别名表，靠「意思近」
================================================================

要解决的问题
------------
模板上写「合同总金额」，用户嘴上说「合同金额」「合同总价」「合同款」；
源文档表头可能写「总价（元）」「合同金额合计」。

原来靠两张手工表在撑：
    · agent/字段别名表.json      —— 5 条人工映射，换行业要重配一遍
    · matcher/retrieve.py 的打分 —— 纯字面（两字片段命中率）
两张表的共同毛病：**只认字面重合**。「总价」和「合同总金额」只有一个「价」字，
片段命中率极低，召回不到；只能人工补一条映射。换个行业（医疗/物流/教育）
字段名全换，别名表要从头配。

本文件的做法
------------
把「字段名」当成一句很短的话，用中文向量模型算语义相似度：
    「合同总价」 vs 「合同总金额」  → 0.83（语义近）
    「合同总价」 vs 「供应商名称」  → 0.31（不近）
候选清单不是人工写的，而是**从数据里长出来的**（文档列头、模板列名、库里
已有字段名），所以换行业不需要配任何东西 —— 新行业的数据进来，候选空间
自动变成那个行业的词。

三路信号，按可信度从高到低融合：
    ① 别名表命中        分数 1.00  —— 手工确认过的，最可信（退成兜底/覆盖）
    ② 字面命中          分数 ≈1.0  —— 完全相等 / 互相包含，不用模型也不该错
    ③ 语义向量          0~1       —— 主力，同一意思不同说法走这一路

关键工程约束（都是踩过的）
--------------------------
· **离线可跑**：模型是本地 ONNX（models/bge-small-zh-v1.5），不联网、不花钱、
  不受接口额度影响。赛题环境断网也能跑。
· **模型缺失不许崩**：只装了 openpyxl 的同门 clone 下来，没有 95MB 模型，
  这时自动降级成「字面 + 别名」，功能不消失，只是少了语义那一路。
· **候选向量要缓存**：339 个列头每次重编码约 0.3s，多列并发会叠加；
  按候选清单的指纹存 npz，第二次直接 load（几毫秒）。
· **线程安全**：api_server 是多线程，ONNX session 可并发读，索引构建要加锁。

对外只暴露三样东西：
    match(query, candidates, aliases)  → [(候选, 分数, 依据), ...]
    resolve(query, candidates, aliases) → 最佳候选 or None（低于阈值就 None）
    FieldSpace(...)                     → 把一堆说法自动聚成同义簇（看"字段空间"）
"""
import hashlib
import json
import math
import os
import re
import threading
import time
from pathlib import Path

BASE = Path(__file__).resolve().parents[1]        # = 研一项目/
MODEL_DIR = Path(os.environ.get("A23_SEMANTIC_MODEL") or (BASE / "models" / "bge-small-zh-v1.5"))
CACHE_DIR = Path(os.environ.get("A23_SEMANTIC_CACHE") or (BASE / "agent" / "语义缓存"))

# ---- 阈值（用真实模板校准过，见 收尾对照/语义匹配-*）----
# bge-small-zh 对「短词 vs 短词」的余弦分布普遍偏高（无关词也有 0.3~0.5），
# 所以不能拿 0.5 当"像"；实测：
#   同一字段不同说法（合同总价/合同金额）0.80~0.92
#   同领域但不同字段（合同金额/付款方式）  0.48~0.62
#   跨领域无关（合同金额/设备型号）        0.25~0.45
# 取 0.68 为"敢用"，0.60 为"可作候选"，0.60 以下不当真（宁可留空，不猜）。
SEM_STRONG = 0.68
SEM_WEAK = 0.60

# "直接改写用户说法"的门槛（比"敢用"严得多）。
# 改动用户输入的字段名是有风险的动作 —— 改错了会去抽另一列，比不改更糟。
# 实测（真实 42 个列头 × 15 组说法）：
#   改对了的：入职时间→入职日期 0.92、设备名字→设备名称 0.98、材料名称→物料名称 0.91
#   改错了的：库存数量→期初库存 0.75（正解期末库存 0.72）、完成度→计划完成 0.71
#            （正解进度 0.71）、岗位工资→岗位 0.91（正解岗位津贴 0.73）
# 0.85 这条线能把上面前两类分开；"岗位工资"靠第二条"领先第 2 名 ≥0.10"挡掉。
SEM_REWRITE = 0.85
MAX_LEN = 64                 # 字段名很短，64 个 token 足够；再长只是浪费算力

# 中文里"看起来不同但意思相同"的尾巴：单位、括号说明
_UNIT_RE = re.compile(r"[（(][^）)]{0,12}[)）]\s*$")
_PUNCT_RE = re.compile(r"[\s\u3000·・,，。、;；:：!！?？\"'“”‘’()（）\[\]【】]")


# =================================================================
# 文本归一
# =================================================================

def norm(text: str) -> str:
    """归一化字段名：去空白标点、全角转半角、统一大小写。

    不做"去单位"—— 单位是有信息的（「处理时长（天）」和「处理时长（小时）」
    是两回事）。单位只在字面比较时作为弱信号处理，见 lexical()。
    """
    if not text:
        return ""
    t = str(text)
    # 全角字母数字 → 半角（ＨＰ → HP）
    t = "".join(chr(ord(c) - 0xFEE0) if 0xFF01 <= ord(c) <= 0xFF5E else c for c in t)
    t = _PUNCT_RE.sub("", t)
    return t.strip().lower()


def core_name(text: str) -> str:
    """去掉括号里的单位/说明：「单价（元）」→「单价」。给字面比较用。"""
    return _UNIT_RE.sub("", str(text or "")).strip()


def _bigrams(s: str):
    s = norm(s)
    return [s[i:i + 2] for i in range(len(s) - 1)] if len(s) >= 2 else ([s] if s else [])


# =================================================================
# 字面相似（模型不可用时的兜底，也是"完全相等"的快路径）
# =================================================================

def lexical(query: str, cand: str) -> float:
    """字面相似度 0~1。四种情况各自给分：

    1.00  完全相等（归一后）
    0.92  一方包含另一方（且短的一方 ≥2 字）—— 「合同金额」⊂「合同金额合计」
    0~1   两字片段 Dice 系数 —— 有共同字但顺序/长度不同
    0.85  去单位后相等 —— 「单价（元）」vs「单价」

    为什么用 Dice 而不是 Jaccard：短词里共同片段的信息量更大，
    「合同金额」vs「合同总金额」（3/4 片段相同）应该比 Jaccard 给的 0.6 更高。
    """
    a, b = norm(query), norm(cand)
    if not a or not b:
        return 0.0
    if a == b:
        return 1.0
    ca, cb = norm(core_name(query)), norm(core_name(cand))
    if ca and cb and ca == cb:
        return 0.85
    short, long_ = (a, b) if len(a) <= len(b) else (b, a)
    if len(short) >= 2 and short in long_:
        # 越短的一方占比越高，越可信：「金额」⊂「合同金额」不如「合同金额」⊂「合同金额合计」
        return 0.78 + 0.14 * (len(short) / len(long_))
    ba, bb = _bigrams(a), _bigrams(b)
    sa, sb = set(ba), set(bb)
    if not sa or not sb:
        return 0.0
    inter = len(sa & sb)
    if not inter:
        return 0.0
    return 2.0 * inter / (len(sa) + len(sb))


# =================================================================
# 向量底座（本地 ONNX，懒加载）
# =================================================================

class Embedder:
    """中文句向量。bge-small-zh-v1.5：24M 参数，512 维，CPU 上 300 条/秒左右。

    输出做 L2 归一化 → 余弦相似度就是点积，批量算极快（矩阵乘法）。
    """

    def __init__(self, model_dir: Path = None, threads: int = None):
        import numpy as np
        import onnxruntime as ort
        from tokenizers import Tokenizer

        d = Path(model_dir or MODEL_DIR)
        model_path = d / "model.onnx"
        tok_path = d / "tokenizer.json"
        if not model_path.exists() or not tok_path.exists():
            raise FileNotFoundError("语义模型缺失：%s" % d)

        so = ort.SessionOptions()
        so.graph_optimization_level = ort.GraphOptimizationLevel.ORT_ENABLE_ALL
        so.intra_op_num_threads = int(threads or max(1, (os.cpu_count() or 4) // 2))
        so.inter_op_num_threads = 1
        self.sess = ort.InferenceSession(str(model_path), so,
                                         providers=["CPUExecutionProvider"])
        self.tok = Tokenizer.from_file(str(tok_path))
        self.tok.enable_truncation(max_length=MAX_LEN)
        self.tok.enable_padding(pad_id=0, pad_token="[PAD]")

        self._in_names = {i.name for i in self.sess.get_inputs()}
        self._out_names = [o.name for o in self.sess.get_outputs()]
        self._np = np
        self.dim = None

    def encode(self, texts, batch_size: int = 32):
        """一批文本 → (n, dim) 的 float32 单位向量。"""
        np = self._np
        texts = [str(t or "")[:200] for t in texts]
        if not texts:
            return np.zeros((0, self.dim or 512), dtype=np.float32)

        out = []
        for i in range(0, len(texts), batch_size):
            chunk = texts[i:i + batch_size]
            encs = self.tok.encode_batch(chunk)
            ids = np.array([e.ids for e in encs], dtype=np.int64)
            mask = np.array([e.attention_mask for e in encs], dtype=np.int64)
            feed = {"input_ids": ids, "attention_mask": mask}
            if "token_type_ids" in self._in_names:
                feed["token_type_ids"] = np.zeros_like(ids)

            res = self.sess.run(None, feed)
            # bge 系列用 CLS 池化（不是 mean pooling）：取第 0 个 token。
            # 兼容两种导出图：3 维 (batch, seq, hidden) 取 CLS；
            # 已经 pool 过的 2 维 (batch, hidden) 直接用。
            hidden = res[0]
            vec = (hidden[:, 0, :] if hidden.ndim == 3 else hidden).astype(np.float32)
            norms = np.linalg.norm(vec, axis=1, keepdims=True)
            norms[norms == 0] = 1.0
            out.append(vec / norms)

        m = np.vstack(out)
        self.dim = m.shape[1]
        return m


_EMB = {"obj": None, "err": None, "tried": False}
_EMB_LOCK = threading.Lock()


def get_embedder():
    """取全局向量器。返回 (Embedder 或 None, 错误信息)。

    只尝试一次：模型不存在时不要每个请求都去 stat 一次磁盘（也别反复打日志）。
    """
    with _EMB_LOCK:
        if not _EMB["tried"]:
            _EMB["tried"] = True
            try:
                _EMB["obj"] = Embedder()
            except Exception as e:
                _EMB["err"] = "%s: %s" % (type(e).__name__, str(e)[:160])
        return _EMB["obj"], _EMB["err"]


def status() -> dict:
    """给前端/接口看的自检信息（不加载模型也能回答）。"""
    emb, err = get_embedder()
    dim = emb.dim if emb else None
    if dim is None:
        # 还没编码过任何文本时 dim 是空的；从 config.json 读 hidden_size 补上，
        # 免得接口返回 null 让人以为模型没加载成功。
        try:
            dim = json.loads((MODEL_DIR / "config.json").read_text(encoding="utf-8")).get("hidden_size")
        except Exception:
            dim = None
    return {
        "available": emb is not None,
        "model": str(MODEL_DIR),
        "model_size_mb": round(
            sum(p.stat().st_size for p in MODEL_DIR.glob("*") if p.is_file()) / 1024 / 1024, 1
        ) if MODEL_DIR.exists() else 0,
        "dim": dim,
        "error": err,
        "thresholds": {"strong": SEM_STRONG, "weak": SEM_WEAK, "rewrite": SEM_REWRITE},
    }


# =================================================================
# 候选索引（带磁盘缓存）
# =================================================================

def _fingerprint(terms) -> str:
    h = hashlib.sha1()
    for t in terms:
        h.update(str(t).encode("utf-8"))
        h.update(b"\x00")
    return h.hexdigest()[:16]


class CandidateIndex:
    """一批候选说法的向量索引。

    两种建法：
        CandidateIndex(terms)            现场编码
        CandidateIndex.from_cache(...)   从 npz 读（同指纹）
    """

    def __init__(self, terms, vectors=None):
        import numpy as np
        self.np = np
        self.terms = [str(t) for t in terms]
        self._lock = threading.Lock()
        if vectors is None:
            emb, err = get_embedder()
            if emb is None:
                self.vectors = None
                self.error = err
            else:
                t0 = time.time()
                self.vectors = emb.encode(self.terms)
                self.encode_sec = round(time.time() - t0, 3)
                self.error = None
        else:
            self.vectors = vectors
            self.error = None

    # ---- 磁盘缓存 ----
    def save(self, tag: str = "idx"):
        if self.vectors is None:
            return None
        try:
            CACHE_DIR.mkdir(parents=True, exist_ok=True)
            p = CACHE_DIR / ("%s-%s.npz" % (tag, _fingerprint(self.terms)))
            self.np.savez_compressed(p, vectors=self.vectors,
                                     terms=self.np.array(self.terms, dtype=object))
            return p
        except Exception:
            return None

    @classmethod
    def load(cls, tag: str, terms):
        """指纹一致才用缓存，否则返回 None（候选变了必须重编码）。"""
        try:
            import numpy as np
            p = CACHE_DIR / ("%s-%s.npz" % (tag, _fingerprint(terms)))
            if not p.exists():
                return None
            z = np.load(p, allow_pickle=True)
            saved = [str(x) for x in z["terms"].tolist()]
            if saved != [str(t) for t in terms]:
                return None
            return cls(terms, vectors=z["vectors"])
        except Exception:
            return None

    def score(self, query_vec):
        """查询向量 → 每个候选的余弦分（已归一化，直接点积）。"""
        if self.vectors is None or query_vec is None:
            return None
        return self.vectors @ query_vec


_INDEX_CACHE = {}
_INDEX_LOCK = threading.Lock()
_INDEX_KEEP = 4


def build_index(terms, tag: str = "fields", persist: bool = True) -> CandidateIndex:
    """按候选清单取索引：内存 → 磁盘 → 现场编码，三级命中。

    内存缓存用「指纹」做键而不是 id(list)：调用方每次都会新拼一个 list，
    用 id 永远命中不了。
    """
    key = (tag, _fingerprint(terms))
    with _INDEX_LOCK:
        hit = _INDEX_CACHE.get(key)
    if hit is not None:
        return hit

    idx = CandidateIndex.load(tag, terms) if persist else None
    if idx is None:
        idx = CandidateIndex(terms)
        if persist:
            idx.save(tag)

    with _INDEX_LOCK:
        if len(_INDEX_CACHE) >= _INDEX_KEEP:
            _INDEX_CACHE.pop(next(iter(_INDEX_CACHE)))
        _INDEX_CACHE[key] = idx
    return idx


# =================================================================
# 匹配：三路融合
# =================================================================

def _alias_lookup(query: str, aliases: dict) -> str:
    """别名表命中的目标词（支持大小写/空格差异）。"""
    if not aliases:
        return ""
    a = aliases if not isinstance(aliases, dict) else {
        str(k).strip(): str(v).strip()
        for k, v in aliases.items()
        if not str(k).startswith("_") and isinstance(v, str)
    }
    q = str(query).strip()
    if q in a:
        return a[q]
    qn = norm(q)
    for k, v in a.items():
        if norm(k) == qn:
            return v
    return ""


def match(query: str, candidates, topk: int = 3, aliases: dict = None,
          min_score: float = None, tag: str = "fields", use_index: bool = True):
    """把 query 对到候选清单里最像的几个说法。

    返回 [{"term", "score", "sem", "lex", "via", "note"}, ...]，按分数降序。

    score 是「最终分」，三路取最好的那一路，并留一点加权（语义与字面都对得上
    的候选，给一点点加成，压过只对上一路的）。

    via 说明这一条靠什么命中（给人看的，也是验收依据）：
        alias     别名表（手工确认）
        exact     完全相同
        lexical   字面包含 / 片段重合
        semantic  语义向量
    """
    cq = str(query or "").strip()
    if not cq:
        return []
    cands = [str(c) for c in candidates if str(c).strip()]
    if not cands:
        return []
    if min_score is None:
        min_score = SEM_WEAK

    # ① 别名表：命中就直接把它提到第一位（但候选里没有它也照样返回，
    #    因为调用方可能拿它去别的文档里找 —— 别名是"说法"，不要求已存在）
    alias_hit = _alias_lookup(cq, aliases)

    # ② 字面分（快，不用模型）
    lex = [lexical(cq, c) for c in cands]

    # ③ 语义分
    sem = [0.0] * len(cands)
    idx = None
    if use_index:
        idx = build_index(cands, tag=tag)
        if idx.vectors is not None:
            emb, _err = get_embedder()
            if emb is not None:
                qv = emb.encode([cq])[0]
                s = idx.score(qv)
                if s is not None:
                    sem = [float(x) for x in s]
    else:
        emb, _err = get_embedder()
        if emb is not None:
            vs = emb.encode(cands + [cq])
            sem = [float(vs[i] @ vs[-1]) for i in range(len(cands))]

    out = []
    for i, c in enumerate(cands):
        sc = max(lex[i], sem[i])
        # 两路都说像 → 小加成（避免"语义擦边"压过"字面完全对"）
        if lex[i] >= 0.5 and sem[i] >= SEM_WEAK:
            sc = min(1.0, sc + 0.03)
        via = "semantic"
        if lex[i] >= 0.999:
            via = "exact"
        elif lex[i] >= sc - 1e-6:
            via = "lexical"
        out.append({"term": c, "score": round(sc, 4), "sem": round(sem[i], 4),
                    "lex": round(lex[i], 4), "via": via})

    if alias_hit:
        for o in out:
            if norm(o["term"]) == norm(alias_hit):
                o.update(score=1.0, via="alias",
                         note="别名表：%s → %s" % (cq, alias_hit))
                break
        else:
            out.insert(0, {"term": alias_hit, "score": 1.0, "sem": None, "lex": None,
                           "via": "alias", "note": "别名表：%s → %s（候选里没有，按别名去抽）"
                           % (cq, alias_hit)})

    out.sort(key=lambda x: (-x["score"], x["term"]))
    out = [o for o in out if o["score"] >= min_score or o["via"] == "alias"]
    return out[:topk]


def resolve(query: str, candidates, aliases: dict = None, min_score: float = None,
            tag: str = "fields"):
    """只要最佳答案；不够像就返回 None（诚实留空，不硬猜）。"""
    res = match(query, candidates, topk=1, aliases=aliases,
                min_score=min_score if min_score is not None else SEM_STRONG, tag=tag)
    return res[0] if res else None


# =================================================================
# 字段空间：把一个行业里的说法自动聚成同义簇
# =================================================================

class FieldSpace:
    """从一堆说法里长出「这个行业的字段空间」。

    这是"全行业通用"的落点：没有任何预置的实体类型/字段清单，
    行业数据进来 → 候选说法进来 → 向量一算 → 同义说法自动抱团。
    每一簇选一个"标准名"（最短、最干净的那个），其余是它的说法。

        fs = FieldSpace(["合同金额", "合同总价", "总价（元）", "供应商名称", ...])
        fs.clusters()  -> [{"name": "合同金额", "members": [...]}, ...]
        fs.canonical("合同总价") -> "合同金额"
    """

    def __init__(self, terms, threshold: float = 0.86):
        self.terms = [str(t) for t in dict.fromkeys(terms) if str(t).strip()]
        self.threshold = threshold
        self.idx = build_index(self.terms, tag="fieldspace")
        self._clusters = None

    def clusters(self):
        if self._clusters is not None:
            return self._clusters
        np = self.idx.np
        if self.idx.vectors is None:
            # 没有模型：每个词自成一簇（至少不崩，界面仍能列字段）
            self._clusters = [{"name": t, "members": [t]} for t in self.terms]
            return self._clusters

        V = self.idx.vectors
        sim = V @ V.T
        n = len(self.terms)
        assigned = [-1] * n
        clusters = []
        for i in range(n):
            if assigned[i] != -1:
                continue
            members = [j for j in range(n) if assigned[j] == -1 and sim[i][j] >= self.threshold]
            for j in members:
                assigned[j] = len(clusters)
            clusters.append(members)

        out = []
        for members in clusters:
            names = [self.terms[j] for j in members]
            # 标准名：最短的优先（通常最干净）；同长度取出现顺序靠前的
            canon = sorted(names, key=lambda s: (len(core_name(s)), len(s)))[0]
            out.append({"name": canon, "members": names,
                        "cohesion": round(float(np.mean([sim[a][b] for a in members
                                                         for b in members if a != b])), 3)
                        if len(members) > 1 else 1.0})
        out.sort(key=lambda c: (-len(c["members"]), c["name"]))
        self._clusters = out
        return out

    def canonical(self, term: str) -> str:
        """这个词属于哪一簇，标准名叫什么（找不到就返回原词）。"""
        t = norm(term)
        for c in self.clusters():
            if any(norm(m) == t for m in c["members"]):
                return c["name"]
        # 不在候选里：拿它去匹配一簇，够像就归进去
        best = resolve(term, [c["name"] for c in self.clusters()],
                       min_score=self.threshold, tag="fieldspace")
        return best["term"] if best else term


# =================================================================
# 自检
# =================================================================

def _selftest():
    lines = []
    st = status()
    lines.append("模型：%s（%.1f MB，维度 %s）" % (st["model"], st["model_size_mb"], st["dim"]))
    lines.append("可用：%s%s" % (st["available"], ("" if st["available"] else "  原因：" + str(st["error"]))))
    lines.append("")

    cands = ["合同总金额", "供应商名称", "付款方式", "设备型号", "签订日期",
             "项目名称", "预算金额", "联系人", "验收结论"]
    pairs = [
        ("合同总价", "合同总金额"), ("合同金额", "合同总金额"), ("合同款", "合同总金额"),
        ("供货方", "供应商名称"), ("厂商", "供应商名称"), ("结算方式", "付款方式"),
        ("规格型号", "设备型号"), ("签约时间", "签订日期"), ("项目名", "项目名称"),
        ("采购金额", "预算金额"),
    ]
    ok1 = ok3 = 0
    for q, gold in pairs:
        res = match(q, cands, topk=3)
        top1 = res[0]["term"] if res else "（无）"
        hit1 = top1 == gold
        hit3 = any(r["term"] == gold for r in res)
        ok1 += hit1
        ok3 += hit3
        lines.append("%-8s → %-8s %s  %s" % (
            q, top1, "✓" if hit1 else "✗",
            "  ".join("%s(%.2f/%s)" % (r["term"], r["score"], r["via"]) for r in res[:3])))
    lines.append("")
    lines.append("Top1 %d/%d   Top3 %d/%d" % (ok1, len(pairs), ok3, len(pairs)))

    # 反例：源里真没有的字段，不能硬对
    lines.append("")
    for q in ("是否紧缺", "库存周转天数", "员工满意度"):
        r = resolve(q, cands)
        lines.append("不该硬认：%-10s → %s" % (q, "（留空）" if r is None else r["term"]))

    txt = "\n".join(lines)
    CACHE_DIR.mkdir(parents=True, exist_ok=True)
    (CACHE_DIR / "_selftest.txt").write_text(txt, encoding="utf-8")
    print(txt)


if __name__ == "__main__":
    _selftest()
