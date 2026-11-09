# -*- coding: utf-8 -*-
"""
history.py —— 历史填写参照（2026-11-02）

为什么有这个模块
----------------
组会 PPT 第 11 页写的三件事，前两件（语义匹配、表头自适应）已经做完，
第三件是"把库里历史填写的记录用起来，填新表前先查同类表是怎么填的"。

这件事的价值不在"抄上次的答案" —— 数据每期都会变，抄值必错。
真正能复用、而且复用了就一定更准的，是**映射关系**：

    「模板里的这一列」  ←来源←  「某份文档里某个列头/标签」

比如模板5「库存数量」上次是从《库存盘点表》的「期末库存」列抽的，
这个对应关系验证过一次就是资产：下次同一张表（哪怕换了一批数据文档），
直接拿"期末库存"当锚点去抽，既不用再猜，也不会再错抽成「期初库存」。

三级复用（由强到弱，越强越省事）
------------------------------
  ① 模板指纹一致（同一张模板，列名集合完全一样）
      → 每列的历史来源全部直接可用。这是最常见的场景：一张表每月填一次。
  ② 指纹不一致，但有历史模板的列名高度重合（≥70%）
      → 只复用"列名完全相同"的那些列。模板改版、加了一列也能用。
  ③ 都没有 → 返回空，退化成今天的行为（定源 + 语义），什么都不会更差。

存放位置（两份，互为备份）
------------------------
  · matcher/历史填写.json      主存档。纯本地，MySQL 挂了也不影响填表。
  · MySQL a23_agent 库         fill_templates / fill_columns 两张表（镜像）。
      为什么要写库：赛题模块（2）要求"提取结果入库存储"，历史填写也是同一类
      资产 —— 能按模板、按列名 SQL 查出来。MySQL 不可用时静默跳过，不阻断。

不进历史的东西
--------------
  · 空列（items=[]）：源里没有就是没有，记进去会让下次把"没找到"当成"锚点"。
  · 主观列（摘要/备注/是否紧缺）：每次都是新写的，没有可复用性。
  · 值本身不用于回填，只用来看"这一列历史上是什么形态"（数字/日期/文本），
    抽完对照一下 —— 历史上全是数字、这次抽出一堆文字，多半是抽错了列。
"""
import hashlib
import json
import re
import time
from pathlib import Path

HERE = Path(__file__).resolve().parent
STORE = HERE / "历史填写.json"
PROJECT = HERE.parent
AGENT_DIR = PROJECT / "agent"

KEEP_RUNS = 60          # 只留最近 60 次填写（一张表一年也就 12 次，够用）
SIMILAR_RATE = 0.7      # ② 级复用的列名重合度门槛
SIMILAR_MIN = 2         # 至少重合几列才认为"同类表"（1 列的偶合太多）

# 主观列：每期都重写，没有"上次怎么填"可参照
SUBJECTIVE = ("摘要", "概要", "简述", "概述", "说明", "描述", "备注", "意见",
              "是否", "结论", "评价", "原因")


# ================================================================
#  指纹与归一化
# ================================================================
def _norm(s: str) -> str:
    """列名归一：去空白/标点/全角括号，用来比"是不是同一列"。"""
    s = str(s or "")
    s = re.sub(r"[\s\u3000]", "", s)
    s = s.replace("（", "(").replace("）", ")").replace("：", ":").replace("，", ",")
    return s.strip().lower()


def fingerprint(cols) -> str:
    """模板指纹 = 列名（按表格顺序）归一后拼起来的 md5。

    为什么带顺序：同一张表的列顺序是稳定的；顺序变了说明是另一张表
    （或者改版了），这时降到 ② 级按列名匹配更安全。
    """
    key = "|".join(_norm(c) for c in cols)
    return hashlib.md5(key.encode("utf-8")).hexdigest()


def _is_subjective(col: str) -> bool:
    return any(k in (col or "") for k in SUBJECTIVE)


def _like_head(t):
    """判断一段文字像不像"列头"，像就返回它，不像返回 None。

    为什么要这一层：源表里 A 列顶上常压着文档大标题
    （「原材料库存盘点表（截至2026年3月31日）」），而同列"最上面那条"
    恰好就是它。不筛掉的话，历史里会记成「物料编码 ← 那张大标题」，
    下次拿它当锚点就会把标题行也一起捞进来。
    """
    t = str(t or "").strip()
    n = re.sub(r"[\s\u3000]", "", t)
    if not (2 <= len(n) <= 14):
        return None
    if n.endswith("表") or n.endswith("清单") or n.endswith("台账"):
        return None
    if any(k in n for k in ("截至", "年度", "月份", "公司", "集团", "明细")):
        return None
    if re.fullmatch(r"[\d.,%\-/年月日]+", n):        # 纯数字/日期
        return None
    return t


def _pick_head(items):
    """从一批 item 里定出这一列的历史来源列头。

    优先用 item["colheads"]（同列前 5 个非空文本）里第一个"像列头"的；
    退回到 item["head"]；都不像就不记（只留源文件，宁可少一个锚点，
    也不能记一个错的 —— 错的锚点会让下次把标题行当数据列）。
    """
    counts = {}
    for it in items:
        cand = None
        for h in (it.get("colheads") or []):
            cand = _like_head(h)
            if cand:
                break
        if not cand:
            cand = _like_head(it.get("head"))
        if cand:
            counts[cand] = counts.get(cand, 0) + 1
    return max(counts, key=counts.get) if counts else None


def value_kind(values):
    """猜这一列历史上是什么形态：number / date / percent / text / mixed。"""
    if not values:
        return "none"
    kinds = set()
    for v in values:
        s = str(v).strip()
        if not s:
            continue
        if re.fullmatch(r"-?\d+(\.\d+)?%", s):
            kinds.add("percent")
        elif re.fullmatch(r"-?\d+(\.\d+)?", s.replace(",", "")):
            kinds.add("number")
        elif re.fullmatch(r"\d{4}[-/年]\d{1,2}([-/月]\d{1,2}日?)?", s):
            kinds.add("date")
        else:
            kinds.add("text")
    if not kinds:
        return "none"
    if len(kinds) == 1:
        return kinds.pop()
    if kinds == {"number", "percent"}:
        return "number"
    return "mixed"


# ================================================================
#  读 / 写主存档
# ================================================================
def _load() -> dict:
    try:
        d = json.loads(STORE.read_text(encoding="utf-8"))
        if isinstance(d, dict) and isinstance(d.get("runs"), list):
            return d
    except Exception:
        pass
    return {"version": 1, "runs": []}


def _save(d: dict):
    d["runs"] = d["runs"][-KEEP_RUNS:]
    STORE.write_text(json.dumps(d, ensure_ascii=False, indent=1), encoding="utf-8")


# ================================================================
#  MySQL 镜像（尽力而为，失败静默）
# ================================================================
def _db():
    """拿连接；MySQL 没起来就返回 None（历史参照不因数据库挂掉而失效）。"""
    try:
        import sys
        if str(AGENT_DIR) not in sys.path:
            sys.path.insert(0, str(AGENT_DIR))
        from db import get_conn
        return get_conn(with_db=True)
    except Exception:
        return None


DDL = [
    """CREATE TABLE IF NOT EXISTS fill_templates (
        id          INT AUTO_INCREMENT PRIMARY KEY,
        fp          VARCHAR(64)  NOT NULL COMMENT '模板列名指纹',
        tpl_name    VARCHAR(255) NULL     COMMENT '模板文件名（会变，只作展示）',
        col_names   TEXT         NULL     COMMENT '列名 JSON（按表内顺序）',
        n_cols      INT          DEFAULT 0,
        n_rows      INT          DEFAULT 0,
        src_files   TEXT         NULL     COMMENT '本次用到的源文档 JSON',
        created_at  DATETIME     NOT NULL,
        KEY idx_fp (fp)
    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COMMENT='A23 历史填写·模板'""",
    """CREATE TABLE IF NOT EXISTS fill_columns (
        id          INT AUTO_INCREMENT PRIMARY KEY,
        fp          VARCHAR(64)  NOT NULL COMMENT '模板列名指纹',
        tpl_name    VARCHAR(255) NULL,
        col_name    VARCHAR(128) NOT NULL COMMENT '模板里的列名',
        src_file    VARCHAR(255) NULL     COMMENT '历史来源文档',
        src_head    VARCHAR(128) NULL     COMMENT '历史来源列头/标签（复用锚点）',
        n_items     INT          DEFAULT 0,
        kind        VARCHAR(16)  NULL     COMMENT '值形态：number/date/text/mixed',
        sample      TEXT         NULL     COMMENT '前 3 个值（示例，不回填）',
        created_at  DATETIME     NOT NULL,
        KEY idx_fp_col (fp, col_name(80))
    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COMMENT='A23 历史填写·列映射'""",
]


def mirror_to_db(rec: dict):
    """把一次填写记录镜像进 MySQL。任何异常都不外抛 —— 填表不能因为库挂了就失败。"""
    conn = _db()
    if conn is None:
        return False
    try:
        with conn.cursor() as cur:
            for ddl in DDL:
                cur.execute(ddl)
            cur.execute(
                "INSERT INTO fill_templates (fp,tpl_name,col_names,n_cols,n_rows,"
                "src_files,created_at) VALUES (%s,%s,%s,%s,%s,%s,%s)",
                (rec["fp"], rec["template"],
                 json.dumps(rec["columns"], ensure_ascii=False),
                 rec["n_cols"], rec["n_rows"],
                 json.dumps(rec["source_files"], ensure_ascii=False),
                 rec["ts"]))
            rows = []
            for col, c in rec["cols"].items():
                rows.append((rec["fp"], rec["template"], col, c.get("file"),
                             c.get("head"), c.get("n", 0), c.get("kind"),
                             json.dumps(c.get("sample") or [], ensure_ascii=False),
                             rec["ts"]))
            if rows:
                cur.executemany(
                    "INSERT INTO fill_columns (fp,tpl_name,col_name,src_file,"
                    "src_head,n_items,kind,sample,created_at) "
                    "VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s)", rows)
        conn.commit()
        return True
    except Exception:
        return False
    finally:
        try:
            conn.close()
        except Exception:
            pass


# ================================================================
#  写：一次填写跑完就记一笔
# ================================================================
def record(template_name: str, cols, results, min_items: int = 1) -> dict:
    """记录一次填写。results = [(列名, 抽取结果 dict or None), ...]

    只记有值的列（空列记进去会让下次把"没找到"当锚点）。
    返回写进存档的那条记录（给调用方打印用）。
    """
    cols = [c for c in cols if c]
    fp = fingerprint(cols)
    col_map = {}
    src_files = []
    n_rows = 0

    for col, res in results:
        if not res:
            continue
        items = res.get("items") or []
        if len(items) < min_items:
            continue
        if _is_subjective(col):
            continue
        vals = [str(it.get("value") or "").strip() for it in items]
        vals = [v for v in vals if v]
        if not vals:
            continue
        # 来源：列头取"最像列头"的那个（见 _pick_head），文件取出现最多的
        head = _pick_head(items)
        files = {}
        for it in items:
            f = (it.get("file") or "").strip()
            if f:
                files[f] = files.get(f, 0) + 1
        src = max(files, key=files.get) if files else None
        if src and src not in src_files:
            src_files.append(src)
        col_map[col] = {
            "file": src,
            "head": head,
            "n": len(vals),
            "kind": value_kind(vals),
            "sample": vals[:3],
        }
        n_rows = max(n_rows, len(vals))

    rec = {
        "ts": time.strftime("%Y-%m-%d %H:%M:%S"),
        "template": template_name or "",
        "fp": fp,
        "n_cols": len(cols),
        "n_rows": n_rows,
        "source_files": src_files,
        "columns": cols,
        "cols": col_map,
    }

    d = _load()
    # 同一张模板只留最近一次：映射是一样的，留多了只是把存档撑大、
    # 还会让"填过几次"这个信息淹没在一堆重复记录里。次数单独记。
    old = [r for r in d["runs"] if r.get("fp") == fp]
    rec["times"] = len(old) + 1
    d["runs"] = [r for r in d["runs"] if r.get("fp") != fp]
    d["runs"].append(rec)
    _save(d)

    rec["db_ok"] = mirror_to_db(rec)
    return rec


# ================================================================
#  读：填新表前先查同类表
# ================================================================
def recall(template_name: str, cols):
    """查历史：这张表（或同类表）上次是怎么填的。

    返回 None（没查到）或 dict：
      hit      "exact"（同一张模板）/ "similar"（列名高度重合）
      from     那次填写的时间
      template 那次填的模板名
      cols     {列名: {"file","head","kind","n","sample"}}（只含本次也有的列）
      lines    给人看的报告行
    """
    cols = [c for c in cols if c]
    if not cols:
        return None
    fp = fingerprint(cols)
    runs = _load()["runs"]
    if not runs:
        return None

    # ① 指纹一致：优先取最近一次
    exact = [r for r in runs if r.get("fp") == fp]
    if exact:
        r = exact[-1]
        return _mk(r, cols, "exact", r.get("times", 1))

    # ② 列名重合度
    want = {_norm(c) for c in cols}
    best, best_rate = None, 0.0
    for r in runs:
        got = {_norm(c) for c in (r.get("columns") or [])}
        if not got:
            continue
        inter = want & got
        rate = len(inter) / len(want)
        if rate > best_rate:
            best, best_rate = r, rate
    if best is not None and best_rate >= SIMILAR_RATE:
        shared = want & {_norm(c) for c in (best.get("columns") or [])}
        if len(shared) >= SIMILAR_MIN:
            return _mk(best, cols, "similar", 1, rate=best_rate)
    return None


def _mk(run: dict, cols, hit: str, times: int, rate: float = 1.0):
    """把一次历史记录裁剪成本次模板能用的映射。"""
    src = run.get("cols") or {}
    by_norm = {_norm(k): (k, v) for k, v in src.items()}
    out = {}
    for c in cols:
        kv = by_norm.get(_norm(c))
        if kv and (kv[1].get("head") or kv[1].get("file")):
            out[c] = kv[1]

    lines = []
    if not out:
        return None
    # 能真正当锚点的只有"有列头"的列（txt/段落来源没有列头，只能靠定源锁文件）
    anchor_n = sum(1 for v in out.values() if v.get("head"))
    if hit == "exact":
        lines.append("命中历史：同一张模板填过 %d 次，最近一次 %s（%d 列有映射，%d 列可锚定列头）"
                     % (times, run.get("ts"), len(out), anchor_n))
    else:
        lines.append("命中历史：与《%s》（%s）列名重合 %.0f%%，复用其中 %d 列（%d 列可锚定列头）"
                     % (run.get("template"), run.get("ts"), rate * 100, len(out), anchor_n))
    heads = []
    for c, v in list(out.items())[:8]:
        if v.get("head") and v["head"] != c:
            heads.append("%s←「%s」" % (c, v["head"]))
        elif v.get("head"):
            heads.append("%s←同名列" % c)
    if heads:
        lines.append("列映射：" + "、".join(heads) +
                     ("…" if len(out) > 8 else ""))
    return {"hit": hit, "from": run.get("ts"), "template": run.get("template"),
            "fp": run.get("fp"), "cols": out, "lines": lines}


# ================================================================
#  抽完之后的格式对照
# ================================================================
KIND_CN = {"number": "数字", "percent": "百分比", "date": "日期",
           "text": "文字", "mixed": "混合", "none": "空"}


def check_kind(col: str, res: dict, hist_col: dict):
    """把本次抽到的值和历史形态对一下，不一致就提醒（只提醒，不改值）。

    这是历史参照的第二层用处：不知道"该是什么样"的情况下，模型抽错列
    （把「期初库存」当「库存数量」）很难被程序发现；但历史上这一列从来是
    数字，这次抽出一串文字，就有个客观信号了。
    """
    if not res or not hist_col:
        return None
    want = hist_col.get("kind")
    if want not in ("number", "percent", "date"):
        return None
    vals = [str(it.get("value") or "").strip() for it in (res.get("items") or [])]
    vals = [v for v in vals if v]
    if not vals:
        return None
    got = value_kind(vals)
    if got == want or got == "mixed":
        return None
    if want in ("number", "percent") and got in ("number", "percent"):
        return None
    return "「%s」历史上是%s（如 %s），这次抽成了%s（如 %s）—— 可能抽错了列" % (
        col, KIND_CN.get(want, want), (hist_col.get("sample") or ["?"])[0],
        KIND_CN.get(got, got), vals[0])


# ================================================================
#  命令行：看历史、清历史
# ================================================================
def main():
    import sys
    d = _load()
    runs = d["runs"]
    if "--clear" in sys.argv:
        _save({"version": 1, "runs": []})
        print("已清空历史填写存档（MySQL 镜像不动）")
        return
    if not runs:
        print("还没有历史填写记录。跑一次 pipeline.py 就会自动记一笔。")
        return
    print("历史填写存档：%d 次（%s）" % (len(runs), STORE))
    print("=" * 66)
    for r in runs[-10:]:
        print("%s ｜ %s" % (r.get("ts"), r.get("template") or "(未命名)"))
        print("     %d 列 / %d 行 ｜ 源：%s"
              % (r.get("n_cols", 0), r.get("n_rows", 0),
                 "、".join(r.get("source_files") or []) or "-"))
        for c, v in (r.get("cols") or {}).items():
            print("       %-12s ← %s ｜ %s ｜ %d 条 ｜ %s"
                  % (c, v.get("head") or "(无列头)", v.get("file") or "-",
                     v.get("n", 0), KIND_CN.get(v.get("kind"), v.get("kind"))))


if __name__ == "__main__":
    main()
