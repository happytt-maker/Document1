# -*- coding: utf-8 -*-
"""
run_all.py —— 一键全跑：7 列抽取 + 自动评测

为什么有这个脚本：
  以前：改一次 COL_NAME → F5 → 等结果 → 再改 → 再跑…… 一共 7 轮，慢，而且容易改错地方。
  现在：一次 F5 全跑完，而且 7 列是【并行】调用大模型的 ——
        总耗时 ≈ 最慢的那一列，而不是 7 列相加。

7 列从哪来：
  从 template_reader/模板表结构.json 里读「模板4」的列名 ——
  第③层（模板理解）的产物直接喂给第④层（证据匹配）。
  将来比赛时换任何一张陌生模板表，这里会自动跟上列名（前提是先跑过 template_reader）。

跑法（VSCode 打开 matcher 文件夹）：
  F5 选「跑全部列（并行）+ 评测」
  或终端：
    python run_all.py            ← 并行（默认，推荐）
    python run_all.py --serial   ← 一列一列跑（慢；用于对比"优化前后"）

输出：
  - 7 个「抽取结果-*.json」（和单列跑一模一样）
  - 屏幕：逐列耗时表 + 总耗时（赛题红线：90 秒/份文档）
  - 自动运行 evaluate.py，生成「评测报告-模板4.txt」
"""
import json
import os
import subprocess
import sys
import time
from collections import Counter
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path

import anchor
import direct
import extract
import history
import rowscope

HERE = Path(__file__).resolve().parent
BASE = HERE.parent
# 模板表结构：默认项目内固定位置，可用环境变量指向别处
# （与 writer/fill.py、template_reader/main.py 同一套约定 —
#  真实数据/泛化测试用独立结构文件，不污染 5 张正式模板的结构）
STRUCT_JSON = Path(os.environ.get(
    "A23_STRUCTURE_JSON", str(BASE / "template_reader" / "模板表结构.json")))

# 万一读不到模板表结构，用这份兜底（防呆）
FALLBACK_COLS = ["投诉编号", "客户名称", "投诉日期", "投诉内容摘要",
                 "责任判定", "处理时长（天）", "处理结果"]


def _pick_sheet():
    """按「指定模板 → 模板4 → 第一张」的优先级挑出目标工作表结构。

    load_columns（只要列名）和 load_column_infos（要别名、斜线拆分）共用这一套挑法，
    免得两处逻辑走偏 —— 一旦走偏，直通会对着一张表的列名去另一张表里找。
    """
    want = os.environ.get("A23_TEMPLATE_NAME", "").strip()
    try:
        data = json.loads(STRUCT_JSON.read_text(encoding="utf-8"))
        if want:
            for t in data:
                if t.get("file") == want and t.get("sheets"):
                    return t["sheets"][0]
            print("（A23_TEMPLATE_NAME=%s 不在结构里，退回默认逻辑）" % want)
        for t in data:
            if "模板4" in t.get("file", ""):
                return t["sheets"][0]
        for t in data:                      # 兜底：没有模板4，就用第一张
            if t.get("sheets"):
                return t["sheets"][0]
    except Exception as e:
        print("（读模板表结构失败：%s，改用内置列名）" % e)
    return None


def load_columns():
    """从模板表结构.json 读出目标模板的列名（按表格里的顺序）。

    「序号」是结构列（写回时按记录数自动编号 1..N），不拿去抽。
    """
    sheet = _pick_sheet()
    if not sheet:
        return list(FALLBACK_COLS)
    return [c["title"] for c in sheet["columns"]
            if c.get("title") and c.get("title") != "序号"]


def load_column_infos():
    """整条列描述（title + aliases + slash_parts）。

    直通要按别名找源表同名列（模板写「国家」、源表叫「国家/地区」），
    只拿标题是不够的 —— 这是 2026-11-03 COVID 数据暴露的。
    """
    sheet = _pick_sheet()
    if not sheet:
        return [{"title": c, "aliases": [], "slash_parts": []} for c in FALLBACK_COLS]
    return [c for c in sheet["columns"]
            if c.get("title") and c.get("title") != "序号"]


def load_blocks():
    """模板的「块」清单：docx 的多张表 / xlsx 的多个 sheet，一块就是一张要填的表。

    为什么要这个（2026-11-07，山东空气质量模板）：
      那个 docx 里有 3 张结构相同的表（德州 / 潍坊 / 临沂），每张表前面一句描述
      说明了它要哪个城市、哪个时刻。旧代码只拿 sheets[0] 当"这张模板"，
      于是只有第 1 张表有数，第 2、3 张永远空着 —— 而且第 1 张还是错的。

    返回 [{"no":1,"name":"表格1","caption":"…","n_rows":24}, …]，按块序。
    """
    want = os.environ.get("A23_TEMPLATE_NAME", "").strip()
    try:
        data = json.loads(STRUCT_JSON.read_text(encoding="utf-8"))
    except Exception:
        return []
    target = None
    for t in data:
        if t.get("file") == want:
            target = t
            break
    if target is None:
        for t in data:
            if "模板4" in t.get("file", ""):
                target = t
                break
    if target is None:
        return []

    out = []
    for i, sh in enumerate(target.get("sheets") or [], start=1):
        segs = sh.get("segments") or []
        n = sum(int(s.get("row_count") or 0) for s in segs)
        if not n:
            n = int((sh.get("fill_region") or {}).get("row_count") or 0)
        out.append({"no": i, "name": sh.get("name") or ("表格%d" % i),
                    "caption": sh.get("caption") or "", "n_rows": n})
    return out


_ANCHOR = {}          # {列名: 源文件名 or None}；None = 这一列放开到全池（由 main 填）
_WEAK = set()         # 定源判定"源内字面完全对不上"的列 → 抽取时启用键值段落兜底
_HIST = {}            # {列名: {"file","head",...}} 历史填写参照（history.recall 填）
# 抽取结果写到哪：默认 matcher/；A23_MATCHER_DIR 指到别处时跟随（与 fill.py 同一约定）。
# 逐列路径也必须用它 —— 否则"产物隔离"只管住了整表路径，
# 逐列跑会把结果写回 matcher/，既污染默认目录，又让 fill.py 在目标目录找不到文件。
RES_DIR = Path(os.environ.get("A23_MATCHER_DIR") or HERE)

# 列级并发上限（0 = 不限）。main 里由 --workers 设定；
# 重审/交叉重审这些内部并行也要受它约束，不然一个任务仍会瞬间打出 N 个请求。
_WORKERS = 0


def _par(n):
    """给定"本来要开 n 个线程"，算出实际开几个（受 --workers 限制，至少 1）。"""
    return max(1, min(n, _WORKERS) if _WORKERS > 0 else n)


# 池多大算"大池"（MB）。实测：日常 18 份文档的池是 0.3MB / 702 块；
# 压测的 16 份大文档是 194.8MB / 50 万块。中间几乎没有样本，
# 所以门槛放在 30MB（是日常池的 100 倍），只会命中"真的很大"的情况。
BIG_POOL_MB = float(os.environ.get("A23_POOL_BIG_MB") or 30)


def _auto_workers():
    """按池的大小自动决定"一次跑几列"。

    为什么需要（2026-11-03 实测）：
      小池 702 块，三张模板并发跑，各自"每列一个线程"，全部 20~42 秒，没问题；
      大池 50 万块，三张模板并发，8 列的模板5 被拖到 126 秒（红线 90 秒）。
    原因不是锁，是**同一台机器的核和内存带宽被 24 个线程抢**：
    500k 块的候选打分是内存密集型的，线程越多越快不起来，反而互相拖。
    所以降到 4 列并发 —— 三张模板 × 4 = 12 线程，落在 16 核以内。

    小池返回 0（不限），保持历史行为 —— 快的东西不该被人为压慢。
    """
    try:
        mb = extract.OUTPUT_JSON.stat().st_size / 1024.0 / 1024.0
    except OSError:
        return 0
    if mb >= BIG_POOL_MB:
        return int(os.environ.get("A23_BIG_POOL_WORKERS") or 4)
    if mb >= 8:
        return 6
    return 0


def run_one(col):
    """跑一列（静默），返回 (列名, 结果, 墙钟秒数)。

    初抽不启用拆词兜底（allow_split=False）：拆词兜底放到一致性重审阶段
    在主源内做 —— 全量实测，在"全文档池"里拆词，模型会被同名干扰拖到
    几十秒还只抽一半；在主源内拆词又快又全（2026-10-11 模板5 实测）。

    only_file 来自定源计划（anchor.py，2026-11-02）：非空时这一列只在那份文档
    里检索。这是"同名列跨文档混抽"的根治手段 —— 让模型根本看不到干扰文档。

    hist 来自历史填写参照（history.py，2026-11-02）：同一张模板上次是从哪个
    列头抽的。定源说"这份文档"，历史说"这一列"——粒度更细，所以两级可以叠加。
    """
    t0 = time.time()
    res = extract.run_column(col, quiet=True, only_file=_ANCHOR.get(col),
                             allow_split=False, weak=(col in _WEAK),
                             hist=_HIST.get(col), out_dir=str(RES_DIR))
    return col, res, time.time() - t0


def _write_col_json(col, res):
    """把某列的结果写回它的 json 文件（交叉重审"不采纳"时，恢复原结果用）。"""
    f = RES_DIR / ("抽取结果-%s.json" % col)
    f.write_text(json.dumps(res, ensure_ascii=False, indent=2), encoding="utf-8")


def _run_with_retry(col, only_file=None, allow_split=True):
    """跑一列重审；结果为"空"时自动再试一次。

    2026-10-11 实测："空"有随机性 —— 模板5「库存数量」（目标与源列头"期初/期末库存"
    只是近似）同题两次跑出 0 条 / 10 条。第二次常能出来，两次都空才认账。
    """
    res = extract.run_column(col, quiet=True, only_file=only_file, allow_split=allow_split,
                             out_dir=str(RES_DIR))
    if res is not None and not (res.get("items") or []):
        res2 = extract.run_column(col, quiet=True, only_file=only_file, allow_split=allow_split,
                                  out_dir=str(RES_DIR))
        if res2 is not None and (res2.get("items") or []):
            return res2
    return res


def reaudit(rows):
    """一致性重审（2026-10-11 A 类修复之二）——跨列来源一致性检查。

    背景（五表连测实测的 4 个失分案例）：
      多个文档含同名列时，"这列到底该信哪份文档"单靠字面分数分不出来 ——
      比如「负责人」在会议纪要和项目周报里都出现，模型抄了周报；
      「备注」在库存盘点表里有整整一列，模型把盘点表的备注填进了会议任务表。
    做法（两步）：
      ① 投票：跑完第一轮后，统计每列的值来自哪份文档，出现列数最多的 = 主源
      ② 重审：来源与主源不一致的"少数派"列，重新抽一次 —— 这次只在那份主源文档里检索
    为什么敢这么做：一张待填表的各列，来源大概率集中在一两份文档里；
      少数派列要么是被"同名诱饵"带偏了（如周报的负责人），要么是跨文档污染（如库存表的备注）。
    保守规则：重审失败（模型调用错）→ 保留原结果；
      重审返回空 → 采纳空（"源里没有就该留空"，不许编造）。

    平票怎么办（2026-10-11 模板1 实测：合同×2 列 vs 销售台账×2 列）：
      "设备采购汇总表"的「数量/单价」被销售台账的同名列骗走（销售台账里
      也有一模一样的设备型号和价格，是干扰数据），投票成了 2:2 平票 ——
      这时不做"跳过"，改做**交叉重审**：把平的几方互相试 ——
        "你来源的列，放到我这边重抽；抽得出（非空）就改判归我，抽不出就保持不动"。
      实测：数量/单价能在合同里抽出完整 4 条 → 改判归合同；设备名称/规格型号
      在销售台账里抽不出 → 保持合同不动。这样平票就不再卡死。

    直接在 rows 上原地替换，返回要打印/存档的报告行。
    """
    lines = []

    # ① 每列的主要来源
    col_src = {}
    for col, res, sec in rows:
        if not res:
            continue
        files = [it.get("file") for it in res.get("items", []) if it.get("file")]
        if not files:
            continue
        col_src[col] = Counter(files).most_common(1)[0][0]

    if not col_src:
        lines.append("（没有可统计的列，跳过重审）")
        return lines, {}

    votes = Counter(col_src.values())
    ranked = votes.most_common()
    lines.append("来源分布：" + "；".join("《%s》×%d 列" % (f, n) for f, n in ranked))

    top_file, top_n = ranked[0]
    ties = [f for f, n in ranked if n == top_n]
    if len(ties) > 1:
        # ---- 平票 → 交叉重审（先判"胜方"，再整批归位）----
        # 2026-10-11 模板1 实测的坑：单列独立改判会拆散整表 —— 规格型号在销售台账里
        # 恰好"抽得出"（产品型号列），被错误改判；正确做法是先判哪一方是"胜方"：
        # 把平的几方互相试（你来源的列，放我这边重抽），谁能承接对方的列多，谁胜；
        # 然后"来自胜方的列保持、来自败方的列统一重审到胜方"，保证整表来源不拆家。
        def _num_nonempty(col, s_src):
            for c2, s2, res2 in results:
                if c2 == col and s2 == s_src:
                    return 1 if (res2 is not None and res2.get("items")) else 0
            return 0

        lines.append("！主源平票（%s）→ 交叉重审" % "、".join(ties))
        cross_cols = [(c, s) for c, s in col_src.items() if s in ties]
        tasks = []
        for s_src in ties:
            for c, c_src in cross_cols:
                if c_src != s_src:
                    tasks.append((c, s_src, c_src))
        if not tasks:
            lines.append("（没有可交叉重审的列）")
            return lines, col_src

        def _cross(col, s_src):
            # 探针也不启用拆词兜底：实测教训 —— 拆词会把"产品型号"当成
            # "规格型号"的证据（泛化召回），做出错误的承接判定
            return col, s_src, _run_with_retry(col, only_file=s_src, allow_split=False)

        results = []
        with ThreadPoolExecutor(max_workers=_par(len(tasks))) as pool:
            futs = [pool.submit(_cross, c, s) for c, s, _ in tasks]
            for f in as_completed(futs):
                results.append(f.result())

        # 承接数：S 能把对方多少列抽出非空 → 谁多谁胜
        def _score(s_src):
            return sum(_num_nonempty(c, s_src) for c, c_src in cross_cols if c_src != s_src)

        ranked_t = sorted(ties, key=lambda s: -_score(s))
        for s_src in ranked_t:
            lines.append("  · 承接测试：《%s》能承接 %d 列" % (s_src, _score(s_src)))

        # 只有"一方明显更强"（严格大于第二名）才改判；打不出胜负就全部保持原样
        if len(ranked_t) >= 2 and _score(ranked_t[0]) > _score(ranked_t[1]) and _score(ranked_t[0]) > 0:
            winner = ranked_t[0]
            lines.append("  → 判《%s》胜" % winner)
        else:
            winner = None
            lines.append("  → 打不出胜负，全部保持原样")

        for col, old_src in cross_cols:
            old_res = None
            for i, (c, r, sec) in enumerate(rows):
                if c == col:
                    old_res = r
                    if winner is None or old_src == winner:
                        # 归胜方（或没胜方）：保持原结果；把 json 写回（探针可能覆盖过）
                        if old_res is not None:
                            _write_col_json(col, old_res)
                        lines.append("· %s：保持《%s》（%d 条）"
                                     % (col, old_src, len((old_res or {}).get("items", []))))
                    else:
                        # 来自败方 → 用"在胜方内重审"的结果改判（非空才换；空则保持）
                        hit = None
                        for c2, s2, res2 in results:
                            if c2 == col and s2 == winner and res2 is not None and res2.get("items"):
                                hit = res2
                                break
                        if hit is not None:
                            n_before = len((old_res or {}).get("items", []))
                            rows[i] = (col, hit, sec)
                            lines.append("· %s：%s → 《%s》 ｜ %d 条 → %d 条（改判）"
                                         % (col, "《%s》" % old_src, winner,
                                            n_before, len(hit["items"])))
                        else:
                            if old_res is not None:
                                _write_col_json(col, old_res)
                            lines.append("· %s：在《%s》里重抽为空 → 保持《%s》（%d 条）"
                                         % (col, winner, old_src,
                                            len((old_res or {}).get("items", []))))
                    break

        # 2026-10-11 全量连测收尾（模板1「供应商」实测的漏洞）：
        # 原来这里直接 return —— 漏掉了"非平票来源的少数派列"。供应商列来自
        # 《供应商联系方式.txt》，既不在胜/败双方里，也没机会进下面的 minority 流程。
        # 现在判出胜方后不提前收工：把主源换成胜方、同步 col_src，
        # 继续往下走"少数派重审"，让第三方来源的列也统一在胜方里过一遍。
        if winner is None:
            return lines, col_src
        for col, old_src in cross_cols:
            for i, (c, r, sec) in enumerate(rows):
                if c == col:
                    files = [it.get("file") for it in (r or {}).get("items", [])
                             if it.get("file")]
                    if files:
                        col_src[col] = Counter(files).most_common(1)[0][0]
                    break
        top_file = winner
        top_n = len([c for c, s in col_src.items() if s == top_file])

    lines.append("主源 = 《%s》（%d 列来自它）" % (top_file, top_n))

    minority = [(c, s) for c, s in col_src.items() if s != top_file]
    # 空结果列也补一刀（2026-10-11 实测：模板5「库存数量」初抽为空的根因，
    #   是盘点的"期初/期末库存"列被合同、销售台账里同叫"数量"的列挤出候选 ——
    #   空列不参与上面的投票，但不代表源里没有；放到主源里重抽一次往往就出来了。
    #   "源里真的没有"的列（如「是否紧缺」）重抽后依然为空 —— 保持不变，这是对的）
    empty_cols = [c for c, res, sec in rows
                  if res is None or not (res.get("items") or [])]
    # 拆词兜底列强制复检（2026-10-11 全量实测收的尾）：
    #   拆词兜底抽出来的列质量不稳（同列两次跑：5 条 / 10 条），
    #   且"来源=主源、又不是空列"时两道网都兜不住 —— 一律在主源内再抽一遍。
    split_cols = [c for c, res, sec in rows
                  if res and res.get("split_fallback") and (res.get("items") or [])]
    seen = {c for c, _ in minority}
    todo = list(minority)
    new_empties = [c for c in empty_cols if c not in seen]
    if new_empties:
        lines.append("空结果列补抽：%s" % "、".join(new_empties))
    todo += [(c, "（空）") for c in new_empties]
    seen |= set(new_empties)
    new_splits = [c for c in split_cols if c not in seen]
    if new_splits:
        lines.append("拆词兜底列复检：%s" % "、".join(new_splits))
    todo += [(c, "（拆词兜底）") for c in new_splits]
    if not todo:
        lines.append("各列来源一致 ✓ 无需重审")
        return lines, col_src

    # ② 待重审列：限定主源文档，重抽（空则自动重试一次）
    def _one(col):
        # 定源时被"放开"的列（主源里字面分 <20，属源里没有的那类），补抽也别锁回主源 ——
        # 否则会把它按主源的语境重抽一遍，白白丢掉全池里本来能找到的线索
        return col, _run_with_retry(col, only_file=_ANCHOR.get(col, top_file))

    results = {}
    with ThreadPoolExecutor(max_workers=_par(len(todo))) as pool:
        futs = [pool.submit(_one, c) for c, _ in todo]
        for f in as_completed(futs):
            col, res = f.result()
            results[col] = res

    # ③ 替换结果（写盘由 run_column 自己覆盖同名 json，这里同步 rows 供展示/摘要）
    for col, old_src in todo:
        new_res = results.get(col)
        for i, (c, r, sec) in enumerate(rows):
            if c != col:
                continue
            n_before = len((r or {}).get("items", []))
            if new_res is None:
                lines.append("· %s：重审失败 → 保留原结果（%s，%d 条）"
                             % (col, old_src, n_before))
                break
            n_after = len(new_res.get("items", []))
            new_files = Counter(it.get("file") for it in new_res.get("items", [])
                                if it.get("file"))
            new_src = ("《%s》" % new_files.most_common(1)[0][0]) if new_files else "（空）"
            old_label = old_src if old_src.startswith("（") else "《%s》" % old_src
            rows[i] = (c, new_res, sec)
            lines.append("· %s：%s → %s ｜ %d 条 → %d 条"
                         % (col, old_label, new_src, n_before, n_after))
            break
    return lines, col_src


def _arg_int(name, default=0):
    """读 "--xxx N" / "--xxx=N" 形式的整数参数；没给就返回默认值。

    为什么要自己解析：这个脚本一直只用 `"--serial" in sys.argv` 这种零参判断，
    没有 argparse。为了不惊动现有调用方（pipeline / 各处验收脚本），
    新参数只在命令行显式出现时才生效，默认行为一个字都不变。
    """
    for i, a in enumerate(sys.argv):
        if a == name and i + 1 < len(sys.argv):
            try:
                return int(sys.argv[i + 1])
            except ValueError:
                return default
        if a.startswith(name + "="):
            try:
                return int(a.split("=", 1)[1])
            except ValueError:
                return default
    return default


def main():
    serial = "--serial" in sys.argv
    # 列级并发上限（2026-11-03）：0 = 不限（每列一个线程，历史行为）。
    # 为什么要它：多任务同时跑时，如果每个任务都按"每列一个线程"开，
    # 5 列 × 4 个任务 = 20 个并发请求，免费额度的模型档直接 429。
    workers = _arg_int("--workers", 0)
    workers_src = "--workers"
    if workers <= 0:
        # 没显式指定 → 按池的大小自动定（大池降到 4 列并发，见 _auto_workers）
        workers = _auto_workers()
        workers_src = "按池大小自动" if workers else "小池不限制"
    cols = load_columns()
    n_par = min(len(cols), workers) if workers > 0 else len(cols)
    global _ANCHOR, _WEAK, _HIST, _WORKERS
    _WORKERS = workers          # 让重审/交叉重审的内部并行也受同一个上限约束

    print("=" * 62)
    tpl_tag = os.environ.get("A23_TEMPLATE_NAME", "").strip()
    print("一键全跑 · 共 %d 列 ｜ 模式：%s%s"
          % (len(cols),
             "一列一列（顺序）" if serial else
             ("并行（最多 %d 列同时%s）" % (n_par, "" if workers_src == "--workers"
                                           else " · " + workers_src)),
             (" ｜ 模板：" + tpl_tag) if tpl_tag else ""))
    print("=" * 62)

    # ---------- 结构化直通（2026-11-03 新增；2026-11-07 改为默认先试）----------
    # 什么时候走：**列名能一一对上就走**，不再要求"用户写了行条件"。
    #   以前的门槛是 `if spec`（必须有日期区间才直通），结果是：
    #   山东空气质量那次要求里只有"表一：监测时间…城市：德州市"，没有区间
    #   → 判定"没有条件"→ 走大模型 → 模型从 2 万行里随手给了 14 行枣庄/东营的数据，
    #     84 秒、三张表只填了一张。整列搬运本来就不该过模型。
    # 对不上（列名缺、源不是表格）仍回落大模型，而且日志写明原因。
    # A23_NO_DIRECT=1 可以强制关掉，用于 A/B 对照。
    res_dir = Path(os.environ.get("A23_MATCHER_DIR") or HERE)
    direct_mode = False
    direct_result = {}
    direct_lines = []
    req_text = (os.environ.get("A23_USER_REQ") or "").strip()
    spec = rowscope.parse_requirement(req_text) if req_text else None
    blocks = load_blocks()
    if req_text:
        print("【用户要求】" + req_text.replace("\n", " / ")[:300])
    if blocks:
        print("【模板分块】共 %d 块：%s" % (
            len(blocks),
            " ｜ ".join("%s(%d 行)" % (b["name"], b["n_rows"]) for b in blocks)))
    for b in blocks:
        _c, _src = rowscope.block_conditions(req_text, b["no"], b.get("caption") or "")
        if _c:
            print("【行条件】%s ← %s：%s" % (b["name"], _src,
                                            rowscope.describe_conditions(_c)))
    if os.environ.get("A23_NO_DIRECT") != "1":
        d = direct.run_blocks(load_column_infos(), blocks, req_text, res_dir)
        direct_lines = list(d["lines"])
        if d["ok"]:
            direct_mode = True
            # 直通产物已经落盘，按列名读回来，走和模型路径完全一样的下游
            for c in cols:
                f = res_dir / ("抽取结果-%s.json" % c)
                res = json.loads(f.read_text(encoding="utf-8")) if f.exists() else None
                d.setdefault("rows", []).append((c, res, 0.0))
            direct_result = d
        print()
    else:
        print("直通：已按 A23_NO_DIRECT=1 关闭")

    # ---------- 定源（2026-11-02）：先决定每列从哪份文档抽 ----------
    # 放在调用大模型之前，纯本地打分（不花钱），5 张模板实测 < 1 秒。
    # --no-anchor 可退回"整表全池"的旧行为，用于 A/B 回归对比。
    anchor_lines = []
    if direct_mode:
        # 直通已经按列名确定了唯一来源，定源这一步没东西可定
        anchor_lines = ["直通：主源已按列名锁定 → 《%s》工作表「%s」"
                        % (direct_result.get("file"), direct_result.get("sheet"))]
        for ln in anchor_lines:
            print(ln)
    elif "--no-anchor" in sys.argv:
        print("定源：已关闭（--no-anchor），整表走全文档池")
        anchor_lines = ["定源：已关闭（--no-anchor）"]
    else:
        t_a0 = time.time()
        ap = anchor.plan(extract.load_pool(), cols)
        for ln in ap["lines"]:
            print(ln)
        _ANCHOR = ap["per_col"]
        _WEAK.clear()
        _WEAK.update(ap.get("weak") or [])
        anchor_lines = ap["lines"]
        print("定源用时 %.2f 秒" % (time.time() - t_a0))

    # ---------- 历史填写参照（2026-11-02）：填新表前先查同类表怎么填的 ----------
    # 和定源是两级粒度：定源定"哪份文档"，历史定"哪一列"（精确到列头）。
    # 纯本地读存档，不花钱；没历史就是空，行为和以前完全一样。
    history_lines = []
    if "--no-history" in sys.argv:
        print("历史参照：已关闭（--no-history）")
        history_lines = ["历史参照：已关闭（--no-history）"]
    else:
        hc = history.recall(tpl_tag, cols)
        if hc:
            for ln in hc["lines"]:
                print(ln)
            _HIST.clear()
            _HIST.update(hc["cols"])
            history_lines = list(hc["lines"]) + ["命中方式：%s ｜ 历史指纹 %s"
                                                 % (hc["hit"], hc["fp"][:8])]
            print("历史参照：%d/%d 列有历史映射（%s）"
                  % (len(hc["cols"]), len(cols), hc["from"]))
        else:
            print("历史参照：没有同类模板的填写记录，本次按定源 + 语义走")
    print()

    if direct_mode:
        print("（直通：本次不调用模型）")
    else:
        print("正在调用大模型（%s）…… 每列约 1~4 秒，请稍等。" % extract.MODEL)
    print()

    t0 = time.time()
    rows = list(direct_result.get("rows") or []) if direct_mode else []

    # ---------- 叙述体：整表一次抽取（2026-11-02） ----------
    # 判据和 extract.run_column 里那条完全一致：**全池没有一个块带列头**。
    # 为什么必须在这里分路：叙述体逐列抽，等于把同一份原文完整读 N 遍 ——
    # 真实数据实测（1.1 万字报告 → 5 列 × 100 行）：单列 56.9 秒，5 列 285 秒，
    # 赛题红线 90 秒/文档，差 3 倍多。改成一次抽整表（一段 = 一行）后 44 秒，
    # 且答案不变（正则独立抽一遍对账，400 个数值格 0 差异）。
    # --no-table-once 可强制退回逐列，用于 A/B 对比。
    pool_blocks = [] if direct_mode else extract.load_pool()
    narrative_mode = bool(pool_blocks) and not any(
        (b.get("_head") or "").strip() for b in pool_blocks)
    # table_once 只表示"这一次真的走了整表一次抽取"。
    # 叙述体 + --no-table-once 时 narrative_mode 仍为 True 但实际是逐列，
    # 两者必须分开记，否则显示/跳过重审/算 token 全都会错。
    table_once = False
    if narrative_mode and "--no-table-once" not in sys.argv:
        print("走法：叙述体文档（全池 %d 块、无一处列头）→ 整表一次抽取"
              % len(pool_blocks))
        print("      原文只读一遍，一段成一行，%d 列共用一次调用" % len(cols))
        print()
        t_one = time.time()
        # 结果写到哪：默认 matcher/；A23_MATCHER_DIR 指到别处时跟随（与 fill.py 同一约定）
        res_dir = Path(os.environ.get("A23_MATCHER_DIR") or HERE)
        one = extract.run_table(cols, quiet=True, out_dir=str(res_dir))
        wall = time.time() - t_one
        if one is None:
            print("整表抽取失败 —— 退回逐列模式")
            narrative_mode = False
        else:
            table_once = True
            for c in cols:
                f = res_dir / ("抽取结果-%s.json" % c)
                res = json.loads(f.read_text(encoding="utf-8")) if f.exists() else None
                rows.append((c, res, wall))
            print("整表抽取完成：%d 行 × %d 列，用时 %.2f 秒"
                  % (one["records"], len(cols), wall))
            print()

    if not rows and serial:
        for c in cols:
            rows.append(run_one(c))
    elif not rows:
        with ThreadPoolExecutor(max_workers=n_par) as pool:
            futs = [pool.submit(run_one, c) for c in cols]
            for f in as_completed(futs):
                rows.append(f.result())

    # 按模板列的原始顺序展示
    order = {c: i for i, c in enumerate(cols)}
    rows.sort(key=lambda r: order.get(r[0], 99))

    # ---------- 一致性重审（跨列来源一致性；重审耗时也算进系统时间） ----------
    t_re0 = time.time()
    if direct_mode:
        # 直通按行号整列搬，每一列的来源和行序都一样，跨列不一致这件事不可能发生
        re_lines = ["已跳过：结构化直通（表对表整列搬，各列同源同序，不存在跨列不一致）"]
        re_sec = 0
    elif table_once:
        # 整表抽取本来就"一段一行、所有列同源"，跨列一致性天然成立，
        # 再跑一遍只会拿不同判据去改已经对齐的结果。
        re_lines = ["已跳过：整表抽取天然同源（一段一行，所有列来自同一段落）"]
        re_sec = 0
    else:
        re_lines, _ = reaudit(rows)
    re_sec = round(time.time() - t_re0, 2)

    wall = time.time() - t0

    print("-" * 62)
    sum_sec = 0.0
    sum_tok = 0
    n_ok = 0
    table = []            # 同时攒一份文本，等下写进「运行摘要.txt」
    for col, res, sec in rows:
        if res is None:
            line = "%s ｜ %.2f 秒 ｜ 失败" % (col, sec)
        else:
            n_items = len(res.get("items", []))
            tok = (res.get("usage") or {}).get("total_tokens", 0)
            sum_sec += sec
            sum_tok += tok
            n_ok += 1
            line = "%s ｜ %.2f 秒 ｜ %d 条 ｜ %d tokens" % (col, sec, n_items, tok)
        print(line)
        table.append(line)

    if direct_mode:
        # 直通的耗时是整批一次搬完的，加总没有意义；也不存在"压缩比"（本来就没调模型）
        total_line = ("结构化直通：整列搬运 %.2f 秒（%d 列，按行号对齐，0 次模型调用）"
                      % (direct_result.get("elapsed") or 0.0, len(cols)))
    elif table_once:
        # 整表模式：各列的 elapsed_sec 是同一次调用的耗时，相加毫无意义
        total_line = "整表一次抽取：%.2f 秒（%d 列共用这一次调用，不是相加）" % (wall, len(cols))
    else:
        total_line = "列耗时合计（≈顺序跑要花的总时间）：%.2f 秒" % sum_sec
    wall_line = "墙钟总耗时（本次实际）：%.2f 秒（赛题红线：90 秒/份文档）" % wall

    # ---------- 语义匹配的对照（2026-10-26） ----------
    # 哪些列不是靠字面命中、而是靠语义向量找回来的 —— 这是"换说法也能填"的证据
    sem_lines = []
    sem_rows = [(c, (r or {}).get("semantic_hints")) for c, r, _s in rows
                if r and r.get("semantic_hints")]
    if sem_rows:
        sem_lines.append("【字段语义匹配】以下列在源文档里换了说法（向量匹配，非别名表）：")
        for c, hs in sem_rows:
            sem_lines.append("  %s → %s" % (c, "、".join("「%s」" % h for h in hs)))

    print("-" * 62)
    for ln in sem_lines:
        print(ln)
    print("-" * 62)
    print(total_line)
    print(wall_line)
    speed_line = ""
    if not serial and wall > 0 and not table_once and not direct_mode:
        speed_line = "压缩比：%.1f 倍（%.2f 秒 → %.2f 秒）" % (sum_sec / wall, sum_sec, wall)
        print(speed_line)
    # 整表模式：5 份结果里记的是同一次调用的 usage，相加会把 token 数放成 5 倍
    if table_once and rows and rows[0][1]:
        sum_tok = (rows[0][1].get("usage") or {}).get("total_tokens", 0)
    tok_line = "tokens 合计：%d" % sum_tok
    ok_line = "成功：%d/%d 列" % (n_ok, len(cols))
    print(tok_line)
    print(ok_line)

    # ---------- 一致性重审报告 ----------
    print()
    print("【一致性重审】用时 %.2f 秒" % re_sec)
    for ln in re_lines:
        print("  " + ln)

    # ---------- 摘要落盘（终端乱码也不怕，真身在这个文件里） ----------
    summary = ["运行摘要 · %s ｜ 模式：%s"
               % (time.strftime("%Y-%m-%d %H:%M"),
                  "结构化直通（表对表整列搬）" if direct_mode else
                  ("整表一次（叙述体）" if table_once else ("顺序" if serial else "并行"))), ""]
    if direct_mode:
        summary += ["【直通】" + ln for ln in direct_lines] + [""]
    if table_once:
        summary += ["【走法】全池无列头 → 叙述体整表抽取：原文读 1 遍，一段成一行",
                    "        逐列抽会把同一份原文读 %d 遍（实测 285 秒，超 90 秒红线）" % len(cols),
                    ""]
    if anchor_lines:
        summary += ["【定源】"] + ["  " + ln for ln in anchor_lines] + [""]
    if history_lines:
        summary += ["【历史参照】"] + ["  " + ln for ln in history_lines] + [""]
    summary += table
    summary += ["", total_line, wall_line]
    if speed_line:
        summary.append(speed_line)
    summary += [tok_line, ok_line]
    if sem_lines:
        summary += [""] + sem_lines
    summary += ["", "【一致性重审（%.2f 秒）】" % re_sec]
    summary += ["  " + ln for ln in re_lines]

    # ---------- 历史参照：本列实际用上了没有 ----------
    # recall 只是"知道以前怎么填的"，真正用上是 extract 找到那一列之后。
    # 分开报，是为了不被"查到了历史"糊弄成"用了历史"。
    hist_hit = [(c, (r or {}).get("history")) for c, r, _s in rows
                if r and r.get("history")]
    if hist_hit:
        hl = ["【历史参照】本次实际生效 %d 列（锚定到上次填对的那一列）：" % len(hist_hit)]
        for c, h in hist_hit:
            hl.append("  %s ← 《%s》「%s」%d 块%s"
                      % (c, h.get("file"), h.get("head"), h.get("n", 0),
                         "（跨文件）" if h.get("cross_file") else ""))
        for ln in hl:
            print(ln)
        summary += [""] + hl
    elif history_lines:
        summary += ["", "【历史参照】" + history_lines[0]]

    # ---------- 写历史存档（填完就记，下次这张表直接可用）----------
    if "--no-history" in sys.argv:
        pass
    elif "--no-save" in sys.argv:
        print("（--no-save：本次不留历史记录）")
    else:
        try:
            hrec = history.record(tpl_tag or "(未命名模板)", cols,
                                  [(c, r) for c, r, _s in rows])
            n = len(hrec.get("cols") or {})
            print("【历史填写】已存档：%d 列映射 / %d 行 ｜ 库镜像：%s"
                  % (n, hrec.get("n_rows", 0),
                     "已写入 a23_agent" if hrec.get("db_ok") else "跳过（MySQL 未连上，本地存档已写）"))
            summary += ["", "【历史填写】%d 列映射已存档（%s）"
                        % (n, "MySQL 镜像成功" if hrec.get("db_ok") else "仅本地存档")]
        except Exception as e:
            print("（历史记录写入失败：%s —— 不影响本次结果）" % e)

    (HERE / "运行摘要.txt").write_text("\n".join(summary) + "\n", encoding="utf-8")
    print("（摘要已存：运行摘要.txt）")

    # ---------- 自动评测 ----------
    if "--no-eval" in sys.argv:
        print("（--no-eval：跳过自动评测 —— 由上层脚本自己对分）")
        return
    print()
    print(">>> 自动运行评测（evaluate.py）……")
    print()
    subprocess.run([sys.executable, str(HERE / "evaluate.py")])


if __name__ == "__main__":
    main()
