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

import extract

HERE = Path(__file__).resolve().parent
BASE = HERE.parent
STRUCT_JSON = BASE / "template_reader" / "模板表结构.json"

# 万一读不到模板表结构，用这份兜底（防呆）
FALLBACK_COLS = ["投诉编号", "客户名称", "投诉日期", "投诉内容摘要",
                 "责任判定", "处理时长（天）", "处理结果"]


def load_columns():
    """从模板表结构.json 读出目标模板的列名（按表格里的顺序）。

    找法（由精确到兜底）：
      0. 环境变量 A23_TEMPLATE_NAME 指定了模板文件名 → 精确找它
         （五表连测 / 一次跑多张模板时用；不设变量时行为完全不变）
      1. 没指定（或没找到）→ 优先找「模板4」（老流程不变）
      2. 还找不到 → 取结构里的第一张表

    过滤：「序号」是结构列（写回时按记录数自动编号 1..N），不拿去抽。
    """
    def cols_of(t):
        sheet = t["sheets"][0]
        out = []
        for c in sheet["columns"]:
            ti = c.get("title")
            if ti and ti != "序号":
                out.append(ti)
        return out

    want = os.environ.get("A23_TEMPLATE_NAME", "").strip()
    try:
        data = json.loads(STRUCT_JSON.read_text(encoding="utf-8"))
        if want:
            for t in data:
                if t.get("file") == want and t.get("sheets"):
                    return cols_of(t)
            print("（A23_TEMPLATE_NAME=%s 不在结构里，退回默认逻辑）" % want)
        for t in data:
            if "模板4" in t.get("file", ""):
                return cols_of(t)
        for t in data:                      # 兜底：没有模板4，就用第一张
            if t.get("sheets"):
                return cols_of(t)
    except Exception as e:
        print("（读模板表结构失败：%s，改用内置列名）" % e)
    return list(FALLBACK_COLS)


def run_one(col):
    """跑一列（静默），返回 (列名, 结果, 墙钟秒数)。

    初抽不启用拆词兜底（allow_split=False）：拆词兜底放到一致性重审阶段
    在主源内做 —— 全量实测，在"全文档池"里拆词，模型会被同名干扰拖到
    几十秒还只抽一半；在主源内拆词又快又全（2026-10-11 模板5 实测）。
    """
    t0 = time.time()
    res = extract.run_column(col, quiet=True, allow_split=False)
    return col, res, time.time() - t0


def _write_col_json(col, res):
    """把某列的结果写回它的 json 文件（交叉重审"不采纳"时，恢复原结果用）。"""
    f = HERE / ("抽取结果-%s.json" % col)
    f.write_text(json.dumps(res, ensure_ascii=False, indent=2), encoding="utf-8")


def _run_with_retry(col, only_file=None, allow_split=True):
    """跑一列重审；结果为"空"时自动再试一次。

    2026-10-11 实测："空"有随机性 —— 模板5「库存数量」（目标与源列头"期初/期末库存"
    只是近似）同题两次跑出 0 条 / 10 条。第二次常能出来，两次都空才认账。
    """
    res = extract.run_column(col, quiet=True, only_file=only_file, allow_split=allow_split)
    if res is not None and not (res.get("items") or []):
        res2 = extract.run_column(col, quiet=True, only_file=only_file, allow_split=allow_split)
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
        with ThreadPoolExecutor(max_workers=max(1, len(tasks))) as pool:
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
        return col, _run_with_retry(col, only_file=top_file)

    results = {}
    with ThreadPoolExecutor(max_workers=max(1, len(todo))) as pool:
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


def main():
    serial = "--serial" in sys.argv
    cols = load_columns()

    print("=" * 62)
    tpl_tag = os.environ.get("A23_TEMPLATE_NAME", "").strip()
    print("一键全跑 · 共 %d 列 ｜ 模式：%s%s"
          % (len(cols), "一列一列（顺序）" if serial else "并行",
             (" ｜ 模板：" + tpl_tag) if tpl_tag else ""))
    print("=" * 62)
    print("正在调用大模型（%s）…… 每列约 1~4 秒，请稍等。" % extract.MODEL)
    print()

    t0 = time.time()
    rows = []
    if serial:
        for c in cols:
            rows.append(run_one(c))
    else:
        with ThreadPoolExecutor(max_workers=len(cols)) as pool:
            futs = [pool.submit(run_one, c) for c in cols]
            for f in as_completed(futs):
                rows.append(f.result())

    # 按模板列的原始顺序展示
    order = {c: i for i, c in enumerate(cols)}
    rows.sort(key=lambda r: order.get(r[0], 99))

    # ---------- 一致性重审（跨列来源一致性；重审耗时也算进系统时间） ----------
    t_re0 = time.time()
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

    total_line = "列耗时合计（≈顺序跑要花的总时间）：%.2f 秒" % sum_sec
    wall_line = "墙钟总耗时（本次实际）：%.2f 秒（赛题红线：90 秒/份文档）" % wall
    print("-" * 62)
    print(total_line)
    print(wall_line)
    speed_line = ""
    if not serial and wall > 0:
        speed_line = "压缩比：%.1f 倍（%.2f 秒 → %.2f 秒）" % (sum_sec / wall, sum_sec, wall)
        print(speed_line)
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
               % (time.strftime("%Y-%m-%d %H:%M"), "顺序" if serial else "并行"), ""]
    summary += table
    summary += ["", total_line, wall_line]
    if speed_line:
        summary.append(speed_line)
    summary += [tok_line, ok_line]
    summary += ["", "【一致性重审（%.2f 秒）】" % re_sec]
    summary += ["  " + ln for ln in re_lines]
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
