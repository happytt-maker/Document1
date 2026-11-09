# -*- coding: utf-8 -*-
"""
anchor.py —— 定源：抽取前先决定"这一列该从哪份文档抽"。

【为什么要有这个模块（2026-11-02）】
  reaudit() 是**事后投票纠偏**：跑完看哪份文档贡献的列最多，再把"少数派"列拉回去重抽。
  它能救"一列被带偏"，救不了"一列整体都来自错文档"，也拦不住"模型自己把两份文档的
  同名列都抽出来" —— 2026-10-28 模板3「负责人」抽出 9 条，
  会议纪要 4 条 +「项目总览」C3-C7 5 条，两种来源抢同一批槽位，按行号对齐后整列错乱。

  这不是提示词能治的：模型没有错，是**任务定义**错了 —— 我们没告诉它"源只有一份"。
  所以正解是在抽取前就把池子收窄到源文档（extract.run_column 的 only_file 参数本来就有，
  之前只给 reaudit 用）。

【信号怎么选（全部本地、零配置、不调模型）】
  候选信号在这一天都量过（收尾对照/评测/定源-诊断.txt，5 张模板 × 18 份文档）：

    cover(d)  该文档里「列头与目标列名完全相等」的列数  ← 最强
              真源文档的表头就是我们待填表的列名；干扰文档只能撞上一两个同名列。
              实测：模板3 会议纪要 4/5 vs 进度跟踪表 1/5（后者正是污染源）
    lit(d)    Σ min(该列在该文档内的字面最高分, 100)      ← 补 cover 的漏
              模板4 的源是 txt（没有表头，cover=0），靠字面分 7 列全中才选得出来。
    score(d)  = 60 * cover(d) + lit(d)

  实测排名（正确源 / 次名）：
    模板1 合同 400 / 销售台账 200 ｜ 模板2 员工表 800 / 其他 100
    模板3 会议纪要 400 / 进度跟踪 100 ｜ 模板4 txt 700 / 销售台账 100
    模板5 盘点表 600 / 技术参数 160

【护栏：宁可不定源，也不能定错】
  定错的代价是整表报废，比不定源（= 今天的行为）严重得多。所以：
    ① 整表闸：主源要"明确领先"才锁 ——
       · 有表头的源：列头精确覆盖 ≥ MIN_COVER（2 列）且分数 ≥ 次名 × MARGIN（1.5）
       · 无表头的源（txt/段落）：cover 必为 0，改要求"字面认出 ≥60% 的列"且同样领先
       两条都不满足就整表走全池 —— 今天的行为，不会更差。
    ② 逐列闸：源内字面分 < MIN_SCORE（20）的列**仍然锁在源内**，只标进 weak。

       这里走过弯路，记下来（2026-11-02 第一版踩的）：
       第一版的做法是"源里没有的列放开到全池，让它自己去找"。
       实测立刻踩雷 —— 模板3「备注」在会议纪要里 0 分，一放开就去抽
       《库存盘点表》那一整列备注（★40，列头精确命中），直接错填。
       这正是 2026-10-11 已经踩过的坑：干扰文档里"恰好同名的列"。

       想通了一件事：一列在主源里找不到，只有两种可能 ——
         (a) 源里真没有 → 该诚实留空，放开=错填
         (b) 源里换了说法（模板1「供应商」在合同里叫「乙方（供货方）」）→
             要靠**语义**在源内找，放开到全池只会捡到别的文档的同名列
       两种情况都不该放开。所以定源只干一件事：**收窄池子**。
       换了说法的列，靠 run_column 里本来就默认开的语义召回，在收窄后的源内解决 ——
       池子干净了，语义的排名也更准（候选少、干扰少）。
"""
from collections import Counter

from retrieve import score_blocks, strip_unit

HEAD_WEIGHT = 60      # 列头精确覆盖一列值多少分
CAP = 100             # 单列字面分封顶，防止一个"标签命中 90 + 精确列头 60 = 150"的块
                      # 把整份文档的分数抬到失真
MIN_COVER = 2         # 主源至少精确覆盖 2 列
MARGIN = 1.5          # 主源得分要领先次名 1.5 倍
MIN_SCORE = 20        # 与 extract.MIN_SCORE 一致（retrieve 的筛选线）
MIN_COLS = 2          # 模板只有 1 列时定源没有意义（分不出主次），直接不定
WEAK_RATIO = 0.6      # 无表头源：字面要认出 60% 以上的列才算"真源"


def doc_scores(blocks, cols):
    """算每份文档对这张模板的承接分。返回 [(file, score, cover, lit, per_col), ...] 降序。

    per_col: {列名: 该列在这份文档内的字面最高分} —— 逐列闸要用。
    """
    by_file = {}
    for b in blocks:
        by_file.setdefault(b.get("file") or "?", []).append(b)

    out = []
    for f, bs in by_file.items():
        ctx = [b.get("_ctx") or "" for b in bs]     # 预建过的直接复用，免重建
        cover = 0
        lit = 0
        per_col = {}
        for col in cols:
            core = strip_unit(col)
            scored = score_blocks(bs, core, min_score=0, contexts=ctx,
                                  use_semantic=False)
            best = scored[0][0] if scored else 0
            per_col[col] = best
            lit += min(best, CAP)
            # 精确列头命中：块上带 _colheads_norm（build_column_contexts 预建）
            if any(b.get("_colheads_norm") and core in b["_colheads_norm"] for b in bs):
                cover += 1
        out.append((f, HEAD_WEIGHT * cover + lit, cover, lit, per_col))
    out.sort(key=lambda t: -t[1])
    return out


def plan(blocks, cols, min_cover=MIN_COVER, margin=MARGIN):
    """给出定源计划。

    返回 dict：
      file      主源文件名，None = 不定源（整表走全池）
      scores    doc_scores 的原样结果（给报告/排查用）
      per_col   {列名: 文件名 or None}；None = 这一列走全池（只有"不定源"时才出现）
      weak      源内字面分 < MIN_SCORE 的列（源里可能没有，或换了说法 ——
                抽取阶段会靠语义在源内找；找不到就诚实留空）
      lines     给人看的报告行
      locked    bool，整表是否真的锁了源
    """
    lines = []
    res = {"file": None, "scores": [], "per_col": {c: None for c in cols},
           "lines": lines, "locked": False, "weak": []}

    if len(cols) < MIN_COLS:
        lines.append("定源：模板列数 < %d，不做定源（分不出主次）" % MIN_COLS)
        return res

    scored = doc_scores(blocks, cols)
    res["scores"] = scored
    if not scored:
        lines.append("定源：池子里没有文档")
        return res

    top_f, top_s, top_c, top_l, top_per = scored[0]
    second = scored[1][1] if len(scored) > 1 else 0
    lines.append("定源打分（列头精确覆盖×%d + 字面分）：" % HEAD_WEIGHT)
    for f, s, c, l, _p in scored[:4]:
        lines.append("  %-30s %4d 分 ｜ 覆盖 %d/%d 列 ｜ 字面合计 %d"
                     % (f[:28], s, c, len(cols), l))
    if len(scored) > 4:
        lines.append("  ……（另 %d 份更低）" % (len(scored) - 4))

    # 闸①：整表闸
    # 两条路都要求"领先次名 MARGIN 倍"，区别只在凭什么领先：
    #   A 有表头的源（xlsx / docx 表格）：靠列头精确覆盖 —— 最强，也最不容易误判
    #   B 无表头的源（txt / 纯段落）：cover 一定是 0，只能靠字面分。但要求它
    #     "字面能认出 60% 以上的列"，否则就是某份文档恰好撞上两列高分而已。
    #     模板4 的源就是 .txt（客户投诉处理记录），7/7 列全中、字面合计 700 分封顶，
    #     只按 cover 判会把它挡掉，整表退回全池 —— 那就等于这个功能没生效。
    lit_full = CAP * len(cols)
    margin_ok = (not second) or (top_s >= second * margin)
    if top_c >= min_cover and margin_ok:
        why = "列头覆盖 %d 列" % top_c
    elif top_c < min_cover and margin_ok and top_l >= WEAK_RATIO * lit_full:
        why = "无表头源，字面认出 %d/%d 列（%d/%d 分）" % (
            len([1 for c in cols if top_per.get(c, 0) >= MIN_SCORE]), len(cols),
            top_l, lit_full)
    else:
        if top_c < min_cover:
            lines.append("→ 不定源：首选《%s》列头只覆盖 %d 列（门槛 %d），"
                         "字面也只认出 %d/%d 分，整表走全池"
                         % (top_f, top_c, min_cover, top_l, lit_full))
        else:
            lines.append("→ 不定源：首选 %d 分 vs 次名 %d 分，没拉开 %.1f 倍，整表走全池"
                         % (top_s, second, margin))
        return res

    res["file"] = top_f
    res["locked"] = True
    lines.append("→ 主源 = 《%s》（%d 分，次名 %d 分；%s）" % (top_f, top_s, second, why))

    # 闸②：逐列闸 —— 源内明显没有的列**也锁在源内**，只标记为 weak。
    # （为什么不放开：见模块开头的"走过弯路"。模板3「备注」一放开就抽到库存盘点表的备注）
    for col in cols:
        res["per_col"][col] = top_f
    weak = [c for c in cols if top_per.get(c, 0) < MIN_SCORE]
    res["weak"] = weak
    if weak:
        lines.append("  · 源内字面认不出（靠语义在源内找，找不到就留空）：%s"
                     % "、".join(weak))
    return res
