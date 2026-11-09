# -*- coding: utf-8 -*-
"""
direct.py —— 结构化直通：表对表，整列搬，零模型调用
================================================================

什么时候用它（2026-11-03，COVID-19 真实数据新增）：

  源文档本身就是一张规整的表（xlsx / docx 表格），模板列名在源表里能找到同名列，
  那么"抽"这个动作没有信息量 —— 答案是确定的，让大模型去猜反而会错。

  实测过的两个坑：
    ① **数量级**：源表 15088 行 × 15 列，模板要 2542 行（筛完日期区间）。
       模型一次只肯给 14~15 行（它把表"总结"了一遍），而且给的是区间外的数据。
       这不是提示词能救的，是产品形态不对：整列搬运不该过模型。
    ② **口径**：走模型时它会把表头行当数据（「国家/地区」被填成了第一条国家）、
       病例数从 N15 起抽 —— 十几条里一半错位。直通按行号搬，天然不会错位。

  所以这里做的是"确定性优先"：能直通就直通，直通不了（列名对不上、
  源是叙述体）才回落到大模型。**这条判定是显式的，不偷偷切换。**

和 extract.py 的分工：
  extract.py  用列头做导航，让模型在海量块里"找"一列 —— 适合叙述体、半结构化
  direct.py   列名已经能一一对上，直接按行号整列搬 —— 适合表格对表格
两者产出的 JSON 结构一致，写回阶段（writer/fill.py）不用区分来源。
"""

import json
import re
import time
from collections import Counter
from pathlib import Path

import rowscope


# ---------------------------------------------------------------- 小工具
def _norm(s):
    """列名归一：去空白、去全角空格、小写。不做同义词替换（那是语义层的事）。"""
    return re.sub(r"[\s\u3000]+", "", str(s or "")).lower()


def _is_numberish(s):
    s = str(s or "").strip().replace(",", "")
    if not s:
        return False
    return re.fullmatch(r"-?\d+(?:\.\d+)?%?", s) is not None


def _col_letter(idx):
    """1 → A，27 → AA（不引 openpyxl，免得为了一个函数拖依赖）。"""
    out = ""
    while idx > 0:
        idx, r = divmod(idx - 1, 26)
        out = chr(65 + r) + out
    return out


# ---------------------------------------------------------------- 还原表
def scan_tables(pool):
    """把池里的 xlsx 单元格按 (文件, 工作表) 还原成「表头 + 数据格」。

    为什么不直接读原文件：读文档阶段已经把每个格子连同 locator（sheet/row/col）
    放进池子了，再读一遍文件等于把同一件事做两遍，而且两份可能不一致
    （池子才是这一趟流程的事实依据）。

    返回 { (file, sheet): {"header_row": n, "headers": {col: text},
                          "cells": {(row, col): (text, addr)}, "max_row": n} }
    """
    raw = {}
    for b in pool:
        loc = b.get("locator") or {}
        if loc.get("kind") != "cell":
            continue
        file, sheet = b.get("file"), loc.get("sheet")
        row, col = loc.get("row"), loc.get("col")
        if not (file and sheet and row and col):
            continue
        text = (b.get("text") or "").strip()
        if not text:
            continue
        addr = loc.get("cell") or ("%s%d" % (_col_letter(col), row))
        raw.setdefault((file, sheet), {})[(row, col)] = (text, addr)

    tables = {}
    for key, cells in raw.items():
        rows = sorted({r for r, _c in cells})
        # 表头行：前 5 行里"非空格子最多"的那一行，且至少 2 格、且大多是文字。
        # （真实表偶尔在表头上方还有一行大标题，只占 1 格 —— 用"格数最多"就跳过去了）
        header_row, best = None, 0
        for r in rows[:5]:
            texts = [t for (rr, _c), (t, _a) in cells.items() if rr == r]
            if len(texts) < 2:
                continue
            text_cnt = sum(1 for t in texts if not _is_numberish(t) and not rowscope.parse_date(t))
            if text_cnt >= max(2, len(texts) * 0.6) and len(texts) > best:
                header_row, best = r, len(texts)
        if header_row is None:
            continue
        headers = {c: t for (r, c), (t, _a) in cells.items() if r == header_row}
        data = {(r, c): v for (r, c), v in cells.items() if r > header_row}
        tables[key] = {
            "header_row": header_row,
            "headers": headers,
            "cells": data,
            "max_row": max(r for r, _c in cells),
        }
    return tables


# ---------------------------------------------------------------- 列名对齐
def match_column(table, col_info):
    """模板列 → 源表列号。对不上返回 None。

    三层，从严到宽（顺序很重要：先精确，避免「人口」被「65岁及以上人口占比」抢走）：
      ① 模板列名 / 别名 == 源列头（去空白小写后）
      ② 模板列名 / 别名 出现在源列头里，且**只有一个**源列头满足
      ③ 源列头出现在模板列名里，且只有一个
    """
    heads = {c: _norm(t) for c, t in table["headers"].items()}
    names = [_norm(col_info.get("title"))]
    names += [_norm(a) for a in (col_info.get("aliases") or [])]
    # 斜线表头（国家/地区）拆分后也算：模板那格写「国家」，源表叫「国家/地区」
    for p in (col_info.get("slash_parts") or []):
        names.append(_norm(p))
    names = [n for n in dict.fromkeys(names) if n]

    for n in names:                                   # ① 精确
        for c, h in heads.items():
            if h == n:
                return c
    for n in names:                                   # ② 源列头包含模板列名
        hit = [c for c, h in heads.items() if n and n in h]
        if len(hit) == 1:
            return hit[0]
    for n in names:                                   # ③ 模板列名包含源列头
        hit = [c for c, h in heads.items() if h and h in n and len(h) >= 2]
        if len(hit) == 1:
            return hit[0]
    return None


def find_date_col(table, spec):
    """按条件里的列名找源表的日期列；找不到就当成"直通不了"。"""
    name = _norm(spec.get("col") or "日期")
    heads = {c: _norm(t) for c, t in table["headers"].items()}
    for c, h in heads.items():
        if h == name:
            return c
    for c, h in heads.items():                        # 「日期」匹配「统计日期」「日期时间」
        if name and (name in h or h in name) and len(h) >= 2:
            return c
    return None


# ---------------------------------------------------------------- 行条件
# 条件里的列名是"人话里的列名"（监测时间 / 城市），源表的列头可能叫别的
# （统计时间 / 地市）。这里做一次落位；落不到就**丢掉这条条件并报出来**，
# 绝不"当成没条件"静默放行 —— 那正是山东那次填错的原因。
_FAMILIES = (
    ("时间", "日期", "时刻", "时段"),
    ("城市", "地市", "地区", "区域", "省", "市"),
)


def resolve_cond_col(table, name):
    """条件列名 → 源表列号；对不上返回 None。"""
    heads = {c: _norm(t) for c, t in table["headers"].items()}
    n = _norm(name)
    if not n:
        return None
    for c, h in heads.items():                        # ① 精确
        if h == n:
            return c
    hit = [c for c, h in heads.items() if n in h or h in n]
    if len(hit) == 1:                                 # ② 互相包含且唯一
        return hit[0]
    for fam in _FAMILIES:                             # ③ 同族兜底
        if any(w in n for w in fam):
            hit = [c for c, h in heads.items() if any(w in h for w in fam)]
            if len(hit) == 1:
                return hit[0]
    return None


def keep_rows(table, rows, conds):
    """按一组条件筛源行。

    返回 (命中的行, 没能落位的条件列名, 每条的命中数)。
    一条条件都落不了位时**返回空结果**（调用方据此报错），
    而不是返回全部行 —— "筛不了"和"不用筛"是两回事。
    """
    pairs, dropped = [], []
    for cond in conds or []:
        c = resolve_cond_col(table, cond.get("col"))
        if c is None:
            dropped.append(cond.get("col"))
        else:
            pairs.append((cond, c))
    if not pairs:
        return [], dropped, []

    kept, per_cond = [], [[] for _ in pairs]
    for r in rows:
        ok = True
        for i, (cond, c) in enumerate(pairs):
            v = (table["cells"].get((r, c)) or ("", ""))[0]
            if rowscope.match_value(v, cond):
                per_cond[i].append(r)
            else:
                ok = False
        if ok:
            kept.append(r)
    return kept, dropped, per_cond


# ---------------------------------------------------------------- 主流程
def run_blocks(cols, blocks, req_text, out_dir, pool=None, verbose=True):
    """按表直通：每张模板表各自筛行、各自整列搬。

    blocks 是模板的「块」清单（docx 多张表 / xlsx 多个 sheet），每块：
        {"no": 1, "caption": "本表记录…德州市各监测站点…", "n_rows": 24}

    每块的行条件按 rowscope.block_conditions 的三层优先级取：
        用户要求里写给「表X」的 → 模板表前描述 → 用户要求里没分表的全局条件

    产出仍是「一列一个 抽取结果-*.json」，但每条 item 带 block 字段，
    写回阶段据此分块落位（见 writer/fill.py 的 build_slot_map_blocks）。
    """
    if pool is None:
        import extract
        pool = extract.load_pool()

    lines = []
    def say(s=""):
        lines.append(str(s))
        if verbose:
            print(s)

    t0 = time.time()
    tables = scan_tables(pool)
    if not tables:
        return {"ok": False, "lines": ["直通：池子里没有可识别的表格，回落大模型路径"],
                "reason": "no_table"}

    # 选主表：能对上最多模板列的那张
    best, best_hit, best_map = None, -1, {}
    for key, t in tables.items():
        m, hit = {}, 0
        for col in cols:
            c = match_column(t, col)
            if c is not None:
                m[col["title"]] = c
                hit += 1
        if hit > best_hit:
            best, best_hit, best_map = key, hit, m

    file, sheet = best
    table = tables[best]
    say("直通：主表 = 《%s》· 工作表「%s」 ｜ 表头第 %d 行 ｜ 数据 %d 行 / %d 列"
        % (file, sheet, table["header_row"],
           len({r for r, _c in table["cells"]}), len(table["headers"])))

    miss = [c["title"] for c in cols if c.get("title") not in best_map]
    if miss:
        say("   ⚠ 这 %d 列在源表里找不到同名列：%s" % (len(miss), "、".join(miss)))
        say("   → 直通只在「每一列都对得上」时才用；这次回落到大模型路径")
        return {"ok": False, "lines": lines, "reason": "col_miss", "miss": miss}

    say("   列对齐：" + " ｜ ".join("%s→「%s」" % (k, table["headers"][v])
                                 for k, v in best_map.items()))

    if not blocks:
        blocks = [{"no": 1, "caption": "", "n_rows": 0}]

    # ---------- 逐块筛行 ----------
    rows_all = sorted({r for r, _c in table["cells"]})
    plan = []          # [{"no","conds","src","rows"}]
    warns = []
    for b in blocks:
        no = int(b.get("no") or 1)
        conds, src = rowscope.block_conditions(
            req_text or "", no, b.get("caption") or "")
        if not conds:
            # 没有任何条件：按源表顺序取前 n_rows 行（模板容量），
            # 而不是把 2 万行全塞进 24 个格子里。
            # ⚠ 这是唯一还可能"填错但看起来成功"的口子：源表里有几万行、
            #   用户又没说要哪一段，前 n 行未必是他要的。所以这里必须喊出来
            #   （山东那次就是要求里只有等值条件、老版本判成"无条件"→ 交了别的城市）。
            cap = int(b.get("n_rows") or 0) or len(rows_all)
            kept = rows_all[:cap]
            plan.append({"no": no, "conds": [], "src": "无条件", "rows": kept})
            if len(rows_all) > max(cap * 3, cap + 5):
                w = ("表%d 没有任何筛选条件 → 只取源表前 %d 行（源表共 %d 行）。"
                     "如果这不是你要的数据，请在「填表要求」里写清城市/时间，"
                     "或在模板该表上方写一句『本表记录…』" % (no, len(kept), len(rows_all)))
                warns.append(w)
                say("   ⚠ " + w)
            else:
                say("   表%d：%s → 取源表前 %d 行（源表共 %d 行，数量吻合）"
                    % (no, src, len(kept), len(rows_all)))
            continue
        kept, dropped, _per = keep_rows(table, rows_all, conds)
        tag = "表%d：%s ｜ %s" % (no, src, rowscope.describe_conditions(conds))
        if dropped:
            tag += " ｜ ⚠ 源表没有这些列，条件没用上：%s" % "、".join(dropped)
        cap = int(b.get("n_rows") or 0)
        if cap and len(kept) > cap:
            tag += " ｜ 命中 %d 行，模板只有 %d 行 → 填前 %d 行" % (
                len(kept), cap, cap)
        tag += " ｜ 命中 %d 行" % len(kept)
        plan.append({"no": no, "conds": conds, "src": src, "rows": kept})
        say("   " + tag)
        if not kept:
            say("   ⚠ 表%d 一条都没命中 —— 这张表会留空，不从别处凑数" % no)

    if not any(p["rows"] for p in plan):
        return {"ok": False, "lines": lines + ["所有块都筛空了 → 回落大模型路径"],
                "reason": "no_rows"}

    # ---------- 整列搬（按块顺序） ----------
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    n_by_col, files = {}, []
    for col in cols:
        title = col["title"]
        sc = best_map[title]
        items = []
        for p in plan:
            for r in p["rows"]:
                cell = table["cells"].get((r, sc))
                if not cell:
                    continue                       # 空格子就留空：写回按行号对齐，不会串行
                text, addr = cell
                items.append({
                    "value": text,
                    "evidence": text,
                    "block": p["no"],
                    "file": file,
                    "head": table["headers"].get(sc),
                    "colheads": None,
                    "line": None,
                    "loc_text": "工作表「%s」 %s" % (sheet, addr),
                    "aligned": "direct",
                })
        res = {
            "col": title,
            "time": time.strftime("%Y-%m-%d %H:%M"),
            "elapsed_sec": 0.0,
            "usage": {"total_tokens": 0, "provider": "direct", "model": "column-copy"},
            "candidate_count": sum(len(p["rows"]) for p in plan),
            "items": items,
            "direct": True,
            "blocks": [{"no": p["no"], "rows": len(p["rows"]), "src": p["src"],
                        "cond": rowscope.describe_conditions(p["conds"])}
                       for p in plan],
            "row_filter": " ｜ ".join(
                "表%d %s" % (p["no"], rowscope.describe_conditions(p["conds"]))
                for p in plan if p["conds"]),
        }
        p = out_dir / ("抽取结果-%s.json" % title)
        p.write_text(json.dumps(res, ensure_ascii=False, indent=1), encoding="utf-8")
        files.append(p.name)
        n_by_col[title] = len(items)

    sec = round(time.time() - t0, 2)
    say("   整列搬完：%d 列 × 最多 %d 条 ｜ 用时 %.2f 秒 ｜ 模型调用 0 次"
        % (len(cols), max(n_by_col.values()) if n_by_col else 0, sec))

    return {"ok": True, "lines": lines,
            "records": max((len(p["rows"]) for p in plan), default=0),
            "elapsed": sec, "file": file, "sheet": sheet, "n_by_col": n_by_col,
            "files": files, "blocks": plan, "warnings": warns,
            "desc": " ｜ ".join("%s" % rowscope.describe_conditions(p["conds"])
                               for p in plan if p["conds"])}


def run(cols, spec, out_dir, pool=None, verbose=True):
    """老接口：单块 + 日期区间条件。保留给旧脚本/对照实验用。"""
    req = (spec or {}).get("raw") or ""
    blocks = [{"no": 1, "caption": "", "n_rows": 0}]
    if spec and not req:
        # spec 是从环境变量 A23_ROW_FILTER 来的，没有原话可传给新解析器
        req = "%s：%s 到 %s" % (spec["col"], spec["lo"], spec["hi"])
    return run_blocks(cols, blocks, req, out_dir, pool=pool, verbose=verbose)
