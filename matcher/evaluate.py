# -*- coding: utf-8 -*-
"""
evaluate.py —— 第 4 步 · 4c：评测（把你的标准答案当尺子）

读两份东西：
  - 标准答案-模板4.md     （你手填的：值 + 来源行号）
  - 抽取结果-*.json       （extract.py 的产物，跑了几列就读几列）

逐格比对，输出「评测报告-模板4.txt」：
  - 值对不对（严格：一字不差 / 数字等价；宽松：包含关系、高度相似）
  - 出处对不对（行号一致才算对）
  - 不一致清单（供人工复看 —— 数字之外，眼睛说了算）

跑法（VSCode 打开 matcher 文件夹）：
  方式一：按 F5 选「跑评测」
  方式二：终端里  python evaluate.py
"""
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

MATCHER = Path(__file__).resolve().parent
ANSWER_MD = MATCHER / "标准答案-模板4.md"

ALL_COLS = ["投诉编号", "客户名称", "投诉日期", "投诉内容摘要",
            "责任判定", "处理时长（天）", "处理结果"]

PUNCT = r"[\s，。、；：！？,.;:!?（）()【】\[\]「」『』\"'']+"
UNITS = set("天个台小时分钟元起次月年份工作日")


def norm(s: str) -> str:
    """归一化：去空白、去标点。"""
    return re.sub(PUNCT, "", s or "")


def lcs_len(a: str, b: str) -> int:
    """最长公共子序列的长度（顺序敏感：字的顺序乱掉就不算匹配）。"""
    if not a or not b:
        return 0
    prev = [0] * (len(b) + 1)
    for x in a:
        cur = [0] * (len(b) + 1)
        for j, y in enumerate(b, 1):
            cur[j] = prev[j - 1] + 1 if x == y else max(prev[j], cur[j - 1])
        prev = cur
    return prev[-1]


def lcs_ratio(a: str, b: str) -> float:
    """最长公共子序列占两边总长的比例（0~1）。"""
    if not a or not b:
        return 0.0
    return 2.0 * lcs_len(a, b) / (len(a) + len(b))


def value_check(exp: str, got: str):
    """值比对。返回 (level, note)：2=严格对，1=近似，0=错。

    近似有两档口径：
      - 包含关系：一个答案是另一个的子串
      - 要素覆盖：标准答案里的信息被覆盖了多少（LCS / 标准答案长度）
        专治"摘要"类 —— 人家概括得长一点、短一点都合理，
        只要标准答案的信息没丢，就该算"近似"而不是"错"。
    """
    ne, ng = norm(exp), norm(got)
    if not ne:
        return -1, "标准答案为空（跳过）"
    if not ng:
        return 0, "没抽到"
    if ne == ng:
        return 2, "一字不差"

    # 数字等价：两边数字相同，且多出来的部分只是单位字（"2天" vs "2"）
    de = re.sub(r"[^\d.]", "", ne)
    dg = re.sub(r"[^\d.]", "", ng)
    re_ = re.sub(r"[\d.]", "", ne)
    rg = re.sub(r"[\d.]", "", ng)
    if de and de == dg and ((not re_ and all(c in UNITS for c in rg))
                            or (not rg and all(c in UNITS for c in re_))):
        return 2, "数字等价"

    if ne in ng or ng in ne:
        return 1, "包含关系"

    l = lcs_len(ne, ng)
    if l / len(ne) >= 0.75:
        return 1, "要素覆盖 %.0f%%" % (100.0 * l / len(ne))

    r = 2.0 * l / (len(ne) + len(ng))
    if r >= 0.75:
        return 1, "相似 %.0f%%" % (r * 100)
    return 0, ""


def parse_answer_md(path: Path):
    """解析标准答案：返回 [{列名: (值, 行号)}, ...]，每条记录一个 dict。"""
    text = path.read_text(encoding="utf-8")
    records = []
    cur = None
    for raw in text.split("\n"):
        line = raw.strip()
        m = re.match(r"^##\s*记录([一二三])", line)
        if m:
            cur = {}
            records.append(cur)
            continue
        if cur is None:
            continue
        m2 = re.match(r"^-\s*(.+?)\s*=\s*(.*?)\s*(?:｜|\|)\s*来源：第\s*(\d+)\s*行", line)
        if m2:
            col = m2.group(1).strip()
            val = m2.group(2).strip()
            ln = int(m2.group(3))
            cur[col] = (val, ln)
    return records


def main():
    lines = []

    def say(t=""):
        lines.append(t)
        print(t)

    if not ANSWER_MD.exists():
        print("找不到标准答案：%s" % ANSWER_MD)
        return

    records = parse_answer_md(ANSWER_MD)
    files = sorted(MATCHER.glob("抽取结果-*.json"))

    say("=" * 62)
    say("评测报告 · 标准答案 vs 抽取结果")
    say("生成时间：%s" % time.strftime("%Y-%m-%d %H:%M"))
    say("=" * 62)

    # 读所有抽取结果，按"列的自然顺序"排好（而不是按文件名）
    by_col = {}
    for f in files:
        res = json.loads(f.read_text(encoding="utf-8"))
        by_col[res.get("col")] = res

    # 排顺序 + 去重（防手滑：就算 ALL_COLS 里某列被写了两遍，也只评一次）
    ordered = []
    seen = set()
    for c in list(ALL_COLS) + [k for k in by_col if k not in ALL_COLS]:
        if c in by_col and c not in seen:
            ordered.append((c, by_col[c]))
            seen.add(c)

    if not records:
        say("标准答案解析失败（没找到「## 记录一/二/三」）。")
    elif not ordered:
        say("还没有任何抽取结果 —— 先去跑 extract.py（每列跑一次）。")
    else:
        n_val_ok = n_val_loose = n_src_ok = n_total = 0
        issues = []

        for col, res in ordered:
            items = res.get("items", [])
            say("")
            say("—— 列名：%s（耗时 %s 秒 / 用量 %s tokens）——"
                % (col, res.get("elapsed_sec", "?"),
                   (res.get("usage") or {}).get("total_tokens", "?")))

            for i, rec in enumerate(records):
                tag = "记录%s" % "一二三"[i]
                if col not in rec:
                    say("%s   标准答案里没有这一列，跳过" % tag)
                    continue
                exp_val, exp_line = rec[col]
                got = items[i] if i < len(items) else None

                if got is None:
                    n_total += 1
                    say("%s   值 错（没抽到）      出处 错（没抽到）" % tag)
                    issues.append("%s·%s：漏抽（标准=%s，第 %d 行）"
                                  % (tag, col, exp_val or "?", exp_line))
                    continue

                v_level, v_note = value_check(exp_val, got.get("value", ""))
                if v_level == -1:
                    say("%s   （标准答案为空，跳过）" % tag)
                    continue

                got_line = got.get("line")
                src_ok = (got_line == exp_line)

                n_total += 1
                n_val_ok += 1 if v_level == 2 else 0
                n_val_loose += 1 if v_level >= 1 else 0
                n_src_ok += 1 if src_ok else 0

                mark_v = "对" if v_level == 2 else ("近似" if v_level == 1 else "错")
                mark_s = "对" if src_ok else "错"
                say("%s   值 %s（%s）      出处 %s %s"
                    % (tag, mark_v, v_note or "不一致",
                       mark_s,
                       got.get("loc_text") or "未定位"))

                if v_level < 2:
                    say("      标准：%s" % exp_val)
                    say("      抽取：%s" % got.get("value"))
                if v_level == 0:
                    # 诊断：标准答案的信息被覆盖了多少（帮人工判断"差异可不可接受"）
                    ne_, ng_ = norm(exp_val), norm(got.get("value", ""))
                    cov = lcs_len(ne_, ng_) / max(1, len(ne_))
                    say("      诊断：标准信息被覆盖 %.0f%%（低 = 信息丢得多；高 = 只是说法不同）"
                        % (cov * 100))
                if not src_ok:
                    say("      出处对不上：标准第 %d 行 / 抽取 %s"
                        % (exp_line, got.get("loc_text") or "未定位"))

                if v_level < 1 or not src_ok:
                    issues.append("%s·%s：%s"
                                  % (tag, col,
                                     ("值不一致 " if v_level < 1 else ("值近似 " if v_level == 1 else ""))
                                     + ("出处错" if not src_ok else "")))

        # ---------- 汇总 ----------
        done_cols = list(by_col.keys())
        missing = [c for c in ALL_COLS if c not in done_cols]

        say("")
        say("=" * 62)
        say("【汇总】已评测 %d 列 × %d 条记录 = %d 格" % (len(ordered), len(records), n_total))
        if n_total:
            say("  值 · 严格（一字不差/数字等价）：%d/%d（%.1f%%）"
                % (n_val_ok, n_total, 100.0 * n_val_ok / n_total))
            say("  值 · 宽松（含近似）：%d/%d（%.1f%%）"
                % (n_val_loose, n_total, 100.0 * n_val_loose / n_total))
            say("  出处 · 行号一致：%d/%d（%.1f%%）"
                % (n_src_ok, n_total, 100.0 * n_src_ok / n_total))
        if missing:
            say("")
            say("未跑的列（在 extract.py 里改 COL_NAME 逐个跑）：")
            say("  " + "、".join(missing))
        if issues:
            say("")
            say("【待人工复看】")
            for it in issues:
                say("  - " + it)
        say("=" * 62)

    out_path = MATCHER / "评测报告-模板4.txt"
    out_path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    print()
    print("（报告已存：%s）" % out_path.name)


if __name__ == "__main__":
    main()
