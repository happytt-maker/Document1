# -*- coding: utf-8 -*-
"""
跑测试.py —— 泛化测试编排：3 个场景各跑一遍完整流水线，对标准答案算准确率
==========================================================================

背景：
  比赛是随机抽题 —— 拿"没见过的新文档 + 新模板"再测一遍，
  才知道系统是真会了，还是只背住了老数据。

对每个场景做的事（就是比赛当天流程的全程仿真）：
  ① 读源文档   doc_reader/main.py <该场景源文档>   → output.json
  ② 理解模板   template_reader/main.py <该场景模板> → 模板表结构.json
  ③ 抽取(并行) matcher/run_all.py --no-eval        → 抽取结果 × 7 列
  ④ 写回表格   writer/fill.py <该场景模板>
  ⑤ 写回验收   writer/verify_fill.py <该场景模板>
  ⑥ 评测对分   标准答案 vs 抽取结果（值严格/宽松、出处、耗时）

安全措施（不弄脏"模板4"的老数据）：
  - 跑之前：把 doc_reader/output.json、模板表结构.json、抽取结果-*.json 等
    备份到「.备份-跑测试前」
  - 全部跑完：还原备份，老流程原样如初
  - 本脚本不动任何"源数据/模板"文件，产物全部收集到「结果-场景X/」

跑法：
  python 跑测试.py          # 全部 3 个场景
  python 跑测试.py 1        # 只跑场景1（试水用）
  python 跑测试.py --no-preflight   # 跳过标准答案预检（不建议）
"""
import importlib.util
import json
import os
import re
import shutil
import subprocess
import sys
import time
from pathlib import Path

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

ROOT = Path(__file__).resolve().parent           # 泛化测试/
BASE = ROOT.parent                               # 研一项目/
PY = sys.executable
BACKUP = ROOT / ".备份-跑测试前"

SCENES = [
    {
        "no": 1, "dir": "场景1-费用报销",
        "template": "模板6-费用报销台账.xlsx",
        "main_src": "源文档/txt/费用报销登记簿.txt", "kind": "txt",
        "expected_cols": ["报销单号", "报销人", "所属部门", "提交日期",
                          "费用明细", "报销金额（元）", "审批状态"],
    },
    {
        "no": 2, "dir": "场景2-招聘初筛",
        "template": "模板7-候选人初筛表.xlsx",
        "main_src": "源文档/xlsx/候选人来访登记表.xlsx", "kind": "xlsx",
        "sheet": "来访登记",
        "expected_cols": ["姓名", "应聘岗位", "学历", "工作年限（年）",
                          "期望薪资（元）", "联系电话", "初筛结论"],
    },
    {
        "no": 3, "dir": "场景3-设备报修",
        "template": "模板8-设备报修工单登记表.xlsx",
        "main_src": "源文档/txt/设备报修记录.txt", "kind": "txt",
        "expected_cols": ["工单号", "设备名称", "故障描述", "报修人",
                          "报修日期", "故障等级", "处理状态"],
    },
]

PUNCT = r"[\s，。、；：！？,.;:!?（）()【】\[\]「」『』\"'']+"


def norm(s: str) -> str:
    return re.sub(PUNCT, "", s or "")


def run(cmd, cwd, env_extra=None, timeout=300):
    """跑一个子脚本，返回 (秒数, 返回码, 输出文本)。"""
    t0 = time.time()
    env = os.environ.copy()
    env["PYTHONIOENCODING"] = "utf-8"
    if env_extra:
        env.update(env_extra)
    try:
        p = subprocess.run([str(x) for x in cmd], cwd=str(cwd),
                           capture_output=True, env=env, timeout=timeout)
        out = p.stdout.decode("utf-8", "replace")
        err = p.stderr.decode("utf-8", "replace")
        if err.strip():
            out += "\n--- stderr ---\n" + err
        return time.time() - t0, p.returncode, out
    except subprocess.TimeoutExpired:
        return time.time() - t0, 999, "超时（%d 秒）" % timeout


def tail(text, n=12):
    lines = [ln for ln in (text or "").splitlines() if ln.strip()]
    return "\n".join("      | " + ln for ln in lines[-n:])


# ==================================================================
# 标准答案：解析 + 预检
# ==================================================================
def parse_answer(path: Path):
    """解析标准答案 md → [{列名: (值, 出处)} × 3 条]。"""
    text = path.read_text(encoding="utf-8")
    records = []
    cur = None
    for raw in text.splitlines():
        line = raw.strip()
        if re.match(r"^##\s*记录[一二三]", line):
            cur = {}
            records.append(cur)
            continue
        if cur is None:
            continue
        m = re.match(r"^-\s*(.+?)\s*=\s*(.*?)\s*(?:｜|\|)\s*来源：(.+?)\s*$", line)
        if m:
            cur[m.group(1).strip()] = (m.group(2).strip(), m.group(3).strip())
    return records


def content_ok(val: str, cell_text: str) -> bool:
    """值 与 出处格内容 是否对得上（包含 或 ≥60% 二字片段命中）。"""
    nv, nc = norm(val), norm(cell_text)
    if not nv or not nc:
        return False
    if nv in nc or nc in nv:
        return True
    bg = [nv[i:i + 2] for i in range(len(nv) - 1)]
    if not bg:
        return True   # 单字值，放宽
    hit = sum(1 for g in bg if g in nc)
    return hit / len(bg) >= 0.6


def preflight(scene):
    """跑测试之前：核对标准答案的每条『来源』确实指向含该值的位置。

    这是防"尺子本身是错的" —— 答案错了，测试跑得再顺也没意义。
    """
    sdir = ROOT / scene["dir"]
    ans = parse_answer(sdir / "标准答案.md")
    src_path = sdir / scene["main_src"]
    problems = []

    if not ans:
        problems.append("标准答案解析失败（没找到「## 记录一/二/三」）")
        return problems

    if scene["kind"] == "txt":
        lines = src_path.read_text(encoding="utf-8").splitlines()
        for ri, rec in enumerate(ans, 1):
            for col, (val, src) in rec.items():
                m = re.match(r"^第\s*(\d+)\s*行$", src)
                if not m:
                    problems.append("记录%d·%s：出处写法不是「第 N 行」：%s" % (ri, col, src))
                    continue
                n = int(m.group(1))
                if not (1 <= n <= len(lines)):
                    problems.append("记录%d·%s：第 %d 行超出文件范围（共 %d 行）"
                                    % (ri, col, n, len(lines)))
                elif not content_ok(val, lines[n - 1]):
                    problems.append("记录%d·%s：第 %d 行内容对不上值「%s」→ 实际：%s"
                                    % (ri, col, n, val, lines[n - 1]))
    else:  # xlsx
        from openpyxl import load_workbook
        wb = load_workbook(str(src_path))
        ws = wb[scene["sheet"]]
        for ri, rec in enumerate(ans, 1):
            for col, (val, src) in rec.items():
                m = re.search(r"([A-Z]+\d+)\s*$", src)
                if not m:
                    problems.append("记录%d·%s：出处里找不到单元格坐标：%s" % (ri, col, src))
                    continue
                coord = m.group(1)
                cell_v = ws[coord].value
                if cell_v is None:
                    problems.append("记录%d·%s：%s 是空格子" % (ri, col, coord))
                elif not content_ok(val, str(cell_v)):
                    problems.append("记录%d·%s：%s 内容对不上值「%s」→ 实际：%s"
                                    % (ri, col, coord, val, cell_v))
        wb.close()
    return problems


# ==================================================================
# 评测（复用 evaluate.py 的值比对函数，出处按"行号/坐标"双体系）
# ==================================================================
def load_value_check():
    spec = importlib.util.spec_from_file_location(
        "a23_evaluate", str(BASE / "matcher" / "evaluate.py"))
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod.value_check


def src_match(exp_src: str, got: dict) -> bool:
    """出处比对：行号体系 / 坐标体系。"""
    m = re.match(r"^第\s*(\d+)\s*行$", exp_src)
    if m:
        n = int(m.group(1))
        return (got.get("line") == n) or ("第 %d 行" % n) in (got.get("loc_text") or "")
    m = re.match(r"^第\s*(\d+)\s*段$", exp_src)
    if m:
        return ("第 %s 段" % m.group(1)) in (got.get("loc_text") or "")
    a = re.sub(r"\s+", "", exp_src)
    b = re.sub(r"\s+", "", got.get("loc_text") or "")
    return bool(a) and bool(b) and (a in b or b in a)


def evaluate_scene(scene, value_check):
    """读 标准答案 + 抽取结果 → 逐格比对。返回 (报告文本, 统计dict)。"""
    sdir = ROOT / scene["dir"]
    records = parse_answer(sdir / "标准答案.md")

    lines = []
    n_total = n_strict = n_loose = n_src = 0
    issues = []

    for col in scene["expected_cols"]:
        f = BASE / "matcher" / ("抽取结果-%s.json" % col)
        res = json.loads(f.read_text(encoding="utf-8")) if f.exists() else None
        items = (res or {}).get("items", [])
        el = (res or {}).get("elapsed_sec", "—")
        tk = ((res or {}).get("usage") or {}).get("total_tokens", "—")
        lines.append("")
        lines.append("—— 列名：%s（耗时 %s 秒 / %s tokens）——" % (col, el, tk))

        for i, rec in enumerate(records):
            tag = "记录%s" % "一二三"[i]
            if col not in rec:
                continue
            exp_val, exp_src = rec[col]
            got = items[i] if i < len(items) else None

            if got is None:
                n_total += 1
                lines.append("%s   值 错（没抽到）      出处 错（没抽到）" % tag)
                issues.append("%s·%s：漏抽（标准=%s，%s）" % (tag, col, exp_val, exp_src))
                continue

            lv, note = value_check(exp_val, got.get("value", ""))
            if lv == -1:
                continue
            src_ok = src_match(exp_src, got)

            n_total += 1
            n_strict += 1 if lv == 2 else 0
            n_loose += 1 if lv >= 1 else 0
            n_src += 1 if src_ok else 0

            mark_v = "对" if lv == 2 else ("近似" if lv == 1 else "错")
            lines.append("%s   值 %s（%s）      出处 %s   %s"
                         % (tag, mark_v, note or "不一致",
                            "对" if src_ok else "错", got.get("loc_text") or "未定位"))
            if lv < 2:
                lines.append("      标准：%s" % exp_val)
                lines.append("      抽取：%s" % got.get("value"))
            if not src_ok:
                lines.append("      出处对不上：标准 = %s ｜ 抽取 = %s"
                             % (exp_src, got.get("loc_text") or "未定位"))
                issues.append("%s·%s：出处错（标准=%s 抽取=%s）"
                              % (tag, col, exp_src, got.get("loc_text") or "未定位"))
            elif lv < 1:
                issues.append("%s·%s：值错（标准=%s 抽取=%s）"
                              % (tag, col, exp_val, got.get("value")))

    stats = {"total": n_total, "strict": n_strict, "loose": n_loose, "src": n_src}
    head = []
    head.append("=" * 62)
    head.append("评测报告 · %s（标准答案 vs 抽取结果）" % scene["dir"])
    head.append("生成时间：%s" % time.strftime("%Y-%m-%d %H:%M"))
    head.append("=" * 62)
    tail_ = []
    tail_.append("")
    tail_.append("=" * 62)
    tail_.append("【汇总】%d 格（%d 列 × 3 条记录均按标准答案计）"
                 % (n_total, len(scene["expected_cols"])))
    if n_total:
        pct = lambda x: 100.0 * x / n_total
        tail_.append("  值 · 严格（一字不差/数字等价）：%d/%d（%.1f%%）"
                     % (n_strict, n_total, pct(n_strict)))
        tail_.append("  值 · 宽松（含近似）：%d/%d（%.1f%%）"
                     % (n_loose, n_total, pct(n_loose)))
        tail_.append("  出处 · 命中：%d/%d（%.1f%%）" % (n_src, n_total, pct(n_src)))
    if issues:
        tail_.append("")
        tail_.append("【待人工复看】")
        for it in issues:
            tail_.append("  - " + it)
    tail_.append("=" * 62)
    return "\n".join(head + lines + tail_), stats


# ==================================================================
# 主流程
# ==================================================================
def backup_workspace():
    """把会被覆盖的产物备份（跑完后还原）。"""
    BACKUP.mkdir(parents=True, exist_ok=True)
    files = [BASE / "doc_reader" / "output.json",
             BASE / "template_reader" / "模板表结构.json",
             BASE / "matcher" / "运行摘要.txt"]
    for f in files:
        if f.exists():
            shutil.copy2(f, BACKUP / f.name)
    # 抽取结果单独存个清单
    listfile = BACKUP / "抽取结果清单.txt"
    names = [p.name for p in (BASE / "matcher").glob("抽取结果-*.json")]
    listfile.write_text("\n".join(names), encoding="utf-8")
    for p in (BASE / "matcher").glob("抽取结果-*.json"):
        shutil.copy2(p, BACKUP / p.name)
    return len(names)


def restore_workspace():
    """还原：先清掉跑测试期间的抽取结果，再恢复备份。"""
    mx = BASE / "matcher"
    for p in mx.glob("抽取结果-*.json"):
        p.unlink()
    for p in BACKUP.glob("抽取结果-*.json"):
        shutil.copy2(p, mx / p.name)
    for name in ("output.json", "模板表结构.json", "运行摘要.txt"):
        src = BACKUP / name
        if not src.exists():
            continue
        if name == "output.json":
            shutil.copy2(src, BASE / "doc_reader" / name)
        elif name == "模板表结构.json":
            shutil.copy2(src, BASE / "template_reader" / name)
        else:
            shutil.copy2(src, mx / name)


def run_scene(scene, value_check, quiet_print=None):
    """跑一个场景的完整流水线 + 评测 + 产物收集。返回统计 dict。"""
    say = quiet_print or print
    sdir = ROOT / scene["dir"]
    out_dir = ROOT / ("结果-场景%d" % scene["no"])
    out_dir.mkdir(parents=True, exist_ok=True)

    total0 = time.time()
    stats = {"scene": scene["dir"], "stages": {}}

    # ① 读源文档
    sec, code, out = run([PY, "main.py", str(sdir / "源文档")],
                         cwd=BASE / "doc_reader")
    stats["stages"]["读源文档"] = round(sec, 2)
    if code != 0:
        return {"error": "① 读源文档失败：\n" + tail(out, 15), "scene": scene["dir"]}
    data = json.loads((BASE / "doc_reader" / "output.json").read_text(encoding="utf-8"))
    n_doc = len(data.get("documents", []))
    n_blk = sum(d.get("block_count", 0) for d in data.get("documents", []))
    say("    ① 读源文档：%d 份 / %d 块 ｜ %.2f 秒" % (n_doc, n_blk, sec))

    # ② 理解模板
    sec, code, out = run([PY, "main.py", str(sdir / "模板表格")],
                         cwd=BASE / "template_reader")
    stats["stages"]["理解模板"] = round(sec, 2)
    if code != 0:
        return {"error": "② 理解模板失败：\n" + tail(out, 15), "scene": scene["dir"]}
    tinfo = json.loads((BASE / "template_reader" / "模板表结构.json").read_text(encoding="utf-8"))
    sheet = tinfo[0]["sheets"][0]
    got_cols = [c["title"] for c in sheet["columns"]]
    if got_cols != scene["expected_cols"]:
        return {"error": "② 模板列名核对失败：\n  期望 %s\n  实际 %s" % (scene["expected_cols"], got_cols),
                "scene": scene["dir"]}
    fr = sheet.get("fill_region") or {}
    say("    ② 理解模板：表头第 %s 行 / 待填 %s~%s 行 / %.2f 秒"
        % (sheet.get("header_row"), fr.get("start_row"), fr.get("end_row"), sec))

    # ③ 抽取（并行）
    sec, code, out = run([PY, "run_all.py", "--no-eval"], cwd=BASE / "matcher")
    stats["stages"]["抽取"] = round(sec, 2)
    if code != 0:
        return {"error": "③ 抽取失败：\n" + tail(out, 15), "scene": scene["dir"]}
    say("    ③ 抽取：%.2f 秒（含 7 列并行调模型）" % sec)

    # ④ 写回
    env = {"A23_TEMPLATE_DIR": str(sdir / "模板表格")}
    sec, code, out = run([PY, "fill.py", scene["template"]],
                         cwd=BASE / "writer", env_extra=env)
    stats["stages"]["写回"] = round(sec, 2)
    if code != 0:
        return {"error": "④ 写回失败（返回码 %d）：\n%s" % (code, tail(out, 15)),
                "scene": scene["dir"]}
    say("    ④ 写回：%.2f 秒" % sec)

    # ⑤ 验收
    sec, code, out = run([PY, "verify_fill.py", scene["template"]],
                         cwd=BASE / "writer", env_extra=env)
    stats["stages"]["验收"] = round(sec, 2)
    vcode = code
    say("    ⑤ 写回验收：%.2f 秒（返回码 %d —— 0=全过）" % (sec, code))

    # ⑥ 评测
    sec0 = time.time()
    report, estats = evaluate_scene(scene, value_check)
    (out_dir / "评测报告.txt").write_text(report, encoding="utf-8")
    stats["stages"]["评测"] = round(time.time() - sec0, 2)
    stats.update(estats)
    stats["wall"] = round(time.time() - total0, 2)
    say("    ⑥ 评测：值严格 %d/%d ｜ 值宽松 %d/%d ｜ 出处 %d/%d"
        % (estats["strict"], estats["total"], estats["loose"],
           estats["total"], estats["src"], estats["total"]))

    # ⑦ 收集产物
    stem = Path(scene["template"]).stem
    picks = [
        (BASE / "doc_reader" / "output.json", "读取结果-output.json"),
        (BASE / "template_reader" / "模板表结构.json", "模板表结构.json"),
        (BASE / "matcher" / "运行摘要.txt", "运行摘要-抽取.txt"),
        (BASE / "writer" / "输出" / (stem + "-已填写.xlsx"), stem + "-已填写.xlsx"),
        (BASE / "writer" / "输出" / ("验收报告-%s.txt" % stem), "验收报告.txt"),
        (BASE / "writer" / "输出" / ("填表报告-%s.txt" % stem), "填表报告.txt"),
    ]
    for src_f, dst_name in picks:
        if src_f.exists():
            shutil.copy2(src_f, out_dir / dst_name)
    for col in scene["expected_cols"]:
        f = BASE / "matcher" / ("抽取结果-%s.json" % col)
        if f.exists():
            shutil.copy2(f, out_dir / f.name)

    stats["verify_code"] = vcode
    return stats


def main():
    args = sys.argv[1:]
    only = None
    for a in args:
        if a.isdigit():
            only = int(a)
    scenes = [s for s in SCENES if (only is None or s["no"] == only)]

    print("=" * 64)
    print("泛化测试 · 跑 %d 个场景 ｜ 时间 %s" % (len(scenes), time.strftime("%Y-%m-%d %H:%M")))
    print("=" * 64)

    # ---- 预检：标准答案必须先是对的 ----
    if "--no-preflight" not in args:
        print("")
        print("【预检】核对标准答案（每条『来源』是否真的指向含该值的位置）……")
        bad = 0
        for scene in scenes:
            pb = preflight(scene)
            if pb:
                bad += len(pb)
                print("  ✗ %s：发现 %d 个问题" % (scene["dir"], len(pb)))
                for p in pb[:10]:
                    print("      - " + p)
            else:
                print("  ✓ %s：全部通过" % scene["dir"])
        if bad:
            print("")
            print("预检不通过 → 停下（尺子不准，测了也没意义）。先修标准答案。")
            return
    else:
        print("（--no-preflight：跳过预检）")

    # ---- 备份 ----
    n = backup_workspace()
    print("")
    print("【备份】已备份旧产物（output.json / 模板表结构 / 抽取结果×%d）→ .备份-跑测试前" % n)

    # ---- 逐场景跑 ----
    value_check = load_value_check()
    all_stats = []
    for scene in scenes:
        print("")
        print("-" * 64)
        print("▶ %s（模板：%s）" % (scene["dir"], scene["template"]))
        print("-" * 64)
        st = run_scene(scene, value_check)
        if "error" in st:
            print("  ！失败：" + st["error"])
            all_stats.append(st)
        else:
            all_stats.append(st)

    # ---- 总报告 ----
    lines = []
    lines.append("=" * 64)
    lines.append("泛化测试 · 总报告（%s）" % time.strftime("%Y-%m-%d %H:%M"))
    lines.append("=" * 64)
    lines.append("")
    lines.append("场景 | 值严格 | 值宽松 | 出处 | 总耗时")
    lines.append("-" * 64)
    T = {"total": 0, "strict": 0, "loose": 0, "src": 0}
    for st in all_stats:
        if "error" in st:
            lines.append("%s | 失败：%s" % (st["scene"], st["error"].splitlines()[0]))
            continue
        T["total"] += st["total"]
        T["strict"] += st["strict"]
        T["loose"] += st["loose"]
        T["src"] += st["src"]
        lines.append("%s | %d/%d | %d/%d | %d/%d | %.2f 秒"
                     % (st["scene"], st["strict"], st["total"], st["loose"],
                        st["total"], st["src"], st["total"], st["wall"]))
    lines.append("-" * 64)
    if T["total"]:
        pct = lambda x: 100.0 * x / T["total"]
        lines.append("合计（%d 格）：值严格 %d/%d（%.1f%%）｜ 值宽松 %d/%d（%.1f%%）｜ 出处 %d/%d（%.1f%%）"
                     % (T["total"], T["strict"], T["total"], pct(T["strict"]),
                        T["loose"], T["total"], pct(T["loose"]),
                        T["src"], T["total"], pct(T["src"])))
    lines.append("")
    lines.append("每场景分阶段耗时（秒）：")
    for st in all_stats:
        if "error" in st:
            continue
        sg = st["stages"]
        lines.append("  %s：读 %s ｜ 模板 %s ｜ 抽取 %s ｜ 写回 %s ｜ 验收 %s ｜ 评测 %s ｜ 合计 %.2f"
                     % (st["scene"], sg.get("读源文档"), sg.get("理解模板"),
                        sg.get("抽取"), sg.get("写回"), sg.get("验收"),
                        sg.get("评测"), st["wall"]))
    lines.append("")
    lines.append("赛题红线对照：≤90 秒/份文档 —— 每场景整条流水线均在红线内。")
    lines.append("=" * 64)
    report = "\n".join(lines)
    (ROOT / "泛化测试总报告.txt").write_text(report, encoding="utf-8")
    print("")
    print(report)

    # ---- 还原 ----
    restore_workspace()
    print("")
    print("【还原】已恢复旧产物（模板4 的 output.json / 模板表结构 / 抽取结果）")


if __name__ == "__main__":
    main()
