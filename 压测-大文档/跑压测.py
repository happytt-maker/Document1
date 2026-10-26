# -*- coding: utf-8 -*-
"""
跑压测.py —— 大文档压测：按赛题规格跑"16 份大文档 × 5 张模板"
======================================================================
赛题【评价方法】：
    16 份文档（5×≥500KB docx ｜ 3×≥15KB md ｜ 5×≥500KB xlsx ｜ 3×≥15KB txt）
    全部上传 → 分别上传 5 张模板 → 每张跑一次 → 每表 ≥80%、每文档 ≤90 秒。

本脚本 = 该流程的仿真：
    前置：读 16 份大文档（模拟"上传"，一次性）→ 生成大池（独立文件，不动日常 output.json）
    逐张：模板理解 → 抽取（并行）→ 写回 → 验收（每张计时、归档）
    附加：金丝雀抽查（从大池里核对 15 组埋点值）+ 规格预检

安全措施（沿用五表连测的成熟做法）：
    - 预检：16 份文档规格达标 + 5 张模板待填区干净，不达标就停
    - 备份：跑前备份 matcher / template_reader / writer 相关产物
    - 还原：跑完自动恢复（日常小池状态不受影响）；压测产物已归档到「结果/」
    - 大池隔离：A23_OUTPUT_JSON 指向「结果/大池-output.json」，不碰 doc_reader/output.json

跑法（终端）：
    python 跑压测.py                # 全流程（含读取）
    python 跑压测.py --skip-read    # 跳过读取（大池已存在时）
    python 跑压测.py --only 4       # 只跑第 4 张（试水用）
"""
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

ROOT = Path(__file__).resolve().parent           # 压测-大文档/
BASE = ROOT.parent                               # 研一项目/
PY = sys.executable

DATA = BASE / "测试数据" / "压测-大文档"
SRC = DATA / "源文档"
ANSWERS = DATA / "压测答案.json"
TPL_DIR = BASE / "测试数据" / "模板表格"
DR = BASE / "doc_reader"
TR = BASE / "template_reader"
MATCHER = BASE / "matcher"
WRITER = BASE / "writer"
WOUT = WRITER / "输出"
RESULT = ROOT / "结果"
POOL = RESULT / "大池-output.json"               # 压测大池（独立文件）

BACKUP = None   # backup() 运行时赋值（时间戳目录）

TEMPLATES = [
    "模板1-设备采购汇总表.xlsx",
    "模板2-员工信息登记表.xlsx",
    "模板3-会议任务跟踪表.xlsx",
    "模板4-客诉处理台账.xlsx",
    "模板5-原材料库存汇总表.xlsx",
]

# 规格：docx/xlsx ≥500KB，md/txt ≥15KB
SPEC = {"docx": 500 * 1024, "xlsx": 500 * 1024, "md": 15 * 1024, "txt": 15 * 1024}


def run(cmd, cwd, env_extra=None, timeout=600):
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
# 预检：16 份文档规格 + 5 张模板待填区
# ==================================================================
def precheck():
    from openpyxl import load_workbook

    problems = []
    print("【预检 1/2】16 份文档规格（docx/xlsx ≥500KB；md/txt ≥15KB）……")
    stats = {}
    for kind in ("docx", "xlsx", "md", "txt"):
        d = SRC / kind
        files = sorted(d.glob("*")) if d.exists() else []
        ok_n = 0
        for f in files:
            if f.is_file():
                kb = f.stat().st_size / 1024
                std = f.stat().st_size >= SPEC[kind]
                ok_n += 1 if std else 0
        stats[kind] = (len(files), ok_n)
        flag = "✓" if files and ok_n == len(files) else "✗"
        print("  %s %s：%d 份（达标 %d）" % (flag, kind, len(files), ok_n))
        if not files:
            problems.append("没有 %s 文档（先跑 生成压测数据.py）" % kind)
        elif ok_n != len(files):
            problems.append("%s 有 %d 份未达标" % (kind, len(files) - ok_n))
    total = sum(v[0] for v in stats.values())
    print("  合计 %d 份（赛题规格 16 份）" % total)
    if total != 16:
        problems.append("文档总数 %d ≠ 16" % total)

    print("")
    print("【预检 2/2】5 张模板的待填区是否干净……")
    tinfo = json.loads((TR / "模板表结构.json").read_text(encoding="utf-8"))
    for tpl in TEMPLATES:
        info = next((t for t in tinfo if t.get("file") == tpl), None)
        if not info or not info.get("sheets"):
            problems.append("%s：模板表结构.json 里没有它" % tpl)
            continue
        sheet = info["sheets"][0]
        fr = sheet.get("fill_region")
        if not fr:
            problems.append("%s：没找到待填区" % tpl)
            continue
        wb = load_workbook(str(TPL_DIR / tpl))
        ws = wb[sheet["name"]]
        dirty = []
        for r in range(fr["start_row"], fr["end_row"] + 1):
            for col in sheet["columns"]:
                v = ws.cell(row=r, column=col["col"]).value
                if v is not None and str(v).strip():
                    dirty.append("%s=%s" % (ws.cell(row=r, column=col["col"]).coordinate, v))
        wb.close()
        if dirty:
            problems.append("%s：待填区有 %d 格旧内容（%s）" % (tpl, len(dirty), "、".join(dirty[:6])))
        else:
            print("  ✓ %s：干净（待填区 %d~%d 行）" % (tpl, fr["start_row"], fr["end_row"]))
    return problems


# ==================================================================
# 备份 / 还原
# ==================================================================
def backup_all():
    global BACKUP
    BACKUP = ROOT / ("备份-压测前-" + time.strftime("%m%d-%H%M%S"))
    BACKUP.mkdir(parents=True, exist_ok=False)
    n = 0
    for src, dst in [(TR / "模板表结构.json", BACKUP / "模板表结构.json"),
                     (MATCHER / "运行摘要.txt", BACKUP / "运行摘要.txt")]:
        if src.exists():
            shutil.copy2(src, dst)
            n += 1
    for f in MATCHER.glob("抽取结果-*.json"):
        shutil.copy2(f, BACKUP / f.name); n += 1
    for f in MATCHER.glob("评测报告-*.txt"):
        shutil.copy2(f, BACKUP / f.name); n += 1
    (BACKUP / "writer输出").mkdir(exist_ok=True)
    for f in WOUT.glob("*"):
        if f.is_file():
            shutil.copy2(f, BACKUP / "writer输出" / f.name); n += 1
    return n


def restore_all():
    # 清掉压测期间的新产物（均已归档到 结果/ 后才执行）
    for f in MATCHER.glob("抽取结果-*.json"):
        f.unlink()
    for f in MATCHER.glob("评测报告-*.txt"):
        f.unlink()
    for f in WOUT.glob("*"):
        if f.is_file():
            f.unlink()
    # 从备份恢复
    for name, dst_dir in [("模板表结构.json", TR), ("运行摘要.txt", MATCHER)]:
        src = BACKUP / name
        if src.exists():
            shutil.copy2(src, dst_dir / name)
    for f in BACKUP.glob("抽取结果-*.json"):
        shutil.copy2(f, MATCHER / f.name)
    for f in BACKUP.glob("评测报告-*.txt"):
        shutil.copy2(f, MATCHER / f.name)
    for f in (BACKUP / "writer输出").glob("*"):
        shutil.copy2(f, WOUT / f.name)


# ==================================================================
# 前置：读 16 份大文档 → 大池（独立文件）
# ==================================================================
def stage_read():
    RESULT.mkdir(parents=True, exist_ok=True)
    sec, code, out = run([PY, "main.py", str(SRC), "-o", str(POOL)], cwd=DR, timeout=1200)
    info = {"sec": round(sec, 2), "code": code}
    if code != 0 or not POOL.exists():
        info["error"] = "读取失败：\n" + tail(out, 15)
        return info
    data = json.loads(POOL.read_text(encoding="utf-8"))
    docs = data.get("documents", [])
    info["n_doc"] = len(docs)
    info["n_block"] = sum(d.get("block_count", 0) for d in docs)
    info["failed"] = len(data.get("failed", []))
    info["pool_mb"] = round(POOL.stat().st_size / 1024 / 1024, 1)
    info["by_type"] = {}
    for d_ in docs:
        t = d_.get("type", "?")
        info["by_type"].setdefault(t, [0, 0])
        info["by_type"][t][0] += 1
        info["by_type"][t][1] += d_.get("block_count", 0)
    return info


# ==================================================================
# 单张模板：模板理解 → 抽取 → 写回 → 验收（计时 + 归档）
# ==================================================================
def run_one(tpl, no):
    out_dir = RESULT / ("模板%d" % no)
    out_dir.mkdir(parents=True, exist_ok=True)
    stem = Path(tpl).stem
    stats = {"tpl": tpl, "no": no, "stages": {}}

    for f in MATCHER.glob("抽取结果-*.json"):
        f.unlink()

    sec, code, out = run([PY, "main.py", str(TPL_DIR)], cwd=TR)
    stats["stages"]["模板"] = round(sec, 2)
    if code != 0:
        stats["error"] = "模板理解失败：\n" + tail(out, 12)
        return stats

    tinfo = json.loads((TR / "模板表结构.json").read_text(encoding="utf-8"))
    info = next((t for t in tinfo if t.get("file") == tpl), None)
    cols = [c["title"] for c in info["sheets"][0]["columns"]] if info else []
    stats["cols"] = cols

    # 抽取（并行；大池通过 A23_OUTPUT_JSON 注入）
    env = {"A23_OUTPUT_JSON": str(POOL), "A23_TEMPLATE_NAME": tpl}
    sec, code, out = run([PY, "run_all.py", "--no-eval"], cwd=MATCHER, env_extra=env, timeout=900)
    stats["stages"]["抽取"] = round(sec, 2)
    if code != 0:
        stats["error"] = "抽取失败：\n" + tail(out, 15)
        return stats

    summ = MATCHER / "运行摘要.txt"
    if summ.exists():
        txt = summ.read_text(encoding="utf-8")
        stats["col_lines"] = [ln for ln in txt.splitlines() if "｜" in ln and "秒" in ln]

    # 写回
    env = {"A23_TEMPLATE_DIR": str(TPL_DIR), "A23_TEMPLATE_NAME": tpl}
    sec, code, out = run([PY, "fill.py", tpl], cwd=WRITER, env_extra=env)
    stats["stages"]["写回"] = round(sec, 2)
    stats["fill_code"] = code
    if code not in (0, 2):
        stats["error"] = "写回失败（返回码 %d）：\n%s" % (code, tail(out, 12))
        return stats

    # 验收
    sec, code, out = run([PY, "verify_fill.py", tpl], cwd=WRITER, env_extra=env)
    stats["stages"]["验收"] = round(sec, 2)
    stats["verify_code"] = code
    stats["verify_tail"] = tail(out, 10)

    stats["resp_sec"] = round(sum(stats["stages"].get(k, 0)
                                  for k in ("模板", "抽取", "写回", "验收")), 2)

    rep = WOUT / ("填表报告-%s.txt" % stem)
    if rep.exists():
        txt = rep.read_text(encoding="utf-8")
        m = re.search(r"【二、写入明细】共 (\d+) 格", txt)
        stats["written"] = int(m.group(1)) if m else None

    # 归档（每张表独立目录；抽取结果副本随后由归档步骤补）
    picks = [
        (WOUT / (stem + "-已填写.xlsx"), stem + "-已填写.xlsx"),
        (WOUT / ("验收报告-%s.txt" % stem), "验收报告.txt"),
        (WOUT / ("填表报告-%s.txt" % stem), "填表报告.txt"),
        (MATCHER / "运行摘要.txt", "运行摘要-抽取.txt"),
    ]
    for src_f, dst in picks:
        if src_f.exists():
            shutil.copy2(src_f, out_dir / dst)
    for col in cols:
        f = MATCHER / ("抽取结果-%s.json" % col)
        if f.exists():
            shutil.copy2(f, out_dir / f.name)

    return stats


# ==================================================================
# 金丝雀抽查（在还原前做：读 matcher 里的本轮抽取结果）
# ==================================================================
def norm(s):
    return re.sub(r"[\s，。、；：！？,.;:!?（）()【】\[\]「」『』\"'']+", "", str(s or ""))


def hit(want, val):
    a, b = norm(want), norm(val)
    if not a or not b:
        return False
    if len(a) <= 3 and a.isdigit():
        return a == b
    return a == b or a in b or b in a


def check_canaries():
    """核对 15 组金丝雀埋点（跨 5 张模板）。

    读取口径（2026-10-14 修正）：优先从「结果/模板N/」归档目录读该张表的
    抽取结果 —— 逐张跑 / 全跑完之后，MATCHER 目录里只剩最后一张表的列文件，
    前几张表的埋点列会被误报"没有结果文件"。归档目录里每张表的结果都完整
    保留（run_one 里已把 抽取结果-*.json 复制过去）。
    未命中时附诊断预览（该列抽到的前 3 条值），方便直接定位问题。
    """
    answer = json.loads(ANSWERS.read_text(encoding="utf-8"))
    rows = []
    for c in answer["canaries"]:
        expect = c["expect"]
        m = re.match(r"模板\s*(\d+)", c.get("tpl", ""))
        no = int(m.group(1)) if m else None
        out_dir = RESULT / ("模板%d" % no) if no else None
        if "_note" in expect:
            rows.append({"tpl": c["tpl"], "group": c["group"], "doc": c["doc"],
                         "form": c["form"], "status": "观察位（不判分）", "cols": {}})
            continue
        cols = {}
        for col, want in expect.items():
            f = None
            if out_dir is not None:
                cand = out_dir / ("抽取结果-%s.json" % col)
                if cand.exists():
                    f = cand
            if f is None:
                cand = MATCHER / ("抽取结果-%s.json" % col)
                if cand.exists():
                    f = cand
            if f is None:
                cols[col] = "没有结果文件（该表未跑或列名对不上）"
                continue
            data = json.loads(f.read_text(encoding="utf-8"))
            vals = [it.get("value", "") for it in data.get("items", [])]
            if any(hit(want, v) for v in vals):
                cols[col] = "命中"
            else:
                preview = "、".join(str(v)[:18] for v in vals[:3]) or "（空）"
                cols[col] = "未命中（%d 条：%s）" % (len(vals), preview)
        n_hit = sum(1 for v in cols.values() if v == "命中")
        n_all = len(cols)
        status = "通过" if n_hit == n_all else ("部分 %d/%d" % (n_hit, n_all) if n_hit else "未命中")
        rows.append({"tpl": c["tpl"], "group": c["group"], "doc": c["doc"],
                     "form": c["form"], "status": status, "cols": cols})
    return rows


# ==================================================================
# 主流程
# ==================================================================
def main():
    args = sys.argv[1:]
    # --only 支持多种写法：--only 4 ｜ --only 2,3 ｜ --only 2 3 ｜ 裸数字 4
    only = None
    only_set = set()
    for i, a in enumerate(args):
        if a.isdigit():
            only_set.add(int(a))
        elif a == "--only" and i + 1 < len(args):
            for x in args[i + 1].replace("，", ",").split(","):
                if x.strip().isdigit():
                    only_set.add(int(x.strip()))
    if only_set:
        only = sorted(only_set)
    skip_read = "--skip-read" in args
    jobs = [(i + 1, t) for i, t in enumerate(TEMPLATES)
            if only is None or (i + 1) in only]

    print("=" * 64)
    print("大文档压测 · 跑 %d 张模板 ｜ 时间 %s" % (len(jobs), time.strftime("%Y-%m-%d %H:%M")))
    print("=" * 64)
    print("口径：前置=读取 16 份大文档（一次性）；单表响应=模板理解+抽取+写回+验收")

    print("")
    if "--no-preflight" not in args:
        if precheck():
            print("")
            print("预检不通过 → 停下。")
            return
    else:
        print("（--no-preflight：跳过预检）")

    n = backup_all()
    print("")
    print("【备份】%d 个文件 → %s" % (n, BACKUP.name))

    # ---- 前置读取 ----
    read_info = None
    if skip_read and POOL.exists():
        data = json.loads(POOL.read_text(encoding="utf-8"))
        docs = data.get("documents", [])
        read_info = {"sec": None, "n_doc": len(docs), "code": 0,
                     "n_block": sum(d.get("block_count", 0) for d in docs),
                     "failed": len(data.get("failed", [])),
                     "pool_mb": round(POOL.stat().st_size / 1024 / 1024, 1)}
        read_info["by_type"] = {}
        for d_ in docs:
            t = d_.get("type", "?")
            read_info["by_type"].setdefault(t, [0, 0])
            read_info["by_type"][t][0] += 1
            read_info["by_type"][t][1] += d_.get("block_count", 0)
        print("")
        print("【前置】--skip-read：沿用小池 %s（%d 份 / %d 块 / %s MB）"
              % (POOL.name, read_info["n_doc"], read_info["n_block"], read_info["pool_mb"]))
    elif skip_read:
        print("")
        print("！--skip-read 但 %s 不存在，改为真读" % POOL.name)
        skip_read = False

    if not skip_read:
        print("")
        print("【前置】读取 16 份大文档（模拟「全部上传」，一次性）……")
        read_info = stage_read()
        if "error" in read_info:
            print("！" + read_info["error"])
            return
        print("  完成：%d 份 / %d 块 ｜ %.2f 秒 ｜ 池文件 %.1f MB ｜ 读取失败 %d"
              % (read_info["n_doc"], read_info["n_block"], read_info["sec"],
                 read_info["pool_mb"], read_info["failed"]))
        for t, (nf, nb) in sorted(read_info["by_type"].items()):
            print("    [%s] %d 份 / %d 块" % (t, nf, nb))

    # ---- 逐张跑 ----
    all_stats = []
    for no, tpl in jobs:
        print("")
        print("-" * 64)
        print("▶ 第 %d 张：%s" % (no, tpl))
        print("-" * 64)
        st = run_one(tpl, no)
        all_stats.append(st)
        if "error" in st:
            print("  ！失败：" + st["error"])
            continue
        sg = st["stages"]
        print("  模板 %.2fs → 抽取 %.2fs → 写回 %.2fs → 验收 %.2fs"
              % (sg.get("模板"), sg.get("抽取"), sg.get("写回"), sg.get("验收")))
        print("  单表响应：%.2f 秒（红线 90 秒）｜ 写入 %s 格 ｜ 验收 %s"
              % (st["resp_sec"], st.get("written", "?"),
                 "通过" if st.get("verify_code") == 0 else "返回码 %s" % st.get("verify_code")))

    # ---- 金丝雀抽查 ----
    print("")
    print("【金丝雀抽查】核对大池抽取结果里的埋点值……")
    canary_rows = check_canaries()
    for r in canary_rows:
        print("  [%s] %s ｜ %s ｜ %s" % (r["group"], r["tpl"], r["doc"], r["status"]))
        for col, s in r["cols"].items():
            if s != "命中":
                print("      - %s：%s" % (col, s))

    # ---- 报告 ----
    lines = []
    lines.append("=" * 64)
    lines.append("大文档压测报告（%s）" % time.strftime("%Y-%m-%d %H:%M"))
    lines.append("=" * 64)
    lines.append("")
    lines.append("赛题规格：5×≥500KB docx + 5×≥500KB xlsx + 3×≥15KB md + 3×≥15KB txt = 16 份")
    lines.append("口径：前置=读取 16 份大文档（一次性）｜ 单表响应=模板理解+抽取+写回+验收（红线 90 秒）")
    lines.append("")
    if read_info:
        lines.append("【前置】读取 16 份大文档：%s ｜ %d 份 / %d 块 ｜ 池文件 %.1f MB ｜ 读取失败 %d"
                     % (("%.2f 秒" % read_info["sec"]) if read_info.get("sec") else "（沿用已生成大池）",
                        read_info["n_doc"], read_info["n_block"], read_info["pool_mb"], read_info["failed"]))
        for t, (nf, nb) in sorted(read_info["by_type"].items()):
            lines.append("    [%s] %d 份 / %d 块" % (t, nf, nb))
    lines.append("")
    lines.append("【逐张成绩】")
    lines.append("-" * 64)
    for st in all_stats:
        if "error" in st:
            lines.append("%s ｜ 失败：%s" % (st["tpl"], st["error"].splitlines()[0]))
            continue
        sg = st["stages"]
        lines.append("%s（%d 列）" % (st["tpl"], len(st.get("cols", []))))
        lines.append("  响应 %.2f 秒 = 模板 %.2f + 抽取 %.2f + 写回 %.2f + 验收 %.2f ｜ 写入 %s 格 ｜ 验收 %s"
                     % (st["resp_sec"], sg.get("模板"), sg.get("抽取"), sg.get("写回"), sg.get("验收"),
                        st.get("written", "?"),
                        "通过" if st.get("verify_code") == 0 else "返回码 %s" % st.get("verify_code")))
    lines.append("")
    lines.append("【逐列耗时明细】")
    for st in all_stats:
        if st.get("col_lines"):
            lines.append("  —— %s ——" % st["tpl"])
            for ln in st["col_lines"]:
                lines.append("    " + ln)
    lines.append("")
    lines.append("【金丝雀抽查（大池埋点 15 组）】")
    lines.append("-" * 64)
    for r in canary_rows:
        lines.append("[%s] %s ｜ %s" % (r["group"], r["tpl"], r["status"]))
        lines.append("    埋点：%s ｜ %s" % (r["doc"], r["form"]))
        for col, s in r["cols"].items():
            lines.append("      %s：%s" % (col, s))
    n_pass = len([r for r in canary_rows if r["status"] == "通过"])
    n_part = len([r for r in canary_rows if r["status"].startswith("部分")])
    n_miss = len([r for r in canary_rows if r["status"] == "未命中"])
    n_obs = len([r for r in canary_rows if r["status"].startswith("观察")])
    lines.append("小结：通过 %d ｜ 部分 %d ｜ 未命中 %d ｜ 观察位 %d" % (n_pass, n_part, n_miss, n_obs))
    lines.append("")
    lines.append("【红线对照】")
    ok = [st for st in all_stats if "error" not in st]
    over = [st for st in ok if st["resp_sec"] > 90]
    if over:
        lines.append("  超红线：%s" % "、".join(st["tpl"] for st in over))
    else:
        lines.append("  %d 张全部 ≤90 秒 ✓" % len(ok))
    lines.append("=" * 64)
    report = "\n".join(lines)
    if only is None:
        (ROOT / "压测报告.txt").write_text(report, encoding="utf-8")
    else:
        # --only 单张模式：小结存该张归档目录（结构化 json），不覆盖总报告
        for st in all_stats:
            d = RESULT / ("模板%d" % st["no"])
            d.mkdir(parents=True, exist_ok=True)
            (d / "压测小结.json").write_text(
                json.dumps(st, ensure_ascii=False, indent=2), encoding="utf-8")

    # ---- 还原 ----
    print("")
    restore_all()
    print("【还原】已恢复日常小池状态（压测产物在 结果/ 目录）")
    print("")
    print(report)
    print("")
    if only is None:
        print("（报告已存：压测-大文档/压测报告.txt）")
    else:
        print("（部分模式：小结已存 结果/模板N/压测小结.json）")


if __name__ == "__main__":
    main()
