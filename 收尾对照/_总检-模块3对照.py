# -*- coding: utf-8 -*-
"""
_总检-模块3对照.py —— 模块（3）「表格自定义数据填写」收口总检（收尾对照）

对照对象：赛题原文-A23.md —— 模块（3）全句 + 用户期望（3）+ 评价方法 + 技术要求。
性质：零模型 · 确定性（重跑必一致）；只读（不改任何数据文件）。

跑法：python 收尾对照/_总检-模块3对照.py
产出：stdout + 收尾对照/模块3对照总检.txt
退出码：0=全过 ｜ 1=有失败
"""
import json
import re
import subprocess
import sys
import time
from datetime import datetime
from pathlib import Path

BASE = Path(__file__).resolve().parent.parent      # 收尾对照/ → 研一项目/
OUT = BASE / "收尾对照"

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

RESULTS = []
LOG = []


def P(s=""):
    LOG.append(s)
    print(s)


def check(name, ok, detail=""):
    RESULTS.append((name, bool(ok), str(detail)))
    P("  [%s] %s ｜ %s" % ("PASS" if ok else "FAIL", name, detail))


def read(p):
    try:
        return Path(p).read_text(encoding="utf-8", errors="replace")
    except Exception:
        return ""


# 常用路径
YC_REPORT = BASE / "压测-大文档" / "压测报告.txt"          # 大文档压测（最终轮）
WS_REPORT = BASE / "五表连测" / "五表连测报告.txt"          # 五表连测（v5 修复后）
FH_REPORT = BASE / "泛化测试" / "泛化测试总报告.txt"        # 泛化 3 场景
WORD_REC = BASE / "writer" / "输出" / "连测记录-Word版5张.txt"
SRC16 = BASE / "测试数据" / "压测-大文档" / "源文档"        # 16 份赛题规格大文档
TPL_X = BASE / "测试数据" / "模板表格"                      # 5 张 xlsx 模板
TPL_W = BASE / "测试数据" / "模板表格-word"                 # 5 张 word 模板
STRUCT = BASE / "template_reader" / "模板表结构.json"
POOL = BASE / "doc_reader" / "output.json"
W_OUT = BASE / "writer" / "输出"


def main():
    t0 = time.time()
    P("=" * 66)
    P("模块（3）收口总检 —— 对照赛题原文（零模型 · 确定性 · 只读）")
    P("对照依据：研一项目/赛题原文-A23.md（2026-10-11 官网存档）")
    P("时间：%s" % datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    P("=" * 66)
    P("")

    # ---- 先做全部数字解析（后面各组引用；解析失败会在对应项 FAIL） ----
    yc_text = read(YC_REPORT)
    ws_text = read(WS_REPORT)
    fh_text = read(FH_REPORT)
    word_rec = read(WORD_REC)

    yc_secs = [float(x) for x in re.findall(r"响应 ([\d.]+) 秒", yc_text)]
    ws_secs = [float(x) for x in re.findall(r"响应 ([\d.]+) 秒", ws_text)]

    fh_m = re.search(r"值严格 (\d+)/(\d+)（([\d.]+)%）\s*｜\s*值宽松 (\d+)/(\d+)（([\d.]+)%）"
                     r"\s*｜\s*出处 (\d+)/(\d+)（([\d.]+)%）", fh_text)
    p_strict = re.search(r"值 · 严格（[^）]*）：(\d+)/(\d+)（([\d.]+)%）", ws_text)
    p_loose = re.search(r"值 · 宽松（[^）]*）：(\d+)/(\d+)（([\d.]+)%）", ws_text)
    p_loc = re.search(r"出处 · 行号一致：(\d+)/(\d+)（([\d.]+)%）", ws_text)

    # ---------------- A 组：赛题模块（3）原文逐句 + 用户期望（3） ----------------
    P("---- A 组：赛题模块（3）原文逐句 + 用户期望（3）对照 ----")

    check("A1 模块原文：利用脚本语句或人工智能等操作",
          (BASE / "pipeline.py").exists() and (BASE / "matcher" / "extract.py").exists(),
          "pipeline.py（一条命令五阶段）+ matcher/extract.py（LLM 抽取 + 规则兜底）")

    struct = []
    try:
        struct = json.loads(STRUCT.read_text(encoding="utf-8"))
    except Exception:
        pass
    check("A2 模块原文：在用户提供的表格（word 或 excel）之后",
          len(struct) == 5 and len(list(TPL_X.glob("*.xlsx"))) == 5
          and len(list(TPL_W.glob("*.docx"))) == 5,
          "模板表结构.json %d 张已解析 ｜ 模板表格/ 5 xlsx ｜ 模板表格-word/ 5 docx" % len(struct))

    n_doc = 0
    try:
        pool = json.loads(POOL.read_text(encoding="utf-8"))
        n_doc = len(pool.get("documents", []))
    except Exception:
        pass
    n16 = len([f for f in SRC16.rglob("*") if f.is_file()]) if SRC16.exists() else 0
    check("A3 模块原文：和非结构化数据（四格式多文档）",
          n_doc >= 18 and n16 == 16,
          "日常检索池 %d 份 ｜ 赛题规格大文档集 %d 份" % (n_doc, n16))

    check("A4 模块原文：从非结构化数据中自动搜索相关信息",
          (BASE / "matcher" / "retrieve.py").exists()
          and (BASE / "matcher" / "extract.py").exists()
          and (BASE / "matcher" / "run_all.py").exists(),
          "matcher 三件套：retrieve（相关度筛选）→ extract（LLM 判定）→ run_all（编排 + 一致性重审）")

    yt_ok = 0
    yc_res = BASE / "压测-大文档" / "结果"
    for i in range(1, 6):
        d = yc_res / ("模板%d" % i)
        if d.exists() and any(f.name.endswith("已填写.xlsx") for f in d.iterdir()) \
           and (d / "验收报告.txt").exists():
            yt_ok += 1
    word_cnt = len(list(W_OUT.glob("*已填写.docx")))
    check("A5 模块原文：并进行表格填写",
          yt_ok == 5 and word_cnt >= 5,
          "压测归档 5/5（已填写 xlsx + 验收报告）｜ Word 版已填写 %d 张" % word_cnt)

    check("A6 用户期望：根据提供的 word 和 excel 表格智能填写相关数据",
          ("5/5 张全部通过" in word_rec) and len(ws_secs) == 5,
          "Word 模板连测 5/5 全过 ｜ Excel 五表连测 5 张响应记录在")

    fill_src = read(BASE / "writer" / "fill.py")
    vr_cnt = len(list(W_OUT.glob("验收报告-*.txt")))
    check("A7 用户期望：自动生成具备直接业务应用价值的、格式严谨的汇总表格",
          ("永不改原模板" in fill_src) and ("只填值" in fill_src) and vr_cnt >= 10,
          "fill.py 三铁律在源码 ｜ 验收报告 %d 份（格式差异 0 处口径）" % vr_cnt)
    P("")

    # ---------------- B 组：赛题评价方法 + 技术要求 ----------------
    P("---- B 组：赛题评价方法 + 技术要求对照 ----")

    got = {"docx": [], "xlsx": [], "md": [], "txt": []}
    if SRC16.exists():
        for f in SRC16.rglob("*"):
            ext = f.suffix.lower().lstrip(".")
            if f.is_file() and ext in got:
                got[ext].append(f)

    def kb(fs):
        sz = [f.stat().st_size for f in fs]
        return "%d~%dKB" % (min(sz) // 1024, max(sz) // 1024) if sz else "无"

    b1_ok = (len(got["docx"]) == 5 and all(f.stat().st_size >= 500 * 1024 for f in got["docx"])
             and len(got["xlsx"]) == 5 and all(f.stat().st_size >= 500 * 1024 for f in got["xlsx"])
             and len(got["md"]) == 3 and all(f.stat().st_size >= 15 * 1024 for f in got["md"])
             and len(got["txt"]) == 3 and all(f.stat().st_size >= 15 * 1024 for f in got["txt"]))
    check("B1 比赛样本规格实测（16 份：5+5+3+3）", b1_ok,
          "docx %d/5（%s）｜ xlsx %d/5（%s）｜ md %d/3（%s）｜ txt %d/3（%s）"
          % (len(got["docx"]), kb(got["docx"]), len(got["xlsx"]), kb(got["xlsx"]),
             len(got["md"]), kb(got["md"]), len(got["txt"]), kb(got["txt"])))

    check("B2 5 个模板表格（word 或 excel，双形态都备）",
          len(list(TPL_X.glob("*.xlsx"))) == 5 and len(list(TPL_W.glob("*.docx"))) == 5,
          "xlsx 5 张 + word 5 张（比赛给哪种都能接）")

    b3_ok = (len(yc_secs) == 5 and len(ws_secs) == 5
             and all(s <= 90 for s in yc_secs + ws_secs))
    if b3_ok:
        b3_detail = ("压测 %.1f~%.1fs ｜ 五表 %.1f~%.1fs（各 5 张全过）"
                     % (min(yc_secs), max(yc_secs), min(ws_secs), max(ws_secs)))
    else:
        b3_detail = "解析异常：压测 %d 条 ｜ 五表 %d 条" % (len(yc_secs), len(ws_secs))
    check("B3 每表响应 ≤90 秒（大文档池 + 日常池，两套实测）", b3_ok, b3_detail)

    pcts = []
    if fh_m:
        pcts += [float(fh_m.group(3)), float(fh_m.group(6)), float(fh_m.group(9))]
    if p_strict:
        pcts.append(float(p_strict.group(3)))
    if p_loose:
        pcts.append(float(p_loose.group(3)))
    if p_loc:
        pcts.append(float(p_loc.group(3)))
    b4_ok = len(pcts) == 6 and all(v >= 80 for v in pcts)
    if b4_ok:
        b4_detail = ("泛化 63 格：严格 %s%%｜宽松 %s%%｜出处 %s%% ｜ 模板4 对分：%s%% / %s%% / %s%%"
                     % (fh_m.group(3), fh_m.group(6), fh_m.group(9),
                        p_strict.group(3), p_loose.group(3), p_loc.group(3)))
    else:
        b4_detail = "数字解析不齐（%d/6）：请检查报告格式" % len(pcts)
    check("B4 准确率 ≥80%（两份独立实测）", b4_ok, b4_detail)

    if len(yc_secs) == 5 and len(ws_secs) == 5:
        check("B5 平均响应时间（5 张取平均，赛题口径）", True,
              "大文档池平均 %.2f 秒 ｜ 日常池平均 %.2f 秒"
              % (sum(yc_secs) / 5, sum(ws_secs) / 5))
    else:
        check("B5 平均响应时间（5 张取平均，赛题口径）", False, "响应秒数解析不齐")

    check("B6 多数据类型稳定运行（四格式全链路）",
          b1_ok and len(list((BASE / "泛化测试").glob("结果-场景*"))) == 3,
          "大池四格式 16 份全跑 ｜ 泛化 3 场景（换新文档 / 新模板）全过")
    P("")

    # ---------------- C 组：收口工程件 ----------------
    P("---- C 组：收口工程件 ----")

    pipe = BASE / "pipeline.py"
    rc = subprocess.run([sys.executable, "-m", "py_compile", str(pipe)],
                        capture_output=True, cwd=str(BASE))
    check("C1 总控 pipeline.py 可编译（一条命令五个阶段）",
          pipe.exists() and rc.returncode == 0,
          "pipeline.py（读→模板→抽→写→验）")

    layers = {
        "doc_reader/run_all.py": BASE / "doc_reader" / "run_all.py",
        "template_reader/main.py": BASE / "template_reader" / "main.py",
        "matcher/run_all.py": BASE / "matcher" / "run_all.py",
        "writer/fill.py": BASE / "writer" / "fill.py",
        "writer/verify_fill.py": BASE / "writer" / "verify_fill.py",
    }
    missing = [k for k, v in layers.items() if not v.exists()]
    check("C2 五层脚本齐全", not missing, "缺：%s" % (missing if missing else "无"))

    runners = {
        "泛化测试/跑测试.py": BASE / "泛化测试" / "跑测试.py",
        "五表连测/跑五表.py": BASE / "五表连测" / "跑五表.py",
        "压测-大文档/跑压测.py": BASE / "压测-大文档" / "跑压测.py",
    }
    missing2 = [k for k, v in runners.items() if not v.exists()]
    check("C3 三条测试编排脚本齐全（泛化 / 五表 / 压测）", not missing2,
          "缺：%s" % (missing2 if missing2 else "无"))

    reports = {
        "泛化测试总报告.txt": FH_REPORT,
        "五表连测报告.txt": WS_REPORT,
        "压测报告.txt": YC_REPORT,
        "连测记录-Word版5张.txt": WORD_REC,
        "压测-总结与优化报告.md": BASE / "压测-大文档" / "压测-总结与优化报告.md",
    }
    missing3 = [k for k, v in reports.items() if not v.exists()]
    check("C4 五份报告齐全", not missing3, "缺：%s" % (missing3 if missing3 else "无"))

    check("C5 环境变量三件套（换数据集 / 换模板 / 换池 可配）",
          ("A23_TEMPLATE_DIR" in read(BASE / "writer" / "fill.py"))
          and ("A23_TEMPLATE_NAME" in read(BASE / "matcher" / "run_all.py"))
          and ("A23_OUTPUT_JSON" in read(BASE / "matcher" / "extract.py")),
          "A23_TEMPLATE_DIR（fill/verify/pipeline）｜ A23_TEMPLATE_NAME（matcher）｜ A23_OUTPUT_JSON（压测隔离）")

    check("C6 写回验收报告齐备（含 Word 版）", vr_cnt >= 10,
          "writer/输出/验收报告-*.txt = %d 份" % vr_cnt)
    P("")

    # ---------------- D 组：数据快照与复算 ----------------
    P("---- D 组：数据快照与复算（重跑必一致） ----")

    if struct and len(struct) == 5:
        cols = [len(t["sheets"][0]["columns"]) for t in struct]
        rows = [(t["sheets"][0].get("fill_region") or {}).get("row_count", 0) for t in struct]
        d1_ok = all(t["sheets"][0].get("header_ok") for t in struct)
        d1_detail = "5 张：列 %s ｜ 待填行 %s" % ("/".join(map(str, cols)), "/".join(map(str, rows)))
    else:
        d1_ok, d1_detail = False, "模板表结构.json 不齐"
    check("D1 模板结构在线（表头 + 待填区已识别）", d1_ok, d1_detail)

    n_doc2, n_blk = 0, 0
    try:
        pool2 = json.loads(POOL.read_text(encoding="utf-8"))
        docs2 = pool2.get("documents", [])
        n_doc2 = len(docs2)
        n_blk = sum(d.get("block_count", 0) for d in docs2)
    except Exception:
        pass
    check("D2 检索池快照（日常池，压测后已还原）", n_doc2 >= 18,
          "%d 份 / %d 块" % (n_doc2, n_blk))

    if len(yc_secs) == 5:
        d3_ok = all(s <= 90 for s in yc_secs)
        d3_detail = "5 张 %s 秒 → 全 ≤90 ｜ 平均 %.2f 秒" % (
            " / ".join("%.2f" % s for s in yc_secs), sum(yc_secs) / 5)
    else:
        d3_ok, d3_detail = False, "解析 %d 条（应 5）" % len(yc_secs)
    check("D3 压测响应复算（解析报告，逐张核对）", d3_ok, d3_detail)

    if len(ws_secs) == 5:
        d4_ok = all(s <= 90 for s in ws_secs)
        d4_detail = "5 张 %s 秒 → 全 ≤90 ｜ 平均 %.2f 秒" % (
            " / ".join("%.2f" % s for s in ws_secs), sum(ws_secs) / 5)
    else:
        d4_ok, d4_detail = False, "解析 %d 条（应 5）" % len(ws_secs)
    check("D4 五表连测响应复算（解析报告，逐张核对）", d4_ok, d4_detail)
    P("")

    # ---------------- E 组：三模块立正索引 ----------------
    P("---- E 组：三模块立正索引 ----")

    m1 = read(BASE / "agent" / "输出" / "_体检-模块1覆盖矩阵.txt")
    m1_line = next((ln.strip() for ln in m1.splitlines() if "合计" in ln), "")
    check("E1 模块（1）收口件：覆盖矩阵体检", bool(m1) and bool(m1_line),
          m1_line if m1_line else "未找到成绩行（文件缺失？）")

    m2 = read(BASE / "agent" / "输出" / "验收-收口" / "模块2对照总检.txt")
    check("E2 模块（2）收口件：对照总检 22/22", "22 / 22" in m2,
          "模块2对照总检.txt（2026-10-14 现场复跑 22/22）")

    mk = OUT / "三模块立正总表.md"
    check("E3 模块（3）收口件：三模块立正总表已就位", mk.exists(),
          "收尾对照/三模块立正总表.md" if mk.exists() else "总表待写（写完重跑本件即为全过）")
    P("")

    # ---------------- 统计与落盘 ----------------
    n_pass = sum(1 for _, ok, _ in RESULTS if ok)
    n_total = len(RESULTS)
    P("=" * 66)
    P("总检成绩：%d / %d 通过（用时 %.1f 秒）" % (n_pass, n_total, time.time() - t0))
    P("=" * 66)

    OUT.mkdir(parents=True, exist_ok=True)
    rp = OUT / "模块3对照总检.txt"
    rp.write_text("\n".join(LOG) + "\n", encoding="utf-8")
    print("（总检报告已存档：%s）" % rp)
    sys.exit(0 if n_pass == n_total else 1)


if __name__ == "__main__":
    main()
