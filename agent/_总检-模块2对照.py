# -*- coding: utf-8 -*-
"""
_总检-模块2对照.py —— 模块（2）「非结构化文档信息提取」收口总检（模块2 · 第6步）

对照对象：赛题原文-A23.md —— 模块（2）全句 + 技术要求 + 评价方法。
性质：零模型 · 确定性（重跑必一致）；只读（不改任何数据）。

跑法：python agent/_总检-模块2对照.py
产出：stdout + 输出/验收-收口/模块2对照总检.txt
退出码：0=全过 ｜ 1=有失败 ｜ 2=环境不可用
"""
import json
import os
import subprocess
import sys
import time
from datetime import datetime
from pathlib import Path

BASE = Path(__file__).resolve().parent
PROJECT = BASE.parent
OUT = BASE / "输出"
ARC = OUT / "验收-收口"
sys.path.insert(0, str(BASE))

try:
    sys.stdout.reconfigure(errors="replace")
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


def main():
    t0 = time.time()
    P("=" * 66)
    P("模块（2）收口总检 —— 对照赛题原文（零模型 · 确定性 · 只读）")
    P("对照依据：研一项目/赛题原文-A23.md（2026-10-11 官网存档）")
    P("时间：%s" % datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    P("=" * 66)
    P("")

    import store
    import quality
    import pool

    try:
        s = store.stats()
    except Exception as e:
        P("【致命】数据库不可用：%s" % e)
        sys.exit(2)
    st = pool.status()
    audit = quality.audit_all()
    scores = quality.score_all()
    summ = quality.summarize(audit, scores)
    a = summ["audit"]
    sc = summ["scores"]
    by = st.get("by_type") or {}

    src_api = read(BASE / "api_server.py")
    src_web = read(BASE / "web" / "index.html")
    src_pool = read(BASE / "pool.py")

    # ---------------- A 组：模块（2）原文逐句 ----------------
    P("---- A 组：赛题模块（2）原文逐句对照 ----")

    check("A1 基于 Web 网站部署（服务 + 网页操作台）",
          ("from fastapi" in src_api) and ("/api/upload" in src_api)
          and ("导入文档" in src_web) and ("数据仓库" in src_web),
          "api_server.py（FastAPI）+ web/index.html（上传/指令/仓库在）")

    check("A2 自动识别用户导入的非结构化文档",
          ("from readers import" in src_pool) and ("read_one" in src_pool)
          and st.get("total", 0) >= 18
          and all(by.get(k, 0) > 0 for k in ("docx", "xlsx", "md", "txt")),
          "pool.py 复用读取器自动识别（read_one/SUPPORTED）｜ 池 %d 份（docx %d · xlsx %d · md %d · txt %d）"
          % (st.get("total", 0), by.get("docx", 0), by.get("xlsx", 0), by.get("md", 0), by.get("txt", 0)))

    check("A3 提取文件文本关键信息（文档画像）",
          (BASE / "summary_ops.py").exists() and s["profiles"] > 0,
          "summary_ops.py 在 ｜ 库内画像 %d 份（明细 %d 条）" % (s["profiles"], s["profile_items"]))

    check("A4 提取实体数据（7 类）",
          (BASE / "entity_ops.py").exists() and s["entities"] > 0,
          "entity_ops.py 在 ｜ 库内实体 %d 条" % s["entities"])

    check("A5 提取用户指定内容（点名字段）",
          (BASE / "extract_ops.py").exists() and s["extractions"] > 0,
          "extract_ops.py 在 ｜ 库内字段提取 %d 条" % s["extractions"])

    check("A6 进行数据库存储操作（五张表）",
          s["documents"] > 0 and s["entities"] > 0 and s["profiles"] > 0 and s["extractions"] > 0,
          "documents %d · entities %d · profiles %d · profile_items %d · extractions %d"
          % (s["documents"], s["entities"], s["profiles"], s["profile_items"], s["extractions"]))

    check("A7 不存在数据识别误差（保真度 100%）",
          a["fidelity"] == 100.0,
          "全库 %d 条 → 可追溯 %d 条 ｜ 保真度 %s%% ｜ 片段命中 %s%%"
          % (a["total"], a["traced"], a["fidelity"], a["frag_rate"]))

    r1 = store._q("SELECT COUNT(*) AS n FROM entities "
                  "WHERE quote IS NULL OR quote='' OR loc IS NULL OR loc=''")[0]["n"]
    r2 = store._q("SELECT COUNT(*) AS n FROM extractions "
                  "WHERE evidence IS NULL OR evidence='' OR loc IS NULL OR loc=''")[0]["n"]
    check("A8 入库规范性（每条带原文摘抄 + 出处）",
          r1 + r2 == 0,
          "缺摘抄/出处的条数：实体 %d · 字段 %d" % (r1, r2))
    P("")

    # ---------------- B 组：技术要求 + 评价方法 ----------------
    P("---- B 组：赛题技术要求 + 评价方法对照 ----")

    check("B1 准确率高于 80%（赛题硬指标）",
          all(v is not None and v >= 80 for v in (sc["entities"], sc["profile"], sc["fields"])),
          "三线查全：实体 %s%% ｜ 画像 %s%% ｜ 字段 %s%%（对照人工标准答案）"
          % (sc["entities"], sc["profile"], sc["fields"]))

    check("B2 智能识别模块提供 API 接口、支持异步调用",
          ("ThreadPoolExecutor" in src_api) and ("/api/job/" in src_api)
          and ("/api/command" in src_api) and ("/api/recognize" in src_api),
          "POST /api/recognize（识别分类专属异步接口）→ job_id → GET /api/job/{id}（线程池 4 worker）")

    check("B3 多数据类型稳定运行（四格式 ≥3 份）",
          all(by.get(k, 0) >= 3 for k in ("docx", "xlsx", "md", "txt")),
          "池内 docx %d · xlsx %d · md %d · txt %d"
          % (by.get("docx", 0), by.get("xlsx", 0), by.get("md", 0), by.get("txt", 0)))

    ans_path = BASE / "标准答案-模块2.json"
    try:
        ans = json.loads(ans_path.read_text(encoding="utf-8"))
        n_ent = sum(len(v) for doc in (ans.get("entities") or {}).values() for v in doc.values())
        n_prof = sum(len(v) for v in (ans.get("profile") or {}).values())
        n_fld = sum(len(v) for v in (ans.get("fields") or {}).values())
        n_fld_vals = sum(len(f.get("expect") or [])
                         for v in (ans.get("fields") or {}).values() for f in v)
        ok_ans = all(k in ans for k in ("entities", "profile", "fields"))
    except Exception:
        n_ent = n_prof = n_fld = n_fld_vals = 0
        ok_ans = False
    check("B4 人工标注与多轮审核（标准答案档案）",
          ok_ans and n_ent > 0 and n_prof > 0 and n_fld > 0,
          "标准答案-模块2.json：实体 %d 项 · 画像 %d 点 · 字段 %d 条目（%d 个值）"
          % (n_ent, n_prof, n_fld, n_fld_vals))

    check("B5 便于部署和使用（双击即用入口）",
          (PROJECT / "启动操作台.bat").exists() and (BASE / "演示-模块2.py").exists(),
          "启动操作台.bat（网页）+ 演示-模块2.py（命令行全链路）")
    P("")

    # ---------------- C 组：收口工程件 ----------------
    P("---- C 组：收口工程件 ----")

    demo = BASE / "演示-模块2.py"
    rc_c = subprocess.run([sys.executable, "-m", "py_compile", str(demo)],
                          capture_output=True, cwd=str(PROJECT))
    check("C1 演示脚本存在且可编译", demo.exists() and rc_c.returncode == 0,
          "agent/演示-模块2.py")

    env = {**os.environ, "PYTHONIOENCODING": "utf-8"}
    try:
        r = subprocess.run([sys.executable, str(demo), "--fast", "--no-save"],
                           capture_output=True, text=True, encoding="utf-8",
                           errors="replace", cwd=str(PROJECT), env=env, timeout=120)
        rc, so = r.returncode, (r.stdout or "")
    except Exception as e:
        rc, so = -1, str(e)
    check("C2 演示脚本 --fast 跑通（退出码 0）",
          rc == 0 and "全链路通过" in so, "退出码 %s" % rc)

    check("C3 演示报告已存档", (OUT / "演示-模块2报告.txt").exists(),
          "输出/演示-模块2报告.txt")

    qh = OUT / "质量报告" / "质量报告.html"
    qj = OUT / "质量报告" / "质量明细.json"
    html = read(qh)
    check("C4 质量报告双件（HTML + 明细 JSON）",
          qh.exists() and qj.exists() and ("保真度" in html),
          "输出/质量报告/质量报告.html（%d 字节）" % (qh.stat().st_size if qh.exists() else 0))

    exp_scripts = ["_验收-格式版.py", "_验收-提取版.py", "_验收-会话版.py", "_验收-API版.py",
                   "_验收-xlsx操作.py", "_验收-格式扩展版.py", "_验收-导入版.py",
                   "_验收-画像版.py", "_验收-实体版.py", "_验收-入库版.py", "_验收-质量版.py"]
    missing = [n for n in exp_scripts if not (BASE / n).exists()]
    check("C5 11 个验收脚本齐全（模块1 六 + 模块2 五）", not missing,
          "缺：%s" % (missing if missing else "无"))
    P("")

    # ---------------- D 组：数据现状 + 一致性 ----------------
    P("---- D 组：数据现状与一致性（快照）----")

    check("D1 检索池在线（标准 18 份基线）", st.get("total", 0) >= 18,
          "池 %d 份 ｜ 上传件 %d ｜ 解析失败 %d"
          % (st.get("total", 0), len(st.get("imported") or []), st.get("failed", 0)))

    check("D2 数据仓库五表在线",
          s["documents"] >= 9 and s["entities"] > 0 and s["profiles"] > 0 and s["extractions"] > 0,
          "文档 %d · 实体 %d · 画像 %d · 字段 %d（最近入库 %s）"
          % (s["documents"], s["entities"], s["profiles"], s["extractions"], s.get("last_ts")))

    check("D3 质量复算（零模型反向审计）",
          a["fidelity"] == 100.0 and a["issues"] == 0,
          "保真度 %s%% ｜ 可疑条数 %d ｜ 复算 %.2f 秒" % (a["fidelity"], a["issues"], audit["elapsed"]))

    try:
        j = json.loads(qj.read_text(encoding="utf-8"))
        rep_total = ((j.get("summary") or {}).get("audit") or {}).get("total")
    except Exception:
        rep_total = None
    check("D4 质量报告与当前库对齐（数字一致）", rep_total == a["total"],
          "报告 total=%s ｜ 复算 total=%s" % (rep_total, a["total"]))
    P("")

    # ---------------- 统计与落盘 ----------------
    n_pass = sum(1 for _, ok, _ in RESULTS if ok)
    n_total = len(RESULTS)
    P("=" * 66)
    P("总检成绩：%d / %d 通过（用时 %.1f 秒）" % (n_pass, n_total, time.time() - t0))
    P("=" * 66)

    ARC.mkdir(parents=True, exist_ok=True)
    rp = ARC / "模块2对照总检.txt"
    rp.write_text("\n".join(LOG) + "\n", encoding="utf-8")
    print("（总检报告已存档：%s）" % rp)
    sys.exit(0 if n_pass == n_total else 1)


if __name__ == "__main__":
    main()
