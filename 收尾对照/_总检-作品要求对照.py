# -*- coding: utf-8 -*-
"""
_总检-作品要求对照.py —— 赛题原文【用户期望】+【技术要求与指标】逐条对照
（零模型 · 确定性 · 只读 —— 除本报告外不改动任何文件）

对照依据：研一项目/赛题原文-A23.md
  §5 项目说明【用户期望】
      （1）智能化应用：足够智能 + 合理的工具使用引导 + 初步接触即可轻松使用
      （2）文本信息自动化识别系统：读文档 → 提取关键内容/用户需求内容 → 存储（保真 + 规范）
      （3）智能信息整合系统：根据 word / excel 表格智能填写 → 格式严谨的汇总表格
  §6 任务要求【技术要求与指标】
      ① 程序形态不限（H5 / App / Web / PC）· 可基于第三方商业 AI 平台 · 可自研
      ② 准确识别数据类型 + 多类型稳定运行 · 准确率高于 80%
      ③ 训练与验证样本均经过人工标注与多轮审核
      ④ 智能识别分类模块提供 API 接口，支持异步调用

口径：每项给出「实现载体 + 证据 + PASS/FAIL」。可随时复跑（零模型、几秒级）。
用法：python 收尾对照/_总检-作品要求对照.py
"""
import re
import sys
import time
from datetime import datetime
from pathlib import Path

HERE = Path(__file__).resolve().parent          # = 收尾对照/
PROJECT = HERE.parent                           # = 研一项目/
AGENT = PROJECT / "agent"
sys.path.insert(0, str(AGENT))

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

T0 = time.time()
REPORT = []
N_PASS = 0
N_ALL = 0


def P(s=""):
    REPORT.append(str(s))
    print(s)


def check(name, ok, detail=""):
    global N_PASS, N_ALL
    N_ALL += 1
    ok = bool(ok)
    if ok:
        N_PASS += 1
    P("  [%s] %s%s" % ("PASS" if ok else "FAIL", name,
                       ("  ｜ " + str(detail)) if detail else ""))
    return ok


def read(p):
    try:
        return Path(p).read_text(encoding="utf-8", errors="replace")
    except Exception:
        return ""


def main():
    P("=" * 66)
    P("A23 作品要求对照总检 ——【用户期望】3 条 +【技术要求与指标】4 条")
    P("（零模型 · 确定性 · 只读）")
    P("对照依据：研一项目/赛题原文-A23.md（§5 用户期望 / §6 技术要求与指标）")
    P("时间：%s" % datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    P("=" * 66)
    P("")

    # ---------------- 载入源码与真值 ----------------
    src_intent = read(AGENT / "parse_intent.py")
    src_edit = read(AGENT / "edit.py")
    src_api = read(AGENT / "api_server.py")
    src_web = read(AGENT / "web" / "index.html")
    src_runall = read(PROJECT / "matcher" / "run_all.py")

    m_model = re.search(r'MODEL\s*=\s*"([^"]+)"', src_intent)
    model_name = m_model.group(1) if m_model else "（未识别）"

    # routes 用 import 真值（比抓源码字符串更硬）；导入失败退化为文本检查
    route_set = set()
    api_import_ok = False
    has_submit_rec = has_submit_fill = False
    try:
        import api_server
        route_set = {getattr(r, "path", "") for r in api_server.app.routes}
        api_import_ok = True
        has_submit_rec = hasattr(api_server.Hub, "submit_recognize")
        has_submit_fill = hasattr(api_server.Hub, "submit_fill")
    except Exception as e:
        P("（提示：api_server 导入失败——%s；改用源码文本检查）" % e)
    has_submit_rec = has_submit_rec or ("def submit_recognize" in src_api)
    has_submit_fill = has_submit_fill or ("def submit_fill" in src_api)

    def route(p):
        return (p in route_set) or (p in src_api)

    # 数据库 / 池 / 质量（不可用时相关项 FAIL 并注明原因）
    db_ok = False
    db_msg = ""
    s = {}
    a = {}
    sc = {}
    pool_st = {}
    by = {}
    store = quality = None
    try:
        import store as _store
        import quality as _quality
        import pool as _pool
        store, quality = _store, _quality
        s = store.stats()
        pool_st = _pool.status()
        by = pool_st.get("by_type") or {}
        # 一次复算：反向审计（保真度）+ 三线查全（对照人工标准答案）
        summ = quality.summarize(quality.audit_all(), quality.score_all())
        a, sc = summ["audit"], summ["scores"]
        db_ok = True
    except Exception as e:
        db_msg = "%s: %s" % (type(e).__name__, e)

    # ================= A 组：用户期望 =================
    P("---- A 组：用户期望（赛题原文 §5） ----")

    check("A1 智能化应用 · 智能底座已接入（大模型 + 规则引擎）",
          ("deepseek" in src_intent.lower()) and ("def parse_one" in src_intent)
          and ("def run_edit" in src_edit),
          "parse_intent.py：DeepSeek（%s）解析指令 ｜ edit.py：执行引擎 + 规则兜底" % model_name)

    check("A2 智能化应用 · 工具使用引导就位（30 秒上手指南 + 双击即用）",
          all(k in src_web for k in ("操作指南", "三步就会", "说话小窍门"))
          and (PROJECT / "启动操作台.bat").exists()
          and (PROJECT / "演示模块2.bat").exists(),
          "网页「操作指南 / 三步就会 / 说话小窍门」｜ 双击入口：启动操作台.bat + 演示模块2.bat")

    check("A3 智能化应用 · 识别 / 填写入口已上操作台与指南（本轮补强）",
          ('id="rec-toggle"' in src_web) and ('id="fill-toggle"' in src_web)
          and ("识别入库" in src_web) and ("表格填写" in src_web),
          "顶栏「🧠 识别入库」「📋 表格填写」两个新入口 + 指南新条目")

    check("A4 文本自动识别 · 识别分类专属 API + 异步调用（本轮补强）",
          route("/api/recognize") and has_submit_rec and ("ThreadPoolExecutor" in src_api),
          "POST /api/recognize → 立即返回 job_id → GET /api/job/{id} 轮询（线程池 4 worker）")

    check("A5 文本自动识别 · 三线提取（画像 / 实体 / 指定字段）+ 数据库存储",
          all((AGENT / f).exists() for f in ("summary_ops.py", "entity_ops.py",
                                             "extract_ops.py", "store.py"))
          and db_ok and s.get("documents", 0) > 0 and s.get("entities", 0) > 0
          and s.get("profiles", 0) > 0 and s.get("extractions", 0) > 0,
          "summary / entity / extract_ops + store 在位 ｜ 库：文档 %d · 实体 %d · 画像 %d · 字段 %d"
          % (s.get("documents", 0), s.get("entities", 0), s.get("profiles", 0),
             s.get("extractions", 0)) if db_ok else ("数据库不可用：" + db_msg))

    if db_ok:
        try:
            r1 = store._q("SELECT COUNT(*) AS n FROM entities "
                          "WHERE quote IS NULL OR quote='' OR loc IS NULL OR loc=''")[0]["n"]
            r2 = store._q("SELECT COUNT(*) AS n FROM extractions "
                          "WHERE evidence IS NULL OR evidence='' OR loc IS NULL OR loc=''")[0]["n"]
        except Exception:
            r1 = r2 = -1
        check("A6 文本自动识别 · 保真度与入库规范性",
              (a.get("fidelity") == 100.0) and (r1 + r2 == 0),
              "保真度 %s%%（%d/%d 条完全可溯）｜ 缺摘抄/出处：实体 %d · 字段 %d"
              % (a.get("fidelity"), a.get("traced"), a.get("total"), r1, r2))
    else:
        check("A6 文本自动识别 · 保真度与入库规范性", False, "数据库不可用：" + db_msg)

    filled = [f for f in (PROJECT / "writer" / "输出").glob("*-已填写.*")]
    fx = [f for f in filled if f.suffix.lower() == ".xlsx"]
    fw = [f for f in filled if f.suffix.lower() == ".docx"]
    check("A7 智能信息整合 · 双形态填写能力（word + excel）",
          (PROJECT / "writer" / "fill.py").exists()
          and (PROJECT / "writer" / "verify_fill.py").exists() and fx and fw,
          "fill.py + verify_fill.py 在位 ｜ 已填写产物：excel %d 份 · word %d 份"
          % (len(fx), len(fw)))

    check("A8 智能信息整合 · 填写入口与 API（本轮补强）",
          route("/api/fill") and route("/api/fill/templates")
          and route("/api/fill/download") and ('id="fill-overlay"' in src_web),
          "POST /api/fill（异步）+ 模板清单 + 下载端点 ｜ 操作台「📋 表格填写」面板")

    reps = sorted((PROJECT / "writer" / "输出").glob("验收报告-*.txt"))
    n_rep = 0
    bad = []
    for f in reps:
        m = re.search(r"单元格样式差异：(\d+) 处", read(f))
        if not m:
            continue
        n_rep += 1
        if int(m.group(1)) != 0:
            bad.append(f.name)
    check("A9 智能信息整合 · 格式严谨（写回零格式差异）",
          n_rep > 0 and not bad,
          "%d 份验收报告「单元格样式差异：0 处」%s"
          % (n_rep, ("；例外：" + "、".join(bad)) if bad else ""))
    P("")

    # ================= B 组：技术要求与指标 =================
    P("---- B 组：技术要求与指标（赛题原文 §6） ----")

    check("B1 程序形态 · Web 网站 + PC 端（浏览器即用 + 双击启动）",
          ("from fastapi" in src_api) and ("FileResponse" in src_api)
          and (PROJECT / "启动操作台.bat").exists(),
          "FastAPI 服务 + 网页操作台（http://127.0.0.1:8000 浏览器即用）｜ 双击启动脚本 ｜ CLI 脚本入口")

    check("B2 技术支撑 · 第三方商业 AI 平台 + 自研算法",
          ("deepseek" in src_intent.lower())
          and (PROJECT / "matcher" / "retrieve.py").exists() and ("重审" in src_runall),
          "DeepSeek 开放 API（指令解析 / 信息抽取）｜ 自研：相关度打分（retrieve.py）+ 一致性重审（run_all.py）")

    readers = [f for f in (PROJECT / "doc_reader" / "readers").glob("*.py")
               if f.name not in ("__init__.py", "base.py")]
    cov = read(AGENT / "输出" / "_体检-模块1覆盖矩阵.txt")
    check("B3 数据类型 · 四格式识别 + 多类型稳定运行",
          len(readers) >= 4 and ("8 / 8" in cov)
          and all(by.get(k, 0) >= 3 for k in ("docx", "xlsx", "md", "txt")),
          "读取器：%s ｜ 覆盖矩阵 8 / 8（四操作 × 四格式）｜ 池内 docx %d · xlsx %d · md %d · txt %d"
          % ("、".join(sorted(f.name for f in readers)),
             by.get("docx", 0), by.get("xlsx", 0), by.get("md", 0), by.get("txt", 0)))

    gen = read(PROJECT / "泛化测试" / "泛化测试总报告.txt")
    five = read(PROJECT / "五表连测" / "五表连测报告.txt")
    m_s = re.search(r"值严格 \d+/\d+（([\d.]+)%）", gen)
    m_l = re.search(r"值宽松 \d+/\d+（([\d.]+)%）", gen)
    m_ds = re.search(r"严格（一字不差/数字等价）：\d+/\d+（([\d.]+)%）", five)
    m_dl = re.search(r"宽松（含近似）：\d+/\d+（([\d.]+)%）", five)
    pcts = [float(x.group(1)) for x in (m_s, m_l, m_ds, m_dl) if x]
    tri = [sc.get(k) for k in ("entities", "profile", "fields")]
    tri_ok = db_ok and all(v is not None and v >= 80 for v in tri)
    check("B4 硬指标 · 信息提取准确率高于 80%",
          len(pcts) == 4 and all(p >= 80 for p in pcts) and tri_ok,
          "泛化 63 格：严格 %s%% / 宽松 %s%% ｜ 对分 21 格：严格 %s%% / 宽松 %s%% ｜ 模块2 三线查全：%s"
          % (m_s.group(1) if m_s else "?", m_l.group(1) if m_l else "?",
             m_ds.group(1) if m_ds else "?", m_dl.group(1) if m_dl else "?",
             " · ".join("%s%%" % v for v in tri) if db_ok else "（数据库不可用）"))

    ans_files = [
        AGENT / "标准答案-模块2.json",
        PROJECT / "matcher" / "标准答案-模板4.md",
        PROJECT / "泛化测试" / "场景1-费用报销" / "标准答案.md",
        PROJECT / "泛化测试" / "场景2-招聘初筛" / "标准答案.md",
        PROJECT / "泛化测试" / "场景3-设备报修" / "标准答案.md",
        PROJECT / "测试数据" / "压测-大文档" / "压测答案.json",
    ]
    n_ans = sum(1 for f in ans_files if f.exists())
    doc_ok = (HERE / "样本标注与审核说明.md").exists()
    rev_ok = (AGENT / "输出" / "_复核-测试用例执行记录.txt").exists()
    check("B5 样本 · 人工标注与多轮审核（档案 + 说明 + 复核记录）",
          n_ans == len(ans_files) and doc_ok and rev_ok,
          "标准答案档案 %d/%d 类 ｜ 样本标注与审核说明.md %s ｜ 复核记录 %s"
          % (n_ans, len(ans_files), "在位" if doc_ok else "缺",
             "在位" if rev_ok else "缺"))

    check("B6 智能识别分类模块 · API 接口 + 异步调用（提交→轮询范式）",
          route("/api/recognize") and route("/api/command") and route("/api/job/{job_id}")
          and ("ThreadPoolExecutor(max_workers=4" in src_api),
          "POST /api/recognize、/api/command → job_id → GET /api/job/{job_id}"
          "（4 worker 线程池，长任务不阻塞界面）")
    P("")

    # ================= C 组：快照 =================
    P("---- C 组：快照（只读汇总，供人工复核） ----")

    check("C1 快照 · 检索池与数据仓库",
          db_ok and pool_st.get("total", 0) > 0,
          "池 %d 份（docx %d · xlsx %d · md %d · txt %d ｜ 上传件 %d）｜ "
          "库：文档 %d · 实体 %d · 画像 %d · 字段 %d"
          % (pool_st.get("total", 0), by.get("docx", 0), by.get("xlsx", 0),
             by.get("md", 0), by.get("txt", 0), len(pool_st.get("imported") or []),
             s.get("documents", 0), s.get("entities", 0), s.get("profiles", 0),
             s.get("extractions", 0)))

    tx = list((PROJECT / "测试数据" / "模板表格").glob("*.xlsx"))
    tw = list((PROJECT / "测试数据" / "模板表格-word").glob("*.docx"))
    check("C2 快照 · 模板双形态清点",
          len(tx) >= 5 and len(tw) >= 5,
          "Excel 模板 %d 张 + Word 模板 %d 张（比赛任选形态均可）" % (len(tx), len(tw)))
    P("")

    # ---------------- 归档 ----------------
    elapsed = time.time() - T0
    P("=" * 66)
    P("总检成绩：%d / %d 通过（用时 %.1f 秒）" % (N_PASS, N_ALL, elapsed))
    P("=" * 66)

    (HERE / "作品要求对照总检.txt").write_text("\n".join(REPORT) + "\n",
                                                encoding="utf-8")
    print("")
    print("（报告已存档：收尾对照/作品要求对照总检.txt）")
    sys.exit(0 if N_PASS == N_ALL else 1)


if __name__ == "__main__":
    main()
