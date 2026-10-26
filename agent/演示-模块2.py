# -*- coding: utf-8 -*-
"""
演示-模块2.py —— 模块（2）「非结构化文档信息提取」全链路一条命令演示
（模块2 · 第 6 步「收口」交付件）

给谁看：导师 / 评委 / 自己自检 —— 一条命令看到「文档 → 提取 → 入库 → 质量」全闭环。

跑什么（四个阶段）：
  【0】环境自检    检索池（几份文档/几块）+ 数据仓库（现在有多少数据）
  【1】三线提取    对演示文档依次跑 画像 / 实体 / 字段 —— 走真实链路：
                   一句话 → 模型解析 → 提取 → 出处定位 → 自动入库
  【2】库核对      该文档在数据库里的分布（并演示幂等：重跑不膨胀）
  【3】质量验证    零模型（确定性）：全库反向审计（保真度）+ 对照人工标准答案（查全率）

用法：
  python agent/演示-模块2.py                 # 默认：设备采购合同（三线齐全的样板）
  python agent/演示-模块2.py --doc 项目周报   # 换一份文档（内置 4 份预设指令）
  python agent/演示-模块2.py --fast          # 跳过三线提取（只做 池/库/质量，秒级）
  python agent/演示-模块2.py --no-save       # 不落盘报告文件

产物：agent/输出/演示-模块2报告.txt（覆盖式，最新一次）
退出码：0=全通过 ｜ 1=有三线提取失败 ｜ 2=环境不可用（数据库/池）
"""
import argparse
import sys
import time
from datetime import datetime
from pathlib import Path

BASE = Path(__file__).resolve().parent          # = agent/
sys.path.insert(0, str(BASE))

try:
    sys.stdout.reconfigure(errors="replace")    # 终端表示不了的字符不崩
except Exception:
    pass

# ------------------------------------------------------------------
# 演示文档 → 三线指令预设（未收录的文档自动降级为"画像+实体"两线）
# ------------------------------------------------------------------
PRESETS = {
    "设备采购合同": {
        "profile": "提取设备采购合同的关键信息",
        "entities": "识别设备采购合同里的实体数据",
        "fields": "从设备采购合同里提取甲乙双方名称和合同总金额",
    },
    "员工信息表": {
        "profile": "提取员工信息表的关键信息",
        "entities": "识别员工信息表里的实体数据",
        "fields": "把员工信息表里所有员工的姓名和工号整理出来",
    },
    "项目周报": {
        "profile": "提取项目周报的关键信息",
        "entities": "识别项目周报里的实体数据",
        "fields": "把项目周报的汇报人和提交日期提取出来",
    },
    "设备验收记录": {
        "profile": "提取设备验收记录的关键信息",
        "entities": "识别设备验收记录里的实体数据",
        "fields": "从设备验收记录里提取验收编号和验收日期",
    },
    "项目推进会会议纪要": {
        "profile": "提取项目推进会会议纪要的关键信息",
        "entities": "识别项目推进会会议纪要里的实体数据",
        "fields": "把项目推进会会议纪要里的待办事项列出来",
    },
}


def _find_preset(doc_arg: str):
    """按名字片段找预设（双向包含匹配）。"""
    for k, v in PRESETS.items():
        if k in doc_arg or doc_arg in k:
            return k, v
    return None, None


def _pick_lines(report: str) -> list:
    """从完整报告里提炼关键行：解析 / 摘要（◆，跳过 ◆ 【类型】分组明细）/ 入库 / 失败。"""
    picked = []
    for l in report.splitlines():
        ls = l.strip()
        if ls.startswith("【解析】") or ls.startswith("【入库】") or ls.startswith("【失败】"):
            picked.append(ls)
        elif ls.startswith("◆") and "【" not in ls:
            picked.append(ls)
    return picked


def main():
    ap = argparse.ArgumentParser(description="模块（2）全链路演示（一条命令）")
    ap.add_argument("--doc", default="设备采购合同", help="演示文档名（默认 设备采购合同）")
    ap.add_argument("--fast", action="store_true", help="跳过三线提取，只做 池/库/质量")
    ap.add_argument("--no-save", action="store_true", help="不落盘报告文件")
    args = ap.parse_args()

    t_all = time.time()
    t_extract = 0.0
    n_fail = 0
    R = []                                       # 报告行（屏幕 + 落盘同一份）

    def P(s=""):
        R.append(s)
        print(s)

    P("=" * 66)
    P("模块（2）「非结构化文档信息提取」全链路演示")
    P("=" * 66)
    P("演示文档：%s ｜ 演示模式：%s" % (args.doc, "快速（跳过提取）" if args.fast else "完整"))
    P("时间：%s" % datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    P("")

    preset_key, preset = _find_preset(args.doc)
    if not preset:
        preset_key = args.doc
        preset = {"profile": "提取%s的关键信息" % args.doc,
                  "entities": "识别%s里的实体数据" % args.doc}

    # ---------- 【0】环境自检 ----------
    P("【0】环境自检")
    try:
        import pool
        st = pool.status()
        by = st.get("by_type") or {}
        P("  · 检索池：%d 份文档（docx %d · xlsx %d · md %d · txt %d）｜ 上传件 %d"
          % (st.get("total", 0), by.get("docx", 0), by.get("xlsx", 0),
             by.get("md", 0), by.get("txt", 0), len(st.get("imported") or [])))
    except Exception as e:
        P("  · 检索池：读取失败（%s）" % e)
        n_fail += 1

    import store
    try:
        s0 = store.stats()
    except Exception as e:
        P("  · 数据仓库：连接失败（%s）" % e)
        P("    → 先启动 MySQL（net start MySQL80）再重跑本演示。")
        sys.exit(2)
    P("  · 数据仓库：a23_agent ｜ 文档 %d · 实体 %d · 画像 %d · 字段 %d"
      % (s0["documents"], s0["entities"], s0["profiles"], s0["extractions"]))
    P("")

    # ---------- 【1】三线提取（真实链路） ----------
    if not args.fast:
        P("【1】三线提取（真实链路：一句话 → 解析 → 提取 → 出处 → 自动入库）")
        import edit
        docs = edit.DEFAULT_DOCS_DIRS
        for label, key in (("画像", "profile"), ("实体", "entities"), ("字段", "fields")):
            cmd = preset.get(key)
            if not cmd:
                P("")
                P("  ▶ 字段线：该文档未配置演示指令，跳过")
                continue
            P("")
            P("  ▶ 指令：%s" % cmd)
            t1 = time.time()
            meta = {}
            ok, report = edit.run_edit(cmd, docs, meta=meta)
            if not ok and "请重跑一次" in report:      # 网络抖动：自动重试 1 次
                ok, report = edit.run_edit(cmd, docs, meta=meta)
            t_extract += time.time() - t1
            for l in _pick_lines(report):
                P("    " + l)
            if not ok:
                n_fail += 1
        P("")
        P("  三线提取合计用时 %.1f 秒" % t_extract)
        P("")
    else:
        P("【1】三线提取（--fast 已跳过）")
        P("")

    # ---------- 【2】库核对 ----------
    P("【2】数据仓库核对（同步式落库：库 = 最新提取结果，不堆积）")
    s1 = store.stats()
    P("  · 全库：文档 %d · 实体 %d · 画像 %d · 字段 %d"
      % (s1["documents"], s1["entities"], s1["profiles"], s1["extractions"]))
    P("  · 对比演示前：文档 %+d ｜ 实体 %+d ｜ 画像 %+d ｜ 字段 %+d（同步式：不堆积，与最新提取对齐）"
      % (s1["documents"] - s0["documents"], s1["entities"] - s0["entities"],
         s1["profiles"] - s0["profiles"], s1["extractions"] - s0["extractions"]))
    hit = None
    for d in store.list_docs():
        if preset_key and preset_key in d["doc_name"]:
            hit = d
            break
    if hit:
        P("  · 本文档《%s》：实体 %d 条 · 字段 %d 条 · 画像 %d 份"
          % (hit["doc_name"], hit["n_ent"], hit["n_fld"], hit["n_prof"]))
        if hit["n_prof"]:
            p, items = store.get_profile(hit["doc_name"])
            if p:
                n_point = sum(1 for it in items if it["kind"] == "point")
                n_elem = sum(1 for it in items if it["kind"] == "element")
                P("      画像明细：核心要点 %d 条 + 关键要素 %d 条 ｜ 主题：%s"
                  % (n_point, n_elem, (p.get("topic") or "")[:40]))
    P("")

    # ---------- 【3】质量验证（零模型） ----------
    P("【3】质量验证（零模型 · 确定性 —— 重跑必一致）")
    import quality
    t_q = time.time()
    audit = quality.audit_all()
    scores = quality.score_all()
    summ = quality.summarize(audit, scores)
    t_q = time.time() - t_q
    a, sc = summ["audit"], summ["scores"]
    n_ans = sum((scores[k]["overall"].get("total") or 0)
                for k in ("entities", "profile", "fields"))
    P("  · 反向审计（库 → 原文）：%d 条 → 可追溯 %d 条 ｜ 保真度 %s%% ｜ 片段命中 %s%%"
      % (a["total"], a["traced"], a["fidelity"], a["frag_rate"]))
    P("  · 三线查全（对照人工标准答案 %d 项）：实体 %s%% ｜ 画像 %s%% ｜ 字段 %s%%"
      % (n_ans, sc["entities"], sc["profile"], sc["fields"]))
    P("  · 有问题的数据条数：%d ｜ 验证耗时 %.2f 秒" % (a["issues"], t_q))
    P("")

    # ---------- 【4】一页总结 ----------
    t_total = time.time() - t_all
    t_other = max(0.0, t_total - t_extract - t_q)
    P("【4】一页总结")
    P("  · 全链路耗时：%.1f 秒（三线提取 %.1f ｜ 质量验证 %.2f ｜ 核对与其他 %.1f）"
      % (t_total, t_extract, t_q, t_other))
    P("  · 结论：文档 → 提取（画像/实体/字段）→ 入库 → 质量验证 全链路%s"
      % ("通过 ✓" if n_fail == 0 else "跑完（有 %d 处失败，见上）" % n_fail))
    P("")
    P("  接下来看什么：")
    P("    · 网页操作台：双击项目根目录「启动操作台.bat」→「🗄 数据仓库」→「🔍 质量体检」")
    P("    · 质量报告  ：agent/输出/质量报告/质量报告.html（可给导师看）")
    P("    · 单条指令  ：python agent/chat.py（交互模式，随手试一句）")
    P("=" * 66)

    if not args.no_save:
        rp = BASE / "输出" / "演示-模块2报告.txt"
        rp.parent.mkdir(parents=True, exist_ok=True)
        P("")
        P("（本报告已存档：%s）" % rp)
        rp.write_text("\n".join(R) + "\n", encoding="utf-8")

    sys.exit(1 if n_fail else 0)


if __name__ == "__main__":
    main()
