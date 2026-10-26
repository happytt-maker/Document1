# -*- coding: utf-8 -*-
"""第 5 步验收：质量验证引擎（quality.py）自检。

验证方式：本地直调 quality 模块（不经 HTTP、不经模型）——
  确定性对分 + 负向必抓（注入假数据必须被审计抓出来）。

分组：
  A 组（8 项）：反向审计正确性 —— 纯函数真/假/拼接三分 + 标注剥离 + markdown 归一
                + 端到端负向注入 + 真库全量数字核对 + 分线合计一致
  B 组（6 项）：计分正确性 —— 实体/字段/画像对分 + 负向删值必扣 + aliases 合并统计
  C 组（4 项）：报告完整性 + 幂等（重跑逐数字一致）
  D 组（3 项）：只读性 —— 库零改动 / 源文件零改动 / 解析池零改动
  E 组（5 项）：CLI 与边界 —— --doc 过滤 / 子进程运行 / 空库不崩 / 不存在文档不崩 / 单跑模式

跑法（在项目根目录）：
  python agent/_验收-质量版.py
"""
import hashlib
import json
import os
import subprocess
import sys
import time
from pathlib import Path

BASE = Path(r"C:\Users\覃敬涛\Desktop\研一项目")
sys.path.insert(0, str(BASE / "agent"))
sys.stdout.reconfigure(errors="replace")

import quality  # noqa: E402

ARC_DIR = BASE / "agent" / "输出" / "验收-质量"
RESULTS = []


def check(no, name, ok, detail=""):
    RESULTS.append({"no": no, "name": name, "ok": bool(ok), "detail": detail})
    print("[%s] %s ｜ %s" % ("PASS" if ok else "FAIL", name, detail))


def md5(p: Path) -> str:
    return hashlib.md5(p.read_bytes()).hexdigest()


def snap_db(db: dict) -> str:
    return json.dumps(db, sort_keys=True, default=str)


def main():
    ARC_DIR.mkdir(parents=True, exist_ok=True)
    print("=" * 66)
    print("模块（2）· 第 5 步「质量验证与准确率报告」验收（零模型 · 确定性）")
    print("=" * 66)
    print()

    db = quality.fetch_db()
    answers = quality.load_answers()

    # ---------- 跑前快照（D 组用） ----------
    before_db = snap_db(db)
    src_files = [BASE / "测试数据" / "源文档" / "txt" / "设备验收记录.txt",
                 BASE / "测试数据" / "源文档" / "xlsx" / "员工信息表.xlsx",
                 BASE / "测试数据" / "源文档" / "docx" / "设备采购合同.docx",
                 BASE / "测试数据" / "导入样本" / "食堂值日安排.txt"]
    src_md5_before = {str(p): md5(p) for p in src_files if p.exists()}
    pool_path = BASE / "doc_reader" / "output.json"
    pool_md5_before = md5(pool_path)

    # ================= A 组：反向审计正确性 =================
    print("---- A 组：反向审计（负向必抓 + 真数据对分）----")
    blocks, src = quality.load_blocks_any("设备验收记录.txt")
    check("A1", "装载原文块（池内）", bool(blocks) and src == "解析池",
          "来源=%s ｜ %d 块" % (src, len(blocks or [])))

    r1 = quality.audit_one("验收编号：YS-2026-0089", "YS-2026-0089", blocks)
    check("A2", "真句子 → traced", r1["status"] == "traced", "status=%s" % r1["status"])

    r2 = quality.audit_one("这段话绝对不存在阿巴阿巴xyz", "阿巴阿巴", blocks)
    check("A3", "编造句子 → miss（负向必抓）", r2["status"] == "miss",
          "status=%s miss=%s" % (r2["status"], r2["miss_parts"]))

    r3 = quality.audit_one("验收编号：YS-2026-0089 外加一句不存在的阿巴阿巴", "x", blocks)
    check("A4", "真+假拼接 → partial（假片段被列出）",
          r3["status"] == "partial" and any("阿巴阿巴" in p for p in r3["miss_parts"]),
          "status=%s miss=%s" % (r3["status"], r3["miss_parts"]))

    kb, _ = quality.load_blocks_any("库存盘点表.xlsx")
    r4 = quality.audit_one("（该格所在列的列头：期初库存）\n120", "120", kb)
    stripped = quality.strip_notes("（该格所在列的列头：期初库存）\n120")
    check("A5", "系统列头标注剥离（且列头名仍须可溯）",
          r4["status"] == "traced" and "期初库存" in stripped,
          "status=%s ｜ stripped=%r" % (r4["status"], stripped))

    check("A6", "markdown 标记归一（** 与 ` 不影响命中）",
          quality.norm("**结论**：用 `id()` 去重") == quality.norm("结论：用 id() 去重"),
          "")

    # 端到端负向：向真库快照注入一条假数据 → 审计必须抓出
    fake_db = {k: list(v) for k, v in db.items()}
    fake_db["entities"].append({"id": 999999, "doc_name": "设备验收记录.txt",
                                "etype": "code", "etype_cn": "编号", "value": "FAKE-X",
                                "freq": 1, "loc": "第 1 行", "quote": "这段话绝对不存在阿巴阿巴xyz",
                                "verify": "na", "source": "model", "run_id": "rFAKE", "ts": None})
    aud_f = quality.audit_all(db=fake_db)
    caught = any("阿巴阿巴" in (i.get("quote") or "") for i in aud_f["issues"])
    check("A7", "端到端负向注入被抓出（审计有检出能力）",
          aud_f["overall"]["miss"] >= 1 and caught,
          "miss=%d ｜ 命中 issue=%s" % (aud_f["overall"]["miss"], caught))

    # 真库全量审计 + 内部一致性
    aud = quality.audit_all(db=db)
    n_sum = len(db["entities"]) + len(db["profile_items"]) + len(db["extractions"])
    line_sum = sum(v["total"] for v in aud["by_line"].values())
    n_q = aud["overall"]["total"] - aud["overall"]["no_quote"]
    fid = (aud["overall"]["traced"] / n_q * 100) if n_q else 0
    check("A8", "真库全量审计（%d 条）+ 分线合计一致" % n_sum,
          aud["overall"]["total"] == n_sum and line_sum == n_sum,
          "overall=%d ｜ 分线合计=%d ｜ 保真度=%.1f%%" % (aud["overall"]["total"], line_sum, fid))

    # ================= B 组：计分正确性 =================
    print("")
    print("---- B 组：计分正确性（对分 + 负向）----")
    ents = [r for r in db["entities"] if r["doc_name"] == "供应商联系方式.txt"]
    s1 = quality.score_entities(ents, answers["entities"]["供应商联系方式.txt"])
    check("B1", "实体计分对分（供应商 30/30）",
          s1["hit"] == 30 and s1["total"] == 30 and s1["type_ok"] == 30,
          "%d/%d ｜ 类型正确 %d" % (s1["hit"], s1["total"], s1["type_ok"]))

    cut = [r for r in ents if r["value"] not in ("赵雪松", "王建平")]
    s2 = quality.score_entities(cut, answers["entities"]["供应商联系方式.txt"])
    miss_ok = sorted(m.split("：")[-1] for m in s2["miss"]) == ["王建平", "赵雪松"]
    check("B2", "实体负向：删 2 人 → 28/30 且缺项列出",
          s2["hit"] == 28 and miss_ok, "%d/%d ｜ 缺 %s" % (s2["hit"], s2["total"], s2["miss"]))

    rows_kc = [r for r in db["extractions"] if r["doc_name"] == "库存盘点表.xlsx"]
    s3 = quality.score_fields(rows_kc, answers["fields"]["库存盘点表.xlsx"])
    check("B3", "字段计分对分（库存表 20/20）",
          s3["hit"] == 20 and s3["total"] == 20, "%d/%d" % (s3["hit"], s3["total"]))

    cut2 = [r for r in rows_kc if not (r["field"] == "期初库存" and r["value"] in ("120", "45"))]
    s4 = quality.score_fields(cut2, answers["fields"]["库存盘点表.xlsx"])
    check("B4", "字段负向：删 2 值 → 18/20",
          s4["hit"] == 18, "%d/%d" % (s4["hit"], s4["total"]))

    rows_e = [r for r in db["extractions"] if r["doc_name"] == "员工信息表.xlsx"]
    s5 = quality.score_fields(rows_e, answers["fields"]["员工信息表.xlsx"])
    x5 = [z for z in s5["by_field"] if z["key"] == "姓名"][0]
    check("B5", "aliases 合并去重（姓名 raw=16 → got=8，命中 8/8）",
          x5["raw_n"] == 16 and x5["got_n"] == 8 and x5["hit"] == 8,
          "raw=%d got=%d hit=%d/%d" % (x5["raw_n"], x5["got_n"], x5["hit"], x5["total"]))

    prof = [p for p in db["profiles"] if p["doc_name"] == "员工信息表.xlsx"][0]
    items = [r for r in db["profile_items"] if r["profile_id"] == prof["id"]]
    s6 = quality.score_profile(prof["topic"], items, answers["profile"]["员工信息表.xlsx"])
    check("B6", "画像计分对分（员工表 9/10）",
          s6["hit"] == 9 and s6["total"] == 10, "%d/%d ｜ 缺 %s" % (s6["hit"], s6["total"], s6["miss"]))

    # ================= C 组：报告完整性 + 幂等 =================
    print("")
    print("---- C 组：报告完整性 + 幂等 ----")
    scores = quality.score_all(db=db)
    hp, jp = quality.build_report(audit=aud, scores=scores)
    check("C1", "报告双文件生成（HTML + 明细 JSON）",
          hp.exists() and jp.exists() and hp.stat().st_size > 4000,
          "HTML %.1fKB ｜ JSON %.1fKB" % (hp.stat().st_size / 1024, jp.stat().st_size / 1024))

    html = hp.read_text(encoding="utf-8")
    need = ["质量验证与准确率报告", "反向审计", "保真度", "诚实清单", "质量明细"]
    lack = [k for k in need if k not in html]
    check("C2", "HTML 关键内容齐全", not lack, "缺：%s" % lack if lack else "6 个关键锚点全在")

    detail = json.loads(jp.read_text(encoding="utf-8"))
    d_ok = (all(k in detail for k in ("summary", "scores", "audit"))
            and detail["summary"]["audit"]["total"] == aud["overall"]["total"]
            and detail["scores"]["entities"]["overall"]["total"] == 101)
    check("C3", "明细 JSON 结构与数字一致", d_ok,
          "audit.total=%s ｜ 实体应含=%s" % (detail["summary"]["audit"]["total"],
                                            detail["scores"]["entities"]["overall"]["total"]))

    aud2 = quality.audit_all(db=db)
    scores2 = quality.score_all(db=db)
    same = (json.dumps(aud2["overall"], sort_keys=True) == json.dumps(aud["overall"], sort_keys=True)
            and json.dumps(scores2["entities"]["overall"], sort_keys=True) == json.dumps(scores["entities"]["overall"], sort_keys=True)
            and json.dumps(scores2["profile"]["overall"], sort_keys=True) == json.dumps(scores["profile"]["overall"], sort_keys=True)
            and json.dumps(scores2["fields"]["overall"], sort_keys=True) == json.dumps(scores["fields"]["overall"], sort_keys=True))
    check("C4", "幂等（重跑逐数字一致）", same, "二次审计/计分与首次全等=%s" % same)

    # ================= D 组：只读性 =================
    print("")
    print("---- D 组：只读性（质量验证不许改任何东西）----")
    db_after = quality.fetch_db()
    check("D1", "数据库零改动（三表全量快照比对）",
          snap_db(db_after) == before_db, "快照全等=%s" % (snap_db(db_after) == before_db))

    bad_src = [p for p, h in src_md5_before.items() if md5(Path(p)) != h]
    check("D2", "源文件零改动（含导入样本，%d 份）" % len(src_md5_before),
          not bad_src, "变动文件：%s" % bad_src if bad_src else "4 份指纹全同")

    check("D3", "解析池零改动（现场解析不入池）",
          md5(pool_path) == pool_md5_before, "")

    # ================= E 组：CLI 与边界 =================
    print("")
    print("---- E 组：CLI 与边界 ----")
    aud_doc = quality.audit_all(db=db, doc="设备采购")
    only = aud_doc["by_doc"] and all("设备采购" in k for k in aud_doc["by_doc"])
    check("E1", "audit --doc 过滤（只含匹配文档）",
          only, "文档=%s" % list(aud_doc["by_doc"].keys()))

    env = {**os.environ, "PYTHONIOENCODING": "utf-8"}
    r_cli = subprocess.run([sys.executable, str(BASE / "agent" / "quality.py"), "--audit"],
                           capture_output=True, text=True, encoding="utf-8",
                           env=env, cwd=str(BASE), timeout=120)
    check("E2", "CLI --audit 子进程运行（退出码 0 + 输出正常）",
          r_cli.returncode == 0 and "反向审计" in r_cli.stdout and "保真度" in r_cli.stdout,
          "rc=%s" % r_cli.returncode)

    empty = {"entities": [], "profiles": [], "profile_items": [], "extractions": []}
    a3 = quality.audit_all(db=empty)
    s3b = quality.score_all(db=empty)
    all_no_data = all(s.get("error") for s in s3b["entities"]["by_doc"].values())
    check("E3", "空数据不崩（audit=0 条 / 计分全标「无数据」）",
          a3["overall"]["total"] == 0 and s3b["entities"]["overall"]["total"] == 0 and all_no_data,
          "audit=0 ｜ 计分=0（全标无数据=%s）" % all_no_data)

    a4 = quality.audit_all(db=db, doc="绝不存在文档名XYZ")
    check("E4", "不存在文档过滤 → 空结果不崩", a4["overall"]["total"] == 0, "")

    s5b = quality.score_all(db=db, doc="绝不存在文档名XYZ")
    check("E5", "计分同样支持过滤且不崩",
          all(s5b[k]["overall"]["total"] == 0 for k in ("entities", "profile", "fields")),
          "")

    # ================= 汇总 =================
    n_pass = sum(1 for r in RESULTS if r["ok"])
    n_total = len(RESULTS)
    print("")
    print("=" * 66)
    print("成绩：%d / %d 通过" % (n_pass, n_total))
    print("=" * 66)

    # 归档
    lines = ["模块（2）· 第 5 步「质量验证与准确率报告」验收报告",
             "时间：%s ｜ 方式：本地直调 quality 模块（零模型 · 确定性）" % time.strftime("%Y-%m-%d %H:%M:%S"),
             "-" * 60]
    for r in RESULTS:
        lines.append("[%s] %s ｜ %s" % ("PASS" if r["ok"] else "FAIL", r["name"], r["detail"]))
    lines.append("-" * 60)
    lines.append("成绩：%d / %d" % (n_pass, n_total))
    lines.append("")
    lines.append("附：本次验收实测业务数字（信息性，非断言）——")
    lines.append("  · 全库审计 %d 条 ｜ 完全可溯 %d ｜ 保真度 %.1f%%"
                 % (aud["overall"]["total"], aud["overall"]["traced"], fid))
    lines.append("  · 三线查全：实体 %d/%d ｜ 画像 %d/%d ｜ 字段 %d/%d"
                 % (scores["entities"]["overall"]["hit"], scores["entities"]["overall"]["total"],
                    scores["profile"]["overall"]["hit"], scores["profile"]["overall"]["total"],
                    scores["fields"]["overall"]["hit"], scores["fields"]["overall"]["total"]))
    (ARC_DIR / "质量验收报告.txt").write_text("\n".join(lines), encoding="utf-8")
    (ARC_DIR / "验收明细.json").write_text(
        json.dumps({"results": RESULTS, "pass": n_pass, "total": n_total},
                   ensure_ascii=False, indent=2), encoding="utf-8")

    sys.exit(0 if n_pass == n_total else 1)


if __name__ == "__main__":
    main()
