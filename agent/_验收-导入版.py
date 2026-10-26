# -*- coding: utf-8 -*-
"""
_验收-导入版.py —— 模块（2）· 第 1 步「文档导入与自动识别」验收（真实 HTTP 层）

覆盖 17 项检查：
  A-1~A-4   HTTP 上传 4 种格式样本 → 自动识别 + 解析 + 入池
  B-1~B-2   边界：不支持格式拒绝（并留痕）、重名自动改名
  C-1~C-4   上传后立即可操作：docx 格式 / txt 提取 / xlsx 格式 / md 编辑
  D-1~D-5   数据库与池核对：imports 表、commands 表、文档库清单、池状态、池内关键词
  E-1~E-2   基线恢复往返：baseline 复位 → rebuild 恢复

前置：操作台服务已启动（127.0.0.1:8000）
产物：agent/输出/验收-导入/导入验收报告.txt + 验收明细.json + 指令产物归档
"""
import json
import shutil
import sys
import time
from pathlib import Path

import requests

BASE = Path(__file__).resolve().parents[1]
AGENT = BASE / "agent"
SAMPLE_DIR = BASE / "测试数据" / "导入样本"
OUT_ROOT = AGENT / "输出"
OUT_DIR = OUT_ROOT / "验收-导入"
OUT_DIR.mkdir(parents=True, exist_ok=True)
sys.path.insert(0, str(AGENT))
import pool  # noqa: E402

URL = "http://127.0.0.1:8000"
SID = "test-import-verify"
results = []


def rec(no, name, ok, actual):
    results.append({"no": no, "name": name, "ok": bool(ok), "actual": str(actual)[:300]})
    print("%s %-5s %s" % ("[PASS]" if ok else "[FAIL]", no, name))
    print("       实际：%s" % str(actual)[:200])


def api_upload(path: Path):
    with open(path, "rb") as f:
        r = requests.post(URL + "/api/upload", files={"file": (path.name, f)}, timeout=120)
    try:
        return r.status_code, r.json()
    except Exception:
        return r.status_code, {"raw": r.text[:200]}


def api_command(text: str):
    r = requests.post(URL + "/api/command", json={"text": text, "session_id": SID}, timeout=30)
    j = r.json()
    jid = j.get("job_id")
    if not jid:
        return {"status": "failed", "result": {"ok": False, "report": str(j)}}
    for _ in range(240):          # 最长 2 分钟
        time.sleep(0.5)
        try:
            j2 = requests.get(URL + "/api/job/" + jid, timeout=15).json()
        except Exception:
            continue
        if j2.get("status") in ("done", "failed"):
            return j2
    return {"status": "timeout", "result": {"ok": False, "report": "轮询超时"}}


def api_get(path: str):
    return requests.get(URL + path, timeout=20).json()


def main():
    t_all = time.time()

    # 0) 服务探测
    try:
        st = api_get("/api/status")
        assert st.get("service") == "ok"
        print("服务在线：docs=%s ｜ db=%s\n" % (st.get("docs_count"), st.get("database")))
    except Exception as e:
        print("【中止】服务不在线：%s\n先启动：python agent/api_server.py" % e)
        sys.exit(2)

    # ---------- 阶段 A ----------
    print("=" * 66)
    print("阶段 A：HTTP 上传 4 种格式样本（自动识别 + 解析 + 入池）")
    print("=" * 66)
    expect = [
        ("供应商合作备忘录.docx", "docx", 21),
        ("系统运维手册.md", "md", 10),
        ("备件库存清单.xlsx", "xlsx", 30),
        ("食堂值日安排.txt", "txt", 3),
    ]
    for i, (name, ftype, blocks) in enumerate(expect, 1):
        code, j = api_upload(SAMPLE_DIR / name)
        ok = (code == 200 and j.get("ok") and j.get("filetype") == ftype
              and j.get("blocks") == blocks)
        rec("A-%d" % i, "上传《%s》→ 识别 %s ｜ %s 块" % (name, ftype, blocks), ok,
            "HTTP %s ｜ type=%s ｜ blocks=%s ｜ %.2fs" %
            (code, j.get("filetype"), j.get("blocks"), j.get("elapsed") or 0))

    # ---------- 阶段 B ----------
    print()
    print("=" * 66)
    print("阶段 B：边界（不支持格式 / 重名）")
    print("=" * 66)

    pdf = OUT_DIR / "_负向-不支持格式.pdf"
    pdf.write_bytes(b"%PDF-1.4 fake file for negative test")
    code, j = api_upload(pdf)
    b1 = (code == 400 and "不支持" in json.dumps(j, ensure_ascii=False))
    rec("B-1", "上传 .pdf → 人话拒绝（HTTP 400）", b1,
        "HTTP %s ｜ %s" % (code, (j.get("detail") or "")[:90]))

    code, j = api_upload(SAMPLE_DIR / "供应商合作备忘录.docx")
    b2 = (code == 200 and j.get("ok") and j.get("stored_name") == "供应商合作备忘录(2).docx")
    rec("B-2", "重传同名 → 自动改名不覆盖", b2,
        "stored_name=%s" % j.get("stored_name"))

    # ---------- 阶段 C ----------
    print()
    print("=" * 66)
    print("阶段 C：上传后立即可操作（4 条指令 · 真 HTTP 链路）")
    print("=" * 66)
    c_start = time.time()
    cases = [
        ("C-1", "把供应商合作备忘录的标题居中", ["居中"], "docx 格式"),
        ("C-2", "从食堂值日安排里提取值日组长", ["赵桂芬"], "txt 提取"),
        ("C-3", "把备件库存清单的表头加粗", ["加粗"], "xlsx 格式"),
        ("C-4", "把系统运维手册里的\u2018巡检\u2019改成\u2018巡查\u2019", ["巡查"], "md 编辑"),
    ]
    for no, text, kws, tag in cases:
        j = api_command(text)
        rep = ((j.get("result") or {}).get("report")) or ""
        ok = (j.get("status") == "done" and (j.get("result") or {}).get("ok")
              and all(k in rep for k in kws))
        rec(no, "[%s] 指令：%s" % (tag, text), ok,
            "耗时 %ss ｜ %s" % (j.get("elapsed"),
                                rep.replace("\n", " ／ ")[-160:]))

    # 指令产物归档（C 阶段之后生成的文件）
    arc = []
    for p in sorted(OUT_ROOT.rglob("*")):
        if p.is_file() and p.parent != OUT_DIR and "中间抽取" not in str(p):
            try:
                if p.stat().st_mtime >= c_start - 1:
                    shutil.copy2(p, OUT_DIR / p.name)
                    arc.append(p.name)
            except Exception:
                pass
    print("（C 阶段产物归档 %d 个 → 验收-导入/）" % len(arc))

    # ---------- 阶段 D ----------
    print()
    print("=" * 66)
    print("阶段 D：数据库与池核对")
    print("=" * 66)

    im = api_get("/api/imports?limit=50")
    rows = im.get("rows", [])
    n_okimp = sum(1 for r in rows if r.get("ok"))
    d1 = len(rows) >= 6 and n_okimp >= 5 and any(not r.get("ok") for r in rows)
    rec("D-1", "imports 表：≥6 条记录（含成功 5 + 失败留痕 1）", d1,
        "共 %d 条 ｜ 成功 %d ｜ 失败留痕 %s" %
        (len(rows), n_okimp, any(not r.get("ok") for r in rows)))

    his = api_get("/api/history?limit=60")
    cmds = [r.get("command") or "" for r in his.get("rows", [])]
    hits = {k: any(k in c for c in cmds) for k in ["标题居中", "值日组长", "表头加粗", "巡查"]}
    d2 = all(hits.values())
    rec("D-2", "commands 表：本轮 4 条指令全部有记录", d2, str(hits))

    docs = api_get("/api/docs").get("docs", [])
    need = ["供应商合作备忘录.docx", "供应商合作备忘录(2).docx",
            "系统运维手册.md", "备件库存清单.xlsx", "食堂值日安排.txt"]
    miss = [n for n in need if n not in docs]
    d3 = (not miss) and len(docs) == 23
    rec("D-3", "文档库清单：18 + 5 = 23 份（含全部导入）", d3,
        "共 %d 份 ｜ 缺：%s" % (len(docs), miss or "无"))

    pst = pool.status()
    d4 = pst["total"] == 23 and len(pst["imported"]) == 5
    rec("D-4", "检索池：23 份（导入 5 份已并入）", d4,
        "total=%s ｜ imported=%s" % (pst["total"], len(pst["imported"])))

    data = json.loads(pool.POOL_JSON.read_text(encoding="utf-8"))
    all_text = "\n".join(b.get("text", "") for d in data["documents"] for b in d.get("blocks", []))
    kws = ["HZ-MOU-2026-015", "灯塔运维监控平台", "BJ-1001", "赵桂芬"]
    hit = [k for k in kws if k in all_text]
    rec("D-5", "池内全文可检索：4 份导入文档关键词命中", len(hit) == len(kws),
        "命中 %d/%d：%s" % (len(hit), len(kws), "、".join(hit)))

    # ---------- 阶段 E ----------
    print()
    print("=" * 66)
    print("阶段 E：基线恢复往返（可复位 / 可恢复）")
    print("=" * 66)

    r1 = pool.baseline()
    rec("E-1", "baseline 复位 → 18 份（导入清出池）", r1["total"] == 18,
        "total=%s ｜ %.2fs" % (r1["total"], r1["elapsed"]))

    r2 = pool.rebuild()
    rec("E-2", "rebuild 恢复 → 23 份（导入重新并入）", r2["total"] == 23 and r2["imported"] == 5,
        "total=%s ｜ imported=%s ｜ %.2fs" % (r2["total"], r2["imported"], r2["elapsed"]))

    # ---------- 汇总 ----------
    n_ok = sum(1 for r in results if r["ok"])
    total = len(results)
    dt = time.time() - t_all

    lines = []
    lines.append("=" * 66)
    lines.append("模块（2）· 第 1 步「文档导入与自动识别」验收报告")
    lines.append("=" * 66)
    lines.append("时间：%s ｜ 总用时 %.1f 秒" % (time.strftime("%Y-%m-%d %H:%M:%S"), dt))
    lines.append("成绩：%d / %d" % (n_ok, total))
    lines.append("")
    for r in results:
        lines.append("%s %-5s %s" % ("[PASS]" if r["ok"] else "[FAIL]", r["no"], r["name"]))
        lines.append("      实际：%s" % r["actual"])
    lines.append("")
    lines.append("现场说明：")
    lines.append("  · 导入文档 5 份在 导入文档/ ｜ 检索池 23 份（18 标准 + 5 导入）")
    lines.append("  · 指令产物已归档本目录 ｜ 数据库记录见 a23_agent.imports / commands")
    lines.append("  · 演示前复位：删导入文件 + pool.baseline()（见战报「演示复位」一节）")
    report = "\n".join(lines)

    (OUT_DIR / "导入验收报告.txt").write_text(report, encoding="utf-8")
    (OUT_DIR / "验收明细.json").write_text(
        json.dumps({"ts": time.strftime("%Y-%m-%d %H:%M:%S"), "score": "%d/%d" % (n_ok, total),
                    "results": results}, ensure_ascii=False, indent=2), encoding="utf-8")

    print()
    print(report)
    print()
    print("报告：%s" % (OUT_DIR / "导入验收报告.txt"))
    sys.exit(0 if n_ok == total else 1)


if __name__ == "__main__":
    main()
