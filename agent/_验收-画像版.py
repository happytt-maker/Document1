# -*- coding: utf-8 -*-
"""
_验收-画像版.py —— 模块（2）· 第 2 步「关键信息提取（文档画像）」验收

29 项检查，走真 HTTP 链路（服务需在 8000 端口运行）：
  A 组（4 份文档 × 6 项）：四格式画像 —— 执行成功 / 主题 / 要点条数 /
      出处定位率≥80% / 覆盖度≥80%（对照人工清单）/ 数字全部可溯（不编造硬检查）
  B 组（2 项）：多轮指代 —— "这份文档的关键信息整理一下"跟着上一句的文档走
  C 组（3 项）：边界与安全 —— 不存在文档失败 / 8 个产物齐全 / 4 份源文件零改动

人工清单：验收脚本里写死的"必含信息点"（写脚本前人工通读源文档圈定），
每个信息点给多个同义 token，任一命中即算覆盖。

产物归档：agent/输出/验收-画像/（报告 + 明细 + 4 份画像产物副本）
"""
import hashlib
import json
import re
import shutil
import socket
import sys
import time
import urllib.request
from pathlib import Path

try:
    sys.stdout.reconfigure(errors="replace")
except Exception:
    pass

BASE = Path(__file__).resolve().parents[1]
OUT_DIR = BASE / "agent" / "输出"
ARC_DIR = OUT_DIR / "验收-画像"
RESULT_FILE = ARC_DIR / "画像验收报告.txt"

API = "http://127.0.0.1:8000"
OPENER = urllib.request.build_opener(urllib.request.ProxyHandler({}))

# ==================== 人工清单（必含信息点） ====================

CHECKLIST = {
    "设备采购合同.docx": [
        ["HT-2026-0417"],
        ["华景智造"],
        ["泰隆机电"],
        ["1,800,000", "壹佰捌拾万"],
        ["2026年3月12日", "3月12日", "2026-03-12"],
        ["30%"],
        ["6月30日"],
        ["24个月", "24 个月"],
        ["VMC-850B", "数控加工中心"],
        ["江宁区"],
    ],
    "员工信息表.xlsx": [
        ["2026年第一季度", "第一季度"],
        ["8名", "8人", "8 名", "8 位", "8位"],
        ["E1001"],
        ["E1008"],
        ["制造中心"],
        ["财务部"],
        ["研发中心"],
        ["19800"],
        ["8800"],
        ["绩效系数"],
    ],
    "项目周报-第12周.md": [
        ["第 12 周", "第12周"],
        ["2026-03-16"],
        ["吴晓敏"],
        ["单元测试"],
        ["xlsx 解析器", "xlsx解析器"],
        ["公式"],
        ["合并单元格"],
        ["04-15", "4月15"],
        ["测试服务器"],
        ["报销单据"],
    ],
    "设备验收记录.txt": [
        ["YS-2026-0089"],
        ["数控加工中心", "VMC-850B"],
        ["2026年6月28日", "6月28日"],
        ["泰隆机电"],
        ["合格"],
        ["0.008mm", "0.008"],
        ["0.005mm", "0.005"],
        ["郑凯"],
        ["李博文"],
        ["第一生产车间"],
    ],
}

CASES = [
    ("设备采购合同.docx", "提取设备采购合同的关键信息"),
    ("员工信息表.xlsx", "提取员工信息表的关键信息"),
    ("项目周报-第12周.md", "把项目周报-第12周的关键信息整理出来"),
    ("设备验收记录.txt", "提取设备验收记录的关键信息"),
]

RESULTS = []


def check(no, name, ok, detail=""):
    RESULTS.append({"no": no, "name": name, "ok": bool(ok), "detail": detail})
    print("[%s] %s ｜ %s" % ("PASS" if ok else "FAIL", name, detail))


def wait_ready(timeout=60):
    t0 = time.time()
    while time.time() - t0 < timeout:
        s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        s.settimeout(2)
        try:
            s.connect(("127.0.0.1", 8000))
            s.close()
            return True
        except Exception:
            s.close()
            time.sleep(2)
    return False


def post(path, payload):
    req = urllib.request.Request(API + path, data=json.dumps(payload).encode("utf-8"),
                                 headers={"Content-Type": "application/json"}, method="POST")
    return json.loads(OPENER.open(req, timeout=30).read().decode("utf-8"))


def get(path):
    return json.loads(OPENER.open(API + path, timeout=20).read().decode("utf-8"))


def wait_job(job_id, timeout=150):
    t0 = time.time()
    while time.time() - t0 < timeout:
        j = get("/api/job/" + job_id)
        if j.get("status") in ("done", "failed", "error"):
            return j
        time.sleep(1.5)
    return {"status": "timeout"}


def md5(p: Path) -> str:
    return hashlib.md5(p.read_bytes()).hexdigest()


def norm_text(s: str) -> str:
    s = (s or "").replace(",", "").replace("，", "")
    return re.sub(r"\s+", "", s)


def num_tokens(s: str):
    toks = []
    toks += re.findall(r"\d{2,}", s)
    toks += re.findall(r"\d{1,3}%", s)
    toks += re.findall(r"\d{1,2}月\d{1,2}日", s)
    return sorted(set(toks))


def doc_fulltext(doc_name: str) -> str:
    d = json.loads((BASE / "doc_reader" / "output.json").read_text(encoding="utf-8"))
    for doc in d.get("documents", []):
        if doc.get("file") == doc_name:
            return norm_text("".join((b.get("text") or "") for b in doc.get("blocks", [])))
    return ""


def main():
    ARC_DIR.mkdir(parents=True, exist_ok=True)
    print("=" * 66)
    print("模块（2）· 第 2 步「关键信息提取（文档画像）」验收")
    print("=" * 66)

    if not wait_ready():
        print("【失败】服务没就绪（8000 端口）——先跑 启动操作台.bat 再验收。")
        sys.exit(1)

    # 源文件指纹（跑前）
    src_dir = BASE / "测试数据" / "源文档"
    fp_before = {}
    for name, _ in CASES:
        for p in src_dir.rglob(name):
            fp_before[name] = md5(p)

    # ---------- A 组：四份文档画像 ----------
    no = 0
    for doc_name, cmd in CASES:
        print("")
        print("---- A 组：%s ----" % doc_name)
        try:
            j0 = post("/api/command", {"text": cmd})
            j = wait_job(j0["job_id"])
            res = j.get("result") or {}
            report = res.get("report") or ""
        except Exception as e:
            no += 1
            check(no, "%s：执行" % doc_name, False, "异常 %s" % e)
            continue

        # 1) 成功 + 报告
        ok1 = (j.get("status") == "done") and res.get("ok") and ("【画像】" in report)
        no += 1
        check(no, "%s：执行成功且报告为画像" % doc_name, ok1,
              "status=%s ok=%s" % (j.get("status"), res.get("ok")))

        # 读 json 产物
        jp = OUT_DIR / (Path(doc_name).stem + "_画像.json")
        data = json.loads(jp.read_text(encoding="utf-8")) if jp.exists() else {}

        # 2) 主题非空
        no += 1
        topic = (data.get("topic") or "").strip()
        check(no, "%s：主题非空" % doc_name, len(topic) >= 10,
              topic[:60])

        # 3) 要点 3~7 条
        pts = data.get("points") or []
        no += 1
        check(no, "%s：要点 3~7 条" % doc_name, 3 <= len(pts) <= 7, "实际 %d 条" % len(pts))

        # 4) 出处定位率 ≥80%
        els = data.get("elements") or []
        all_items = pts + els
        n_hit = sum(1 for x in all_items if x.get("loc_ok"))
        rate = (n_hit / len(all_items) * 100) if all_items else 0
        no += 1
        check(no, "%s：出处定位率 ≥80%%" % doc_name, rate >= 80 and len(all_items) > 0,
              "%d/%d = %.0f%%" % (n_hit, len(all_items), rate))

        # 5) 覆盖度 ≥80%（对照人工清单）
        merged = topic + " ".join(
            [p.get("text", "") for p in pts] +
            ["%s %s" % (e.get("name", ""), e.get("value", "")) for e in els])
        cl = CHECKLIST.get(doc_name, [])
        hit_pts = [c for c in cl if any(t in merged for t in c)]
        cov = (len(hit_pts) / len(cl) * 100) if cl else 0
        miss = [c[0] for c in cl if c not in hit_pts]
        no += 1
        check(no, "%s：信息覆盖 ≥80%%" % doc_name, cov >= 80,
              "%d/%d = %.0f%%%s" % (len(hit_pts), len(cl), cov,
                                    (" ｜ 缺：" + "、".join(miss)) if miss else ""))

        # 6) 数字全部可溯（不编造硬检查）
        full = doc_fulltext(doc_name)
        toks = num_tokens(" ".join([p.get("text", "") for p in pts] +
                                   [e.get("value", "") for e in els]))
        bad = [t for t in toks if t not in full]
        no += 1
        check(no, "%s：转述中的数字全部可在原文溯源" % doc_name, len(bad) == 0,
              "检查 %d 个数字%s" % (len(toks), (" ｜ 溯源失败：" + "、".join(bad[:6])) if bad else ""))

    # ---------- B 组：多轮指代 ----------
    print("")
    print("---- B 组：多轮指代 ----")
    sid = "acc-profile-b-%d" % int(time.time())
    try:
        j1 = wait_job(post("/api/command", {"text": "把设备采购合同的标题居中",
                                            "session_id": sid})["job_id"])
        j2 = wait_job(post("/api/command", {"text": "这份文档的关键信息整理一下",
                                            "session_id": sid})["job_id"])
        r2 = (j2.get("result") or {}).get("report") or ""
        no += 1
        check(no, "B-1 第 1 轮（标题居中）完成", j1.get("status") == "done",
              "status=%s" % j1.get("status"))
        no += 1
        ok2 = (j2.get("status") == "done") and ("设备采购合同.docx" in r2) and ("【画像】" in r2)
        check(no, "B-2 「这份文档」指代到上轮文档并出画像", ok2,
              "报告含文档名=%s ｜ 含【画像】=%s"
              % ("设备采购合同.docx" in r2, "【画像】" in r2))
    except Exception as e:
        no += 2
        check(no - 1, "B-1 第 1 轮（标题居中）完成", False, "异常 %s" % e)
        check(no, "B-2 「这份文档」指代", False, "异常 %s" % e)

    # ---------- C 组：边界与安全 ----------
    print("")
    print("---- C 组：边界与安全 ----")
    try:
        j3 = wait_job(post("/api/command", {"text": "提取不存在的文档ABC的关键信息"})["job_id"])
        r3 = (j3.get("result") or {}).get("report") or ""
        ok3 = (j3.get("status") in ("done", "failed")) and (
            (not (j3.get("result") or {}).get("ok")) and
            ("找不到" in r3 or "反问" in r3 or "不明确" in r3))
        no += 1
        check(no, "C-1 不存在的文档 → 失败/反问（不瞎猜）", ok3,
              "status=%s ｜ %s" % (j3.get("status"), r3.strip().splitlines()[-1][:60] if r3 else ""))
    except Exception as e:
        no += 1
        check(no, "C-1 不存在的文档 → 失败/反问", False, "异常 %s" % e)

    # C-2 产物齐全
    missing = []
    for doc_name, _ in CASES:
        for ext in (".md", ".json"):
            p = OUT_DIR / (Path(doc_name).stem + "_画像" + ext)
            if not p.exists():
                missing.append(p.name)
    no += 1
    check(no, "C-2 画像产物齐全（4 份 × md+json）", not missing,
          "缺：" + "、".join(missing) if missing else "8 个产物全部就位")

    # C-3 源文件零改动
    changed = []
    for name, _ in CASES:
        for p in src_dir.rglob(name):
            if fp_before.get(name) != md5(p):
                changed.append(name)
    no += 1
    check(no, "C-3 源文档零改动（4 份指纹）", not changed,
          "改动：" + "、".join(changed) if changed else "md5 全部一致")

    # ---------- 归档 ----------
    for doc_name, _ in CASES:
        for ext in (".md", ".json"):
            src = OUT_DIR / (Path(doc_name).stem + "_画像" + ext)
            if src.exists():
                shutil.copy2(src, ARC_DIR / src.name)
    (ARC_DIR / "验收明细.json").write_text(
        json.dumps(RESULTS, ensure_ascii=False, indent=2), encoding="utf-8")

    n_pass = sum(1 for r in RESULTS if r["ok"])
    n_all = len(RESULTS)
    lines = []
    lines.append("=" * 66)
    lines.append("模块（2）· 第 2 步「关键信息提取（文档画像）」验收报告")
    lines.append("=" * 66)
    lines.append("时间：%s ｜ 成绩：%d / %d" % (time.strftime("%Y-%m-%d %H:%M:%S"), n_pass, n_all))
    lines.append("")
    for r in RESULTS:
        lines.append("[%s] %s ｜ %s" % ("PASS" if r["ok"] else "FAIL", r["name"], r["detail"]))
    lines.append("")
    lines.append("现场说明：")
    lines.append("  · 四格式画像产物已归档本目录（4 份 × md+json）")
    lines.append("  · 覆盖度对照清单为人工通读源文档后圈定（脚本内置）")
    lines.append("  · 数字溯源 = 把画像中转述的数字回原文全文核对（不编造硬检查）")
    RESULT_FILE.write_text("\n".join(lines), encoding="utf-8")

    print("")
    print("=" * 66)
    print("验收结果：%d / %d 通过%s" % (n_pass, n_all, "  ✅ 全部通过" if n_pass == n_all else ""))
    print("报告：%s" % RESULT_FILE)
    sys.exit(0 if n_pass == n_all else 1)


if __name__ == "__main__":
    main()
