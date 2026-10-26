# -*- coding: utf-8 -*-
"""
_验收-实体版.py —— 模块（2）· 第 3 步「实体数据提取」验收

35 项检查，走真 HTTP 链路（服务需在 8000 端口运行）：
  A 组（5 份文档 × 6 项）：跨格式 + 表格密度文档 —— 执行成功 / 7 类结构合法 /
      出处定位率≥80% / 查全率≥90%（对照人工清单）/ 类型正确率≥85% / 去重生效
  B 组（2 项）：多轮指代 —— "识别这份文档里的实体数据"跟着上一句的文档走
  C 组（3 项）：边界与安全 —— 不存在文档失败 / 10 个产物齐全 / 5 份源文件零改动

人工清单：写脚本前人工通读源文档圈定「必含实体」；每个实体给多个可接受形态
（如金额的￥/大写写法、公司全称/简称），归一后与模型输出"双向包含"即算命中。
  查全率 = 清单命中数 / 清单总数
  类型正确率 = 值命中中「类型也归对」的占比（考核 7 类分类能力）

产物归档：agent/输出/验收-实体/（报告 + 明细 + 5 份实体表副本）
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
ARC_DIR = OUT_DIR / "验收-实体"
RESULT_FILE = ARC_DIR / "实体验收报告.txt"

API = "http://127.0.0.1:8000"
OPENER = urllib.request.build_opener(urllib.request.ProxyHandler({}))

TYPE_ORDER = ["person", "org", "date", "money", "code", "contact", "product"]

# ==================== 人工清单（必含实体；每项 = 可接受形态列表） ====================

CHECKLIST = {
    "设备采购合同.docx": {
        "org": [["华景智造科技有限公司", "华景智造"],
                ["泰隆机电设备有限公司", "泰隆机电"]],
        "date": [["2026年3月12日"],
                 ["2026 年 6 月 30 日", "2026年6月30日"]],
        "money": [["1,800,000", "￥1,800,000.00", "壹佰捌拾万元", "人民币壹佰捌拾万元整"]],
        "code": [["HT-2026-0417"]],
        "product": [["数控加工中心"], ["VMC-850B"], ["激光切割机"], ["LC-3015F"],
                    ["工业机器人"], ["IRB-2600"], ["检测工作台"], ["QC-2000"]],
    },
    "员工信息表.xlsx": {
        "person": [["王振国"], ["李慧敏"], ["张新民"], ["赵雪松"],
                   ["陈思彤"], ["刘宏伟"], ["孙雅静"], ["周文彬"]],
        "org": [["某某科技有限公司"], ["制造中心"], ["财务部"], ["研发中心"],
                ["采购部"], ["信息化管理部"], ["综合管理部"]],
        "date": [["2018-03-05"], ["2019-07-16"], ["2017-11-02"], ["2021-04-19"],
                 ["2022-08-01"], ["2020-02-24"], ["2023-06-12"], ["2022-09-05"]],
        "code": [["E1001"], ["E1002"], ["E1003"], ["E1004"],
                 ["E1005"], ["E1006"], ["E1007"], ["E1008"]],
        "contact": [["13805170001"], ["13805170002"], ["13805170003"], ["13805170004"],
                    ["13805170005"], ["13805170006"], ["13805170007"], ["13805170008"]],
    },
    "设备验收记录.txt": {
        "person": [["郑凯"], ["李博文"], ["周立"]],
        "org": [["泰隆机电设备有限公司", "泰隆机电"]],
        "date": [["2026年6月28日"]],
        "code": [["YS-2026-0089"], ["TL-850-20260512"], ["HT-2026-0417"]],
        "product": [["数控加工中心"], ["VMC-850B"]],
    },
    "项目周报-第12周.md": {
        "person": [["吴晓敏"]],
        "date": [["2026-03-16"], ["2026-03-22"], ["2026-03-23"],
                 ["04-15"], ["03-25"], ["04-18"]],
        "org": [["财务部"]],
    },
    "供应商联系方式.txt": {
        "person": [["赵雪松"], ["王建平"], ["刘美芳"], ["陈国强"],
                   ["周文杰"], ["孙丽华"], ["吴德明"]],
        "org": [["采购部"], ["泰隆机电设备有限公司"], ["苏州精诚自动化有限公司"],
                ["无锡恒信精密制造有限公司"], ["常州华翔智能装备有限公司"],
                ["扬州天工重工有限公司"], ["南通海通装备有限公司"]],
        "date": [["2026年3月15日"]],
        "contact": [["13901234501"], ["13901234502"], ["13901234503"],
                    ["13901234504"], ["13901234505"], ["13901234506"]],
        "product": [["数控机床"], ["激光设备"], ["工业机器人集成"], ["精密零部件加工"],
                    ["检测设备"], ["工装夹具"], ["重型机械配件"], ["液压系统"], ["密封件"]],
    },
}

CASES = [
    ("设备采购合同.docx", "识别设备采购合同里的实体数据"),
    ("员工信息表.xlsx", "识别员工信息表里的实体数据"),
    ("设备验收记录.txt", "识别设备验收记录里的实体数据"),
    ("项目周报-第12周.md", "识别项目周报-第12周里的实体数据"),
    ("供应商联系方式.txt", "识别供应商联系方式里的实体数据"),
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


# ==================== 值归一 + 比对 ====================
_FULL2HALF = str.maketrans(
    "０１２３４５６７８９ＡＢＣＤＥＦＧＨＩＪＫＬＭＮＯＰＱＲＳＴＵＶＷＸＹＺ"
    "ａｂｃｄｅｆｇｈｉｊｋｌｍｎｏｐｑｒｓｔｕｖｗｘｙｚ",
    "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz")


def norm_v(s: str) -> str:
    s = re.sub(r"\s+", "", s or "")
    s = s.translate(_FULL2HALF)
    for a, b in (("，", ","), ("：", ":"), ("（", "("), ("）", ")")):
        s = s.replace(a, b)
    s = s.replace("￥", "").replace("¥", "")
    s = s.replace(",", "")
    s = re.sub(r"\.00$", "", s)
    return s


def find_hit(alt_list, entities):
    """对清单项（多形态）在模型实体中找命中。返回 (命中?, 命中实体的 type)。"""
    for e in entities:
        ev = norm_v(e.get("value"))
        if not ev:
            continue
        for alt in alt_list:
            av = norm_v(alt)
            if not av:
                continue
            if av == ev or (len(av) >= 2 and av in ev) or (len(ev) >= 2 and ev in av):
                return True, e.get("type")
    return False, None


def main():
    ARC_DIR.mkdir(parents=True, exist_ok=True)
    print("=" * 66)
    print("模块（2）· 第 3 步「实体数据提取」验收")
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

    # ---------- A 组：5 份文档实体提取 ----------
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
            no += 6
            check(no - 5, "%s：执行" % doc_name, False, "异常 %s" % e)
            continue

        # 1) 执行成功 + 报告为【实体】
        ok1 = (j.get("status") == "done") and res.get("ok") and ("【实体】" in report)
        no += 1
        check(no, "%s：执行成功且报告为实体表" % doc_name, ok1,
              "status=%s ok=%s" % (j.get("status"), res.get("ok")))

        # 读 json 产物
        jp = OUT_DIR / (Path(doc_name).stem + "_实体.json")
        data = json.loads(jp.read_text(encoding="utf-8")) if jp.exists() else {}
        entities = data.get("entities") or []

        # 2) 7 类结构合法
        bad = []
        for e in entities:
            if e.get("type") not in TYPE_ORDER:
                bad.append("type=%s" % e.get("type"))
            if not str(e.get("value") or "").strip():
                bad.append("空值")
            if e.get("verify") not in ("ok", "suspect", "na"):
                bad.append("verify=%s" % e.get("verify"))
            if "loc" not in e:
                bad.append("缺 loc")
        no += 1
        check(no, "%s：7 类结构合法（type/value/verify/loc）" % doc_name,
              len(entities) >= 5 and not bad,
              "实体 %d 条%s" % (len(entities), (" ｜ 异常：" + "、".join(bad[:4])) if bad else ""))

        # 3) 出处定位率 ≥80%
        n_hit = sum(1 for e in entities if e.get("loc_ok"))
        rate = (100.0 * n_hit / len(entities)) if entities else 0
        no += 1
        check(no, "%s：出处定位率 ≥80%%" % doc_name, rate >= 80,
              "%d/%d = %.0f%%" % (n_hit, len(entities), rate))

        # 4) 查全率 ≥90%（对照人工清单）
        cl = CHECKLIST.get(doc_name, {})
        total_cl = sum(len(v) for v in cl.values())
        value_hits, type_hits, misses, mism = 0, 0, [], []
        for typ, items in cl.items():
            for alts in items:
                hit_ok, mtype = find_hit(alts, entities)
                if hit_ok:
                    value_hits += 1
                    if mtype == typ:
                        type_hits += 1
                    else:
                        mism.append("%s→%s" % (alts[0], mtype))
                else:
                    misses.append(alts[0])
        recall = 100.0 * value_hits / total_cl if total_cl else 0
        no += 1
        check(no, "%s：查全率 ≥90%%（对照人工清单）" % doc_name, recall >= 90,
              "%d/%d = %.0f%%%s" % (value_hits, total_cl, recall,
                                    (" ｜ 缺：" + "、".join(misses[:6])) if misses else ""))

        # 5) 类型正确率 ≥85%（值命中的里，类型归对的占比）
        tacc = (100.0 * type_hits / value_hits) if value_hits else 0
        no += 1
        check(no, "%s：类型正确率 ≥85%%" % doc_name, tacc >= 85,
              "%d/%d = %.0f%%%s" % (type_hits, value_hits, tacc,
                                    (" ｜ 错归：" + "、".join(mism[:4])) if mism else ""))

        # 6) 去重生效（归一后无重复值）
        seen, dups = set(), []
        for e in entities:
            k = (e.get("type"), norm_v(e.get("value")))
            if k in seen:
                dups.append(e.get("value"))
            seen.add(k)
        no += 1
        check(no, "%s：去重生效（无重复值）" % doc_name, not dups,
              "重复：" + "、".join(dups[:4]) if dups else "归一后 %d 个唯一值" % len(seen))

    # ---------- B 组：多轮指代 ----------
    print("")
    print("---- B 组：多轮指代 ----")
    sid = "acc-entity-b-%d" % int(time.time())
    try:
        j1 = wait_job(post("/api/command", {"text": "把设备采购合同的标题居中",
                                            "session_id": sid})["job_id"])
        j2 = wait_job(post("/api/command", {"text": "识别这份文档里的实体数据",
                                            "session_id": sid})["job_id"])
        r2 = (j2.get("result") or {}).get("report") or ""
        no += 1
        check(no, "B-1 第 1 轮（标题居中）完成", j1.get("status") == "done",
              "status=%s" % j1.get("status"))
        no += 1
        ok2 = (j2.get("status") == "done") and ("设备采购合同.docx" in r2) and ("【实体】" in r2)
        check(no, "B-2 「这份文档」指代到上轮文档并出实体表", ok2,
              "报告含文档名=%s ｜ 含【实体】=%s"
              % ("设备采购合同.docx" in r2, "【实体】" in r2))
    except Exception as e:
        no += 2
        check(no - 1, "B-1 第 1 轮（标题居中）完成", False, "异常 %s" % e)
        check(no, "B-2 「这份文档」指代", False, "异常 %s" % e)

    # ---------- C 组：边界与安全 ----------
    print("")
    print("---- C 组：边界与安全 ----")
    try:
        j3 = wait_job(post("/api/command", {"text": "识别不存在的文档ABC里的实体数据"})["job_id"])
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
            p = OUT_DIR / (Path(doc_name).stem + "_实体" + ext)
            if not p.exists():
                missing.append(p.name)
    no += 1
    check(no, "C-2 实体表产物齐全（5 份 × md+json）", not missing,
          "缺：" + "、".join(missing) if missing else "10 个产物全部就位")

    # C-3 源文件零改动
    changed = []
    for name, _ in CASES:
        for p in src_dir.rglob(name):
            if fp_before.get(name) != md5(p):
                changed.append(name)
    no += 1
    check(no, "C-3 源文档零改动（5 份指纹）", not changed,
          "改动：" + "、".join(changed) if changed else "md5 全部一致")

    # ---------- 归档 ----------
    for doc_name, _ in CASES:
        for ext in (".md", ".json"):
            src = OUT_DIR / (Path(doc_name).stem + "_实体" + ext)
            if src.exists():
                shutil.copy2(src, ARC_DIR / src.name)
    (ARC_DIR / "验收明细.json").write_text(
        json.dumps(RESULTS, ensure_ascii=False, indent=2), encoding="utf-8")

    n_pass = sum(1 for r in RESULTS if r["ok"])
    n_all = len(RESULTS)
    lines = []
    lines.append("=" * 66)
    lines.append("模块（2）· 第 3 步「实体数据提取」验收报告")
    lines.append("=" * 66)
    lines.append("时间：%s ｜ 成绩：%d / %d" % (time.strftime("%Y-%m-%d %H:%M:%S"), n_pass, n_all))
    lines.append("")
    for r in RESULTS:
        lines.append("[%s] %s ｜ %s" % ("PASS" if r["ok"] else "FAIL", r["name"], r["detail"]))
    lines.append("")
    lines.append("现场说明：")
    lines.append("  · 5 份实体表产物已归档本目录（5 份 × md+json）")
    lines.append("  · 查全率对照人工通读源文档圈定的清单（脚本内置，共 101 项）")
    lines.append("  · 类型正确率 = 值命中中「7 类归类也对」的占比")
    lines.append("  · 格式核验（正则二次校验）结果见产物 json 的 verify 字段")
    RESULT_FILE.write_text("\n".join(lines), encoding="utf-8")

    print("")
    print("=" * 66)
    print("验收结果：%d / %d 通过%s" % (n_pass, n_all, "  ✅ 全部通过" if n_pass == n_all else ""))
    print("报告：%s" % RESULT_FILE)
    sys.exit(0 if n_pass == n_all else 1)


if __name__ == "__main__":
    main()
