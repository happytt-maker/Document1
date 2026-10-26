# -*- coding: utf-8 -*-
"""
_验收-入库版.py —— 模块（2）· 第 4 步「提取结果入库」验收

29 项检查，走真 HTTP 链路（服务需在 8000 端口运行）+ 真 CLI：
  0 组（2 项）：前置 —— 服务就绪 / 数据库就绪
  A 组（10 项）：三线落库 —— 实体 / 画像 / 字段 各跑一条真指令，
      对分「产物 json ↔ 数据库」条数与值集一致、保真度（quote/loc 非空率 100%）
  B 组（2 项）：幂等 —— 同一产物重复落库不膨胀、第二次"新增 0"
  C 组（4 项）：补录与查询 —— store.py --import-all 幂等 / --search / --docs
  D 组（6 项）：API 端点 —— stats / entities 过滤 / fields / search / 边界
  E 组（5 项）：边界与安全 —— 源文件零改动 / 编辑指令不进库 / 全表 quote 保真度 /
      报告完整性

产物归档：agent/输出/验收-入库/（报告 + 明细 json）
"""
import hashlib
import json
import socket
import subprocess
import sys
import time
import urllib.request
from pathlib import Path
from urllib.parse import quote

try:
    sys.stdout.reconfigure(errors="replace")
except Exception:
    pass

BASE = Path(__file__).resolve().parents[1]
OUT_DIR = BASE / "agent" / "输出"
ARC_DIR = OUT_DIR / "验收-入库"
RESULT_FILE = ARC_DIR / "入库-验收报告.txt"

API = "http://127.0.0.1:8000"
OPENER = urllib.request.build_opener(urllib.request.ProxyHandler({}))
PY = sys.executable
STORE = BASE / "agent" / "store.py"

DOC = "设备采购合同"
SRC_NAMES = ["设备采购合同.docx", "员工信息表.xlsx", "设备验收记录.txt",
             "项目周报-第12周.md", "供应商联系方式.txt"]

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


def wait_job(job_id, timeout=180):
    t0 = time.time()
    while time.time() - t0 < timeout:
        j = get("/api/job/" + job_id)
        if j.get("status") in ("done", "failed", "error"):
            return j
        time.sleep(1.5)
    return {"status": "timeout"}


def submit(text, sid):
    d = post("/api/command", {"text": text, "session_id": sid})
    return wait_job(d["job_id"])


def job_rep(j):
    """取任务的报告文本（注意：job 的 ok/report 在 result 子字典里）。"""
    return (j.get("result") or {}).get("report") or ""


def job_ok(j):
    return bool((j.get("result") or {}).get("ok"))


def md5(p: Path) -> str:
    return hashlib.md5(p.read_bytes()).hexdigest()


def load_json(name):
    return json.loads((OUT_DIR / name).read_text(encoding="utf-8"))


def run_store(args, timeout=300):
    """跑一次 store.py 命令行，返回 (退出码, 输出文本)。"""
    r = subprocess.run([PY, str(STORE)] + args, capture_output=True, text=True,
                       encoding="utf-8", errors="replace", timeout=timeout,
                       cwd=str(BASE))
    return r.returncode, (r.stdout or "") + (r.stderr or "")


def main():
    ARC_DIR.mkdir(parents=True, exist_ok=True)
    print("=" * 66)
    print("模块（2）· 第 4 步「提取结果入库」验收")
    print("=" * 66)

    # ---------------- 0. 前置 ----------------
    ok_ready = wait_ready()
    check("0-1", "服务就绪（8000 端口）", ok_ready)
    if not ok_ready:
        return finish()
    st0 = get("/api/db/stats")
    check("0-2", "数据库就绪（五表可用）",
          st0.get("ok") and all(st0["stats"].get(k, 0) >= 0 for k in
                                ("documents", "entities", "profiles", "extractions")),
          "documents=%s entities=%s" % (st0["stats"].get("documents"), st0["stats"].get("entities")))

    # 源文件指纹（E 组用）
    fp0 = {}
    for n in SRC_NAMES:
        for p in (BASE / "测试数据").rglob(n):
            fp0[n] = md5(p)
            break

    sid = "验收入库-%d" % int(time.time() % 100000)

    # ---------------- A 组：三线落库 ----------------
    print("\n---- A 组：三线落库（真指令 → 对分数据库）----")
    # A1 实体
    j = submit("识别设备采购合同里的实体数据", sid)
    rep = job_rep(j)
    check("A1-1", "实体指令：执行成功且报告含【入库】行",
          j.get("status") == "done" and job_ok(j) and "【入库】" in rep,
          "状态=%s" % j.get("status"))
    n_prod = len(load_json("设备采购合同_实体.json").get("entities") or [])
    d = get("/api/db/entities?doc=" + quote(DOC) + "&limit=500")
    rows = d.get("rows") or []
    check("A1-2", "实体条数对分（产物 json ↔ 数据库）",
          len(rows) == n_prod, "产物 %d 条 ｜ 库 %d 条" % (n_prod, len(rows)))
    nq = sum(1 for r in rows if (r.get("quote") or "").strip())
    nl = sum(1 for r in rows if (r.get("loc") or "").strip())
    check("A1-3", "保真度：库中每条带原文摘抄 + 出处（100%）",
          len(rows) > 0 and nq == len(rows) and nl == len(rows),
          "quote %d/%d ｜ loc %d/%d" % (nq, len(rows), nl, len(rows)))
    dvals_chk = {r["value"] for r in rows}
    pvals_chk = {e["value"] for e in load_json("设备采购合同_实体.json").get("entities") or []}
    check("A1-4", "库-产物值集完全一致（同步式落库）",
          dvals_chk == pvals_chk,
          "库 %d 个 ｜ 产物 %d 个 ｜ 差集 %s" % (len(dvals_chk), len(pvals_chk),
                                                list(dvals_chk ^ pvals_chk)[:4]))

    # A2 画像
    j2 = submit("提取设备采购合同的关键信息", sid)
    rep2 = job_rep(j2)
    check("A2-1", "画像指令：执行成功且报告含【入库】行",
          j2.get("status") == "done" and job_ok(j2) and "【入库】" in rep2,
          "状态=%s" % j2.get("status"))
    pj = load_json("设备采购合同_画像.json")
    n_items = len(pj.get("points") or []) + len(pj.get("elements") or [])
    d2 = get("/api/db/profile?doc=" + quote(DOC))
    check("A2-2", "画像明细对分（要点+要素 ↔ 库）",
          d2.get("found") and len(d2.get("items") or []) == n_items,
          "产物 %d 条 ｜ 库 %d 条" % (n_items, len(d2.get("items") or [])))
    topic_ok = bool((d2.get("profile") or {}).get("topic"))
    check("A2-3", "画像主题非空", topic_ok)

    # A3 字段
    j3 = submit("从设备采购合同里提取甲乙双方名称和合同总金额", sid)
    rep3 = job_rep(j3)
    check("A3-1", "字段指令：执行成功且报告含【入库】行",
          j3.get("status") == "done" and job_ok(j3) and "【入库】" in rep3,
          "状态=%s" % j3.get("status"))
    fj = load_json("设备采购合同_提取结果.json")
    n_f = sum(len(f.get("items") or []) for f in fj.get("fields") or [])
    d3 = get("/api/db/fields?doc=" + quote(DOC) + "&limit=500")
    frows = d3.get("rows") or []
    check("A3-2", "字段条数对分（产物 json ↔ 数据库）",
          len(frows) == n_f, "产物 %d 条 ｜ 库 %d 条" % (n_f, len(frows)))
    fields_have = {r["field"] for r in frows}
    check("A3-3", "三个字段齐全（甲方名称/乙方名称/合同总金额）",
          {"甲方名称", "乙方名称", "合同总金额"} <= fields_have,
          "库中字段=%s" % sorted(fields_have))

    # ---------------- B 组：幂等 ----------------
    print("\n---- B 组：幂等（同一产物重复落库不膨胀）----")
    sys.path.insert(0, str(BASE / "agent"))
    import store as store_mod
    jp = OUT_DIR / "设备采购合同_实体.json"
    store_mod.save_any(jp)                      # 预热：确保库与产物对齐
    n_before = len((get("/api/db/entities?doc=" + quote(DOC) + "&limit=500").get("rows")) or [])
    r2 = store_mod.save_any(jp)                 # 同产物再落一次
    n_after = len((get("/api/db/entities?doc=" + quote(DOC) + "&limit=500").get("rows")) or [])
    check("B1", "重复落库：库中条数不变", n_after == n_before,
          "前 %d ｜ 后 %d" % (n_before, n_after))
    check("B2", "第二次落库：新增 0 · 移除 0（全部为更新）",
          r2["new"] == 0 and r2.get("removed", 0) == 0 and r2["updated"] == n_after,
          "new=%d updated=%d removed=%d" % (r2["new"], r2["updated"], r2.get("removed", 0)))

    # ---------------- C 组：补录与查询 CLI ----------------
    print("\n---- C 组：补录与查询（store.py CLI）----")
    rc, out = run_store(["--import-all"])
    check("C1", "--import-all 跑通且失败 0",
          rc == 0 and "失败 0" in out, "退出码=%s" % rc)
    st1 = get("/api/db/stats")["stats"]
    rc2, out2 = run_store(["--import-all"])
    st2 = get("/api/db/stats")["stats"]
    same = all(st1.get(k) == st2.get(k) for k in ("documents", "entities", "profiles", "extractions"))
    check("C2", "补录幂等：再跑一次各表条数不变", same and rc2 == 0,
          "entities %s→%s" % (st1.get("entities"), st2.get("entities")))
    rc3, out3 = run_store(["--search", "华景"])
    check("C3", "--search 跨表搜到（实体+字段）",
          "华景智造科技有限公司" in out3 and "机构" in out3, "退出码=%s" % rc3)
    rc4, out4 = run_store(["--docs"])
    check("C4", "--docs 文档分布含关键文档",
          "设备采购合同.docx" in out4 and "员工信息表.xlsx" in out4,
          "退出码=%s" % rc4)

    # ---------------- D 组：API 端点 ----------------
    print("\n---- D 组：API 端点 ----")
    ds = get("/api/db/stats")
    check("D1", "GET /api/db/stats 五表数字就绪",
          ds.get("ok") and ds["stats"]["documents"] >= 5 and ds["stats"]["entities"] >= 100,
          "documents=%d entities=%d" % (ds["stats"]["documents"], ds["stats"]["entities"]))
    dp = get("/api/db/entities?type=person&q=" + quote("王"))
    vals = [r["value"] for r in dp.get("rows") or []]
    check("D2", "GET /api/db/entities 类型+关键词过滤", len(vals) >= 2 and all("王" in v for v in vals),
          "命中=%s" % vals[:5])
    dc = get("/api/db/entities?doc=" + quote("供应商") + "&type=contact")
    check("D3", "GET /api/db/entities 文档+类型过滤（联系方式）",
          len(dc.get("rows") or []) >= 4, "条数=%d" % len(dc.get("rows") or []))
    df = get("/api/db/fields?q=E1001")
    check("D4", "GET /api/db/fields 关键词查询",
          len(df.get("rows") or []) >= 1 and "E1001" in df["rows"][0]["value"])
    dse = get("/api/db/search?q=" + quote("泰隆"))
    kinds = {r["kind"] for r in dse.get("rows") or []}
    check("D5", "GET /api/db/search 跨实体+字段", len(dse.get("rows") or []) >= 4 and len(kinds) >= 2,
          "kinds=%s" % sorted(kinds))
    dn = get("/api/db/entities?q=" + quote("绝不可能存在的词XYZ"))
    dnf = get("/api/db/profile?doc=" + quote("不存在的文档ABC"))
    check("D6", "边界：不存在 → 空结果不报错",
          len(dn.get("rows") or []) == 0 and dnf.get("found") is False)

    # ---------------- E 组：边界与安全 ----------------
    print("\n---- E 组：边界与安全 ----")
    ok_fp = True
    detail_fp = []
    for n, h0 in fp0.items():
        p = None
        for q in (BASE / "测试数据").rglob(n):
            p = q
            break
        h1 = md5(p) if p else ""
        detail_fp.append("%s %s" % (n, "✓" if h1 == h0 else "XX"))
        ok_fp = ok_fp and h1 == h0
    check("E1", "5 份源文件零改动（md5）", ok_fp, " ｜ ".join(detail_fp))

    je = submit("把设备采购合同的标题居中", sid)
    rep_e = job_rep(je)
    check("E2", "编辑类指令不进库（报告无【入库】行）",
          je.get("status") == "done" and job_ok(je) and "【入库】" not in rep_e,
          "状态=%s" % je.get("status"))

    all_ents = get("/api/db/entities?limit=500").get("rows") or []
    n_all = len(all_ents)
    n_q = sum(1 for r in all_ents if (r.get("quote") or "").strip())
    check("E3", "全表保真度：实体 quote 非空率 100%",
          n_all >= 100 and n_q == n_all,
          "quote %d/%d（全库 %d 条）" % (n_q, n_all, n_all))

    check("E4", "编辑类报告仍含【输出】行（入库改造无退化）",
          "【输出】" in rep_e)

    check("E-extra", "报告含入库行时也含【输出】产物路径（实体报告完整性）",
          "【输出】" in rep and "已存入数据库" in rep)

    return finish()


def finish():
    print("\n" + "=" * 66)
    n_pass = sum(1 for r in RESULTS if r["ok"])
    n_all = len(RESULTS)
    print("验收结果：%d / %d %s" % (n_pass, n_all, "✅ 全部通过" if n_pass == n_all else "❌ 有未过项"))
    print("=" * 66)
    ARC_DIR.mkdir(parents=True, exist_ok=True)
    lines = ["=" * 66,
             "模块（2）· 第 4 步「提取结果入库」验收报告",
             "时间：%s ｜ 成绩：%d / %d" % (time.strftime("%Y-%m-%d %H:%M:%S"), n_pass, n_all),
             "=" * 66, ""]
    for r in RESULTS:
        lines.append("[%s] %s ｜ %s" % ("PASS" if r["ok"] else "FAIL", r["name"], r["detail"]))
    lines += ["", "现场说明：",
              "  · 三线（实体/画像/字段）落库对分口径：产物 json 条数/值集 == 数据库",
              "  · 保真度 = 库中每条数据带原文摘抄 quote + 出处 loc 的比例",
              "  · 落库为同步式：产物是新快照 → 新增/更新/移除三步对齐（幂等、不膨胀）",
              "  · 归档：本目录（报告 + 明细 json）"]
    RESULT_FILE.write_text("\n".join(lines), encoding="utf-8")
    (ARC_DIR / "验收明细.json").write_text(
        json.dumps(RESULTS, ensure_ascii=False, indent=2), encoding="utf-8")
    print("报告已写入：%s" % RESULT_FILE)
    sys.exit(0 if n_pass == n_all else 1)


if __name__ == "__main__":
    main()
