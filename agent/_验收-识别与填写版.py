# -*- coding: utf-8 -*-
"""
_验收-识别与填写版.py —— 作品要求补强验收：识别分类专属异步 API + 表格填写入口

覆盖：
  A. GET  /api/fill/templates        模板清单（Excel 5 + Word 5，双形态）
  B. POST /api/recognize             识别分类（异步）：提交立即返回 + 轮询完成 +
                                     字段提取正确 + 入库可查 + 批量多文档
  C. POST /api/recognize             负向：找不到的文档 404 / 空 docs 400 /
                                     勾字段线但没给字段名 400
  D. POST /api/fill                  表格填写（异步）：提交 → 轮询完成（写入 / 格式差异 0）
                                     + GET /api/fill/download 下载 + 负向 404
  E. 指纹校验                        跑前跑后：源文档 / 模板表格 / 模板表格-word 零改动

流程：
  1. 记录三个目录的整树指纹（跑前）
  2. 启动 api_server.py 子进程（临时端口）
  3. 依次验收（requests）
  4. 停服务，指纹校验（跑后）
  5. 输出报告（控制台 + agent/输出/验收-识别与填写.txt）
"""
import hashlib
import os
import socket
import subprocess
import sys
import time
from pathlib import Path

import requests

BASE = Path(__file__).resolve().parents[1]
AGENT = BASE / "agent"
OUT_DIR = AGENT / "输出"
OUT_DIR.mkdir(exist_ok=True)

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

REPORT = []
N_PASS = 0
N_ALL = 0


def check(name, ok, detail=""):
    global N_PASS, N_ALL
    N_ALL += 1
    ok = bool(ok)
    if ok:
        N_PASS += 1
    line = ("  [PASS] " if ok else "  [FAIL] ") + name + (("  | " + str(detail)) if detail else "")
    REPORT.append(line)
    print(line)
    return ok


def tree_md5(root: Path) -> dict:
    h = {}
    for p in sorted(root.rglob("*")):
        if p.is_file():
            h[str(p.relative_to(root))] = hashlib.md5(p.read_bytes()).hexdigest()
    return h


def free_port(preferred: int = 8981) -> int:
    for port in range(preferred, preferred + 30):
        with socket.socket() as s:
            if s.connect_ex(("127.0.0.1", port)) != 0:
                return port
    raise RuntimeError("找不到空闲端口")


def wait_job(s: requests.Session, base: str, jid: str, timeout: float = 300):
    """轮询任务直到 done/failed。返回 (最后响应, 状态序列)。"""
    states = []
    t0 = time.time()
    while time.time() - t0 < timeout:
        j = s.get(f"{base}/api/job/{jid}", timeout=10).json()
        st = j.get("status")
        if not states or states[-1] != st:
            states.append(st)
        if st in ("done", "failed"):
            return j, states
        time.sleep(0.8)
    return None, states


def main() -> int:
    t_all = time.time()
    print("=" * 66)
    print("作品要求补强验收：识别分类异步 API + 表格填写入口")
    print("=" * 66)

    # ---------- 指纹（跑前） ----------
    trees = {
        "源文档": BASE / "测试数据" / "源文档",
        "模板表格": BASE / "测试数据" / "模板表格",
        "模板表格-word": BASE / "测试数据" / "模板表格-word",
    }
    before = {k: tree_md5(v) for k, v in trees.items()}

    port = free_port()
    base = f"http://127.0.0.1:{port}"
    log_path = OUT_DIR / "api_server-识别与填写验收.log"
    env = dict(os.environ)
    env["PYTHONIOENCODING"] = "utf-8"
    log_f = open(log_path, "w", encoding="utf-8")
    proc = subprocess.Popen(
        [sys.executable, str(AGENT / "api_server.py"), "--port", str(port)],
        cwd=str(BASE), stdout=log_f, stderr=subprocess.STDOUT, env=env,
    )

    s = requests.Session()
    s.trust_env = False

    def finish():
        try:
            proc.terminate()
            try:
                proc.wait(timeout=8)
            except Exception:
                proc.kill()
        except Exception:
            pass
        log_f.close()

        # ---------- 指纹（跑后） ----------
        print("")
        print("[E] 指纹校验（源文档 / 模板：系统应零改动）")
        for k, v in trees.items():
            after = tree_md5(v)
            same = (after == before[k])
            check("E·%s 跑前跑后一致（%d 个文件）" % (k, len(after)),
                  same, "" if same else "有文件被改动！")

        elapsed = time.time() - t_all
        print("")
        print("=" * 66)
        print("验收成绩：%d / %d 通过（用时 %.1f 秒）" % (N_PASS, N_ALL, elapsed))
        print("=" * 66)
        REPORT.append("")
        REPORT.append("=" * 66)
        REPORT.append("验收成绩：%d / %d 通过（用时 %.1f 秒）" % (N_PASS, N_ALL, elapsed))
        REPORT.append("=" * 66)
        rp = OUT_DIR / "验收-识别与填写.txt"
        rp.write_text("\n".join(REPORT) + "\n", encoding="utf-8")
        print("（报告已存档：agent/输出/验收-识别与填写.txt）")
        return 0 if N_PASS == N_ALL else 1

    try:
        # ---------- 就绪等待 ----------
        t0 = time.time()
        ready = False
        while time.time() - t0 < 30:
            if proc.poll() is not None:
                break
            try:
                r = s.get(base + "/api/status", timeout=2)
                if r.status_code == 200:
                    ready = True
                    break
            except Exception:
                time.sleep(0.5)
        if not check("服务启动就绪", ready, "端口 %d，%.1fs" % (port, time.time() - t0)):
            tail = log_path.read_text(encoding="utf-8", errors="replace")[-800:]
            print("---- 服务日志尾部 ----")
            print(tail)
            return finish()

        # ---------- A. 模板清单 ----------
        print("\n[A] 模板清单（GET /api/fill/templates）")
        tps = s.get(base + "/api/fill/templates", timeout=10).json().get("templates", [])
        n_x = sum(1 for t in tps if t["kind"] == "xlsx")
        n_w = sum(1 for t in tps if t["kind"] == "docx")
        check("双形态模板清单齐（Excel %d + Word %d）" % (n_x, n_w),
              n_x >= 5 and n_w >= 5, "共 %d 张" % len(tps))

        # ---------- B. 识别分类（异步） ----------
        print("\n[B] 识别分类异步 API（POST /api/recognize）")
        t_sub = time.time()
        r = s.post(base + "/api/recognize", timeout=30, json={
            "docs": ["设备采购合同"],
            "tasks": ["profile", "entities", "fields"],
            "fields": ["甲方名称", "合同总金额"],
        })
        sub_sec = time.time() - t_sub
        j0 = r.json()
        check("提交成功（HTTP %d）" % r.status_code, r.status_code == 200, str(j0)[:120])
        check("提交立即返回（%.2fs < 3s，异步不阻塞）" % sub_sec, sub_sec < 3.0, "")
        check("返回任务号与类型", j0.get("job_id") and j0.get("kind") == "recognize",
              "job=%s kind=%s" % (j0.get("job_id"), j0.get("kind")))
        check("文档名解析为全名", (j0.get("docs") or [""])[0] == "设备采购合同.docx",
              "docs=%s" % j0.get("docs"))

        j_first = s.get(base + "/api/job/" + j0["job_id"], timeout=5).json()
        check("首次轮询＝未完成（真异步）", j_first.get("status") in ("queued", "running"),
              "status=%s" % j_first.get("status"))

        j, states = wait_job(s, base, j0["job_id"], timeout=240)
        check("任务完成（状态轨迹 %s）" % "→".join(states),
              j is not None and j.get("status") == "done",
              "" if j else "超时未完成")
        if j:
            res = j.get("result") or {}
            meta = res.get("meta") or {}
            lines = ((meta.get("docs") or [{}])[0].get("lines") or {})
            check("三线全部成功（画像 / 实体 / 字段）",
                  res.get("ok") and all(lines.get(k, {}).get("ok") for k in ("profile", "entities", "fields")),
                  "耗时 %ss" % j.get("elapsed"))
            check("结果报告含识别摘要", "智能识别分类" in (res.get("report") or ""),
                  (res.get("report") or "").splitlines()[1][:60] if res.get("report") else "")

        # 入库核对：字段值可查
        rows = s.get(base + "/api/db/fields", timeout=10,
                     params={"doc": "设备采购合同", "q": "甲方名称"}).json().get("rows", [])
        hit_a = any("华景智造" in (x.get("value") or "") for x in rows)
        rows2 = s.get(base + "/api/db/fields", timeout=10,
                      params={"doc": "设备采购合同", "q": "合同总金额"}).json().get("rows", [])
        hit_b = any("1,800,000" in (x.get("value") or "") for x in rows2)
        check("入库可查：甲方名称＝华景智造…", hit_a, "")
        check("入库可查：合同总金额＝1,800,000", hit_b, "")

        # 批量：2 份文档 × 实体线
        r = s.post(base + "/api/recognize", timeout=30, json={
            "docs": ["项目周报", "设备验收记录"], "tasks": ["实体"]})
        j0b = r.json()
        check("批量提交（2 份文档 × 实体线）", r.status_code == 200 and len(j0b.get("docs") or []) == 2,
              str(j0b.get("docs")))
        jb, states_b = wait_job(s, base, j0b["job_id"], timeout=240)
        if jb:
            metas = (jb.get("result") or {}).get("meta") or {}
            docs_meta = metas.get("docs") or []
            ok_docs = len(docs_meta) == 2 and all(
                d.get("lines", {}).get("entities", {}).get("ok") for d in docs_meta)
            check("批量完成：两文档实体线全成功", (jb.get("result") or {}).get("ok") and ok_docs,
                  "耗时 %ss" % jb.get("elapsed"))
        else:
            check("批量完成：两文档实体线全成功", False, "超时未完成")

        # ---------- C. 负向分支 ----------
        print("\n[C] 负向分支")
        r = s.post(base + "/api/recognize", timeout=15, json={"docs": ["不存在的文档XYZ"]})
        check("找不到文档 → 404", r.status_code == 404, r.json().get("detail", "")[:60])
        r = s.post(base + "/api/recognize", timeout=15, json={"docs": []})
        check("空 docs → 400", r.status_code == 400, r.json().get("detail", "")[:60])
        r = s.post(base + "/api/recognize", timeout=15, json={"docs": ["设备采购合同"], "tasks": ["fields"]})
        check("勾字段线但没给字段名 → 400", r.status_code == 400, r.json().get("detail", "")[:60])

        # ---------- D. 表格填写（异步） ----------
        print("\n[D] 表格填写（POST /api/fill → 轮询 → 下载）")
        t_sub = time.time()
        r = s.post(base + "/api/fill", timeout=30,
                   json={"template": "模板4-客诉处理台账.xlsx", "skip_read": True})
        sub2 = time.time() - t_sub
        j1 = r.json()
        check("提交成功且立即返回（%.2fs）" % sub2,
              r.status_code == 200 and sub2 < 3.0, "job=%s" % j1.get("job_id"))
        jf, states_f = wait_job(s, base, j1["job_id"], timeout=420)
        check("填写任务完成（状态轨迹 %s）" % "→".join(states_f),
              jf is not None and jf.get("status") == "done" and (jf.get("result") or {}).get("ok"),
              "" if jf else "超时未完成")
        if jf:
            res = jf.get("result") or {}
            rep = res.get("report") or ""
            meta = res.get("meta") or {}
            check("报告含关键证据（写入 / 格式差异 0）",
                  ("写入" in rep) and ("格式差异 0 处" in rep),
                  " ｜ ".join(rep.splitlines()[1:3])[:100])
            check("产出文件落盘", bool(meta.get("output_file")) and Path(meta["output_file"]).exists(),
                  meta.get("output_file", ""))
            check("分阶段计时齐全", len(meta.get("stages") or {}) >= 3,
                  str(meta.get("stages"))[:100])

            dl = s.get(base + "/api/fill/download", timeout=30,
                       params={"template": "模板4-客诉处理台账"})
            check("下载端点可用（HTTP %d，%d 字节）" % (dl.status_code, len(dl.content)),
                  dl.status_code == 200 and len(dl.content) > 2000,
                  dl.headers.get("content-disposition", "")[:70])

            dl2 = s.get(base + "/api/fill/download", timeout=15, params={"template": "根本没有的模板"})
            check("下载负向：没跑过的模板 → 404", dl2.status_code == 404, "")

        # 负向：不在清单里的模板
        r = s.post(base + "/api/fill", timeout=15, json={"template": "根本不存在的模板.xlsx"})
        check("填写负向：模板不在清单 → 404", r.status_code == 404, r.json().get("detail", "")[:60])

        return finish()
    except Exception as e:
        print("验收中断：%s: %s" % (type(e).__name__, e))
        return finish()


if __name__ == "__main__":
    sys.exit(main())
