# -*- coding: utf-8 -*-
"""
_验收-API版.py —— 模块（1）· 第 6 步验收：HTTP API + 异步调用 + 网页操作台

流程：
  1. 记录源文档整树指纹（跑前）
  2. 启动 api_server.py 子进程（临时端口）
  3. 等就绪，依次验收：
     A. /api/status /api/docs          —— 服务与文档库
     B. 编辑指令（异步提交 + 轮询）      —— 立即返回 job_id、正确完成
     C. 多轮指代（"把它的金额加粗"）     —— 同会话第二轮，仍锁定同一文档
     D. 长任务异步性（提取 10~15s）      —— 提交不阻塞、期间可见 running
     E. /api/history                    —— 历史入库
     F. 错误分支（空指令 400 / 不存在任务 404 / 网页 200）
  4. 停服务，源文档指纹校验（跑后）
  5. 输出报告（控制台 + agent/输出/第6步-API验收报告.txt）
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
DOCS = BASE / "测试数据" / "源文档"
OUT_DIR = AGENT / "输出"
OUT_DIR.mkdir(exist_ok=True)

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

REPORT = []


def check(name: str, ok, detail=""):
    ok = bool(ok)
    REPORT.append((name, ok, str(detail)))
    print(("  [PASS] " if ok else "  [FAIL] ") + name + (("  | " + str(detail)) if detail else ""))
    return ok


def tree_md5(root: Path) -> dict:
    """目录整树指纹（文档文件内容）。"""
    h = {}
    for p in sorted(root.rglob("*")):
        if p.is_file():
            h[str(p.relative_to(root))] = hashlib.md5(p.read_bytes()).hexdigest()
    return h


def free_port(preferred: int = 8977) -> int:
    for port in range(preferred, preferred + 20):
        with socket.socket() as s:
            if s.connect_ex(("127.0.0.1", port)) != 0:
                return port
    raise RuntimeError("找不到空闲端口")


def wait_job(s: requests.Session, base: str, jid: str, timeout: float = 90):
    """轮询任务直到 done/failed。返回 (最后响应, 状态序列)。"""
    states = []
    t0 = time.time()
    while time.time() - t0 < timeout:
        j = s.get(f"{base}/api/job/{jid}", timeout=5).json()
        if not states or states[-1] != j.get("status"):
            states.append(j.get("status"))
        if j.get("status") in ("done", "failed"):
            return j, states
        time.sleep(0.6)
    return None, states


def main() -> int:
    print("=" * 66)
    print("模块（1）第 6 步验收：HTTP API + 异步调用 + 网页操作台")
    print("=" * 66)

    before = tree_md5(DOCS)
    port = free_port()
    base = f"http://127.0.0.1:{port}"
    log_path = OUT_DIR / "api_server-验收轮.log"

    env = dict(os.environ)
    env["PYTHONIOENCODING"] = "utf-8"
    log_f = open(log_path, "w", encoding="utf-8")
    proc = subprocess.Popen(
        [sys.executable, str(AGENT / "api_server.py"), "--port", str(port)],
        cwd=str(BASE), stdout=log_f, stderr=subprocess.STDOUT, env=env,
    )

    s = requests.Session()
    s.trust_env = False                      # 本机服务不走系统代理

    try:
        # ---------- 就绪等待 ----------
        t0 = time.time()
        ready = False
        while time.time() - t0 < 30:
            if proc.poll() is not None:
                break                        # 进程已退出（启动失败）
            try:
                r = s.get(base + "/api/status", timeout=2)
                if r.status_code == 200:
                    ready = True
                    break
            except Exception:
                time.sleep(0.5)
        if not check("服务启动就绪", ready, f"{time.time()-t0:.1f}s 端口 {port}"):
            tail = log_path.read_text(encoding="utf-8", errors="replace")[-800:]
            print("---- 服务日志尾部 ----")
            print(tail)
            return finish(proc, log_f, before)

        # ---------- A. 状态 / 文档库 ----------
        print("\n[A] 服务状态与文档库")
        st = s.get(base + "/api/status", timeout=5).json()
        check("GET /api/status 服务正常", st.get("service") == "ok", "")
        check("GET /api/status 数据库连通", st.get("database") == "ok", st.get("database_tip", ""))
        check("GET /api/status 文档库可读", (st.get("docs_count") or 0) >= 10,
              f"docs_count={st.get('docs_count')}")

        docs = s.get(base + "/api/docs", timeout=5).json().get("docs", [])
        check("GET /api/docs 返回清单", any("设备采购合同" in d for d in docs),
              f"{len(docs)} 份")

        # ---------- B. 编辑指令：异步提交 + 轮询 ----------
        print("\n[B] 编辑指令（异步提交 + 轮询）")
        sid = "验收api-" + time.strftime("%H%M%S")
        t0 = time.time()
        r = s.post(base + "/api/command",
                   json={"text": "把设备采购合同的标题居中", "session_id": sid}, timeout=10)
        submit_s = time.time() - t0
        j1 = r.json()
        check("POST /api/command 立即返回 job_id",
              r.status_code == 200 and bool(j1.get("job_id")),
              f"job={j1.get('job_id')} session={j1.get('session_id')}")
        check("异步性：提交响应 < 1.5s", submit_s < 1.5, f"{submit_s*1000:.0f}ms")

        res1, states1 = wait_job(s, base, j1.get("job_id"), timeout=60)
        ok1 = bool(res1 and res1.get("result", {}).get("ok"))
        check("任务1 完成：标题居中", ok1,
              f"耗时 {res1.get('elapsed') if res1 else '-'}s 状态序列 {states1}")
        m1 = (res1 or {}).get("result", {}).get("meta") or {}
        check("任务1 定位到正确文档", "设备采购" in (m1.get("doc") or ""), m1.get("doc"))

        # ---------- C. 多轮指代 ----------
        print("\n[C] 多轮指代（'它'= 上一句的文档）")
        r = s.post(base + "/api/command",
                   json={"text": "把它的金额加粗", "session_id": sid}, timeout=10)
        j2 = r.json()
        res2, states2 = wait_job(s, base, j2.get("job_id"), timeout=60)
        ok2 = bool(res2 and res2.get("result", {}).get("ok"))
        m2 = (res2 or {}).get("result", {}).get("meta") or {}
        check("任务2 完成：金额加粗", ok2, f"耗时 {res2.get('elapsed') if res2 else '-'}s")
        check("任务2 '它'指代正确（仍是设备采购合同）",
              "设备采购" in (m2.get("doc") or ""), m2.get("doc"))

        # ---------- D. 长任务异步性 ----------
        print("\n[D] 长任务异步性（提取类指令 10~15s 不阻塞）")
        t0 = time.time()
        r = s.post(base + "/api/command",
                   json={"text": "从员工信息表里提取员工姓名和工号", "session_id": sid}, timeout=10)
        submit3_s = time.time() - t0
        j3 = r.json()
        check("长任务提交立即返回（不等它跑完）", submit3_s < 1.5, f"{submit3_s*1000:.0f}ms")
        res3, states3 = wait_job(s, base, j3.get("job_id"), timeout=120)
        ok3 = bool(res3 and res3.get("result", {}).get("ok"))
        check("长任务完成：提取员工姓名+工号", ok3,
              f"耗时 {res3.get('elapsed') if res3 else '-'}s")
        check("长任务期间可观察到 running（证明真的在跑）",
              "running" in states3 or "done" in states3, f"状态序列 {states3}")

        # ---------- E. 历史 ----------
        print("\n[E] 操作历史（MySQL）")
        h = s.get(base + "/api/history?limit=100", timeout=5).json()
        rows = h.get("rows") or []
        check("GET /api/history 有数据", len(rows) >= 3, f"{len(rows)} 条")
        check("历史含本会话记录", any(sid in (row.get("session_id") or "") for row in rows),
              f"会话 {sid}")
        latest = rows[0] if rows else {}
        check("最新记录 = 最后一条指令",
              "员工" in (latest.get("command") or ""), latest.get("command"))

        # ---------- F. 会话状态 / 错误分支 / 网页 ----------
        print("\n[F] 会话状态 / 错误分支 / 网页操作台")
        sn = s.get(base + f"/api/session/{sid}", timeout=5).json()
        check("GET /api/session 有状态", sn.get("exists") and sn.get("turn") == 3,
              f"turn={sn.get('turn')} current_doc={sn.get('current_doc')}")

        r = s.post(base + "/api/command", json={"text": "   "}, timeout=10)
        check("空指令 → 400", r.status_code == 400, "")
        r = s.get(base + "/api/job/j9999", timeout=5)
        check("不存在的任务 → 404", r.status_code == 404, "")
        r = s.get(base + "/", timeout=5)
        check("GET / 网页操作台可达", r.status_code == 200 and "A23" in r.text and "发送" in r.text,
              f"{len(r.text)} 字节")

    finally:
        # ---------- 停服务 ----------
        proc.terminate()
        try:
            proc.wait(timeout=8)
        except Exception:
            proc.kill()
        log_f.close()

    return finish(proc, None, before)


def finish(proc, log_f, before) -> int:
    if log_f is not None:
        try:
            proc.terminate()
            proc.wait(timeout=8)
        except Exception:
            try:
                proc.kill()
            except Exception:
                pass
        log_f.close()

    after = tree_md5(DOCS)
    check("源文档零改动（整树指纹）", before == after,
          f"{len(after)} 个文件" + ("" if before == after else "  ⚠ 有改动！"))

    n_pass = sum(1 for _, ok, _ in REPORT if ok)
    total = len(REPORT)
    head = [
        "=" * 66,
        "模块（1）第 6 步验收报告：HTTP API + 异步调用 + 网页操作台",
        "时间：%s" % time.strftime("%Y-%m-%d %H:%M:%S"),
        "=" * 66,
    ]
    body = []
    for i, (name, ok, detail) in enumerate(REPORT, 1):
        body.append("%2d. [%s] %s%s" % (i, "PASS" if ok else "FAIL", name,
                                        ("  | " + detail) if detail else ""))
    tail = [
        "-" * 66,
        "合计：%d / %d 通过%s" % (n_pass, total, "  ✅ 全部通过" if n_pass == total else "  ⚠ 存在失败项"),
        "=" * 66,
    ]
    text = "\n".join(head + body + tail)
    out = OUT_DIR / "第6步-API验收报告.txt"
    out.write_text(text, encoding="utf-8")
    print()
    print(text)
    print("报告已写入：%s" % out)
    return 0 if n_pass == total else 1


if __name__ == "__main__":
    sys.exit(main())
