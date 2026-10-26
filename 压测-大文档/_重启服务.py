# -*- coding: utf-8 -*-
"""重启 8000 操作台服务：杀旧进程 → detached 启动新版 → 验证新路由"""
import subprocess, sys, time, os, json, urllib.request
from pathlib import Path

sys.stdout.reconfigure(errors="replace")

AGENT = Path(r"C:\Users\覃敬涛\Desktop\研一项目\agent")
PY = r"C:\Users\覃敬涛\.workbuddy\binaries\python\envs\default\Scripts\python.exe"
PORT = 8000

# ---- 1. 找占用 8000 的 PID ----
out = subprocess.run(["netstat", "-ano"], capture_output=True)
txt = out.stdout.decode("gbk", errors="replace")
pids = set()
for line in txt.splitlines():
    if (":%d " % PORT) in line and "LISTENING" in line:
        pids.add(line.split()[-1])
print("[1] 旧进程 PID:", pids or "（无）")

# ---- 2. 杀掉 ----
for pid in pids:
    r = subprocess.run(["taskkill", "/F", "/PID", pid], capture_output=True)
    print("    taskkill", pid, "→", r.returncode)

if pids:
    time.sleep(1)

# ---- 3. detached 启动新版 ----
log_path = AGENT / "输出" / "api_server-运行.log"
log = open(log_path, "ab")
DETACHED = 0x00000008 | 0x00000200  # DETACHED_PROCESS | CREATE_NEW_PROCESS_GROUP
p = subprocess.Popen(
    [PY, "api_server.py"],
    cwd=str(AGENT),
    stdout=log, stderr=log, stdin=subprocess.DEVNULL,
    creationflags=DETACHED,
)
print("[2] 已拉起新进程 PID:", p.pid, "（日志 → 输出/api_server-运行.log）")

# ---- 4. 等就绪 + 验证 ----
ok = False
for i in range(12):
    time.sleep(1)
    try:
        r = urllib.request.urlopen("http://127.0.0.1:%d/openapi.json" % PORT, timeout=3)
        spec = json.loads(r.read().decode("utf-8"))
        paths = spec.get("paths", {})
        need = ["/api/recognize", "/api/fill", "/api/fill/templates", "/api/fill/download"]
        missing = [k for k in need if k not in paths]
        print("[3] 服务就绪（%.1fs）｜ 路由 %d 个 ｜ 新路由缺失: %s" % (i + 1, len(paths), missing or "无 ✓"))
        ok = not missing
        break
    except Exception:
        continue

if not ok:
    print("[3] 服务尚未就绪（12s 超时）——稍后手动检查")

# 顺带确认旧路由还在
try:
    r = urllib.request.urlopen("http://127.0.0.1:%d/openapi.json" % PORT, timeout=3)
    paths = json.loads(r.read().decode("utf-8")).get("paths", {})
    for k in ("/api/command", "/api/job/{job_id}", "/api/docs", "/api/store/summary"):
        print("    旧路由", k, "→", "在 ✓" if k in paths else "缺 ✗")
except Exception as e:
    print("    旧路由复查失败:", e)
