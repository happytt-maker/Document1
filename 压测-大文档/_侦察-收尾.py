# -*- coding: utf-8 -*-
"""收尾侦察：MEMORY 字符数 / 收尾对照目录 / 8000 端口服务状态 / 今日日志"""
import sys, socket, urllib.request, json
from pathlib import Path

sys.stdout.reconfigure(errors="replace")

MEM = Path(r"C:\Users\覃敬涛\WorkBuddy\2026-09-24-10-48-03\.workbuddy\memory")
m = MEM / "MEMORY.md"
t = m.read_text(encoding="utf-8")
print("== MEMORY.md ==")
print("字符数:", len(t))
print("预算: 4000 → 需压缩到", max(0, len(t) - 3800), "字符以内（目标 ~3300）")

print()
print("== memory 目录文件 ==")
for f in sorted(MEM.iterdir()):
    print(" ", f.name, f.stat().st_size, "字节")

print()
print("== 收尾对照 目录 ==")
d = Path(r"C:\Users\覃敬涛\Desktop\研一项目\收尾对照")
for f in sorted(d.iterdir()):
    print(" ", f.name, f.stat().st_size, "字节")

print()
print("== 8000 端口服务 ==")
sock = socket.socket()
sock.settimeout(2)
try:
    sock.connect(("127.0.0.1", 8000))
    print("端口 8000: 在监听")
    try:
        r = urllib.request.urlopen("http://127.0.0.1:8000/api/health", timeout=3)
        print("  /api/health →", r.status, r.read().decode("utf-8", "replace")[:200])
    except Exception as e:
        print("  /api/health 失败:", e)
    # 检查是否新版（是否有 /api/recognize 路由）
    try:
        r = urllib.request.urlopen("http://127.0.0.1:8000/openapi.json", timeout=3)
        spec = json.loads(r.read().decode("utf-8"))
        paths = sorted(spec.get("paths", {}).keys())
        print("  openapi 路由数:", len(paths))
        for k in ("/api/recognize", "/api/fill", "/api/fill/templates", "/api/fill/download"):
            print("   ", k, "→", "有" if k in spec.get("paths", {}) else "缺")
    except Exception as e:
        print("  openapi 读取失败:", e)
except OSError as e:
    print("端口 8000: 未监听（", e, "）")
finally:
    sock.close()
