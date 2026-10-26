# -*- coding: utf-8 -*-
"""压测前侦察：目录状态 / 大池 / MEMORY.md 大小"""
import json, os
from pathlib import Path

BASE = Path(__file__).parent
MEM = Path(r"C:\Users\覃敬涛\WorkBuddy\2026-09-24-10-48-03\.workbuddy\memory\MEMORY.md")

lines = []
lines.append("=== 压测目录 ===")
for p in sorted(BASE.iterdir()):
    if p.is_dir():
        n = sum(1 for _ in p.rglob("*") if _.is_file())
        lines.append(f"[DIR ] {p.name}/  ({n} files)")
    else:
        lines.append(f"[FILE] {p.name}  {p.stat().st_size} bytes")

src = BASE / "源文档"
if src.exists():
    files = sorted(src.iterdir())
    lines.append(f"\n=== 源文档 {len(files)} 份 ===")
    for f in files:
        lines.append(f"{f.name}  {f.stat().st_size/1024:.1f} KB")

res = BASE / "结果"
lines.append("\n=== 结果目录 ===")
if res.exists():
    for p in sorted(res.rglob("*")):
        if p.is_file():
            lines.append(f"{p.relative_to(res)}  {p.stat().st_size} bytes")
else:
    lines.append("(不存在)")

lines.append("\n=== MEMORY.md ===")
if MEM.exists():
    t = MEM.read_text(encoding="utf-8")
    lines.append(f"{len(t)} chars / {len(t.encode('utf-8'))} bytes / {len(t.splitlines())} lines")

out = BASE / "_recon.txt"
out.write_text("\n".join(lines), encoding="utf-8")
print("OK")
