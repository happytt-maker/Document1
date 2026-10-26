# -*- coding: utf-8 -*-
"""冒烟 v2：验证 FastDocx（XML 直出）生成的 docx 能被读取器读回 + 测生成速度。"""
import importlib
import sys
import time
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
G = importlib.import_module("生成压测数据")

t0 = time.time()
G.build_devices_docx(300)
t1 = time.time()
print("build_devices_docx(300) → %.2f 秒" % (t1 - t0))
G.build_complaints_docx(300)
G.build_meetings_docx(300)
G.build_employees_docx(300)
G.build_report_docx(300)
print("5 份 docx（300 规模）总用时 %.2f 秒" % (time.time() - t0))

# 读取器读回验证
ROOT = HERE.parent.parent
sys.path.insert(0, str(ROOT / "doc_reader"))
readers = importlib.import_module("readers")

for name in ["01-设备采购档案.docx", "02-客户投诉处理档案.docx", "05-年度经营分析报告.docx"]:
    p = G.DOCX_DIR / name
    doc = readers.read_one(p)
    n_para = doc.get("paragraph_count")
    n_tbl = doc.get("table_count")
    n_blk = doc.get("block_count")
    print("%s → %d 段 / %d 表 / %d 块 ｜ %.1f KB" % (name, n_para, n_tbl, n_blk, p.stat().st_size / 1024))

# 抽查内容
doc = readers.read_one(G.DOCX_DIR / "01-设备采购档案.docx")
texts = [b["text"] for b in doc["blocks"]]
print("首 6 块：")
for t in texts[:6]:
    print("   ", t[:50])
found = [t for t in texts if "超精密五轴磨床" in t]
print("金丝雀搜索 [超精密五轴磨床]：", "找到" if found else "!! 没找到", found[:1])

print("SMOKE2 OK")
