# -*- coding: utf-8 -*-
"""生成后验证：① 16 份文件规格达标 ② 金丝雀值确实写进了对应文档（zip 直读 XML 搜索，快）。"""
import json
import re
import zipfile
from pathlib import Path

DATA = Path(__file__).resolve().parent
SRC = DATA / "源文档"
OUT = DATA / "_check_gen_out.txt"
SPEC = {"docx": 500, "xlsx": 500, "md": 15, "txt": 15}

lines = []


def strip_tags(xml):
    return re.sub(r"<[^>]+>", "", xml)


def read_text_of(path):
    """把 docx/xlsx/txt/md 的可读文本抽出来（不解析，直接读 XML）。

    注意：openpyxl 输出的中文是 XML 数字实体（&#nnnnn;），要解码后再搜。
    """
    if path.suffix in (".txt", ".md"):
        return path.read_text(encoding="utf-8", errors="replace")
    text = ""
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        cand = [n for n in names if n.endswith(".xml") and
                ("document" in n or "sharedStrings" in n or "sheet" in n)]
        for n in cand:
            try:
                text += strip_tags(z.read(n).decode("utf-8", "ignore"))
            except Exception:
                pass
    text = re.sub(r"&#(\d+);", lambda m: chr(int(m.group(1))), text)
    return text


# ---- [1] 规格 ----
lines.append("=== [1] 文件规格 ===")
ok_cnt = 0
tot = 0
for kind in ("docx", "xlsx", "md", "txt"):
    for f in sorted((SRC / kind).glob("*")):
        if not f.is_file():
            continue
        tot += 1
        kb = f.stat().st_size / 1024
        std = SPEC[kind]
        flag = "OK" if kb >= std else "!! 未达标"
        ok_cnt += 1 if kb >= std else 0
        lines.append("  %-5s %-34s %8.1f KB  %s" % (kind, f.name, kb, flag))
lines.append("  合计 %d 份，达标 %d 份" % (tot, ok_cnt))

# ---- [2] 金丝雀存在性 ----
lines.append("")
lines.append("=== [2] 金丝雀值存在性（在文档文本里能不能搜到） ===")
checks = [
    ("docx/01-设备采购档案.docx", ["超精密五轴磨床", "XM-9000A", "常州精工精密机械有限公司"]),
    ("xlsx/01-设备采购明细总表.xlsx", ["真空离子镀膜机", "VCM-700T"]),
    ("txt/01-设备验收记录汇总.txt", ["全自动影像测量仪", "VMS-650"]),
    ("docx/04-员工管理与培训档案.docx", ["易湛卢", "EMP-90017", "戚砚舟"]),
    ("xlsx/02-员工花名册全量.xlsx", ["沈寄澜", "EMP-90098"]),
    ("docx/03-项目会议纪要汇编.docx", ["完成大模型推理加速方案评审", "戚玥", "沈砚清"]),
    ("xlsx/03-项目任务分解与进度表.xlsx", ["编制赛题答辩材料初稿", "顾今朝"]),
    ("docx/02-客户投诉处理档案.docx", ["TS-2026-7001", "TS-2026-7088", "常州龙城精锻有限公司"]),
    ("docx/05-年度经营分析报告.docx", ["高纯钼粉", "MAT-7A001"]),
    ("xlsx/05-原材料库存盘点明细.xlsx", ["单晶硅棒", "MAT-7B003", "高纯石墨坩埚"]),
]
for rel, keys in checks:
    p = SRC / rel
    if not p.exists():
        lines.append("  [缺失] %s" % rel)
        continue
    text = read_text_of(p)
    bad = [k for k in keys if k not in text]
    lines.append("  %s %s" % ("OK " if not bad else "!! ", rel))
    if bad:
        lines.append("      没搜到：%s" % "、".join(bad))

OUT.write_text("\n".join(lines), encoding="utf-8")
print("WROTE", OUT)
