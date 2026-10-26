# -*- coding: utf-8 -*-
"""第 4 步验收脚本：4 条提取指令逐条真跑（完整链路）+ 自动核对 + 源文件零改动验证。

跑法（在项目根目录）：
  python agent/_验收-提取版.py
"""
import hashlib
import json
import sys
import time
from pathlib import Path

BASE = Path(r"C:\Users\覃敬涛\Desktop\研一项目")
sys.path.insert(0, str(BASE / "agent"))
sys.stdout.reconfigure(errors="replace")

from edit import run_edit  # noqa: E402

DOCS = BASE / "测试数据" / "源文档"
OUT = BASE / "agent" / "输出"

CASES = [
    {"id": 13, "text": "把员工信息表里所有员工的姓名和工号整理出来",
     "doc": "员工信息表.xlsx",
     "checks": [
         ("姓名", ["王振国", "李慧敏", "张新民", "赵雪松", "陈思彤", "刘宏伟", "孙雅静", "周文彬"]),
         ("工号", ["E1001", "E1002", "E1003", "E1004", "E1005", "E1006", "E1007", "E1008"]),
     ]},
    {"id": 14, "text": "从库存盘点表里把每种物料的名称和期初库存列出来",
     "doc": "库存盘点表.xlsx",
     "checks": [
         ("物料名称", ["冷轧钢板", "铝合金型材", "轴承", "伺服电机", "密封圈",
                       "紧固螺栓", "液压油", "控制柜", "电缆", "防锈漆"]),
         ("期初库存", ["120", "45", "2400", "36", "5600", "12800", "800", "8", "3200", "150"]),
     ]},
    {"id": 15, "text": "从设备采购合同里提取甲乙双方名称和合同总金额",
     "doc": "设备采购合同.docx",
     "checks": [
         ("甲方名称", ["华景智造科技有限公司"]),
         ("乙方名称", ["泰隆机电设备有限公司"]),
         ("合同总金额", ["1,800,000", "壹佰捌拾万"]),
     ]},
    {"id": 16, "text": "把项目推进会会议纪要里的待办事项列出来",
     "doc": "项目推进会会议纪要.docx",
     "checks": [
         ("待办事项", ["完成数据接入模块开发", "编写接口文档", "算法调优实验",
                       "整理测试用例集", "服务器资源申请与环境部署"]),
     ]},
]

# 附加演示（不计分）：口语别名「工资」→「基本工资」
DEMO = {"text": "把员工信息表里所有员工的姓名和工资列出来",
        "doc": "员工信息表.xlsx",
        "checks": [("基本工资", ["12800", "14500", "18600", "9200", "13400", "10600", "8800", "19800"])]}


def sha(p: Path) -> str:
    return hashlib.md5(p.read_bytes()).hexdigest()


def find_doc(name: str):
    for p in DOCS.rglob("*"):
        if p.is_file() and p.name == name:
            return p
    return None


def check_json(doc_name: str, checks: list):
    """读结构化结果，逐字段核对。返回 [(字段, 是否过, 说明), ...]。"""
    jp = OUT / (Path(doc_name).stem + "_提取结果.json")
    if not jp.exists():
        return [(c[0], False, "结果文件不存在") for c in checks]
    data = json.loads(jp.read_text(encoding="utf-8"))
    out = []
    for fname, expect in checks:
        target = None
        for f in data.get("fields", []):
            if f["field"] == fname:
                target = f
                break
        if target is None:
            out.append((fname, False, "字段缺失"))
            continue
        got = [str(it["value"]).strip() for it in target["items"]]
        missing = [e for e in expect if not any(e == v or e in v for v in got)]
        if missing:
            out.append((fname, False, "缺: %s" % missing))
        else:
            out.append((fname, True, "%d/%d 全中 ｜ 实际 %d 条" % (len(expect), len(expect), len(got))))
    return out


def main():
    print("=" * 62)
    print("第 4 步验收 · 内容提取接入（4 条指令完整链路真跑）")
    print("=" * 62)
    print()

    hashes = {c["doc"]: sha(find_doc(c["doc"])) for c in CASES}

    all_ok = True
    n_pass = n_total = 0
    t_all = time.time()
    costs = []

    for c in CASES:
        t0 = time.time()
        ok, report = run_edit(c["text"], DOCS)
        cost = time.time() - t0
        costs.append(cost)
        (OUT / ("第%d条-提取报告.txt" % c["id"])).write_text(report, encoding="utf-8")

        rows = check_json(c["doc"], c["checks"])
        print("---- 第 %d 条 ｜ 全链路 %.1f 秒 ----" % (c["id"], cost))
        print("  指令：%s" % c["text"])
        print("  执行：%s" % ("成功" if ok else "失败"))
        for fname, okf, msg in rows:
            n_total += 1
            n_pass += int(okf)
            all_ok = all_ok and okf
            print("  [%s] %s ｜ %s" % ("✓" if okf else "✗", fname, msg))
        print()

    print("---- 源文件零改动核对（提取是只读操作）----")
    for c in CASES:
        same = sha(find_doc(c["doc"])) == hashes[c["doc"]]
        all_ok = all_ok and same
        print("  [%s] %s" % ("✓" if same else "✗", c["doc"]))

    print()
    print("---- 附加演示：口语别名「工资」→「基本工资」（不计分）----")
    ok, report = run_edit(DEMO["text"], DOCS)
    (OUT / "第13条b-别名演示报告.txt").write_text(report, encoding="utf-8")
    rows = check_json(DEMO["doc"], DEMO["checks"])
    for fname, okf, msg in rows:
        print("  [%s] %s ｜ %s" % ("✓" if okf else "✗", fname, msg))

    print()
    print("=" * 62)
    print("结果：字段核对 %d/%d 通过 ｜ 4 条全链路耗时 %s 秒"
          % (n_pass, n_total, " + ".join("%.1f" % x for x in costs)))
    print("=" * 62)
    return 0 if all_ok else 1


if __name__ == "__main__":
    sys.exit(main())
