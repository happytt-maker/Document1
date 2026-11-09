# -*- coding: utf-8 -*-
"""
复杂表头-演示-1026.py
=====================
用「像真的」业务数据填两张复杂模板，产出一眼能看懂的效果文件：

    模板6-多层表头设备采购台账-已填写.xlsx   ← 两级表头，列名靠父名区分
    模板8-跨页原材料入库台账-已填写.xlsx     ← 重点：第 10~12 条落在第 2 页

为什么单独做演示（而不是直接看验收产物）：
    验收用的是编码值（V4_0 这种），它能证明"位置对"，但看不出"业务上对不对"。
    演示用真实含义的数据（数控车床 / 入库数量 / 供应商），
    打开文件就能判断第二页有没有真的接上。

说明：这里的"抽取结果"是脚本手写的，模拟 matcher 产出（格式完全相同），
      因此走的是和真实流程一模一样的 fill.py 写回路径，不是另写一套。
      演示用的抽取结果放在临时目录，不污染 matcher/ 里的真实结果。

用法：
    python 复杂表头-演示-1026.py
"""

import json
import os
import subprocess
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
BASE = HERE.parent
TMP = HERE / "_复杂表头演示"
MATCHER = TMP / "matcher"
STRUCT = TMP / "结构.json"
TPL_DIR = BASE / "测试数据" / "模板-复杂表头"
OUT_DIR = BASE / "writer" / "输出"

# ----------------------------------------------------------------------
# 模板6：多层表头（两级，靠父名区分重复子名）
# ----------------------------------------------------------------------
TPL6 = "模板6-多层表头设备采购台账.xlsx"
DATA6 = {
    "设备名称": ["数控车床", "立式加工中心", "激光切割机"],
    "规格型号": ["CK6140", "VMC-850B", "HG-3015"],
    "上半年采购-数量": ["4", "2", "1"],
    "上半年采购-金额（元）": ["372000", "536000", "288000"],
    "下半年采购-数量": ["6", "3", "2"],
    "下半年采购-金额（元）": ["558000", "804000", "576000"],
    "供应商名称": ["江苏精工机床有限公司", "沈阳中捷机床厂", "武汉华工激光"],
    "联系方式": ["0512-6688-1201", "024-8899-3302", "027-8755-6108"],
}

# ----------------------------------------------------------------------
# 模板8：跨页（9 + 9 个位置，给 12 条 → 第 10~12 条必须出现在第 2 页）
# ----------------------------------------------------------------------
TPL8 = "模板8-跨页原材料入库台账.xlsx"
DATA8 = {
    "物料编码": ["RM-1001", "RM-1002", "RM-1003", "RM-1004", "RM-1005", "RM-1006",
                 "RM-1007", "RM-1008", "RM-1009", "RM-1010", "RM-1011", "RM-1012"],
    "物料名称": ["冷轧钢板", "不锈钢管", "铝合金型材", "轴承钢", "铜排", "橡胶密封圈",
                 "尼龙扎带", "环氧树脂", "紧固螺栓", "润滑油", "砂纸", "焊丝"],
    "规格型号": ["SPCC 1.2×1250", "304 Φ32×2", "6063 T5 40×40", "GCr15 Φ40",
                 "T2 40×4", "NBR Φ50", "4×200mm", "E-44", "M8×30", "46# 200L",
                 "P400", "ER50-6 Φ1.2"],
    "入库数量": ["120", "86", "240", "60", "150", "800",
                 "1500", "25", "3000", "12", "200", "35"],
    "单位": ["张", "根", "支", "根", "条", "个",
             "包", "桶", "盒", "桶", "包", "盘"],
    "入库日期": ["2026-09-02", "2026-09-03", "2026-09-03", "2026-09-05", "2026-09-06", "2026-09-08",
                 "2026-09-10", "2026-09-11", "2026-09-12", "2026-09-15", "2026-09-16", "2026-09-18"],
    "供应商": ["宝钢股份", "久立特材", "南山铝业", "兴澄特钢", "金龙铜业", "中鼎密封",
               "华通塑胶", "巴陵石化", "晋亿实业", "长城润滑", "金鹰砂纸", "大西洋焊材"],
}


def write_extracts(data, n):
    if MATCHER.exists():
        for f in MATCHER.glob("*.json"):
            f.unlink()
    MATCHER.mkdir(parents=True, exist_ok=True)
    for col, values in data.items():
        items = [{"value": v, "file": "演示源数据.xlsx",
                  "loc_text": "第 %d 行" % (i + 2)}
                 for i, v in enumerate(values)]
        (MATCHER / ("抽取结果-%s.json" % col)).write_text(
            json.dumps({"column": col, "items": items},
                       ensure_ascii=False, indent=2), encoding="utf-8")


def main():
    sys.path.insert(0, str(BASE / "template_reader"))
    from reader import read_folder

    docs = read_folder(TPL_DIR)
    STRUCT.parent.mkdir(parents=True, exist_ok=True)
    STRUCT.write_text(json.dumps(docs, ensure_ascii=False, indent=2), encoding="utf-8")

    env = dict(os.environ)
    env["A23_STRUCTURE_JSON"] = str(STRUCT)
    env["A23_MATCHER_DIR"] = str(MATCHER)
    env["A23_TEMPLATE_DIR"] = str(TPL_DIR)
    env["PYTHONIOENCODING"] = "utf-8"

    for tpl, data in ((TPL6, DATA6), (TPL8, DATA8)):
        n = max(len(v) for v in data.values())
        write_extracts(data, n)
        out = OUT_DIR / (Path(tpl).stem + "-已填写.xlsx")
        if out.exists():
            out.unlink()

        print("\n" + "=" * 60)
        print("填写 %s（%d 条记录）" % (tpl, n))
        print("=" * 60)
        p = subprocess.run([sys.executable, str(BASE / "writer" / "fill.py"), tpl],
                           capture_output=True, text=True, encoding="utf-8",
                           errors="replace", env=env, timeout=300)
        # 只挑关键几行打印，不刷屏
        for line in (p.stdout or "").splitlines():
            if any(k in line for k in ("待填区", "写入 ", "统计", "还有", "已生成", "序号")):
                print("  " + line.strip())
        if p.returncode != 0:
            print("  失败：%s" % (p.stderr or "")[-400:])
            return 1
        print("  → %s" % out)

    print("\n提示：打开这两份文件，重点看模板8 的第 2 页（第 15~17 行）"
          "是否接着填了第 10~12 条。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
