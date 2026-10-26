# -*- coding: utf-8 -*-
r"""
template_reader/verify_templates.py
-----------------------------------
验收脚本：拿「人工答案」核对程序输出。

人工答案来自 研一项目/模板表人工答案.md
（我用眼睛看过 5 张表，手写的标准答案）

核对三件事：
    1. 表头行号对不对
    2. 列数对不对
    3. 待填行数对不对
"""

import json
from pathlib import Path

from reader import read_folder

TPL_DIR = Path(r"C:\Users\覃敬涛\Desktop\研一项目\测试数据\模板表格")
OUT = Path(__file__).parent / "模板表验收报告.txt"
JSON_OUT = Path(__file__).parent / "模板表结构.json"

# -----------------------------------------------------------------
# 人工答案（从 模板表人工答案.md 抄过来的）
# 格式：文件名 -> (表头行号, 列数, 待填行数)
# -----------------------------------------------------------------
EXPECTED = {
    "模板1-设备采购汇总表.xlsx": (4, 6, 8),
    "模板2-员工信息登记表.xlsx": (3, 8, 12),
    "模板3-会议任务跟踪表.xlsx": (4, 6, 9),
    "模板4-客诉处理台账.xlsx": (3, 7, 6),
    "模板5-原材料库存汇总表.xlsx": (3, 8, 12),
}

lines = []


def w(s=""):
    lines.append(str(s))
    print(s)


def main():
    results = read_folder(TPL_DIR)

    w("=" * 70)
    w("模板表结构验收")
    w("=" * 70)
    w()

    total = 0
    passed = 0

    for doc in results:
        if "error" in doc:
            w("[失败] {0} —— {1}".format(doc["file"], doc["error"]))
            w()
            continue

        expect = EXPECTED.get(doc["file"])
        if expect is None:
            w("[跳过] {0}（人工答案里没有这张表）".format(doc["file"]))
            w()
            continue

        exp_header, exp_cols, exp_fill = expect

        # 一张表可能有多个 sheet，这里只核对第一个
        sheet = doc["sheets"][0]

        got_header = sheet["header_row"]
        got_cols = len([c for c in sheet["columns"] if c["title"] is not None])
        got_fill = sheet["fill_region"]["row_count"] if sheet["fill_region"] else 0

        w("#" * 70)
        w("文件：{0}".format(doc["file"]))
        w("工作表：{0}   （{1} 行 × {2} 列）".format(
            sheet["name"], sheet["max_row"], sheet["max_column"]))
        w()

        # 逐项核对
        checks = [
            ("表头行号", exp_header, got_header),
            ("列数", exp_cols, got_cols),
            ("待填行数", exp_fill, got_fill),
        ]

        all_ok = True
        for name, exp, got in checks:
            total += 1
            ok = (exp == got)
            if ok:
                passed += 1
            else:
                all_ok = False
            w("  {0:<10} 人工答案 {1:>3}   程序输出 {2:>3}   {3}".format(
                name, exp, got, "OK" if ok else "★不符"))

        w()
        w("  列标题：{0}".format(
            " | ".join(c["title"] for c in sheet["columns"] if c["title"])))
        w("  待填区域：第 {0} 行 到 第 {1} 行".format(
            sheet["fill_region"]["start_row"],
            sheet["fill_region"]["end_row"]) if sheet["fill_region"] else "  待填区域：未识别")
        w("  表头特征检查：{0}".format(sheet["header_note"]))
        w()
        w("  → 这张表：{0}".format("全部通过" if all_ok else "有不符合项"))
        w()

    w("=" * 70)
    w("合计核对 {0} 项，通过 {1} 项".format(total, passed))
    w()

    if total > 0 and passed == total:
        w("结论：5 张模板表的结构全部识别正确。")
        w("      表头行号、列数、待填行数都与人工答案一致。")
    else:
        w("结论：有 {0} 项不符，需要检查判断逻辑。".format(total - passed))

    OUT.write_text("\n".join(lines), encoding="utf-8")
    JSON_OUT.write_text(
        json.dumps(results, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    print("\n已写出：" + str(OUT))
    print("已写出：" + str(JSON_OUT))


if __name__ == "__main__":
    main()
