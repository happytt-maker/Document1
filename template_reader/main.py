# -*- coding: utf-8 -*-
"""
main.py
-------
模板表结构读取器 —— 主程序。

跑一遍，把模板表的「填表说明书」打印出来。

    python main.py                        # 默认目录：测试数据/模板表格
    python main.py "某个模板文件夹"         # 读指定文件夹（泛化测试用）

输出：
    终端      —— 直接给人看的可读版
    模板表结构.json —— 给后面程序用的机器版
"""

import json
import sys
from pathlib import Path

from reader import read_folder

TPL_DIR = Path(r"C:\Users\覃敬涛\Desktop\研一项目\测试数据\模板表格")
JSON_OUT = Path(__file__).parent / "模板表结构.json"


def main():
    # 用法：python main.py [模板文件夹]
    #   不带参数 = 原来的默认目录（测试数据/模板表格）
    #   带参数   = 读指定目录（泛化测试 / 比赛换数据集时用）
    folder = Path(sys.argv[1]) if len(sys.argv) > 1 else TPL_DIR
    results = read_folder(folder)

    print("=" * 70)
    print("模板表结构 —— 填表说明书")
    print("=" * 70)
    print()

    for doc in results:
        if "error" in doc:
            print("[失败] {0} —— {1}".format(doc["file"], doc["error"]))
            print()
            continue

        print("文件：{0}".format(doc["file"]))
        print("-" * 70)

        for sheet in doc["sheets"]:
            print("  工作表：{0}   （{1} 行 × {2} 列）".format(
                sheet["name"], sheet["max_row"], sheet["max_column"]))
            print("    表头在第 {0} 行".format(sheet["header_row"]))
            print()

            print("    这一列要填什么：")
            for col in sheet["columns"]:
                title = col["title"] if col["title"] else "（无标题）"
                print("      第 {0} 列  {1}".format(col["col"], title))
            print()

            region = sheet["fill_region"]
            if region:
                print("    要填几行：第 {0} 行 到 第 {1} 行，共 {2} 行".format(
                    region["start_row"], region["end_row"], region["row_count"]))
            else:
                print("    要填几行：没找到空白区")
            print()

        print()

    # 存一份 JSON 给后面的程序用
    JSON_OUT.write_text(
        json.dumps(results, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    print("已写出：" + str(JSON_OUT))


if __name__ == "__main__":
    main()
