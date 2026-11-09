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
import os
import sys
from pathlib import Path

from reader import read_folder

TPL_DIR = Path(r"C:\Users\覃敬涛\Desktop\研一项目\测试数据\模板表格")
# 输出位置：默认项目内的 模板表结构.json；
# 用 --out 或环境变量 A23_STRUCTURE_JSON 指到别处（验收脚本不污染真实产物）
JSON_OUT = Path(os.environ.get(
    "A23_STRUCTURE_JSON", str(Path(__file__).parent / "模板表结构.json")))


def _split_args(argv):
    """把 --out / --only 之类的开关与普通参数分开（普通参数仍只有「模板文件夹」一个）。

    --only <文件名>：只读这一张模板。
      为什么要它（2026-11-07）：真实数据常常把「源文档」和「模板」放在同一个文件夹里
      （用户就是这么放的），于是模板读取器会把那份 2 万行的源表也当模板解析一遍 ——
      光 openpyxl 加载就 4.7 秒，而这一趟流水线根本不需要它。
      指定 --only 之后，② 从 7.1 秒降到 0.3 秒。
    """
    plain, out, only = [], None, None
    i = 0
    while i < len(argv):
        a = argv[i]
        if a == "--out" and i + 1 < len(argv):
            out = Path(argv[i + 1])
            i += 2
            continue
        if a.startswith("--out="):
            out = Path(a.split("=", 1)[1])
            i += 1
            continue
        if a == "--only" and i + 1 < len(argv):
            only = argv[i + 1]
            i += 2
            continue
        if a.startswith("--only="):
            only = a.split("=", 1)[1]
            i += 1
            continue
        if not a.startswith("--"):
            plain.append(a)
        i += 1
    return plain, out, only


def main():
    global JSON_OUT
    # 用法：python main.py [模板文件夹] [--out 结构json路径] [--only 模板文件名]
    #   不带参数 = 原来的默认目录（测试数据/模板表格）
    #   带参数   = 读指定目录（泛化测试 / 比赛换数据集时用）
    plain, out, only = _split_args(sys.argv[1:])
    if out:
        JSON_OUT = out
    folder = Path(plain[0]) if plain else TPL_DIR
    results = read_folder(folder, only=only)
    if only:
        print("（只读指定模板：%s —— 同目录下的源文档不解析）" % only)
        print()

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
            rows = sheet.get("header_rows") or (
                [sheet["header_row"]] if sheet.get("header_row") else [])
            print("  工作表：{0}   （{1} 行 × {2} 列）".format(
                sheet["name"], sheet["max_row"], sheet["max_column"]))
            if rows:
                if len(rows) == 1:
                    print("    表头在第 {0} 行（单层）".format(rows[0]))
                else:
                    print("    表头在第 {0}~{1} 行（{2} 层）".format(
                        rows[0], rows[-1], len(rows)))
            else:
                print("    表头：没找到")
            print()

            print("    这一列要填什么：")
            for col in sheet["columns"]:
                title = col["title"] if col["title"] else "（无标题）"
                extra = []
                if col.get("slash"):
                    extra.append("斜线格，可拆为 " + " / ".join(col["slash_parts"]))
                if col.get("aliases"):
                    extra.append("别名 " + "、".join(col["aliases"][:4]))
                suffix = ("    ← " + "；".join(extra)) if extra else ""
                print("      第 {0} 列  {1}{2}".format(col["col"], title, suffix))
            print()

            labels = sheet.get("label_cols") or []
            if labels:
                print("    行标签列（不填数据）：第 {0} 列".format(
                    "、".join(str(c) for c in labels)))

            segs = sheet.get("segments") or []
            if segs:
                desc = "＋".join(
                    "第 {0} 行 到 第 {1} 行（{2} 行）".format(
                        s["start_row"], s["end_row"], s["row_count"]) for s in segs)
                total = sum(s["row_count"] for s in segs)
                print("    要填几行：{0}，共 {1} 行".format(desc, total))
                if len(segs) > 1:
                    print("      （跨页续表：{0} 段连起来填，第 i 条记录依次往下排）".format(
                        len(segs)))
            else:
                print("    要填几行：没找到空白区")
            print()

        print()

    # 存一份 JSON 给后面的程序用
    JSON_OUT.parent.mkdir(parents=True, exist_ok=True)
    JSON_OUT.write_text(
        json.dumps(results, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    print("已写出：" + str(JSON_OUT))


if __name__ == "__main__":
    main()
