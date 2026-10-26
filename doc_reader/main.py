# -*- coding: utf-8 -*-
"""
main.py
-------
文档读取器 · 主程序

用法（在 doc_reader 目录下运行）：
    python main.py                         读默认的测试数据文件夹
    python main.py "某个文件夹路径"          读指定文件夹
    python main.py "路径" -o 结果.json       指定输出文件名

做的事：
    1. 扫描文件夹里所有 txt / md 文件
    2. 逐个读出来，每个内容块都带上位置信息
    3. 汇总成一个 JSON 文件
"""

import argparse
import json
import sys
from pathlib import Path

from readers import read_one, is_supported, SUPPORTED

# 默认读这个文件夹（就是我给你生成的那批测试数据）1
DEFAULT_DIR = Path(r"C:\Users\覃敬涛\Desktop\研一项目\测试数据\源文档")


def collect_files(folder):
    """
    扫描文件夹，挑出所有能读的文件。

    用 rglob("*") 会递归进子文件夹，这样 docx/ xlsx/ txt/ md/ 都能扫到。
    结果排个序，让每次运行顺序一致（方便对比）。
    """
    folder = Path(folder)
    files = []

    for p in sorted(folder.rglob("*")):
        if p.is_file() and is_supported(p):
            files.append(p)

    return files


def read_all(folder):
    """把文件夹里所有能读的文件读出来。"""
    files = collect_files(folder)

    documents = []
    skipped = []
    failed = []

    for path in files:
        try:
            doc = read_one(path)
            if doc is None:
                skipped.append(str(path))
            else:
                documents.append(doc)
        except Exception as e:
            # 单个文件读失败不能拖垮整个流程，记下来继续
            failed.append({"file": str(path), "error": repr(e)})

    return {
        "source_dir": str(folder),
        "supported_types": SUPPORTED,
        "file_count": len(files),
        "ok_count": len(documents),
        "failed_count": len(failed),
        "skipped_count": len(skipped),
        "documents": documents,
        "failed": failed,
    }


def print_summary(result):
    """在终端里打一份人能看懂的摘要。"""
    print()
    print("=" * 66)
    print("文档读取器 · 运行结果")
    print("=" * 66)
    print("源文件夹 : " + result["source_dir"])
    print("支持格式 : " + ", ".join(result["supported_types"]))
    print("扫到文件 : {0} 个".format(result["file_count"]))
    print("成功读取 : {0} 个".format(result["ok_count"]))
    if result["failed_count"]:
        print("读取失败 : {0} 个".format(result["failed_count"]))
    print("-" * 66)

    print()
    print("逐个文件明细：")
    print("  {0:<34} {1:>6} {2:>8} {3:>8}".format("文件名", "行数", "块数", "格式"))
    print("  " + "-" * 62)

    for doc in result["documents"]:
        line_count = doc.get("line_count", "-")
        block_count = doc.get("block_count", len(doc.get("blocks", [])))
        print("  {0:<34} {1:>6} {2:>8} {3:>8}".format(
            doc["file"][:33], str(line_count), str(block_count), doc["type"]
        ))

    total_blocks = sum(len(d.get("blocks", [])) for d in result["documents"])
    print("  " + "-" * 62)
    print("  合计 {0} 个文件，{1} 个内容块".format(len(result["documents"]), total_blocks))

    if result["failed"]:
        print()
        print("读取失败的文件：")
        for item in result["failed"]:
            print("  {0}\n     {1}".format(item["file"], item["error"]))

    print()


def main():
    parser = argparse.ArgumentParser(description="把一文件夹的 txt/md 读成 JSON")
    parser.add_argument("folder", nargs="?", default=str(DEFAULT_DIR),
                        help="要扫描的文件夹（默认：测试数据/源文档）")
    parser.add_argument("-o", "--output", default="output.json",
                        help="输出 JSON 文件名（默认 output.json）")

    args = parser.parse_args()

    folder = Path(args.folder)
    if not folder.exists():
        print("[错误] 文件夹不存在：" + str(folder))
        sys.exit(1)

    result = read_all(folder)

    # 输出 JSON：ensure_ascii=False 才能让中文原样显示
    out_path = Path(args.output)
    out_path.write_text(
        json.dumps(result, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    print_summary(result)
    print("JSON 已写入: " + str(out_path.resolve()))
    print()


if __name__ == "__main__":
    main()
