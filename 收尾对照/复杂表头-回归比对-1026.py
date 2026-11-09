# -*- coding: utf-8 -*-
"""
复杂表头-回归比对-1026.py
=========================
改造解析器之后，必须回答一个问题：**原来那 10 张规整模板，结论有没有被带偏？**

做法（不靠记忆、不靠肉眼看）：
    1. 从 git HEAD 取出改造前的 reader.py，单独加载成 old_reader
    2. 用它和新的 reader 各读一遍 测试数据/模板表格 与 测试数据/模板表格-word
    3. 逐张模板、逐张表比对：表头行、列名、待填区

比对时忽略新增的 `table` 字段（新结构里每个待填段都带"第几张表"，
规整模板恒为 1，不影响语义）。

结论只有两种：SAME（完全一致）/ DIFF（列出哪一项从什么变成了什么）。
改造的标准是「只有不规整的表会变，规整的表必须一模一样」。

用法：
    python 复杂表头-回归比对-1026.py
"""

import importlib.util
import json
import subprocess
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
BASE = HERE.parent
TMP = HERE / "_regress"
DIRS = ["测试数据/模板表格", "测试数据/模板表格-word"]


def load_old_reader():
    """从 git HEAD 取改造前的 reader.py 并加载（模块名区分开，互不干扰）。"""
    TMP.mkdir(parents=True, exist_ok=True)
    p = subprocess.run(["git", "show", "HEAD:template_reader/reader.py"],
                       cwd=str(BASE), capture_output=True, text=True,
                       encoding="utf-8", errors="replace")
    if p.returncode != 0:
        raise SystemExit("取不到改造前的 reader.py：%s" % (p.stderr or "")[:200])
    src = TMP / "old_reader.py"
    src.write_text(p.stdout, encoding="utf-8")

    spec = importlib.util.spec_from_file_location("old_reader", src)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def flat(doc):
    """把一张模板的结构拍平成便于比对的列表。"""
    out = []
    for sh in doc.get("sheets", []):
        fr = sh.get("fill_region") or {}
        out.append({
            "name": sh.get("name"),
            "header_row": sh.get("header_row"),
            "titles": [c.get("title") for c in sh.get("columns", [])],
            # 只比三要素，忽略表结构新增的 table 字段
            "fill": [fr.get("start_row"), fr.get("end_row"), fr.get("row_count")],
        })
    return out


def main():
    print("=" * 72)
    print("回归比对：改造前后，原来的 10 张规整模板结论是否一致")
    print("=" * 72)

    old_mod = load_old_reader()
    sys.path.insert(0, str(BASE / "template_reader"))
    from reader import read_folder as new_read

    old_docs = []
    new_docs = []
    for d in DIRS:
        old_docs += old_mod.read_folder(BASE / d)
        new_docs += new_read(BASE / d)

    oldmap = {x["file"]: x for x in old_docs}
    lines, n_same, n_diff = [], 0, 0

    for doc in new_docs:
        fname = doc["file"]
        if fname not in oldmap:
            lines.append("NEW  %s（改造前没有这张）" % fname)
            n_diff += 1
            continue
        a, b = flat(oldmap[fname]), flat(doc)
        if a == b:
            lines.append("SAME %-34s %d 张表" % (fname, len(b)))
            n_same += 1
            continue

        n_diff += 1
        lines.append("DIFF %s" % fname)
        for i, (x, y) in enumerate(zip(a, b), start=1):
            for key in ("header_row", "titles", "fill"):
                if x.get(key) != y.get(key):
                    lines.append("     第%d张表 %-10s 旧=%s" % (i, key, x.get(key)))
                    lines.append("     第%d张表 %-10s 新=%s" % (i, key, y.get(key)))

    text = "\n".join(lines)
    print(text)
    print()
    print("一致 %d 张 / 有差异 %d 张" % (n_same, n_diff))

    (HERE / "复杂表头-回归比对-1026.txt").write_text(
        "回归比对（改造前 vs 改造后）\n一致 %d 张 / 差异 %d 张\n\n%s"
        % (n_same, n_diff, text), encoding="utf-8")
    return 0 if n_diff == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
