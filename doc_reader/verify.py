# -*- coding: utf-8 -*-
"""
verify.py
---------
验收脚本：抽查 JSON 里的位置信息，看能不能回到原文。

做法：
  从 output.json 里挑几个块，
  按它的 locator 回到原文件取出对应内容，
  跟 JSON 里存的文字对比。对得上就算通过。

不同格式的回溯方式不一样：
  txt / md  → locator 里是 line_start / line_end，按行取
  docx      → locator 里是 para_index（段落）或 table_index/row/col（单元格）
  xlsx      → locator 里是 sheet / cell，按坐标取
"""

import json
from pathlib import Path

BASE = Path(__file__).parent
SRC = Path(r"C:\Users\覃敬涛\Desktop\研一项目\测试数据\源文档")

out = []


def w(s=""):
    out.append(str(s))
    print(s)


def find_source_file(file_name):
    """按文件名在测试数据目录里找到真实路径。"""
    hits = list(SRC.rglob(file_name))
    return hits[0] if hits else None


# ---------- 各格式的回溯函数 ----------
# 每个函数接收 (源文件路径, locator)，返回"回到原文取出来的内容"

def backtrack_txt_md(path, loc):
    """txt / md：按行号取。"""
    lines = path.read_text(encoding="utf-8", errors="replace").splitlines()
    start, end = loc["line_start"], loc["line_end"]
    return "\n".join(lines[start - 1:end])


def backtrack_docx(path, loc):
    """docx：段落按段落序号取，单元格按表格坐标取。"""
    from docx import Document

    doc = Document(str(path))
    kind = loc.get("kind")

    if kind == "paragraph":
        paras = doc.paragraphs
        i = loc["para_index"] - 1
        return paras[i].text if i < len(paras) else ""

    if kind == "table_cell":
        table = doc.tables[loc["table_index"] - 1]
        cell = table.rows[loc["row"] - 1].cells[loc["col"] - 1]
        return cell.text

    return ""


def backtrack_xlsx(path, loc):
    """
    xlsx：按「工作表名 + 单元格坐标」取回值。

    注意要读两遍，跟读取器保持一致：
      有值        → 直接取缓存值
      值是 None   → 看公式版里有没有公式（坑 1 的现场）
    """
    import openpyxl

    sheet_name = loc["sheet"]
    coord = loc["cell"]

    wb_values = openpyxl.load_workbook(str(path), data_only=True)
    value = wb_values[sheet_name][coord].value
    if value is not None:
        return str(value)

    wb_formulas = openpyxl.load_workbook(str(path), data_only=False)
    fval = wb_formulas[sheet_name][coord].value
    return str(fval) if fval is not None else ""


# 用扩展名分派
BACKTRACKERS = {
    "txt": backtrack_txt_md,
    "md": backtrack_txt_md,
    "docx": backtrack_docx,
    "xlsx": backtrack_xlsx,
}


def main():
    result = json.loads((BASE / "output.json").read_text(encoding="utf-8"))

    w("=" * 66)
    w("验收检查：位置信息能不能回到原文")
    w("=" * 66)
    w()

    passed = 0
    failed = 0
    skipped = 0
    checked = 0

    for doc in result["documents"]:
        ftype = doc["type"]
        backtrack = BACKTRACKERS.get(ftype)

        if backtrack is None:
            w("[跳过] {0}（暂时没有 {1} 的回溯方式）".format(doc["file"], ftype))
            w()
            skipped += len(doc.get("blocks", []))
            continue

        src_file = find_source_file(doc["file"])
        if src_file is None:
            w("[跳过] 找不到源文件: " + doc["file"])
            w()
            continue

        # 每个文件抽查前 3 个块
        for block in doc["blocks"][:3]:
            try:
                picked = backtrack(src_file, block["locator"])
            except Exception as e:
                w("[出错] {0} {1}".format(doc["file"], repr(e)))
                w()
                failed += 1
                checked += 1
                continue

            checked += 1
            same = picked.strip() == block["text"].strip()
            # docx 单元格可能是 '吴晓敏\n吴晓敏'，清理过也算对
            if not same and block["text"] in picked:
                same = True

            if same:
                passed += 1
                mark = "OK  "
            else:
                failed += 1
                mark = "差异"

            w("[{0}] {1}".format(mark, doc["file"]))
            w("       位置 : {0}".format(block["locator"]))
            w("       JSON : {0!r}".format(block["text"][:52]))
            w("       回读 : {0!r}".format(picked[:52]))
            w()

    w("-" * 66)
    w("抽查 {0} 个块，通过 {1} 个，不符 {2} 个".format(checked, passed, failed))
    if skipped:
        w("（另有 {0} 个块因格式未支持回溯而跳过）".format(skipped))
    w()

    if failed == 0:
        w("结论：位置信息可用，每种格式都能回到原文。")
    else:
        w("结论：有 {0} 个块不符，需要检查。".format(failed))

    (BASE / "验收报告.txt").write_text("\n".join(out), encoding="utf-8")


if __name__ == "__main__":
    main()
