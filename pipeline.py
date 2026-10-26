# -*- coding: utf-8 -*-
"""
pipeline.py —— 总控流水线：一条命令跑完「读 → 模板 → 抽 → 写 → 验」

为什么有这个脚本：
  以前跑完整流程要开 4 个文件夹、按 4 次 F5、来回跳。
  比赛当天时间紧，而且流程应该是"文档进去、表格出来"一气呵成。
  这个脚本就是比赛当天真实流程的预演：
      新文档 + 一张模板  →  一条命令  →  一份"已填写"的表 + 验收报告

五个阶段（每个阶段还是一个独立脚本，总控只负责按顺序调用、收集结果、出一页日志）：
  ① 读源文档     doc_reader/run_all.py       → output.json（全部内容 + 位置）
  ② 理解模板     template_reader/main.py     → 模板表结构.json（表头/列/待填区）
  ③ 抽取 + 评测  matcher/run_all.py          → 抽取结果 × 7 列 + 评测报告
  ④ 写回表格     writer/fill.py              → writer/输出/xxx-已填写.xlsx
  ⑤ 写回验收     writer/verify_fill.py       → writer/输出/验收报告-xxx.txt

用法（VSCode 打开「研一项目」文件夹后，F5 下拉里选）：
    python pipeline.py                        # 全流程（默认模板4）
    python pipeline.py --template 模板3-会议任务跟踪表.xlsx
    python pipeline.py --skip-read            # 跳过①（output.json 没变时提速）
    python pipeline.py --overwrite            # ④ 发现待填区有内容时，覆盖写入

出错怎么办：
  - 哪一步失败，就在哪一步停下（不做"带病往下跑"）
  - ④ 的规矩不变：待填区有内容 → 默认停下不写，两条路（清模板 / --overwrite）由人拍板

输出：
  屏幕 + 运行日志.txt（一页看完：每阶段状态 / 耗时 / 关键数字 / 总耗时红线对照）
"""
import json
import os
import re
import subprocess
import sys
import time
from pathlib import Path

# 让中文在终端正常显示（失败了也不影响；日志以文件为准）
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

HERE = Path(__file__).resolve().parent        # = 研一项目/
PY = sys.executable                            # 当前解释器（与 F5 配置一致）
LOG_PATH = HERE / "运行日志.txt"

DEFAULT_TEMPLATE = "模板4-客诉处理台账.xlsx"

# 模板目录：默认 测试数据/模板表格；可用环境变量 A23_TEMPLATE_DIR 覆盖
# （泛化测试 / Word 模板目录 / 比赛换数据集时用；fill.py、verify_fill.py 同款）
TEMPLATE_DIR = os.environ.get("A23_TEMPLATE_DIR",
                              str(HERE / "测试数据" / "模板表格"))


# ==================================================================
# 工具
# ==================================================================
def run_stage(cwd: Path, cmd: list, extra_env: dict = None):
    """跑一个子脚本。强制子进程按 UTF-8 输出，返回 (秒数, 返回码, 输出文本)。

    为什么设 PYTHONIOENCODING：
      子脚本打印的中文在 Windows 终端默认走 GBK，很容易乱码；
      统一强制 UTF-8，总控这边再按 UTF-8 解码，两头就对齐了。

    extra_env：给某一步单独加的环境变量（如 ③ 抽列要用的
      A23_TEMPLATE_NAME=当前模板名 —— 让 matcher 抽「这张模板」的列）。
    """
    t0 = time.time()
    env = os.environ.copy()
    env["PYTHONIOENCODING"] = "utf-8"
    if extra_env:
        env.update(extra_env)
    proc = subprocess.run(
        [str(PY)] + cmd,
        cwd=str(cwd),
        capture_output=True,
        env=env,
    )
    sec = time.time() - t0
    out = proc.stdout.decode("utf-8", errors="replace")
    err = proc.stderr.decode("utf-8", errors="replace")
    text = out
    if err.strip():
        text += "\n--- stderr ---\n" + err
    return sec, proc.returncode, text


def tail(text: str, n: int = 28) -> str:
    """取输出文本的最后 n 行（失败/注意时附在日志里，方便定位）。"""
    lines = [ln for ln in text.splitlines() if ln.strip()]
    return "\n".join("      | " + ln for ln in lines[-n:])


def read_json(path: Path):
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return None


# ==================================================================
# 各阶段「关键数字」提取（读产物文件；抓不到就少写一行，不报错）
# ==================================================================
def brief_read():
    """① 之后：从 doc_reader/output.json 数文件/块/失败。"""
    data = read_json(HERE / "doc_reader" / "output.json")
    if not data:
        return "（没读到 output.json）"
    docs = data.get("documents", [])
    n_fail = sum(1 for d in docs if "error" in d)
    n_blocks = sum(d.get("block_count", 0) for d in docs if "error" not in d)
    return "%d 份文件 / %d 个内容块 / 读取失败 %d" % (len(docs), n_blocks, n_fail)


def brief_template(template_name):
    """② 之后：从 模板表结构.json 报模板张数 + 目标模板结构。"""
    data = read_json(HERE / "template_reader" / "模板表结构.json")
    if not data:
        return "（没读到 模板表结构.json）"
    s = "共 %d 张模板" % len(data)
    for item in data:
        if item.get("file") == template_name:
            sh = item["sheets"][0]
            fr = sh.get("fill_region") or {}
            s += " ｜ 目标：表头第 %s 行 / %d 列 / 待填 %s 行" % (
                sh.get("header_row"), len(sh.get("columns", [])),
                fr.get("row_count", "?"))
    return s


def brief_extract():
    """③ 之后：从 运行摘要.txt + 最新评测报告 抓数字。"""
    bits = []

    p = HERE / "matcher" / "运行摘要.txt"
    if p.exists():
        t = p.read_text(encoding="utf-8")
        m = re.search(r"墙钟总耗时.*?：([\d.]+) 秒", t)
        if m:
            bits.append("抽取墙钟 %s 秒" % m.group(1))
        m = re.search(r"成功：(\d+)/(\d+) 列", t)
        if m:
            bits.append("%s/%s 列成功" % (m.group(1), m.group(2)))

    reps = sorted((HERE / "matcher").glob("评测报告-*.txt"),
                  key=lambda f: f.stat().st_mtime)
    if reps:
        t = reps[-1].read_text(encoding="utf-8")
        m = re.search(r"值 · 严格.*?：(\d+)/(\d+)", t)
        if m:
            bits.append("值(严格) %s/%s" % (m.group(1), m.group(2)))
        m = re.search(r"值 · 宽松.*?：(\d+)/(\d+)", t)
        if m:
            bits.append("值(宽松) %s/%s" % (m.group(1), m.group(2)))
        m = re.search(r"出处 .*?：(\d+)/(\d+)", t)
        if m:
            bits.append("出处 %s/%s" % (m.group(1), m.group(2)))

    return " ｜ ".join(bits) if bits else "（没抓到统计行，详见 matcher 下的报告）"


def brief_fill(template_name):
    """④ 之后：从填表报告抓 写入/覆盖 数。"""
    stem = Path(template_name).stem
    tag = "-word" if str(template_name).lower().endswith(".docx") else ""
    p = HERE / "writer" / "输出" / ("填表报告-%s%s.txt" % (stem, tag))
    if not p.exists():
        return "（没读到填表报告）"
    t = p.read_text(encoding="utf-8")
    bits = []
    m = re.search(r"【二、写入明细】共 (\d+) 格", t)
    if m:
        bits.append("写入 %s 格" % m.group(1))
    m = re.search(r"覆盖 (\d+) 格", t)
    if m:
        bits.append("覆盖旧值 %s 格" % m.group(1))
    m = re.search(r"模式：(.+)", t)
    if m:
        bits.append(m.group(1).strip())
    return " ｜ ".join(bits) if bits else "（详见填表报告）"


def brief_verify(template_name):
    """⑤ 之后：从验收报告抓 值/格式/结论。"""
    stem = Path(template_name).stem
    tag = "-word" if str(template_name).lower().endswith(".docx") else ""
    p = HERE / "writer" / "输出" / ("验收报告-%s%s.txt" % (stem, tag))
    if not p.exists():
        return "（没读到验收报告）"
    t = p.read_text(encoding="utf-8")
    bits = []
    m = re.search(r"与抽取结果一致：(\d+) / (\d+) 格", t)
    if m:
        bits.append("值一致 %s/%s 格" % (m.group(1), m.group(2)))
    m = re.search(r"单元格样式差异：(\d+) 处", t)
    if m:
        bits.append("格式差异 %s 处" % m.group(1))
    if "全部通过" in t:
        bits.append("结论：可以交卷")
    return " ｜ ".join(bits) if bits else "（详见验收报告）"


# ==================================================================
# 主流程
# ==================================================================
def main():
    args = sys.argv[1:]
    skip_read = "--skip-read" in args
    overwrite = "--overwrite" in args
    template_name = DEFAULT_TEMPLATE
    if "--template" in args:
        i = args.index("--template")
        if i + 1 < len(args):
            template_name = args[i + 1]

    t_start = time.time()
    stem = Path(template_name).stem
    log = []

    def w(s=""):
        log.append(str(s))
        print(s)

    def write_log():
        LOG_PATH.write_text("\n".join(log) + "\n", encoding="utf-8")

    def stop_and_report(reason):
        """中途停下：写日志 + 收尾提示。"""
        w("")
        w("-" * 64)
        w("流程在中途停下：%s" % reason)
        w("修复后重跑即可（本日志每跑一次覆盖一份）。")
        w("=" * 64)
        write_log()

    w("=" * 64)
    w("A-23 总控流水线 · 运行日志")
    w("开始：%s ｜ 目标模板：%s" % (time.strftime("%Y-%m-%d %H:%M:%S"), template_name))
    w("命令：python pipeline.py %s" % " ".join(args))
    w("=" * 64)
    w()

    # ---------------- ① 读源文档 ----------------
    if skip_read:
        w("【① 读源文档】已跳过（--skip-read，沿用现有 output.json）")
        w()
    else:
        sec, code, out = run_stage(HERE / "doc_reader", ["run_all.py"])
        if code != 0:
            w("【① 读源文档】失败（返回码 %d）" % code)
            w(tail(out))
            w()
            return stop_and_report("读文档失败")
        w("【① 读源文档】完成 ｜ %.2f 秒" % sec)
        w("    " + brief_read())
        w()

    # ---------------- ② 理解模板 ----------------
    # 把「模板目录」传给读取器：读的是同一批模板（默认目录 / A23_TEMPLATE_DIR）
    sec, code, out = run_stage(HERE / "template_reader", ["main.py", TEMPLATE_DIR])
    if code != 0:
        w("【② 理解模板】失败（返回码 %d）" % code)
        w(tail(out))
        w()
        return stop_and_report("模板理解失败")
    w("【② 理解模板】完成 ｜ %.2f 秒" % sec)
    w("    " + brief_template(template_name))
    w()

    # ---------------- ③ 抽取 + 评测 ----------------
    # 传 A23_TEMPLATE_NAME：让 matcher 抽「当前这张模板」的列（而非默认模板4）
    sec, code, out = run_stage(HERE / "matcher", ["run_all.py"],
                               {"A23_TEMPLATE_NAME": template_name})
    if code != 0:
        w("【③ 抽取 + 评测】失败（返回码 %d）" % code)
        w(tail(out))
        w()
        return stop_and_report("抽取失败")
    w("【③ 抽取 + 评测】完成 ｜ %.2f 秒" % sec)
    w("    " + brief_extract())
    w()

    # ---------------- ④ 写回表格 ----------------
    cmd = ["fill.py", template_name] + (["--overwrite"] if overwrite else [])
    sec, code, out = run_stage(HERE / "writer", cmd)
    if code == 2:
        w("【④ 写回表格】注意：待填区已有内容，脚本已停下、没有写入")
        w(tail(out, 20))
        w()
        w("    两条路：A. 打开模板清掉旧数据后重跑  ｜  B. 加 --overwrite 覆盖（旧值会记进报告）")
        w()
        return stop_and_report("写回被『先查后写』拦下（这是设计好的保护，不是坏）")
    if code != 0:
        w("【④ 写回表格】失败（返回码 %d）" % code)
        w(tail(out))
        w()
        return stop_and_report("写回失败")
    w("【④ 写回表格】完成 ｜ %.2f 秒" % sec)
    w("    " + brief_fill(template_name))
    w()

    # ---------------- ⑤ 写回验收 ----------------
    sec, code, out = run_stage(HERE / "writer", ["verify_fill.py", template_name])
    if code == 2:
        w("【⑤ 写回验收】注意：验收发现问题，请打开报告查看")
        w("    " + brief_verify(template_name))
        w()
        w("    报告：writer/输出/验收报告-%s.txt" % stem)
        w()
    elif code != 0:
        w("【⑤ 写回验收】失败（返回码 %d）" % code)
        w(tail(out))
        w()
        return stop_and_report("验收脚本执行失败")
    else:
        w("【⑤ 写回验收】完成 ｜ %.2f 秒" % sec)
        w("    " + brief_verify(template_name))
        w()

    # ---------------- 总结 ----------------
    total = time.time() - t_start
    w("-" * 64)
    w("总耗时：%.2f 秒（赛题红线：≤90 秒/份文档；具体计时口径待主办方确认，此处按整条流水线对照）" % total)
    w("产物：")
    if str(template_name).lower().endswith(".docx"):
        w("  交卷文件 → writer/输出/%s-已填写.docx" % stem)
        w("  验收报告 → writer/输出/验收报告-%s-word.txt" % stem)
    else:
        w("  交卷文件 → writer/输出/%s-已填写.xlsx" % stem)
        w("  验收报告 → writer/输出/验收报告-%s.txt" % stem)
    w("  评测报告 → matcher/评测报告-*.txt")
    w("=" * 64)
    write_log()


if __name__ == "__main__":
    main()
