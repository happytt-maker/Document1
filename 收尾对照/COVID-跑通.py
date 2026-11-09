# -*- coding: utf-8 -*-
"""
COVID-19 真实数据 · 带环境隔离的分步跑批
用法：
    python COVID-跑通.py read        只跑 ①②（读文档 + 读模板）
    python COVID-跑通.py extract     跑 ③（抽取）
    python COVID-跑通.py fill        跑 ④⑤（写回 + 验收）
    python COVID-跑通.py all         全跑
可以加 --tag 名字 区分产物目录（默认 base）
"""
import os, sys, time, shutil, subprocess, json
from pathlib import Path

HERE = Path(r"C:\Users\覃敬涛\Desktop\研一项目")
WS = HERE / "测试数据" / "真实数据" / "07-COVID-19"
SRC = WS / "源文档"
TPL = WS / "模板"
RUN = WS / "结果"
PY = sys.executable
TPL_NAME = "COVID-19 模板.xlsx"


def setup():
    SRC.mkdir(parents=True, exist_ok=True)
    TPL.mkdir(parents=True, exist_ok=True)
    RUN.mkdir(parents=True, exist_ok=True)
    # 源文档：只放这两份，避免 16 份文档混池
    pairs = [
        (HERE / "测试数据" / "源文档" / "xlsx" / "COVID-19全球数据集（节选）.xlsx", SRC / "COVID-19全球数据集（节选）.xlsx"),
        (HERE / "测试数据" / "源文档" / "docx" / "中国COVID-19新冠疫情情况.docx", SRC / "中国COVID-19新冠疫情情况.docx"),
    ]
    for a, b in pairs:
        if a.exists() and (not b.exists() or a.stat().st_size != b.stat().st_size):
            shutil.copy2(a, b)
    t = HERE / "测试数据" / "模板上传" / TPL_NAME
    if t.exists() and not (TPL / TPL_NAME).exists():
        shutil.copy2(t, TPL / TPL_NAME)


def run(cwd, cmd, extra=None, label=""):
    env = os.environ.copy()
    env["PYTHONIOENCODING"] = "utf-8"
    env["PYTHONUTF8"] = "1"
    if extra:
        env.update({k: str(v) for k, v in extra.items()})
    t0 = time.time()
    p = subprocess.run([PY] + cmd, cwd=str(cwd), capture_output=True, env=env)
    sec = time.time() - t0
    out = p.stdout.decode("utf-8", "replace") + p.stderr.decode("utf-8", "replace")
    print(f"\n===== {label} ｜ {sec:.1f}s ｜ rc={p.returncode} =====")
    print(out[-4000:])
    return sec, p.returncode, out


def main():
    step = sys.argv[1] if len(sys.argv) > 1 else "all"
    setup()

    if step in ("read", "all"):
        run(HERE / "doc_reader", ["main.py", str(SRC), "-o", str(RUN / "output.json")], label="① 读源文档")
        # 结构写到工作区自己的文件里 —— 别覆盖 template_reader/模板表结构.json（那是 5 张正式模板的）
        run(HERE / "template_reader", ["main.py", str(TPL), "--out", str(RUN / "模板表结构.json")],
            label="② 读模板")

    if step in ("extract", "all"):
        run(HERE / "matcher", ["run_all.py", "--no-eval"], {
            "A23_TEMPLATE_NAME": TPL_NAME,
            "A23_MATCHER_DIR": str(RUN),
            "A23_OUTPUT_JSON": str(RUN / "output.json"),
            "A23_STRUCTURE_JSON": str(RUN / "模板表结构.json"),
            "A23_USER_REQ": "智能填表，将两个文件的内容中日期一列从2020/7/1到2020/8/31的数据填入模板中.",
        }, label="③ 抽取")

    if step in ("fill", "all"):
        run(HERE / "writer", ["fill.py", TPL_NAME], {
            "A23_TEMPLATE_DIR": str(TPL),
            "A23_OUT_DIR": str(RUN),
            "A23_MATCHER_DIR": str(RUN),
            "A23_STRUCTURE_JSON": str(RUN / "模板表结构.json"),
        }, label="④ 写回")
        run(HERE / "writer", ["verify_fill.py", TPL_NAME], {
            "A23_TEMPLATE_DIR": str(TPL),
            "A23_OUT_DIR": str(RUN),
            "A23_MATCHER_DIR": str(RUN),
            "A23_STRUCTURE_JSON": str(RUN / "模板表结构.json"),
        }, label="⑤ 验收")


if __name__ == "__main__":
    main()
