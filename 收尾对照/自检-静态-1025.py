# -*- coding: utf-8 -*-
"""静态自检 v2：按项目真实结构检查"""
import os
import py_compile
import sys
import tempfile

ROOT = r"C:\Users\覃敬涛\Desktop\研一项目"
AGENT = os.path.join(ROOT, "agent")
out = []
ok = 0
bad = 0


def chk(name, cond, detail=""):
    global ok, bad
    if cond:
        ok += 1
        out.append("  [OK]   %s %s" % (name, detail))
    else:
        bad += 1
        out.append("  [FAIL] %s %s" % (name, detail))


out.append("=== 1. 三个模块与总控 ===")
paths = {
    "pipeline.py（总控）": os.path.join(ROOT, "pipeline.py"),
    "doc_reader/（文档读取包）": os.path.join(ROOT, "doc_reader"),
    "doc_reader/readers/docx_reader.py": os.path.join(ROOT, "doc_reader", "readers", "docx_reader.py"),
    "doc_reader/readers/xlsx_reader.py": os.path.join(ROOT, "doc_reader", "readers", "xlsx_reader.py"),
    "template_reader/（模板读取包）": os.path.join(ROOT, "template_reader"),
    "template_reader/reader.py": os.path.join(ROOT, "template_reader", "reader.py"),
    "matcher/retrieve.py": os.path.join(ROOT, "matcher", "retrieve.py"),
    "matcher/extract.py": os.path.join(ROOT, "matcher", "extract.py"),
    "matcher/run_all.py": os.path.join(ROOT, "matcher", "run_all.py"),
    "writer/fill.py": os.path.join(ROOT, "writer", "fill.py"),
    "writer/verify_fill.py": os.path.join(ROOT, "writer", "verify_fill.py"),
}
for k, p in paths.items():
    chk(k, os.path.exists(p), "" if os.path.exists(p) else "  缺失")

out.append("")
out.append("=== 2. agent 与网页 ===")
for k, p in {
    "api_server.py": os.path.join(AGENT, "api_server.py"),
    "asr.py": os.path.join(AGENT, "asr.py"),
    "digital_tts.py": os.path.join(AGENT, "digital_tts.py"),
    "support.py": os.path.join(AGENT, "support.py"),
    "feedback.py": os.path.join(AGENT, "feedback.py"),
    "auth.py": os.path.join(AGENT, "auth.py"),
    "db.py": os.path.join(AGENT, "db.py"),
    "parse_intent.py": os.path.join(AGENT, "parse_intent.py"),
    "web/index.html": os.path.join(AGENT, "web", "index.html"),
    "web/login.html": os.path.join(AGENT, "web", "login.html"),
    "db_config.json": os.path.join(AGENT, "db_config.json"),
    "models/faster-whisper-small/model.bin": os.path.join(AGENT, "models", "faster-whisper-small", "model.bin"),
}.items():
    chk(k, os.path.exists(p))

out.append("")
out.append("=== 3. 测试数据 ===")
src = os.path.join(ROOT, "测试数据", "源文档")
n_src = 0
detail = []
if os.path.isdir(src):
    for sub in sorted(os.listdir(src)):
        sp = os.path.join(src, sub)
        if os.path.isdir(sp):
            fs = [f for f in os.listdir(sp) if not f.startswith(".")]
            n_src += len(fs)
            detail.append("%s=%d" % (sub, len(fs)))
chk("源文档总数 >= 16", n_src >= 16, "（%d 份：%s）" % (n_src, ", ".join(detail)))

tpl = os.path.join(ROOT, "测试数据", "模板表格")
tpls = [f for f in os.listdir(tpl) if f.endswith((".xlsx", ".xls"))] if os.path.isdir(tpl) else []
chk("模板表格 = 5 张", len(tpls) == 5, "（%d 张）" % len(tpls))

tplw = os.path.join(ROOT, "测试数据", "模板表格-word")
tplws = [f for f in os.listdir(tplw) if f.endswith(".docx")] if os.path.isdir(tplw) else []
chk("Word 模板 = 5 张", len(tplws) == 5, "（%d 张）" % len(tplws))

out.append("")
out.append("=== 4. Python 语法（全部 .py，正确写法） ===")
py_files = []
for dirpath, dirnames, filenames in os.walk(ROOT):
    dirnames[:] = [d for d in dirnames if d not in (".git", "__pycache__", "models", ".wheels")]
    for f in filenames:
        if f.endswith(".py"):
            py_files.append(os.path.join(dirpath, f))
errs = []
tmpdir = tempfile.mkdtemp()
for i, p in enumerate(py_files):
    try:
        py_compile.compile(p, cfile=os.path.join(tmpdir, "t%d.pyc" % i), doraise=True)
    except Exception as e:
        errs.append("%s -> %s" % (os.path.relpath(p, ROOT), str(e).replace("\n", " ")[:180]))
chk("全部 .py 语法正确", not errs, "（检查 %d 个文件）" % len(py_files))
for e in errs[:12]:
    out.append("       " + e)

out.append("")
out.append("=== 5. 模块 import（按真实运行方式） ===")
sys.path.insert(0, AGENT)
sys.path.insert(0, os.path.join(ROOT, "matcher"))
sys.path.insert(0, ROOT)

for name in ("doc_reader", "template_reader"):
    try:
        __import__(name)
        chk("import " + name, True)
    except Exception as e:
        chk("import " + name, False, "%s: %s" % (type(e).__name__, str(e)[:180]))

for name in ("retrieve", "extract", "run_all"):
    try:
        __import__(name)
        chk("import matcher." + name, True)
    except Exception as e:
        chk("import matcher." + name, False, "%s: %s" % (type(e).__name__, str(e)[:180]))

for name in ("parse_intent", "asr", "digital_tts", "support", "feedback", "auth", "db"):
    try:
        m = __import__(name)
        info = ""
        if name == "asr":
            st = getattr(m, "status", None)
            if callable(st):
                try:
                    s = st()
                    info = "（status=%s）" % str(s)[:100]
                except Exception as e2:
                    info = "（status() 报错：%s）" % str(e2)[:80]
        chk("import agent." + name, True, info)
    except Exception as e:
        chk("import agent." + name, False, "%s: %s" % (type(e).__name__, str(e)[:180]))

out.append("")
out.append("=== 6. 写出依赖库 ===")
for mod, label in [("docx", "python-docx"), ("openpyxl", "openpyxl"), ("fastapi", "FastAPI"),
                   ("uvicorn", "uvicorn"), ("pymysql", "PyMySQL"), ("faster_whisper", "faster-whisper"),
                   ("edge_tts", "edge-tts"), ("PIL", "Pillow")]:
    try:
        m = __import__(mod)
        chk(label, True, getattr(m, "__version__", ""))
    except Exception as e:
        chk(label, False, str(e)[:100])

out.append("")
out.append("总计：%d 项通过，%d 项失败" % (ok, bad))
open(r"C:\Users\覃敬涛\WorkBuddy\2026-09-24-10-48-03\_static2.txt", "w", encoding="utf-8").write("\n".join(out))
print("PASS=%d FAIL=%d" % (ok, bad))
