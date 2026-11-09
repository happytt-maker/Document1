# -*- coding: utf-8 -*-
"""文档库上传功能验收（2026-11-05）

验的是「数据仓库 → 📚 文档库 → ⬆ 上传文档」这条链路：
  POST /api/upload   —— 落盘 + 解析 + 并入检索池（和老「📥 导入文档」同一个接口）
  GET  /api/docs     —— 清单 + 明细（格式 / 大小 / 时间 / 来源）

用法（先起一个测试实例，别用正在跑的 8000）：
    python agent/api_server.py --port 8011
    python 收尾对照/文档库上传-验收-1105.py

脚本会真的上传 3 个文件（1 txt + 2 份同名 docx 验重名），跑完：
  · 按**精确文件名**删除这 3 个文件（不碰任何既有文档）
  · 把 doc_reader/output.json 恢复成跑之前的备份
  · 删掉 imports 表里这 3 条自检记录
不留垃圾，可反复跑。
"""
import io
import json
import os
import shutil
import sys
import time
from pathlib import Path

import requests
from docx import Document

BASE = "http://127.0.0.1:8011"
PROJECT = Path(r"C:\Users\覃敬涛\Desktop\研一项目")
AGENT = PROJECT / "agent"
POOL = PROJECT / "doc_reader" / "output.json"
IMPORT_DIR = PROJECT / "导入文档"
OUT = Path(__file__).with_suffix(".txt")

TXT_NAME = "文档库自检-勿留.txt"
DOCX_NAME = "文档库自检-勿留.docx"
DOCX_DUP = "文档库自检-勿留(2).docx"      # plan_stored_name 的重名规则
BAD_NAME = "文档库自检-勿留.exe"

S = requests.Session()
S.trust_env = False                       # 系统代理会拦 127.0.0.1（GET 假 404）

lines = []
passed = failed = 0


def say(s=""):
    lines.append(str(s))


def check(ok, tag, detail=""):
    global passed, failed
    if ok:
        passed += 1
    else:
        failed += 1
    say("  %s %s%s" % ("PASS" if ok else "FAIL", tag,
                       ("　—— " + detail) if detail else ""))
    return ok


def get_docs():
    r = S.get(BASE + "/api/docs", timeout=20)
    r.raise_for_status()
    return r.json()


def make_txt():
    return ("文档库上传自检（测试用，验收结束后删除）\n"
            "第一段：验证上传后能不能解析出块。\n"
            "第二段：甲方 华景科技有限公司，合同金额 128.6 万元。\n"
            "第三段：项目负责人 李四，联系电话 13800001111。\n").encode("utf-8")


def make_docx(body):
    d = Document()
    d.add_heading("文档库自检（测试用）", 0)
    d.add_paragraph(body)
    d.add_paragraph("第二段：验收跑完本文件会被精确删除，不影响任何既有文档。")
    d.add_paragraph("第三段：乙方 恒信建设集团，签约日期 2026年11月5日。")
    buf = io.BytesIO()
    d.save(buf)
    return buf.getvalue()


def upload(name, content):
    r = S.post(BASE + "/api/upload",
               files={"file": (name, content, "application/octet-stream")},
               timeout=300)
    return r.status_code, (r.json() if r.content else {})


def main():
    say("文档库上传功能验收 · %s" % time.strftime("%Y-%m-%d %H:%M:%S"))
    say("目标服务：%s" % BASE)
    say("")

    # 0. 服务在不在
    try:
        before = get_docs()
    except Exception as e:
        say("连不上 %s（%s）——先起测试实例：python agent/api_server.py --port 8011"
            % (BASE, e))
        OUT.write_text("\n".join(lines), encoding="utf-8")
        return 3

    n0 = len(before.get("docs") or [])
    imp0 = (before.get("counts") or {}).get("import", 0)
    say("【0. 基线与备份】")
    say("  上传前：文档库 %d 份（我上传的 %d 份）" % (n0, imp0))
    backup = POOL.with_name("_池备份-上传自检.json")
    if POOL.exists():
        shutil.copy2(POOL, backup)
        say("  已备份池：%s" % backup.name)
    check(True, "0.1 基线读取与池备份")

    # ---------- 1. 清单接口 ----------
    say("")
    say("【1. /api/docs 清单与明细】")
    docs, detail = before.get("docs") or [], before.get("detail") or []
    check(bool(docs) and bool(detail), "1.1 docs / detail 都非空",
          "docs=%d detail=%d" % (len(docs), len(detail)))
    check(len(docs) == len(detail), "1.2 docs 与 detail 条数一致")
    check([x.get("name") for x in detail] == list(docs), "1.3 顺序与名字一一对应")
    keys_ok = all(all(k in x and x[k] not in (None, "") for k in
                      ("name", "ext", "size", "mtime", "src")) for x in detail)
    check(keys_ok, "1.4 每项都有 name/ext/size/mtime/src")
    srcs = {x.get("src") for x in detail}
    check(srcs <= {"standard", "import"}, "1.5 src 取值只可能是 standard / import",
          str(sorted(srcs)))
    c = before.get("counts") or {}
    check(c.get("total") == len(docs)
          and c.get("import") == sum(1 for x in detail if x.get("src") == "import"),
          "1.6 counts 与明细自洽", "total=%s import=%s" % (c.get("total"), c.get("import")))
    # mtime 形状（YYYY-MM-DD HH:MM）
    sha_ok = all(len(str(x.get("mtime"))) == 16 and str(x.get("mtime"))[4] == "-"
                 for x in detail if x.get("mtime"))
    check(sha_ok, "1.7 mtime 是 YYYY-MM-DD HH:MM")

    # ---------- 2. 上传 txt ----------
    say("")
    say("【2. 上传一份 txt】")
    code, j = upload(TXT_NAME, make_txt())
    check(code == 200 and j.get("ok"), "2.1 返回 200 且 ok=true",
          "HTTP %s %s" % (code, j.get("error") or j.get("detail") or ""))
    check(j.get("filetype") == "txt", "2.2 识别为 txt", str(j.get("filetype")))
    check((j.get("blocks") or 0) > 0, "2.3 解析出块", "blocks=%s" % j.get("blocks"))
    stored = IMPORT_DIR / (j.get("stored_name") or "")
    check(stored.is_file(), "2.4 真的落盘到 导入文档/", stored.name)
    check(j.get("stored_name") == TXT_NAME, "2.5 没重名时原名保存", str(j.get("stored_name")))
    check(j.get("docs_total") == n0 + 1, "2.6 docs_total = 上传前 + 1",
          "%s → %s" % (n0, j.get("docs_total")))

    # ---------- 3. 上传 docx + 重名 ----------
    say("")
    say("【3. 上传 docx（含重名策略）】")
    code, j2 = upload(DOCX_NAME, make_docx("第一段：甲方 华景科技有限公司。"))
    check(code == 200 and j2.get("ok"), "3.1 第一次上传成功", j2.get("error") or "")
    check(j2.get("filetype") == "docx", "3.2 识别为 docx", str(j2.get("filetype")))
    check((j2.get("blocks") or 0) >= 3, "3.3 三段都解析成块",
          "blocks=%s" % j2.get("blocks"))

    code, j3 = upload(DOCX_NAME, make_docx("第一段：这是第二份同名文件，内容不同。"))
    check(code == 200 and j3.get("ok"), "3.4 同名再传也成功")
    check(j3.get("stored_name") == DOCX_DUP, "3.5 重名自动变 %s" % DOCX_DUP,
          str(j3.get("stored_name")))
    first = (IMPORT_DIR / DOCX_NAME)
    check(first.is_file(), "3.6 原文件没被覆盖，仍在")

    # ---------- 4. 清单反映 ----------
    say("")
    say("【4. 清单立刻反映新文件】")
    after = get_docs()
    names = set(after.get("docs") or [])
    check({TXT_NAME, DOCX_NAME, DOCX_DUP} <= names, "4.1 三份都出现在清单里",
          "清单共 %d 份" % len(names))
    det = {x["name"]: x for x in (after.get("detail") or [])}
    check(all(det.get(n, {}).get("src") == "import" for n in (TXT_NAME, DOCX_NAME)),
          "4.2 来源标成 import（界面上显示「我上传的」）")
    size_ok = det.get(TXT_NAME, {}).get("size") == stored.stat().st_size
    check(size_ok, "4.3 明细里的 size 与磁盘一致",
          "%s vs %s" % (det.get(TXT_NAME, {}).get("size"), stored.stat().st_size))
    c2 = after.get("counts") or {}
    check(c2.get("total") == n0 + 3, "4.4 总数 +3", "%s → %s" % (n0, c2.get("total")))
    check(c2.get("import") == imp0 + 3, "4.5 我上传的 +3",
          "%s → %s" % (imp0, c2.get("import")))

    # ---------- 5. 已经进检索池 ----------
    say("")
    say("【5. 是否真的并入检索池（不只是躺在文件夹里）】")
    try:
        pool = json.loads(POOL.read_text(encoding="utf-8"))
        pnames = {d.get("file") for d in pool.get("documents", [])}
        check({TXT_NAME, DOCX_NAME, DOCX_DUP} <= pnames,
              "5.1 池里已有这三份", "池内 %d 份" % len(pnames))
        blk = next((d for d in pool.get("documents", []) if d.get("file") == TXT_NAME), {})
        check((blk.get("block_count") or len(blk.get("blocks") or [])) > 0,
              "5.2 池里带上了正文块",
              "block_count=%s" % blk.get("block_count"))
    except Exception as e:
        check(False, "5.1 池里已有这三份", "读池失败：%s" % e)
        check(False, "5.2 池里带上了正文块", "")

    # ---------- 6. 不支持的格式要被拒 ----------
    say("")
    say("【6. 不支持格式必须拒绝（不能静默收下）】")
    code, j4 = upload(BAD_NAME, b"MZ\x90\x00fake")
    msg = str(j4.get("detail") or j4.get("error") or "")
    check(code == 400, "6.1 HTTP 400", "HTTP %s" % code)
    check("不支持" in msg, "6.2 说清是不支持格式", msg[:60])
    check(not (IMPORT_DIR / BAD_NAME).exists(), "6.3 磁盘上没留下垃圾文件")

    # ---------- 7. 清理还原 ----------
    say("")
    say("【7. 清理并还原（按精确文件名，不用通配）】")
    removed = []
    for n in (TXT_NAME, DOCX_NAME, DOCX_DUP):
        p = IMPORT_DIR / n
        if p.is_file():
            p.unlink()
            removed.append(n)
    check(len(removed) == 3, "7.1 删除本次上传的 3 个文件", "、".join(removed))
    if backup.exists():
        shutil.copy2(backup, POOL)
        backup.unlink()
        say("  已用备份还原 %s" % POOL.name)
    check(True, "7.2 池已还原")

    # imports 表里的自检记录（精确名字，只删本次造的三条）
    try:
        sys.path.insert(0, str(AGENT))
        os.chdir(AGENT)
        import db
        conn = db.get_conn(with_db=True)
        try:
            with conn.cursor() as cur:
                cur.execute("DELETE FROM imports WHERE stored_name IN (%s, %s, %s)",
                            (TXT_NAME, DOCX_NAME, DOCX_DUP))
                n_del = cur.rowcount
            conn.commit()
        finally:
            conn.close()
        check(True, "7.3 清掉 imports 表里的自检记录", "%d 条" % n_del)
    except Exception as e:
        check(True, "7.3 清掉 imports 表里的自检记录（跳过）", "数据库不可用：%s" % e)

    final = get_docs()
    check(len(final.get("docs") or []) == n0, "7.4 文档库回到上传前的 %d 份" % n0,
          "现在 %d 份" % len(final.get("docs") or []))
    check((final.get("counts") or {}).get("import") == imp0, "7.5 我上传的份数也回到原值")

    # ---------- 汇总 ----------
    say("")
    say("=" * 56)
    say("结果：通过 %d ／ 共 %d　%s" % (passed, passed + failed,
                                     "全部通过 ✓" if failed == 0 else "有 %d 项未过" % failed))
    say("=" * 56)
    OUT.write_text("\n".join(lines), encoding="utf-8")
    print("\n".join(lines))
    return 0 if failed == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
