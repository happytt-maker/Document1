# -*- coding: utf-8 -*-
"""模板上传接口验收（8010 测试端口）"""
import io
import json
import os
import sys
import time
import uuid

import requests

BASE = "http://127.0.0.1:8010"
ROOT = r"C:\Users\覃敬涛\Desktop\研一项目"
PASS, FAIL = [], []
LOG = []


def ok(name, cond, extra=""):
    (PASS if cond else FAIL).append(name)
    line = "%s %s%s" % ("[通过]" if cond else "[失败]", name, ("  —— " + str(extra)) if extra else "")
    LOG.append(line)
    print(line)


s = requests.Session()
s.trust_env = False        # 绕开系统代理（这个环境里 502 的老问题）


def get_tpls():
    r = s.get(BASE + "/api/fill/templates", timeout=30)
    return r.status_code, r.json()


def upload(fname, data, content_type="application/octet-stream"):
    r = s.post(BASE + "/api/fill/templates/upload",
               files={"file": (fname, data, content_type)}, timeout=120)
    try:
        j = r.json()
    except Exception:
        j = {"raw": r.text[:300]}
    return r.status_code, j


def delete(name):
    r = s.delete(BASE + "/api/fill/templates", params={"name": name}, timeout=30)
    try:
        j = r.json()
    except Exception:
        j = {"raw": r.text[:300]}
    return r.status_code, j


def main():
    LOG.append("=" * 62)
    LOG.append("模板上传 · 接口验收")
    LOG.append("=" * 62)

    # ---------- 0 服务在不在
    try:
        code, j = get_tpls()
    except Exception as e:
        LOG.append("服务没起来：%s" % e)
        open(os.path.join(os.path.dirname(os.path.abspath(__file__)),
                          "模板上传-接口验收-1026.txt"), "w", encoding="utf-8").write("\n".join(LOG))
        return 1

    base_tpls = j.get("templates") or []
    base_names = {t["name"] for t in base_tpls}
    LOG.append("【0】原始清单：%d 张（xlsx %d / docx %d）" % (
        len(base_tpls),
        sum(1 for t in base_tpls if t["kind"] == "xlsx"),
        sum(1 for t in base_tpls if t["kind"] == "docx")))
    ok("GET /api/fill/templates 200", code == 200, code)
    ok("清单里每项都有 upload 字段", all("upload" in t for t in base_tpls), len(base_tpls))
    ok("系统自带模板 upload=false", all(t["upload"] is False for t in base_tpls))

    tag = uuid.uuid4().hex[:6]
    xlsx_src = os.path.join(ROOT, "测试数据", "模板表格", "模板1-设备采购汇总表.xlsx")
    docx_src = os.path.join(ROOT, "测试数据", "模板表格-word", "模板1-设备采购汇总表.docx")
    xlsx_bytes = open(xlsx_src, "rb").read()
    docx_bytes = open(docx_src, "rb").read()

    up_xlsx = "验收-采购汇总-%s.xlsx" % tag
    up_docx = "验收-采购汇总-%s.docx" % tag

    # ---------- 1 正常上传 xlsx
    LOG.append("\n【1】上传 xlsx")
    code, j = upload(up_xlsx, xlsx_bytes)
    ok("上传返回 200", code == 200, "%s %s" % (code, str(j)[:120]))
    ok("返回里带着新模板名", j.get("name") == up_xlsx, j.get("name"))
    ok("返回 ok=True", j.get("ok") is True)
    code, j2 = get_tpls()
    names = {t["name"] for t in j2["templates"]}
    ok("清单里出现了它", up_xlsx in names)
    got = next((t for t in j2["templates"] if t["name"] == up_xlsx), None)
    ok("它的 upload 标记为 true", bool(got and got["upload"] is True), got)
    ok("它的 kind 判定为 xlsx", bool(got and got["kind"] == "xlsx"), got and got.get("kind"))
    ok("清单总数 +1", len(j2["templates"]) == len(base_tpls) + 1,
       "%d -> %d" % (len(base_tpls), len(j2["templates"])))

    # ---------- 2 上传 docx
    LOG.append("\n【2】上传 docx")
    code, j = upload(up_docx, docx_bytes)
    ok("docx 上传 200", code == 200, "%s %s" % (code, str(j)[:120]))
    code, j2 = get_tpls()
    got = next((t for t in j2["templates"] if t["name"] == up_docx), None)
    ok("docx 的 kind 判定为 docx", bool(got and got["kind"] == "docx"), got and got.get("kind"))
    ok("docx 也能标 upload=true", bool(got and got["upload"]))

    # ---------- 3 重名拒绝
    LOG.append("\n【3】重名保护")
    code, j = upload(up_xlsx, xlsx_bytes)
    ok("同名再传被拒 409", code == 409, "%s %s" % (code, str(j.get("detail"))[:100]))
    code, j = upload("模板1-设备采购汇总表.xlsx", xlsx_bytes)
    ok("跟系统自带模板重名被拒 409", code == 409, "%s %s" % (code, str(j.get("detail"))[:80]))

    # ---------- 4 格式白名单
    LOG.append("\n【4】格式白名单")
    code, j = upload("坏文件-%s.txt" % tag, b"hello world")
    ok(".txt 被拒 400", code == 400, "%s %s" % (code, str(j.get("detail"))[:90]))
    ok("提示里说明了只支持哪些", "xlsx" in str(j.get("detail", "")).lower(), str(j.get("detail"))[:110])
    code, j = upload("旧格式-%s.xls" % tag, b"\xd0\xcf\x11\xe0")
    ok(".xls 旧格式被拒 400", code == 400, code)
    code, j = upload("无扩展名-%s" % tag, xlsx_bytes)
    ok("无扩展名被拒 400", code == 400, code)

    # ---------- 5 假文件（扩展名对、内容不对）
    LOG.append("\n【5】内容校验")
    code, j = upload("伪装-%s.xlsx" % tag, b"this is not a real xlsx" * 10)
    ok("假 xlsx（非 PK 头）被拒 400", code == 400, "%s %s" % (code, str(j.get("detail"))[:90]))
    ok("提示教人去另存为", "另存为" in str(j.get("detail", "")), str(j.get("detail"))[:120])
    code, j = upload("坏docx-%s.docx" % tag, b"PK\x03\x04" + b"\x00" * 200)
    ok("残缺 docx 被拒 400", code == 400, code)
    code, j = upload("空文件-%s.xlsx" % tag, b"")
    ok("空文件被拒 400", code == 400, code)

    # ---------- 6 超大文件
    LOG.append("\n【6】体积上限")
    big = b"PK\x03\x04" + os.urandom(13 * 1024 * 1024)
    code, j = upload("超大-%s.xlsx" % tag, big)
    ok("13MB 被拒 400", code == 400, "%s %s" % (code, str(j.get("detail"))[:90]))
    ok("提示里带上了 MB 数字", "MB" in str(j.get("detail", "")), str(j.get("detail"))[:100])

    # ---------- 7 路径注入
    LOG.append("\n【7】文件名清洗")
    code, j = upload("../../evil-%s.xlsx" % tag, xlsx_bytes)
    ok("带 ../ 的文件名能传但被洗掉", code == 200 and "/" not in j.get("name", "") and "\\" not in j.get("name", ""),
       "%s -> %s" % (code, j.get("name")))
    ok("没有写到上传目录之外", not os.path.exists(os.path.join(ROOT, "evil-%s.xlsx" % tag)))
    if code == 200:
        delete(j["name"])
    code, j = upload("a:b*c?-%s.xlsx" % tag, xlsx_bytes)
    ok("特殊字符被替换", code == 200 and ":" not in j.get("name", ""), j.get("name"))
    if code == 200:
        delete(j["name"])

    # ---------- 8 删除保护
    LOG.append("\n【8】删除")
    code, j = delete("模板1-设备采购汇总表.xlsx")
    ok("删系统自带模板被拒 403", code == 403, "%s %s" % (code, str(j.get("detail"))[:90]))
    code, j = get_tpls()
    ok("系统模板还在", "模板1-设备采购汇总表.xlsx" in {t["name"] for t in j["templates"]})
    code, j = delete("根本没有这个-%s.xlsx" % tag)
    ok("删不存在的返回 404", code == 404, code)
    code, j = delete(up_docx)
    ok("删自己上传的 docx 成功", code == 200, "%s %s" % (code, str(j)[:100]))
    code, j = get_tpls()
    ok("删完清单里没有了", up_docx not in {t["name"] for t in j["templates"]})
    ok("系统模板数量没受影响",
       len([t for t in j["templates"] if not t["upload"]]) == len(base_tpls),
       len(j["templates"]))

    # ---------- 9 上传的模板真能跑填表
    LOG.append("\n【9】端到端：上传的模板跑一次表格填写")
    t0 = time.time()
    r = s.post(BASE + "/api/fill", json={"template": up_xlsx, "skip_read": True}, timeout=30)
    ok("提交填表任务 200", r.status_code == 200, "%s %s" % (r.status_code, r.text[:150]))
    if r.status_code == 200:
        jid = r.json()["job_id"]
        final = None
        for _ in range(120):
            time.sleep(1.5)
            jr = s.get(BASE + "/api/job/" + jid, timeout=30).json()
            if jr.get("status") in ("done", "error"):
                final = jr
                break
        cost = time.time() - t0
        LOG.append("    任务状态=%s 耗时=%.1fs" % ((final or {}).get("status"), cost))
        if final:
            LOG.append("    报告：" + str((final.get("result") or {}).get("report", ""))[:400].replace("\n", "\n    "))
        ok("任务跑到 done", bool(final and final.get("status") == "done"), (final or {}).get("status"))
        ok("跑成功了（ok=True）", bool(final and (final.get("result") or {}).get("ok")),
           ((final or {}).get("result") or {}).get("report", "")[:200])

    # ---------- 10 收尾清理
    LOG.append("\n【10】清理验收残留")
    code, j = delete(up_xlsx)
    ok("删掉验收用的上传模板", code == 200, code)
    code, j = get_tpls()
    ok("清单恢复原样", len(j["templates"]) == len(base_tpls),
       "%d vs %d" % (len(j["templates"]), len(base_tpls)))
    ok("没有残留的 upload 模板", all(not t["upload"] for t in j["templates"]))

    LOG.append("\n" + "=" * 62)
    LOG.append("通过 %d 项 ｜ 失败 %d 项" % (len(PASS), len(FAIL)))
    if FAIL:
        LOG.append("失败项：" + "；".join(FAIL))
    LOG.append("=" * 62)

    out = os.path.join(os.path.dirname(os.path.abspath(__file__)), "模板上传-接口验收-1026.txt")
    open(out, "w", encoding="utf-8").write("\n".join(LOG))
    print("\n写出：" + out)
    return 0 if not FAIL else 1


if __name__ == "__main__":
    raise SystemExit(main())
