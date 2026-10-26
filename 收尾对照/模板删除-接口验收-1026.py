# -*- coding: utf-8 -*-
"""「删除已上传模板」接口验收（默认 8012 测试端口）

只验删除这条线：
  清单带上传时间 → 上传两张 → 逐张删 → 文件真的没了 → 系统模板拒绝 → 不存在的 404
  → 临时上传文件不冒充模板 → 最后清干净。
"""
import json
import os
import random
import string
import sys
import time

import requests

BASE = "http://127.0.0.1:8012"
ROOT = r"C:\Users\覃敬涛\Desktop\研一项目"
UP_DIR = os.path.join(ROOT, "测试数据", "模板上传")
SYS_XLSX = os.path.join(ROOT, "测试数据", "模板表格", "模板1-设备采购汇总表.xlsx")
SYS_DOCX = os.path.join(ROOT, "测试数据", "模板表格-word", "模板1-设备采购汇总表.docx")

PASS, FAIL, LOG = [], [], []


def ok(name, cond, extra=""):
    (PASS if cond else FAIL).append(name)
    line = "%s %s%s" % ("[通过]" if cond else "[失败]", name,
                        ("  —— " + str(extra)) if extra else "")
    LOG.append(line)
    print(line)


def say(s):
    LOG.append(s)
    print(s)


s = requests.Session()
s.trust_env = False          # 绕开系统代理（这个环境 502 的老毛病）


def tpls():
    r = s.get(BASE + "/api/fill/templates", timeout=30)
    return r.status_code, (r.json().get("templates") or [])


def upload(name, data):
    r = s.post(BASE + "/api/fill/templates/upload",
               files={"file": (name, data, "application/octet-stream")}, timeout=120)
    try:
        j = r.json()
    except Exception:
        j = {"raw": r.text[:200]}
    return r.status_code, j


def delete(name):
    r = s.delete(BASE + "/api/fill/templates", params={"name": name}, timeout=30)
    try:
        j = r.json()
    except Exception:
        j = {"raw": r.text[:200]}
    return r.status_code, j


def main():
    LOG.append("=" * 62)
    LOG.append("模板删除 · 接口验收（%s）" % BASE)
    LOG.append("=" * 62)

    # ---------- 0 服务
    try:
        code, base = tpls()
    except Exception as e:
        say("服务没起来：%s" % e)
        return 1
    base_upload = [t for t in base if t["upload"]]
    say("【0】当前模板 %d 张（系统 %d / 上传 %d）"
        % (len(base), len(base) - len(base_upload), len(base_upload)))
    ok("GET 模板清单 200", code == 200, code)
    ok("每项都带 mtime（用于显示上传时间）", all("mtime" in t for t in base),
       [t["name"] for t in base if "mtime" not in t][:3])

    tag = "".join(random.choice(string.ascii_lowercase + string.digits) for _ in range(6))
    n1 = "删除验收-甲-%s.xlsx" % tag
    n2 = "删除验收-乙-%s.docx" % tag
    c1, j1 = upload(n1, open(SYS_XLSX, "rb").read())
    c2, j2 = upload(n2, open(SYS_DOCX, "rb").read())
    ok("准备用：两张模板都传上去了", c1 == 200 and c2 == 200, (c1, c2))

    # ---------- 1 清单里能看见「我上传的」
    LOG.append("\n【1】清单能认出自己上传的")
    code, lst = tpls()
    mine = {t["name"]: t for t in lst if t["upload"]}
    ok("两张都在上传清单里", n1 in mine and n2 in mine, sorted(mine))
    ok("删之前上传数 = 2", len(mine) == 2, len(mine))
    ok("带 size", bool(mine.get(n1, {}).get("size")))
    ok("mtime 是「年-月-日 时:分」", bool(mine.get(n1, {}).get("mtime")) and
       len(str(mine[n1]["mtime"])) == 16, mine.get(n1, {}).get("mtime"))

    # ---------- 2 正常删除（xlsx）
    LOG.append("\n【2】删掉一张自己上传的 xlsx")
    code, j = delete(n1)
    ok("删除返回 200", code == 200, "%s %s" % (code, str(j)[:120]))
    ok("返回里说明删了哪张", j.get("deleted") == n1, j.get("deleted"))
    ok("返回里带剩余上传数（1）", j.get("uploaded_left") == 1, j.get("uploaded_left"))
    ok("磁盘上的文件真的没了", not os.path.isfile(os.path.join(UP_DIR, n1)))
    code, lst = tpls()
    names = {t["name"] for t in lst}
    ok("清单里也没有它了", n1 not in names)
    ok("清单总数回到原数 +1", len(lst) == len(base) + 1, "%d" % len(lst))

    # ---------- 3 反复删同一张
    LOG.append("\n【3】重复删除")
    code, j = delete(n1)
    ok("再删同一张 → 404", code == 404, "%s %s" % (code, str(j.get("detail"))[:80]))

    # ---------- 4 系统模板保护
    LOG.append("\n【4】系统自带模板删不掉")
    code, j = delete("模板1-设备采购汇总表.xlsx")
    ok("删系统 xlsx → 403", code == 403, "%s %s" % (code, str(j.get("detail"))[:90]))
    ok("提示里说了「系统自带」", "系统自带" in str(j.get("detail")), str(j.get("detail"))[:100])
    ok("系统 xlsx 还在磁盘上", os.path.isfile(SYS_XLSX))
    code, j = delete("模板1-设备采购汇总表.docx")
    ok("删系统 docx → 403", code == 403, code)
    ok("系统 docx 还在磁盘上", os.path.isfile(SYS_DOCX))
    code, lst = tpls()
    ok("系统模板数量没变", len([t for t in lst if not t["upload"]]) == len(base) - len(base_upload),
       len(lst))

    # ---------- 5 名字清洗 / 路径穿越
    LOG.append("\n【5】名字不干净时不会删到别处")
    code, j = delete("../../agent/api_server.py")
    ok("带 ../ 的名字被洗干净后按 404 处理", code == 404, "%s %s" % (code, str(j)[:90]))
    ok("api_server.py 毫发无伤", os.path.isfile(os.path.join(ROOT, "agent", "api_server.py")))
    code, j = delete("")
    ok("空名字 → 400", code == 400, code)

    # ---------- 6 正在上传的临时文件不算模板
    LOG.append("\n【6】半截临时文件不冒充模板")
    tmp = os.path.join(UP_DIR, ".uploading-真空档-%s.xlsx" % tag)
    os.makedirs(UP_DIR, exist_ok=True)
    with open(tmp, "wb") as fh:
        fh.write(b"PK\x03\x04half")
    code, lst = tpls()
    ok("清单里看不到 .uploading- 临时文件",
       all(not t["name"].startswith(".uploading-") for t in lst),
       [t["name"] for t in lst if t["name"].startswith(".uploading-")])
    os.remove(tmp)

    # ---------- 7 删掉最后一张 → 上传池清空
    LOG.append("\n【7】删到一张不剩")
    code, j = delete(n2)
    ok("删 docx 返回 200", code == 200, code)
    ok("剩余上传数 = 0", j.get("uploaded_left") == 0, j.get("uploaded_left"))
    ok("上传目录里没有验收残留",
       not os.path.isfile(os.path.join(UP_DIR, n2)),
       os.listdir(UP_DIR) if os.path.isdir(UP_DIR) else "目录不存在")
    code, lst = tpls()
    ok("清单回到原来的 %d 张" % len(base), len(lst) == len(base), len(lst))
    ok("没有 upload 模板了", all(not t["upload"] for t in lst))

    # ---------- 8 文件被占用时的提示（信息性，不计对错）
    LOG.append("\n【8】（信息）文件被占用时")
    n3 = "占用测试-%s.xlsx" % tag
    c3, _ = upload(n3, open(SYS_XLSX, "rb").read())
    if c3 == 200:
        p = os.path.join(UP_DIR, n3)
        fh = open(p, "rb")
        code, j = delete(n3)
        fh.close()
        say("    文件被本进程打开时调用删除 → HTTP %s，提示：%s"
            % (code, str(j.get("detail"))[:100]))
        if code != 200:
            ok("被占用时给的是 423 + 人话", code == 423, code)
            code2, j2 = delete(n3)
            ok("关掉之后能正常删掉", code2 == 200, code2)
        else:
            say("    （Windows 这个场景没拦住，不影响功能）")
        if os.path.isfile(p):
            delete(n3)

    LOG.append("\n" + "=" * 62)
    LOG.append("通过 %d 项 ｜ 失败 %d 项" % (len(PASS), len(FAIL)))
    if FAIL:
        LOG.append("失败项：" + "；".join(FAIL))
    LOG.append("=" * 62)

    out = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                       "模板删除-接口验收-1026.txt")
    open(out, "w", encoding="utf-8").write("\n".join(LOG))
    print("\n写出：" + out)
    return 0 if not FAIL else 1


if __name__ == "__main__":
    raise SystemExit(main())
