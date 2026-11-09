# -*- coding: utf-8 -*-
"""源文档填表（2026-11-07）自检：接口 + 真跑一次 + 范围生效的证据。

验的事：
  A. GET /api/docs 里的文件，/api/fill 能认（docs 传名字）
  B. 传一个文档库里没有的名字 → 404 + 人话
  C. 指定 1 份源文档真跑一单 → 报告里出现「源文档：…（只在这 1 份里找数）」
  D. 范围真的生效：任务目录里的小池只剩这 1 份文档
"""
import json
import os
import sys
import time
from pathlib import Path

# 本机测试：把代理清掉（系统代理会拦 127.0.0.1 的 GET，返回假 404）
for k in list(os.environ):
    if "PROXY" in k.upper():
        os.environ.pop(k, None)

import requests

BASE = "http://127.0.0.1:8000"
ROOT = Path(r"C:\Users\覃敬涛\Desktop\研一项目")
S = requests.Session()
S.trust_env = False
OK = []


def ck(name, cond, extra=""):
    OK.append((name, bool(cond), extra))
    print(("  [PASS] " if cond else "  [FAIL] ") + name + (("  " + extra) if extra else ""))


def main():
    print("=" * 66)
    print("源文档填表 · 自检")
    print("=" * 66)

    docs = S.get(BASE + "/api/docs", timeout=20).json()
    names = docs.get("docs") or []
    print("\n[A] 文档库 %d 份：%s" % (len(names), "、".join(names[:5]) + "…"))
    ck("文档库非空", len(names) > 0)

    tps = S.get(BASE + "/api/fill/templates", timeout=20).json().get("templates", [])
    print("\n[B] 可填模板 %d 张：%s" % (len(tps), "、".join(t["name"] for t in tps[:4]) + "…"))
    ck("模板非空", len(tps) > 0)

    # 负向：文档库里没有的名字
    r = S.post(BASE + "/api/fill", timeout=20,
               json={"template": tps[0]["name"], "docs": ["根本不存在的文档.docx"]})
    print("\n[C] 传不存在的源文档 → %d %s" % (r.status_code, r.text[:120]))
    ck("不存在的源文档被挡住（404）", r.status_code == 404)
    ck("错误信息是人话", "文档库里没有" in r.text)

    # 正向：挑一份 Excel 模板 + 一份 txt 源文档（快）
    tpl = next((t["name"] for t in tps
                if t["kind"] == "xlsx" and "验收" in t["name"]), None) or tps[0]["name"]
    src = next((n for n in names if n.endswith(".txt")), names[0])
    print("\n[D] 真跑一单：模板=%s ｜ 源文档=%s" % (tpl, src))
    r = S.post(BASE + "/api/fill", timeout=60,
               json={"template": tpl, "docs": [src], "overwrite": True})
    print("    提交 → %d %s" % (r.status_code, r.text[:160]))
    ck("提交成功", r.status_code == 200)
    if r.status_code != 200:
        return finish()
    jid = r.json().get("job_id")

    t0 = time.time()
    j = {}
    while time.time() - t0 < 300:
        j = S.get(BASE + "/api/job/" + jid, timeout=20).json()
        if j.get("status") in ("done", "failed"):
            break
        time.sleep(1.5)
    res = j.get("result") or {}
    rep = res.get("report") or ""
    print("\n[E] 任务 %s → %s ｜ %.1fs" % (jid, j.get("status"), time.time() - t0))
    print("-" * 66)
    print(rep[:1200])
    print("-" * 66)
    ck("任务跑完（不报错）", j.get("status") == "done" and (res.get("ok") is not None))
    ck("报告里写明只在这 1 份里找数", "只在这 1 份里找数" in rep)
    ck("报告里点名了源文档", src.split(".")[0] in rep)

    # 小池只剩这一份
    pool = ROOT / "agent" / "jobs" / jid / "pool.json"
    if pool.exists():
        d = json.loads(pool.read_text(encoding="utf-8"))
        files = [x.get("file") for x in d.get("documents", [])]
        print("\n[F] 任务小池：%s（%d 份）" % ("、".join(files), len(files)))
        ck("小池只剩勾的那 1 份", files == [src], str(files))
    else:
        ck("任务小池已生成", False, str(pool))

    return finish()


def finish():
    n = len(OK)
    p = sum(1 for _k, ok, _e in OK if ok)
    print("\n" + "=" * 66)
    print("合计 %d 项 ｜ 通过 %d ｜ 失败 %d" % (n, p, n - p))
    print("=" * 66)
    return 0 if p == n else 1


if __name__ == "__main__":
    sys.exit(main())
