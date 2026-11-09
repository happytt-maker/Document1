# -*- coding: utf-8 -*-
"""验证 api_server 的并发改造：两个填表任务同时提交。

看三件事：
  ① 还串不串行（以前 fill_lock 锁整条流水线，第二个任务要等第一个跑完）
  ② 产物是否各进各的目录（agent/jobs/<jid>/）
  ③ 结果对不对（成品能不能打开、有没有互相覆盖）

用 8011 端口，不打扰正在 8000 上跑的那个实例。
"""
import json
import time
from pathlib import Path

import requests

BASE = "http://127.0.0.1:8011"
OUT = Path(__file__).resolve().parent / "_服务端并发验证.txt"
TEMPLATES = ["模板1-设备采购汇总表.xlsx", "模板4-客诉处理台账.xlsx"]

s = requests.Session()
s.trust_env = False          # 绕开系统代理（本地 127.0.0.1 的 GET 会被代理拦成假 404）


def _res(j):
    """任务结果在 result 里（ok / report / meta / error），不是顶层。"""
    return (j or {}).get("result") or {}


def main():
    lines = []

    def say(x=""):
        lines.append(x)
        print(x)

    say("=" * 60)
    say("api_server 并发验证（端口 8011，并发闸门 2）")
    say("=" * 60)

    t0 = time.time()
    ids = []
    for tpl in TEMPLATES:
        r = s.post(BASE + "/api/fill", json={"template": tpl, "skip_read": True}, timeout=30)
        say("提交 %-28s → %s %s" % (tpl, r.status_code, r.text[:120]))
        if r.status_code == 200:
            ids.append((r.json()["job_id"], tpl))
    if not ids:
        say("一个都没提交成功")
        OUT.write_text("\n".join(lines), encoding="utf-8")
        return

    # 轮询到全部结束
    done, last = {}, 0.0
    while len(done) < len(ids):
        for jid, tpl in ids:
            if jid in done:
                continue
            r = s.get("%s/api/job/%s" % (BASE, jid), timeout=30)
            if r.status_code != 200:
                continue
            j = r.json()
            if j.get("status") in ("done", "error", "failed") or j.get("ok") is not None \
                    and j.get("status") != "running":
                if j.get("status") in ("done", "error", "failed"):
                    done[jid] = j
        last = time.time() - t0
        if last > 600:
            say("超时（>600 秒），停止等待")
            break
        time.sleep(1.0)

    wall = time.time() - t0
    say("")
    say("%-10s %-30s %-8s %10s" % ("job", "模板", "状态", "耗时(秒)"))
    jobs = {}
    for jid, tpl in ids:
        j = done.get(jid, {})
        jobs[jid] = j
        say("%-10s %-30s %-8s %10s"
            % (jid, tpl, j.get("status"), j.get("elapsed")))
    say("")
    say("总墙钟：%.2f 秒" % wall)

    # 目录隔离：workdir 必须是各自的
    say("")
    say("【产物隔离】")
    dirs = []
    for jid, tpl in ids:
        wd = (_res(jobs.get(jid)).get("meta") or {}).get("workdir")
        dirs.append(wd)
        say("  %s → %s" % (jid, wd))
    if len(set(dirs)) == len(dirs) and all(dirs):
        say("  ✓ 每个任务一套独立目录")
    else:
        say("  ✗ workdir 有重复或缺失")

    # 成品是否各就各位
    say("")
    say("【成品】")
    for jid, tpl in ids:
        wd = (_res(jobs.get(jid)).get("meta") or {}).get("workdir")
        if not wd:
            continue
        stem = Path(tpl).stem
        f = Path(wd) / "out" / (stem + "-已填写.xlsx")
        say("  %s → %s ｜ %s"
            % (Path(wd).name, f.name, "存在" if f.exists() else "缺失"))
        if f.exists():
            say("      %d 字节" % f.stat().st_size)

    # 并发证据：两个任务的执行区间是否重叠
    say("")
    say("【是否真并发】")
    el = [jobs.get(jid, {}).get("elapsed") or 0 for jid, _t in ids]
    sum_el = sum(el)
    say("  各任务耗时之和 = %.2f 秒 ｜ 实际总墙钟 = %.2f 秒" % (sum_el, wall))
    if wall < sum_el * 0.85:
        say("  ✓ 总墙钟明显小于耗时之和 → 两个任务在同时跑（不再是全局串行）")
    else:
        say("  注意：总墙钟接近耗时之和 → 可能仍在排队（检查 A23_FILL_CONCURRENCY）")

    # 报告摘要
    say("")
    say("【任务报告摘要】")
    for jid, tpl in ids:
        j = jobs.get(jid) or {}
        r = _res(j)
        say("  %s %s：ok=%s ｜ %s"
            % (jid, tpl, r.get("ok"),
               (r.get("report") or "").replace("\n", " ｜ ")[:220]))

    OUT.write_text("\n".join(lines), encoding="utf-8")


if __name__ == "__main__":
    main()
