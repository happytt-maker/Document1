# -*- coding: utf-8 -*-
"""网页端（Web API）真机冒烟：起服务 → 带「用户要求」提填表任务 → 轮询到完成 → 看结果
重点验证：require 透传 → A23_USER_REQ → 行条件 → 直通筛行，界面拿到的是 2542 行而不是 15 行。
"""
import json
import os
import subprocess
import sys
import time
import urllib.request

BASE = r"C:\Users\覃敬涛\Desktop\研一项目"
PY = sys.executable
PORT = 8011
URL = "http://127.0.0.1:%d" % PORT

# 本机访问别走系统代理（代理会劫持 127.0.0.1 的 GET，返回假的 404）
for k in list(os.environ):
    if k.lower().endswith("_proxy"):
        os.environ.pop(k, None)
OPENER = urllib.request.build_opener(urllib.request.ProxyHandler({}))


def req(path, data=None, method=None):
    body = None
    headers = {}
    if data is not None:
        body = json.dumps(data).encode("utf-8")
        headers["Content-Type"] = "application/json"
    r = urllib.request.Request(URL + path, data=body, headers=headers,
                               method=method or ("POST" if body else "GET"))
    with OPENER.open(r, timeout=30) as resp:
        return json.loads(resp.read().decode("utf-8"))


def main():
    log = open(r"C:\Users\覃敬涛\Desktop\研一项目\收尾对照\_API冒烟-服务日志.txt", "w", encoding="utf-8")
    srv = subprocess.Popen([PY, os.path.join(BASE, "agent", "api_server.py"), "--port", str(PORT)],
                           cwd=BASE, stdout=log, stderr=subprocess.STDOUT)
    try:
        # 等就绪
        for _ in range(60):
            try:
                tpls = req("/api/fill/templates")
                break
            except Exception:
                time.sleep(1)
        else:
            print("服务没起来，看 _API冒烟-服务日志.txt")
            return 1
        names = [t["name"] for t in tpls.get("templates", tpls if isinstance(tpls, list) else [])]
        print("可填模板：", names)
        target = "COVID-19 模板.xlsx"
        assert target in names, "模板清单里没有 %s" % target

        job = req("/api/fill", {
            "template": target,
            "skip_read": False,
            "require": "智能填表，将两个文件的内容中日期一列从2020/7/1到2020/8/31的数据填入模板中.",
        })
        jid = job["job_id"]
        print("已提交：", jid)

        t0 = time.time()
        while time.time() - t0 < 300:
            st = req("/api/job/%s" % jid)
            if st.get("status") in ("done", "failed"):
                break
            time.sleep(2)
        print("\n状态：", st.get("status"), "｜ 耗时 %s 秒" % st.get("elapsed"))
        res = st.get("result") or {}
        print("\n--- 界面拿到的报告 ---")
        print(res.get("report"))
        print("\n--- meta ---")
        m = res.get("meta") or {}
        for k in ("elapsed", "stages", "output_file", "pool_mb", "gate_weight"):
            print("  %s = %s" % (k, m.get(k)))
        if res.get("error"):
            print("  error =", res["error"])
        return 0 if res.get("ok") else 2
    finally:
        srv.terminate()
        try:
            srv.wait(timeout=10)
        except Exception:
            srv.kill()
        log.close()


if __name__ == "__main__":
    sys.exit(main())
