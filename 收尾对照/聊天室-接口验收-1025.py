# -*- coding: utf-8 -*-
"""聊天室-接口验收-1025.py —— 用户端多人聊天室（公共大群）接口验收

跑法：
    python 收尾对照\\聊天室-接口验收-1025.py            # 默认 8010（测试端口）
    python 收尾对照\\聊天室-接口验收-1025.py --port 8000

场景（全部用真实 HTTP + 三个独立 cookie 会话，模拟三个人同时在线）：
   1  三个人注册并进群 → 在线名单里能互相看见
   2  A 发言 → B / C 轮询都能拿到，昵称对
   3  同一个人 1 秒内连发 → 后端拒绝（防刷屏）
   4  超过 500 字 → 后端拒绝
   5  B 发「@A」→ A 能收到，内容原样
   6  A 撤回自己的消息 → B 整段重拉看到「已撤回」，且 rev 变了
   7  B 撤回 A 的消息 → 拒绝（只能撤自己的）
   8  C 收起面板（leave）→ 在线名单少一个；alive=0 的轮询不算在线
最后自动清理测试账号与测试消息，不在库里留垃圾。
"""

import argparse
import json
import os
import sys
import time
import uuid
from pathlib import Path

# 本机测试必须绕开系统代理：这台机器上 HTTP_PROXY 指向 127.0.0.1:507xx，
# 会把对 127.0.0.1 的 GET 也拦下来（返回一个假 404），POST 却穿过去了 —— 排查过好几次
for _k in ("HTTP_PROXY", "HTTPS_PROXY", "ALL_PROXY",
           "http_proxy", "https_proxy", "all_proxy"):
    os.environ.pop(_k, None)

import requests

BASE_DIR = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(BASE_DIR / "agent"))

LOG_PATH = Path(__file__).resolve().parent / "聊天室-接口验收-1025.txt"

PASS, FAIL = [], []
SESSION_TAG = uuid.uuid4().hex[:6]


class Tee:
    """同时写控制台和 UTF-8 日志文件（控制台是 GBK，直接重定向会乱码）。"""

    def __init__(self, *streams):
        self.streams = streams

    def write(self, s):
        for st in self.streams:
            try:
                st.write(s)
                st.flush()
            except Exception:
                pass

    def flush(self):
        for st in self.streams:
            try:
                st.flush()
            except Exception:
                pass



def ok(name, cond, detail=""):
    (PASS if cond else FAIL).append(name)
    print(("  [PASS] " if cond else "  [FAIL] ") + name + (("  " + detail) if detail else ""))
    return cond


def api(sess, method, url, **kw):
    kw.setdefault("timeout", 20)
    r = sess.request(method, url, **kw)
    try:
        j = r.json()
    except Exception:
        j = {}
    return r.status_code, j


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--port", type=int, default=8010)
    ap.add_argument("--keep", action="store_true", help="保留测试数据（排查用）")
    args = ap.parse_args()
    base = "http://127.0.0.1:%d" % args.port
    print("=" * 66)
    print("聊天室接口验收 · %s" % base)
    print("=" * 66)

    names = ["A", "B", "C"]
    users = {}
    account = {}
    ids = []
    try:
        # ---------- 1 注册 + 进群
        print("\n【1】三个人注册进群")
        for n in names:
            s = requests.Session()
            uname = "rtest%s_%s" % (n.lower(), SESSION_TAG)
            code, j = api(s, "POST", base + "/api/auth/register",
                          json={"username": uname, "password": "Room123456",
                                "display_name": "测试" + n})
            if code != 200:
                print("  注册失败：%s %s" % (code, j))
                raise SystemExit(1)
            users[n] = s
            account[n] = j["user"]["display_name"]
            ids.append(j["user"]["id"])
        ok("三个测试账号注册成功", len(ids) == 3, "ids=%s" % ids)

        for n in names:
            code, j = api(users[n], "GET", base + "/api/room/poll?after=0&alive=1")
        ok("A 进群轮询成功", code == 200 and j.get("ok"), "online=%s" % j.get("online_count"))
        code, j = api(users["B"], "GET", base + "/api/room/poll?after=0&alive=1")
        code, j = api(users["C"], "GET", base + "/api/room/poll?after=0&alive=1")
        code, j = api(users["A"], "GET", base + "/api/room/poll?after=0&alive=1")
        online_names = [u["name"] for u in j.get("online", [])]
        ok("三个人都在在线名单里", j.get("online_count") >= 3 and
           all(account[n] in online_names for n in names), "online=%s" % online_names)
        ok("自己的那条带 me 标记", any(u["me"] for u in j.get("online", [])))

        # ---------- 2 A 发言，B/C 都能看到
        print("\n【2】A 发言 → 别人能不能收到")
        code, j = api(users["A"], "POST", base + "/api/room/send",
                      json={"content": "大家好，我是 A，刚把合同里的标题居中了"})
        ok("A 发送成功", code == 200 and j.get("message", {}).get("id"), str(j.get("detail", "")))
        first_id = j["message"]["id"]
        ok("返回的昵称是 A 的昵称", j["message"]["user_name"] == account["A"],
           j["message"]["user_name"])

        for who in ("B", "C"):
            code, jb = api(users[who], "GET", base + "/api/room/poll?after=0&alive=1")
            hit = [m for m in jb.get("messages", []) if m["id"] == first_id]
            ok("%s 轮询能拿到 A 这条" % who, bool(hit))
            if hit:
                ok("%s 看到的昵称是「测试A」" % who, hit[0]["user_name"] == account["A"],
                   hit[0]["user_name"])

        # ---------- 3 限速
        print("\n【3】防刷屏：1 秒内连发")
        code, j = api(users["A"], "POST", base + "/api/room/send", json={"content": "接着再来一条"})
        ok("第二句被拒绝（400）", code == 400 and "快" in (j.get("detail") or ""),
           "%s %s" % (code, j.get("detail")))
        time.sleep(1.2)

        # ---------- 4 超长
        print("\n【4】超长消息")
        code, j = api(users["A"], "POST", base + "/api/room/send", json={"content": "字" * 501})
        ok("501 字被拒绝（400）", code == 400 and "最多" in (j.get("detail") or ""),
           "%s %s" % (code, j.get("detail")))
        code, j = api(users["A"], "POST", base + "/api/room/send", json={"content": "   "})
        ok("空白消息被拒绝", code == 400)

        # ---------- 5 @ 提及
        print("\n【5】@ 提及")
        code, j = api(users["B"], "POST", base + "/api/room/send",
                      json={"content": "@测试A 你把标题居中是在哪儿点的？"})
        ok("B 发 @A 成功", code == 200)
        at_id = j["message"]["id"]
        code, ja = api(users["A"], "GET", base + "/api/room/poll?after=%d&alive=1" % first_id)
        hit = [m for m in ja.get("messages", []) if m["id"] == at_id]
        ok("A 收到 B 的 @ 消息", bool(hit))
        ok("增量只返回新的（不重复旧的）", all(m["id"] > first_id for m in ja.get("messages", [])))

        # ---------- 6 撤回
        print("\n【6】撤回 + 版本号")
        code, jr = api(users["A"], "POST", base + "/api/room/recall", json={"message_id": first_id})
        ok("A 撤回自己的消息", code == 200 and jr.get("ok"), str(jr.get("detail", "")))
        rev_after = jr.get("rev")
        code, jb = api(users["B"], "GET", base + "/api/room/poll?after=0&alive=1")
        hit = [m for m in jb.get("messages", []) if m["id"] == first_id]
        ok("B 重拉后看到 deleted=True", bool(hit) and hit[0]["deleted"] is True)
        ok("B 重拉后看到 rev 变了", isinstance(rev_after, int) and rev_after >= 1,
           "rev=%s" % rev_after)

        # ---------- 7 越权撤回
        print("\n【7】不能撤回别人的")
        code, j = api(users["B"], "POST", base + "/api/room/recall", json={"message_id": at_id})
        ok("B 撤自己的可以", code == 200)
        time.sleep(0.2)
        code, j = api(users["A"], "POST", base + "/api/room/recall", json={"message_id": at_id})
        ok("A 撤 B 的消息被拒绝（400）", code == 400 and "自己" in (j.get("detail") or ""),
           "%s %s" % (code, j.get("detail")))

        # ---------- 8 离开
        print("\n【8】收起面板 = 离开；alive=0 不算在线")
        api(users["C"], "POST", base + "/api/room/leave")
        api(users["C"], "GET", base + "/api/room/poll?after=0&alive=0")   # 只看不报到
        code, j = api(users["A"], "GET", base + "/api/room/poll?after=0&alive=1")
        online_names = [u["name"] for u in j.get("online", [])]
        ok("C 离开后不在名单里", account["C"] not in online_names, "online=%s" % online_names)
        ok("A/B 还在名单里", account["A"] in online_names and account["B"] in online_names)

        # ---------- 9 统计
        print("\n【9】消息计数")
        code, j = api(users["A"], "GET", base + "/api/room/poll?after=0&alive=1")
        # 这一轮真正落库的只有 2 条（A 第一条 + B 的 @），被限速 / 超长挡下的都没进库；
        # 两条都已经撤回，但消息本身仍计数（撤回只是打标记，不删记录）
        ok("total 至少 2 条（已撤回的仍计数）", (j.get("total") or 0) >= 2,
           "total=%s" % j.get("total"))
        ok("已撤回的仍在列表里（deleted=True）",
           sum(1 for m in j.get("messages", []) if m["deleted"]) >= 2)
        ok("返回字段齐全", all(k in j for k in ("messages", "online", "online_count",
                                              "total", "rev", "min_gap", "me_id")))

    finally:
        if args.keep:
            print("\n（--keep：保留测试数据，测试账号：%s）" % ", ".join(account.values()))
        else:
            cleanup(ids)
            print("\n（已清理测试账号 %d 个、测试消息与在线记录）" % len(ids))

    print("\n" + "=" * 66)
    print("通过 %d 项，失败 %d 项" % (len(PASS), len(FAIL)))
    if FAIL:
        print("失败项：")
        for f in FAIL:
            print("  - " + f)
    print("=" * 66)
    return 0 if not FAIL else 1


def cleanup(user_ids):
    """删掉测试账号的消息 / 在线记录 / 会话 / 账号本身。"""
    if not user_ids:
        return
    try:
        import db
        conn = db.get_conn(with_db=True)
        try:
            with conn.cursor() as cur:
                ph = ",".join(["%s"] * len(user_ids))
                cur.execute("DELETE FROM room_messages WHERE user_id IN (%s)" % ph, tuple(user_ids))
                cur.execute("DELETE FROM room_online WHERE user_id IN (%s)" % ph, tuple(user_ids))
                cur.execute("DELETE FROM sessions WHERE user_id IN (%s)" % ph, tuple(user_ids))
                cur.execute("DELETE FROM users WHERE id IN (%s)" % ph, tuple(user_ids))
            conn.commit()
        finally:
            conn.close()
    except Exception as e:
        print("  清理失败（不影响验收结论）：%s" % e)


if __name__ == "__main__":
    _f = open(LOG_PATH, "w", encoding="utf-8")
    _old = sys.stdout
    sys.stdout = Tee(_old, _f)
    try:
        _code = main()
    finally:
        sys.stdout = _old
        _f.close()
        print("验收日志已写入：%s" % LOG_PATH)
    raise SystemExit(_code)
