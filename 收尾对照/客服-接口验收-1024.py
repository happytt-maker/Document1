# -*- coding: utf-8 -*-
"""AI 客服快捷指令 + 人工客服先留言：接口级端到端验收（不依赖浏览器）。"""
import json, sys, time, requests

BASE = "http://127.0.0.1:8000"
ok, bad = [], []


def check(name, cond, detail=""):
    (ok if cond else bad).append(name)
    print(("  PASS " if cond else "  FAIL ") + name + ("" if cond else "  <<< " + str(detail)[:300]))


def cli(user, pwd, portal):
    s = requests.Session()
    s.trust_env = False          # 关掉系统代理，避免 502 劫持
    r = s.post(BASE + "/api/auth/login",
               json={"username": user, "password": pwd, "portal": portal}, timeout=10)
    r.raise_for_status()
    return s


def fresh(s, mode):
    """清掉进行中的会话，再开一个全新的（一个用户同时只留一个会话）。"""
    for _ in range(6):
        j = s.get(BASE + "/api/support/mine", timeout=10).json()
        sess = j.get("session")
        if not sess:
            break
        s.post(BASE + "/api/support/close", json={"session_id": sess["id"]}, timeout=10)
    j = s.post(BASE + "/api/support/start", json={"mode": mode}, timeout=15).json()
    return j


# ---------------- 1. 快捷指令清单
u = cli("demo_ai", "123456", "user")
items = u.get(BASE + "/api/support/faq", timeout=10).json().get("items", [])
check("faq 列表 6 条", len(items) == 6, items)
check("faq 字段齐全", all({"no", "label", "q"} <= set(i) for i in items), items)
print("  清单：" + " / ".join("%s.%s" % (i["no"], i["label"]) for i in items))

# ---------------- 2. AI 会话：回数字秒答（不走大模型）
j = fresh(u, "ai")
sess = j["session"]; sid = sess["id"]
print("  AI 会话 id=%s mode=%s status=%s" % (sid, sess["mode"], sess["status"]))
check("AI 会话开局 ai/ai", sess["mode"] == "ai" and sess["status"] == "ai", sess)

for no, kw in [(1, "docx"), (3, "拖"), (6, "表格")]:
    t0 = time.time()
    d = u.post(BASE + "/api/support/send", json={"session_id": sid, "content": str(no)}, timeout=40).json()
    ms = int((time.time() - t0) * 1000)
    ai = d.get("ai") or {}
    txt = (ai.get("message") or {}).get("content", "")
    check("回 %d 秒回预设（%dms）" % (no, ms), bool(txt) and kw in txt and ms < 800, (ms, txt[:80]))
    check("回 %d 标记 faq 命中" % no, ai.get("faq") == no, ai)

d = u.post(BASE + "/api/support/send", json={"session_id": sid, "content": "２"}, timeout=40).json()
ai = d.get("ai") or {}
txt = (ai.get("message") or {}).get("content", "")
check("全角「２」也能命中", "导入文档" in txt and ai.get("faq") == 2, (ai.get("faq"), txt[:80]))

d = u.post(BASE + "/api/support/send", json={"session_id": sid, "content": "9"}, timeout=40).json()
txt = ((d.get("ai") or {}).get("message") or {}).get("content", "")
check("回 9 给出范围提示", ("1–6" in txt or "1-6" in txt), txt[:80])

# 原样问句也命中（不算大模型）
d = u.post(BASE + "/api/support/send", json={"session_id": sid, "content": "怎么插入图片？"}, timeout=40).json()
ai = d.get("ai") or {}
check("整句提问命中第 3 条", ai.get("faq") == 3, ai.get("faq"))

# ---------------- 3. 人工会话：不用接入就能发
h = cli("demo_human", "123456", "user")
j = fresh(h, "human")
hs = j["session"]; hid = hs["id"]
print("  人工会话 id=%s mode=%s status=%s" % (hid, hs["mode"], hs["status"]))
check("人工会话开局 human/waiting", hs["mode"] == "human" and hs["status"] == "waiting", hs)

t0 = time.time()
r = h.post(BASE + "/api/support/send", json={"session_id": hid, "content": "我想问导出格式能不能改（客服还没接入）"}, timeout=15)
ms = int((time.time() - t0) * 1000)
d = r.json()
check("waiting 状态可发消息（%dms）" % ms, r.status_code == 200 and d.get("ok"), (r.status_code, d))
check("waiting 发送返回留言提示", bool(d.get("notice")), d.get("notice"))
print("  提示：" + str(d.get("notice")))

# 客服侧：列表里能看到这条留言
kf = cli("kf_demo", "123456", "staff")
board = kf.get(BASE + "/api/staff/sessions", timeout=10).json()
waiting, mine = board.get("waiting", []), board.get("mine", [])
hit = [x for x in waiting + mine if x["id"] == hid]
check("客服工作台能看到该会话", bool(hit), ([x["id"] for x in waiting][:6], [x["id"] for x in mine][:6]))
if hit:
    print("  队列卡片：%s ｜ 留言 %s 条 ｜ 摘要 %s"
          % (hit[0].get("user_name"), hit[0].get("msg_count"), hit[0].get("last_text")))
    check("卡片标出「已留言 N 条」", int(hit[0].get("msg_count") or 0) >= 1, hit[0].get("msg_count"))

r = kf.post(BASE + "/api/staff/accept", json={"session_id": hid}, timeout=10)
check("客服接入", r.status_code == 200 and (r.json().get("session") or {}).get("status") == "active", r.text[:200])
msgs = kf.get(BASE + "/api/staff/poll?session_id=%d&after=0" % hid, timeout=10).json().get("messages", [])
pre = [m for m in msgs if m["sender"] == "user" and "导出格式" in m["content"]]
check("接入后能看到接入前留言", bool(pre), msgs[-3:])

r = kf.post(BASE + "/api/staff/send", json={"session_id": hid, "content": "可以的，导出时能选格式"}, timeout=10)
check("客服回复", r.status_code == 200, r.text[:200])
pm = h.get(BASE + "/api/support/poll?session_id=%d&after=0" % hid, timeout=10).json().get("messages", [])
got = [m for m in pm if m["sender"] == "staff" and "导出时能选格式" in m["content"]]
check("用户端收到客服回复", bool(got), pm[-3:])

# 收尾：结束两个会话
h.post(BASE + "/api/support/close", json={"session_id": hid}, timeout=10)
u.post(BASE + "/api/support/close", json={"session_id": sid}, timeout=10)
print("\n结果：通过 %d 项，失败 %d 项" % (len(ok), len(bad)))
if bad:
    print("失败项：" + "；".join(bad))
sys.exit(1 if bad else 0)
