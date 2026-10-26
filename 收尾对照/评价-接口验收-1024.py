# -*- coding: utf-8 -*-
"""评价 / 意见反馈：接口级验收（用户提交 → 客服端 → 管理员端闭环）。"""
import sys, time, requests

BASE = "http://127.0.0.1:8000"
ok, bad = [], []


def check(name, cond, detail=""):
    (ok if cond else bad).append(name)
    print(("  PASS " if cond else "  FAIL ") + name + ("" if cond else "  <<< " + str(detail)[:280]))


def cli(u, p, portal):
    s = requests.Session(); s.trust_env = False
    r = s.post(BASE + "/api/auth/login", json={"username": u, "password": p, "portal": portal}, timeout=10)
    r.raise_for_status()
    return s


def purge(s):
    for _ in range(8):
        j = s.get(BASE + "/api/support/mine", timeout=10).json()
        if not j.get("session"):
            return
        s.post(BASE + "/api/support/close", json={"session_id": j["session"]["id"]}, timeout=10)


# ---------------- 1. 用户端：各功能点都能提交
u = cli("demo_ai", "123456", "user")
for feat, rating, text, tags in [
    ("doc_edit", 5, "把标题居中、加粗一次就成了，格式没乱", "改得对,速度快"),
    ("extract", 4, "合同金额抽对了，但甲方联系人漏了一位", "漏了字段"),
    ("fill", 5, "模板填写省了半小时，数字和原文对得上", "填得对,省时间"),
    ("other", 0, "希望操作历史能按文档名筛一下", "希望加功能"),
]:
    r = u.post(BASE + "/api/feedback", json={"feature": feat, "rating": rating,
                                             "content": text, "tags": tags}, timeout=10)
    d = r.json()
    check("提交 %s 评价" % feat, r.status_code == 200 and d.get("item", {}).get("id"), r.text[:160])

check("空提交被拦下", u.post(BASE + "/api/feedback",
                          json={"feature": "doc_edit", "rating": 0, "content": ""}, timeout=10).status_code == 400)
check("乱填功能点被拦下", u.post(BASE + "/api/feedback",
                            json={"feature": "abc", "rating": 5}, timeout=10).status_code == 400)
check("星级超范围被拦下", u.post(BASE + "/api/feedback",
                            json={"feature": "fill", "rating": 9}, timeout=10).status_code == 400)

mine = u.get(BASE + "/api/feedback/mine", timeout=10).json().get("items", [])
check("我的评价能查回来", len(mine) >= 4, len(mine))

# ---------------- 2. 客服会话评价（走真实流程：先留言 → 客服接入 → 结束 → 评价）
h = cli("demo_human", "123456", "user")
purge(h)
j = h.post(BASE + "/api/support/start", json={"mode": "human"}, timeout=15).json()
sid = j["session"]["id"]
h.post(BASE + "/api/support/send", json={"session_id": sid, "content": "导出能不能只导一列？"}, timeout=15)

kf = cli("kf_demo", "123456", "staff")
kf.post(BASE + "/api/staff/accept", json={"session_id": sid}, timeout=10)
kf.post(BASE + "/api/staff/send", json={"session_id": sid, "content": "可以，导出时选一列就行"}, timeout=10)
h.post(BASE + "/api/support/close", json={"session_id": sid}, timeout=10)

r = h.post(BASE + "/api/feedback", json={"feature": "chat", "rating": 5,
                                         "content": "客服小张回得快，问题一次解决", "tags": "答得快,解决了问题",
                                         "session_id": sid}, timeout=10)
check("会话结束后可评客服", r.status_code == 200, r.text[:160])

sv = kf.get(BASE + "/api/staff/feedback", timeout=10).json()
check("客服端看到我接待的评价", any(x.get("session_id") == sid for x in sv.get("mine", [])), sv.get("mine"))
check("客服端平均分算出来了", sv.get("mine_avg", 0) > 0, sv.get("mine_avg"))

# 客服会话卡片上带星级
board = kf.get(BASE + "/api/staff/sessions", timeout=10).json()
all_sess = board.get("waiting", []) + board.get("mine", [])
rated = [s for s in all_sess if s.get("rating")]
check("会话卡片带出星级字段", bool(rated) or board.get("mine") == [], [s.get("rating") for s in all_sess][:5])

# ---------------- 3. 管理员端：汇总 + 列表 + 标记处理
a = cli("2482991994", "123456", "admin")
d = a.get(BASE + "/api/admin/feedback?limit=200", timeout=10).json()
st = d.get("stats", {})
check("管理员看到全量评价", len(d.get("items", [])) >= 5, len(d.get("items", [])))
check("总览有平均分", st.get("avg", 0) > 0, st)
check("按功能分组非空", len(st.get("by_feature", [])) >= 4, st.get("by_feature"))
check("星级分布五档", len(st.get("dist", [])) == 5, st.get("dist"))
print("  统计：总数 %s · 平均 %s · 待处理 %s · 今日 %s" % (st.get("total"), st.get("avg"), st.get("unread"), st.get("today")))
print("  分功能：" + " / ".join("%s %.1f分 %d条" % (x["name"], x["avg"], x["count"]) for x in st.get("by_feature", [])))

fid = d["items"][0]["id"]
r = a.post(BASE + "/api/admin/feedback/%d/status" % fid, json={"status": "done"}, timeout=10)
check("标记已处理", r.status_code == 200, r.text[:140])
d2 = a.get(BASE + "/api/admin/feedback?status=done", timeout=10).json()
check("筛选「已处理」能查到", any(x["id"] == fid for x in d2.get("items", [])), [x["id"] for x in d2.get("items", [])][:5])

# 权限边界
check("普通用户进不了客服评价接口", u.get(BASE + "/api/staff/feedback", timeout=10).status_code == 403)
check("普通用户进不了管理员评价接口", u.get(BASE + "/api/admin/feedback", timeout=10).status_code == 403)
check("未登录不能提交", requests.post(BASE + "/api/feedback", json={"feature": "other", "rating": 5},
                                   timeout=10).status_code == 401)

print("\n结果：通过 %d 项，失败 %d 项" % (len(ok), len(bad)))
if bad:
    print("失败项：" + "；".join(bad))
sys.exit(1 if bad else 0)
