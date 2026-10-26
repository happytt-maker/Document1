# -*- coding: utf-8 -*-
"""清掉验收过程中产生的评价，再种一批像样的示范评价（方便演示/截图）。"""
import sys
sys.path.insert(0, r"C:\Users\覃敬涛\Desktop\研一项目\agent")

import requests
import db

conn = db.get_conn(with_db=True)
try:
    with conn.cursor() as cur:
        cur.execute("DELETE FROM feedback")
        n = cur.rowcount
    conn.commit()
    print("已清空 %d 条测试评价" % n)
finally:
    conn.close()

BASE = "http://127.0.0.1:8000"


def cli(u, p):
    s = requests.Session(); s.trust_env = False
    s.post(BASE + "/api/auth/login", json={"username": u, "password": p, "portal": "user"}, timeout=10)
    return s


ai = cli("demo_ai", "123456")
hu = cli("demo_human", "123456")

SEED = [
    (ai, "doc_edit", 5, "把设备采购合同的标题居中加粗，一次就成，格式没乱", "改得对,速度快"),
    (ai, "doc_edit", 4, "改整段格式很快，就是「第二段」这种说法有时候要说得更具体", "位置不准"),
    (ai, "extract", 5, "合同金额、甲乙方都抽对了，出处点一下就能回到原文", "抽得全,数字对"),
    (ai, "extract", 4, "实体抽得挺全，员工信息表里有一个工号漏了", "漏了字段"),
    (ai, "fill", 5, "五张模板填完只花了两分钟，下载下来直接能用", "填得对,省时间"),
    (hu, "chat", 5, "客服回得快，还教我怎么让结果只导一列", "答得快,解决了问题"),
    (hu, "voice", 4, "语音输入比打字快，长的句子偶尔要改一两个字", "识别准"),
    (ai, "image", 5, "图片拖进对话框就选中了，说一句「插到标题下面」，位置和宽度都对", "位置准,插图很快"),
    (hu, "other", 0, "希望操作历史能按文档名筛一下，现在翻起来有点长", "希望加功能"),
]
for s, feat, r, txt, tags in SEED:
    resp = s.post(BASE + "/api/feedback", json={"feature": feat, "rating": r,
                                                "content": txt, "tags": tags}, timeout=10)
    print(resp.status_code, feat, (resp.json().get("item") or {}).get("id"))

# 把那条客服评价挂到 kf_demo 真接待过的一场会话上，客服端「我收到的评价」才有数
conn = db.get_conn(with_db=True)
try:
    with conn.cursor() as cur:
        cur.execute("SELECT id FROM support_sessions WHERE staff_id = "
                    "(SELECT id FROM users WHERE username='kf_demo') AND status='closed' "
                    "ORDER BY id DESC LIMIT 1")
        row = cur.fetchone()
        if row:
            cur.execute("UPDATE feedback SET session_id = %s "
                        "WHERE feature = 'chat' AND session_id IS NULL", (row[0],))
            print("客服评价已挂到会话 #%s（%d 条）" % (row[0], cur.rowcount))
    conn.commit()
finally:
    conn.close()
