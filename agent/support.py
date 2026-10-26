# -*- coding: utf-8 -*-
"""
support.py —— 在线客服：AI 客服 + 人工客服（用户端 ↔ 客服端）

三张表：
  support_sessions  一次咨询（mode=ai 纯机器 / mode=human 转人工；status=ai/waiting/active/closed）
  support_messages  每条消息（sender=user/ai/staff/system）
  （客服在线状态复用 users.last_seen，不额外建表）

消息怎么走：
  用户发一句 → 落 support_messages → 若是 AI 会话，同步调大模型生成回复也落库；
  前端每 1.5 秒带 after_id 轮询，只取 id 更大的消息，所以不会重复、也不会丢。

为什么用轮询而不用 WebSocket：
  这套服务是「一个 uvicorn 进程 + 线程池」的简单形态，轮询两行代码就够，
  断了自动重连，不占长连接；会话量是「一个班次几十个」，1.5 秒延迟完全够用。
"""

import json
import sys
from datetime import datetime, timedelta
from pathlib import Path
from urllib.parse import urlparse

sys.path.insert(0, str(Path(__file__).resolve().parent))

import db  # noqa: E402

try:
    import requests  # noqa: E402
except ImportError:  # pragma: no cover
    requests = None

from llm_config import chat as _llm_chat, MODEL  # noqa: E402

ONLINE_SECONDS = 45          # 客服心跳 15 秒一次；45 秒内算在线
MAX_CONTENT = 2000           # 单条消息字数上限
HISTORY_FOR_AI = 8           # 给大模型带多少条上下文

# AI 客服的人设与事实边界：只许照着这些说，避免它编功能
AI_SYSTEM = """你是「智能文档助手」的在线客服，回答用户关于这套系统的问题。

【产品事实 · 只能按这里回答，不许编】
- 能处理的文件：Word(.docx)、Excel(.xlsx)、Markdown(.md/.markdown)、纯文本(.txt)
- 三块能力：
  1) 自然语言改文档：居中、加粗、字号、字体、颜色、行距、替换文字、插图片
  2) 识别提取：从文档里抽实体 / 字段 / 文档画像，每条都带原文出处，可入库查询
  3) 表格填写：读懂模板 → 找数据 → 填入 → 自检，导出新的 Excel
- 用法：左边导入文档，在输入框说一句话（例如「把标题居中」）；支持语音输入（Edge 浏览器）；
  图片可以直接拖进对话框，再说插到哪里
- 结果文件保存在 agent\\输出\\ 目录，原文件不会被覆盖
- 常见问题：识别不到文档 → 说明是哪份文档；格式没生效 → 说清位置（标题 / 第几段 / 哪一列）
- 目前是内部演示系统，免费使用

【回答要求】
- 中文口语，简洁：最多 3 句话、120 字以内，不要 Markdown 标题、不要长篇清单
- 只答和本系统相关的问题；无关问题一句话带过，并引导回文档操作
- 不知道就直说「这个我不太确定，可以点上面的『人工客服』问同事」，不要瞎猜
"""

# 常见问题：AI 客服面板里列成 1–6，用户只回一个数字就立刻答（不走大模型，0 延迟）
# 内容都按产品事实写，改功能时记得同步改这里
AI_FAQ = [
    {"no": 1, "label": "支持哪些格式", "q": "支持哪些文件格式？",
     "a": "支持 Word（.docx）、Excel（.xlsx）、Markdown（.md）和纯文本（.txt）。"
          "图片可以直接拖进对话框，再说一句插到哪里。"},
    {"no": 2, "label": "怎么改格式", "q": "怎么改文档格式？",
     "a": "先在左边「导入文档」，再在输入框说一句就行，比如「把标题居中」「把第二段加粗、字号改成 14」"
          "「全文行距 1.5 倍」。原文件不会改，结果另存一份。"},
    {"no": 3, "label": "怎么插图片", "q": "怎么插入图片？",
     "a": "把图片直接拖进对话框，松手就选中；再说「插到标题下面」或「放到文末」。"
          "不写位置直接回车，默认放在文末。"},
    {"no": 4, "label": "指令没生效", "q": "识别不到文档 / 指令没生效怎么办？",
     "a": "先确认有没有先导入文档。还不行就把话说具体：是哪一份文档（例如「设备采购合同」）、"
          "改哪里（标题 / 第几段 / 哪一列）、改成什么。"},
    {"no": 5, "label": "结果在哪", "q": "结果文件保存在哪里？原文件会被改吗？",
     "a": "结果都放在 agent\\输出\\ 目录，文件名带时间后缀，原文件不动；处理完页面会给下载入口。"},
    {"no": 6, "label": "提取 / 填表", "q": "提取数据和填写表格怎么用？",
     "a": "顶栏「识别入库」选文档，自动抽实体、字段和文档画像，每条都带原文出处；"
          "「表格填写」选一份模板，它会找数、填好并自检，完成后下载 Excel 就行。"},
]

_FAQ_BY_NO = {f["no"]: f for f in AI_FAQ}
_FULL_DIGITS = str.maketrans("０１２３４５６７８９", "0123456789")

# 用户回了个不存在的编号：不调模型，直接说清范围
_FAQ_MISS = {"no": 0, "label": "", "q": "",
             "a": "我这边只有 1–6 这 6 条快捷问题，回个数字我直接答；"
                  "别的问题直接说就行，需要真人帮忙点上面的「人工客服」。"}


# ---------------------------------------------------------------- 建表

def _now() -> datetime:
    return datetime.now().replace(microsecond=0)


def _ts(dt: datetime | None = None) -> str:
    return (dt or _now()).strftime("%Y-%m-%d %H:%M:%S")


def ensure_tables() -> None:
    """建客服两张表（幂等）。"""
    conn = db.get_conn(with_db=True)
    try:
        with conn.cursor() as cur:
            cur.execute("""
                CREATE TABLE IF NOT EXISTS support_sessions (
                    id          INT AUTO_INCREMENT PRIMARY KEY,
                    user_id     INT          NOT NULL COMMENT '发起咨询的用户',
                    user_name   VARCHAR(50)  NULL     COMMENT '用户显示名（冗余，列表直接显示）',
                    mode        VARCHAR(10)  NOT NULL DEFAULT 'ai' COMMENT 'ai / human',
                    status      VARCHAR(12)  NOT NULL DEFAULT 'ai' COMMENT 'ai / waiting / active / closed',
                    staff_id    INT          NULL     COMMENT '接待客服',
                    staff_name  VARCHAR(50)  NULL     COMMENT '客服显示名',
                    created_at  DATETIME     NOT NULL,
                    accepted_at DATETIME     NULL     COMMENT '接单时间',
                    closed_at   DATETIME     NULL     COMMENT '结束时间',
                    updated_at  DATETIME     NOT NULL COMMENT '最后一条消息时间',
                    INDEX idx_user (user_id),
                    INDEX idx_status (status),
                    INDEX idx_staff (staff_id)
                ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COMMENT='A23 客服会话'
            """)
            cur.execute("""
                CREATE TABLE IF NOT EXISTS support_messages (
                    id          BIGINT AUTO_INCREMENT PRIMARY KEY,
                    session_id  INT          NOT NULL,
                    sender      VARCHAR(10)  NOT NULL COMMENT 'user / ai / staff / system',
                    sender_id   INT          NULL,
                    sender_name VARCHAR(50)  NULL,
                    content     TEXT         NOT NULL,
                    created_at  DATETIME     NOT NULL,
                    INDEX idx_sess (session_id, id)
                ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COMMENT='A23 客服消息'
            """)
        conn.commit()
    finally:
        conn.close()


# ---------------------------------------------------------------- 小工具

def _row_session(r) -> dict:
    return {
        "id": r[0], "user_id": r[1], "user_name": r[2], "mode": r[3], "status": r[4],
        "staff_id": r[5], "staff_name": r[6],
        "created_at": str(r[7] or ""), "accepted_at": str(r[8] or ""),
        "closed_at": str(r[9] or ""), "updated_at": str(r[10] or ""),
    }


_SESS_COLS = ("id, user_id, user_name, mode, status, staff_id, staff_name, "
              "created_at, accepted_at, closed_at, updated_at")


def _row_msg(r) -> dict:
    return {
        "id": r[0], "session_id": r[1], "sender": r[2], "sender_id": r[3],
        "sender_name": r[4] or "", "content": r[5], "created_at": str(r[6] or ""),
    }


def _get_session(cur, sid: int) -> dict | None:
    cur.execute("SELECT " + _SESS_COLS + " FROM support_sessions WHERE id = %s", (int(sid),))
    r = cur.fetchone()
    return _row_session(r) if r else None


def _add_msg(cur, sid: int, sender: str, content: str,
             sender_id: int | None = None, sender_name: str = "") -> dict:
    content = (content or "").strip()
    if not content:
        raise ValueError("消息不能为空")
    if len(content) > MAX_CONTENT:
        raise ValueError("消息太长，最多 %d 字" % MAX_CONTENT)
    now = _ts()
    cur.execute("INSERT INTO support_messages (session_id, sender, sender_id, sender_name, "
                "content, created_at) VALUES (%s, %s, %s, %s, %s, %s)",
                (int(sid), sender, sender_id, sender_name, content, now))
    mid = cur.lastrowid
    cur.execute("UPDATE support_sessions SET updated_at = %s WHERE id = %s", (now, int(sid)))
    return {"id": mid, "session_id": int(sid), "sender": sender, "sender_id": sender_id,
            "sender_name": sender_name, "content": content, "created_at": now}


# ---------------------------------------------------------------- 用户侧

def my_session(user_id: int) -> dict | None:
    """当前进行中的咨询（没有返回 None）。一个用户同时只留一个。"""
    conn = db.get_conn(with_db=True)
    try:
        with conn.cursor() as cur:
            cur.execute("SELECT " + _SESS_COLS + " FROM support_sessions "
                        "WHERE user_id = %s AND status <> 'closed' "
                        "ORDER BY id DESC LIMIT 1", (int(user_id),))
            r = cur.fetchone()
            return _row_session(r) if r else None
    finally:
        conn.close()


def start(user: dict, mode: str = "ai") -> dict:
    """开一次咨询 / 转人工。已有进行中的会话就复用（AI 转人工则原地升级）。"""
    mode = "human" if mode == "human" else "ai"
    uid, uname = int(user["id"]), user["display_name"] or user["username"]
    conn = db.get_conn(with_db=True)
    try:
        with conn.cursor() as cur:
            cur.execute("SELECT " + _SESS_COLS + " FROM support_sessions "
                        "WHERE user_id = %s AND status <> 'closed' "
                        "ORDER BY id DESC LIMIT 1", (uid,))
            r = cur.fetchone()
            if r:
                sess = _row_session(r)
                if mode == "human" and sess["mode"] == "ai":
                    cur.execute("UPDATE support_sessions SET mode = 'human', status = 'waiting', "
                                "updated_at = %s WHERE id = %s", (_ts(), sess["id"]))
                    ahead = _queue_ahead(cur, sess["id"])
                    _add_msg(cur, sess["id"], "system", _wait_text(ahead))
                    sess["mode"], sess["status"] = "human", "waiting"
                return sess

            now = _ts()
            cur.execute("INSERT INTO support_sessions (user_id, user_name, mode, status, "
                        "created_at, updated_at) VALUES (%s, %s, %s, %s, %s, %s)",
                        (uid, uname, mode, "ai" if mode == "ai" else "waiting", now, now))
            sid = cur.lastrowid
            if mode == "ai":
                _add_msg(cur, sid, "system",
                         "我是小橘，可以问我系统怎么用、支持哪些格式、报错怎么办；"
                         "也可以直接回数字 1–6 看常见问题。需要真人帮忙就点上面的「人工客服」。")
            else:
                _add_msg(cur, sid, "system", _wait_text(_queue_ahead(cur, sid)) + "这段时间可以先留言，客服接入后能看到。")
            return _get_session(cur, sid)
    finally:
        conn.close()


def _queue_ahead(cur, sid: int) -> int:
    """我前面还有几个在排队（不含自己）。"""
    cur.execute("SELECT COUNT(*) FROM support_sessions WHERE status = 'waiting' AND id < %s",
                (int(sid),))
    return int(cur.fetchone()[0])


def _wait_text(ahead: int) -> str:
    """排队提示：自己是第一个时不说「前面还有 0 位」。"""
    if ahead <= 0:
        return "正在为你转接人工客服，请稍候…"
    return "正在为你转接人工客服，前面还有 %d 位，请稍候…" % ahead


def user_send(user: dict, session_id: int, content: str) -> dict:
    """用户发消息。AI 会话当场生成回复；人工会话只落库等客服来取。"""
    uid = int(user["id"])
    conn = db.get_conn(with_db=True)
    try:
        with conn.cursor() as cur:
            sess = _get_session(cur, session_id)
            if not sess or sess["user_id"] != uid:
                raise ValueError("会话不存在")
            if sess["status"] == "closed":
                raise ValueError("这次咨询已经结束了，重新发起一个吧")
            msg = _add_msg(cur, session_id, "user", content, uid,
                           user["display_name"] or user["username"])
            mode = sess["mode"]
            # 排队时催一下：告诉用户还在等
            waiting_ahead = _queue_ahead(cur, session_id) if sess["status"] == "waiting" else 0
        result = {"ok": True, "message": msg, "session": sess, "ai": None}
        if mode == "ai":
            hit = faq_hit(content)
            # 只回了个数字（或原样问了常见问题）→ 直接给预设答案，不调模型，秒回
            result["ai"] = _faq_reply(session_id, hit) if hit else ai_reply(user, session_id, content)
        elif sess["status"] == "waiting":
            # 客服还没接入也能发，消息先落下，客服接进来就能看到
            result["notice"] = ("留言已发给客服，接入后会回复你" if waiting_ahead <= 0
                                else "留言已收到，前面还有 %d 位，客服接入后会回复你" % waiting_ahead)
        return result
    finally:
        conn.close()


# ---------------------------------------------------------------- AI 客服

def faq_list() -> list[dict]:
    """给前端渲染快捷指令（只给编号和短标题，答案点下去才回来）。"""
    return [{"no": f["no"], "label": f["label"], "q": f["q"]} for f in AI_FAQ]


def faq_hit(content: str) -> dict | None:
    """用户只回了个数字（1–6，允许全角 / #1 / 「1.」）就命中；原样问句子也算。"""
    t = (content or "").strip().translate(_FULL_DIGITS)
    if not t:
        return None
    t = t.lstrip("#＃").strip().rstrip("。.、,，!！?？ ").strip()
    if t.isdigit():
        return _FAQ_BY_NO.get(int(t)) or _FAQ_MISS
    for f in AI_FAQ:
        if t == f["q"].rstrip("？?"):
            return f
    return None


def _faq_reply(session_id: int, faq: dict) -> dict:
    """落一条预设答案（跟大模型回复同一条链路，前端不用区分）。"""
    conn = db.get_conn(with_db=True)
    try:
        with conn.cursor() as cur:
            msg = _add_msg(cur, session_id, "ai", faq["a"], None, "智能客服")
            return {"message": msg, "error": "", "faq": faq["no"]}
    finally:
        conn.close()


def _history(session_id: int, limit: int = HISTORY_FOR_AI) -> list[dict]:
    conn = db.get_conn(with_db=True)
    try:
        with conn.cursor() as cur:
            cur.execute("SELECT sender, content FROM support_messages "
                        "WHERE session_id = %s AND sender IN ('user','ai','staff') "
                        "ORDER BY id DESC LIMIT %s", (int(session_id), int(limit)))
            rows = cur.fetchall()
            return [{"role": "assistant" if s in ("ai", "staff") else "user", "content": c}
                    for s, c in reversed(rows)]
    finally:
        conn.close()


def _call_model(messages: list[dict]) -> str:
    """调大模型要一段自由文本（不强制 JSON，跟 parse_intent 的解析调用不同）。"""
    text, _usage, _cost, err = _llm_chat(
        [{"role": "system", "content": AI_SYSTEM}] + messages,
        temperature=0.5, max_tokens=400, timeout=30)
    if err:
        raise RuntimeError(err)
    return (text or "").strip()


def ai_reply(user: dict, session_id: int, question: str) -> dict:
    """生成并落库一条 AI 回复。失败给兜底话术，不让用户干等。"""
    msgs = _history(session_id)
    if not msgs or msgs[-1]["content"] != question:
        msgs.append({"role": "user", "content": question})
    text, err = "", ""
    try:
        text = _call_model(msgs)
    except Exception as e:                       # 网络 / 额度 / 限流都走这里
        err = str(e)[:120]
        text = ("人工小助手这边网络有点卡，没能查到答案。"
                "你可以点上面的「人工客服」，让同事直接帮你看。")
    text = text[:MAX_CONTENT]
    conn = db.get_conn(with_db=True)
    try:
        with conn.cursor() as cur:
            msg = _add_msg(cur, session_id, "ai", text, None, "智能客服")
            return {"message": msg, "error": err}
    finally:
        conn.close()


def history(session_id: int, after_id: int = 0, limit: int = 200) -> list[dict]:
    conn = db.get_conn(with_db=True)
    try:
        with conn.cursor() as cur:
            cur.execute("SELECT id, session_id, sender, sender_id, sender_name, content, created_at "
                        "FROM support_messages WHERE session_id = %s AND id > %s "
                        "ORDER BY id LIMIT %s", (int(session_id), int(after_id), int(limit)))
            return [_row_msg(r) for r in cur.fetchall()]
    finally:
        conn.close()


def poll(session_id: int, after_id: int = 0) -> dict:
    """轮询：返回会话最新状态 + 新消息（after_id 之后的）。"""
    conn = db.get_conn(with_db=True)
    try:
        with conn.cursor() as cur:
            sess = _get_session(cur, session_id)
            if not sess:
                raise ValueError("会话不存在")
            cur.execute("SELECT COUNT(*) FROM support_sessions WHERE status = 'waiting' AND id < %s",
                        (int(session_id),))
            ahead = int(cur.fetchone()[0]) if sess["status"] == "waiting" else 0
        return {"ok": True, "session": sess, "messages": history(session_id, after_id),
                "queue_ahead": ahead}
    finally:
        conn.close()


def close_by_user(user_id: int, session_id: int) -> dict:
    conn = db.get_conn(with_db=True)
    try:
        with conn.cursor() as cur:
            sess = _get_session(cur, session_id)
            if not sess or sess["user_id"] != int(user_id):
                raise ValueError("会话不存在")
            if sess["status"] != "closed":
                _add_msg(cur, session_id, "system", "用户已结束本次咨询")
                cur.execute("UPDATE support_sessions SET status = 'closed', closed_at = %s, "
                            "updated_at = %s WHERE id = %s", (_ts(), _ts(), int(session_id)))
            return _get_session(cur, session_id)
    finally:
        conn.close()


# ---------------------------------------------------------------- 客服侧

def touch(user_id: int) -> None:
    """客服端每 15 秒心跳一次，用来判断在不在线。"""
    conn = db.get_conn(with_db=True)
    try:
        with conn.cursor() as cur:
            cur.execute("UPDATE users SET last_seen = %s WHERE id = %s", (_ts(), int(user_id)))
    finally:
        conn.close()


def staff_sessions(staff_id: int) -> dict:
    """客服工作台：待接入队列 + 我正在接待的。"""
    conn = db.get_conn(with_db=True)
    try:
        with conn.cursor() as cur:
            cur.execute("SELECT " + _SESS_COLS + " FROM support_sessions "
                        "WHERE status = 'waiting' ORDER BY id")
            waiting = [_row_session(r) for r in cur.fetchall()]
            cur.execute("SELECT " + _SESS_COLS + " FROM support_sessions "
                        "WHERE status = 'active' AND staff_id = %s ORDER BY updated_at DESC",
                        (int(staff_id),))
            mine = [_row_session(r) for r in cur.fetchall()]
            for s in waiting + mine:
                cur.execute("SELECT content, created_at FROM support_messages "
                            "WHERE session_id = %s ORDER BY id DESC LIMIT 1", (s["id"],))
                last = cur.fetchone()
                s["last_text"] = (last[0] if last else "")[:60]
                s["last_at"] = str(last[1] or "") if last else ""
                # 用户留言条数：客服还没接入时也能看出「他已经说了几句」
                cur.execute("SELECT COUNT(*) FROM support_messages "
                            "WHERE session_id = %s AND sender = 'user'", (s["id"],))
                s["msg_count"] = int(cur.fetchone()[0])
                # 用户对这次会话的评分（没评就是 0），卡片上直接显示
                try:
                    cur.execute("SELECT rating FROM feedback WHERE session_id = %s "
                                "AND rating > 0 ORDER BY id DESC LIMIT 1", (s["id"],))
                    rrow = cur.fetchone()
                    s["rating"] = int(rrow[0]) if rrow else 0
                except Exception:
                    s["rating"] = 0
        return {"ok": True, "waiting": waiting, "mine": mine}
    finally:
        conn.close()


def staff_accept(staff: dict, session_id: int) -> dict:
    """接单：waiting → active。被人抢先接走会说清楚。"""
    conn = db.get_conn(with_db=True)
    try:
        with conn.cursor() as cur:
            sess = _get_session(cur, session_id)
            if not sess:
                raise ValueError("会话不存在")
            if sess["status"] == "active" and sess["staff_id"] != int(staff["id"]):
                raise ValueError("这一单已经被 %s 接了" % (sess["staff_name"] or "其他客服"))
            if sess["status"] == "closed":
                raise ValueError("用户已经结束了这次咨询")
            if sess["status"] != "active":
                name = staff["display_name"] or staff["username"]
                cur.execute("UPDATE support_sessions SET status = 'active', staff_id = %s, "
                            "staff_name = %s, accepted_at = %s, updated_at = %s WHERE id = %s",
                            (int(staff["id"]), name, _ts(), _ts(), int(session_id)))
                _add_msg(cur, session_id, "system",
                         "客服「%s」已接入，请直接描述你的问题。" % name)
            return _get_session(cur, session_id)
    finally:
        conn.close()


def staff_send(staff: dict, session_id: int, content: str) -> dict:
    conn = db.get_conn(with_db=True)
    try:
        with conn.cursor() as cur:
            sess = _get_session(cur, session_id)
            if not sess:
                raise ValueError("会话不存在")
            if sess["staff_id"] != int(staff["id"]):
                raise ValueError("这不是你接待的会话")
            if sess["status"] != "active":
                raise ValueError("会话已结束，不能再回复")
            msg = _add_msg(cur, session_id, "staff", content, int(staff["id"]),
                           staff["display_name"] or staff["username"])
            return {"ok": True, "message": msg}
    finally:
        conn.close()


def staff_close(staff: dict, session_id: int) -> dict:
    conn = db.get_conn(with_db=True)
    try:
        with conn.cursor() as cur:
            sess = _get_session(cur, session_id)
            if not sess:
                raise ValueError("会话不存在")
            if sess["staff_id"] != int(staff["id"]):
                raise ValueError("这不是你接待的会话")
            if sess["status"] != "closed":
                _add_msg(cur, session_id, "system", "客服已结束本次咨询，感谢你的反馈")
                cur.execute("UPDATE support_sessions SET status = 'closed', closed_at = %s, "
                            "updated_at = %s WHERE id = %s", (_ts(), _ts(), int(session_id)))
            return _get_session(cur, session_id)
    finally:
        conn.close()


# ---------------------------------------------------------------- 管理端

def admin_overview() -> dict:
    conn = db.get_conn(with_db=True)
    try:
        with conn.cursor() as cur:
            def one(sql, args=()):
                cur.execute(sql, args)
                return int(cur.fetchone()[0] or 0)

            today = _now().strftime("%Y-%m-%d")
            since = _ts(_now() - timedelta(seconds=ONLINE_SECONDS))
            return {
                "users": one("SELECT COUNT(*) FROM users"),
                "users_today": one("SELECT COUNT(*) FROM users WHERE DATE(created_at) = %s", (today,)),
                "staff": one("SELECT COUNT(*) FROM users WHERE role LIKE '%%staff%%'"),
                "staff_online": one("SELECT COUNT(*) FROM users WHERE role LIKE '%%staff%%' "
                                    "AND last_seen >= %s", (since,)),
                "sessions_total": one("SELECT COUNT(*) FROM support_sessions"),
                "sessions_waiting": one("SELECT COUNT(*) FROM support_sessions WHERE status = 'waiting'"),
                "sessions_active": one("SELECT COUNT(*) FROM support_sessions WHERE status = 'active'"),
                "messages_today": one("SELECT COUNT(*) FROM support_messages WHERE DATE(created_at) = %s",
                                      (today,)),
            }
    finally:
        conn.close()


def admin_users(q: str = "", limit: int = 200) -> list[dict]:
    conn = db.get_conn(with_db=True)
    try:
        with conn.cursor() as cur:
            sql = ("SELECT id, username, display_name, avatar, role, status, created_at, last_login "
                   "FROM users")
            args = []
            if q:
                sql += " WHERE username LIKE %s OR display_name LIKE %s"
                args = ["%" + q + "%", "%" + q + "%"]
            sql += " ORDER BY id DESC LIMIT %s"
            args.append(int(limit))
            cur.execute(sql, tuple(args))
            out = []
            for r in cur.fetchall():
                out.append({
                    "id": r[0], "username": r[1], "display_name": r[2] or r[1],
                    "avatar": ("/api/avatar/" + r[3]) if r[3] else "",
                    "roles": (r[4] or "user").split(","),
                    "status": r[5] or "active",
                    "created_at": str(r[6] or ""), "last_login": str(r[7] or ""),
                })
            return out
    finally:
        conn.close()


def admin_staff() -> list[dict]:
    """客服坐席：在线状态 + 当前接待数 + 今日接待数。"""
    conn = db.get_conn(with_db=True)
    try:
        with conn.cursor() as cur:
            since = _ts(_now() - timedelta(seconds=ONLINE_SECONDS))
            today = _now().strftime("%Y-%m-%d")
            cur.execute("SELECT id, username, display_name, avatar, last_seen FROM users "
                        "WHERE role LIKE '%%staff%%' ORDER BY id")
            rows = cur.fetchall()
            out = []
            for r in rows:
                cur.execute("SELECT COUNT(*) FROM support_sessions WHERE staff_id = %s "
                            "AND status = 'active'", (r[0],))
                active = int(cur.fetchone()[0])
                cur.execute("SELECT COUNT(*) FROM support_sessions WHERE staff_id = %s "
                            "AND DATE(accepted_at) = %s", (r[0], today))
                today_n = int(cur.fetchone()[0])
                out.append({
                    "id": r[0], "username": r[1], "display_name": r[2] or r[1],
                    "avatar": ("/api/avatar/" + r[3]) if r[3] else "",
                    "online": bool(r[4]) and str(r[4]) >= since,
                    "last_seen": str(r[4] or ""),
                    "active": active, "today": today_n,
                })
            return out
    finally:
        conn.close()


def admin_sessions(status: str = "", limit: int = 100) -> list[dict]:
    conn = db.get_conn(with_db=True)
    try:
        with conn.cursor() as cur:
            sql = "SELECT " + _SESS_COLS + " FROM support_sessions"
            args = []
            if status:
                sql += " WHERE status = %s"
                args.append(status)
            sql += " ORDER BY id DESC LIMIT %s"
            args.append(int(limit))
            cur.execute(sql, tuple(args))
            out = [_row_session(r) for r in cur.fetchall()]
            for s in out:
                cur.execute("SELECT COUNT(*) FROM support_messages WHERE session_id = %s", (s["id"],))
                s["msg_count"] = int(cur.fetchone()[0])
            return out
    finally:
        conn.close()


def session_transcript(session_id: int) -> dict:
    """管理员看某次咨询的完整聊天记录。"""
    conn = db.get_conn(with_db=True)
    try:
        with conn.cursor() as cur:
            sess = _get_session(cur, session_id)
            if not sess:
                raise ValueError("会话不存在")
            return {"session": sess, "messages": history(session_id, 0, 500)}
    finally:
        conn.close()


def staff_ahead(session_id: int) -> int:
    conn = db.get_conn(with_db=True)
    try:
        with conn.cursor() as cur:
            return _queue_ahead(cur, session_id)
    finally:
        conn.close()


if __name__ == "__main__":
    ensure_tables()
    print("【客服】support_sessions / support_messages 已就绪")
