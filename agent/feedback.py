# -*- coding: utf-8 -*-
"""
feedback.py —— 用户评价 / 意见反馈（用户端提交 → 客服端 & 管理员端可见）

设计要点：
- 一张表 feedback 存所有评价：功能点 + 星级（1–5，可为 0 表示只写意见）+ 文字 + 快捷标签
- 功能点固定枚举（FEATURES），前端按功能挂入口，统计也按功能分组
- 只有登录用户能提交；客服端只看与本坐席相关的会话评价 + 全站最新，管理员端看全量 + 统计
- 状态：new 待看 / read 已看 / done 已处理（管理员改）
- 数据库不可用时给人话提示，不抛裸异常给前端
"""

from datetime import datetime

import db

# 功能点：键是接口里的 feature，值是给人看的名字
FEATURES = {
    "doc_edit": "文档操作",
    "extract": "识别入库",
    "fill": "表格填写",
    "voice": "语音输入",
    "image": "图片插入",
    "chat": "客服咨询",
    "other": "意见反馈",
}

STATUS_TEXT = {"new": "待处理", "read": "已查看", "done": "已处理"}
MAX_CONTENT = 1000


def _now() -> datetime:
    return datetime.now().replace(microsecond=0)


def _ts(dt: datetime | None = None) -> str:
    return (dt or _now()).strftime("%Y-%m-%d %H:%M:%S")


def ensure_table() -> None:
    """建评价表（幂等）。"""
    conn = db.get_conn(with_db=True)
    try:
        with conn.cursor() as cur:
            cur.execute("""
                CREATE TABLE IF NOT EXISTS feedback (
                    id          INT AUTO_INCREMENT PRIMARY KEY,
                    user_id     INT          NOT NULL COMMENT '提交人',
                    user_name   VARCHAR(50)  NULL     COMMENT '提交人显示名（冗余，列表直接显示）',
                    feature     VARCHAR(20)  NOT NULL COMMENT 'doc_edit/extract/fill/voice/chat/other',
                    rating      TINYINT      NOT NULL DEFAULT 0 COMMENT '1-5 星；0=只写了意见没打分',
                    content     VARCHAR(1000) NULL    COMMENT '意见正文',
                    tags        VARCHAR(200)  NULL    COMMENT '快捷标签，逗号分隔',
                    session_id  INT          NULL     COMMENT '评的是哪次客服会话（可空）',
                    status      VARCHAR(10)  NOT NULL DEFAULT 'new' COMMENT 'new/read/done',
                    handled_by  INT          NULL     COMMENT '处理人',
                    handled_at  DATETIME     NULL     COMMENT '处理时间',
                    created_at  DATETIME     NOT NULL COMMENT '提交时间',
                    KEY idx_fb_time (created_at),
                    KEY idx_fb_feature (feature),
                    KEY idx_fb_status (status)
                ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
                COMMENT='用户评价与意见反馈'
            """)
        conn.commit()
    finally:
        conn.close()


# ---------------------------------------------------------------- 写入

def add(user: dict, feature: str, rating: int = 0, content: str = "",
        session_id: int | None = None, tags: str = "") -> dict:
    """提交一条评价：星级和文字至少要有一样，否则等于空提交。"""
    feature = (feature or "").strip()
    if feature not in FEATURES:
        raise ValueError("功能点不对：%s" % feature)
    try:
        rating = int(rating or 0)
    except (TypeError, ValueError):
        raise ValueError("星级要是 1–5 的整数")
    if rating < 0 or rating > 5:
        raise ValueError("星级要在 1–5 之间")
    content = (content or "").strip()[:MAX_CONTENT]
    tags = (tags or "").strip()[:200]
    if rating == 0 and not content:
        raise ValueError("至少给个星级，或者写两句意见")

    uid = int(user["id"])
    uname = user.get("display_name") or user.get("username") or ""
    conn = db.get_conn(with_db=True)
    try:
        with conn.cursor() as cur:
            cur.execute(
                "INSERT INTO feedback (user_id, user_name, feature, rating, content, tags, "
                "session_id, status, created_at) VALUES (%s, %s, %s, %s, %s, %s, %s, 'new', %s)",
                (uid, uname, feature, int(rating or 0), content, tags,
                 int(session_id) if session_id else None, _ts()))
            fid = cur.lastrowid
        conn.commit()
        return {"id": fid, "feature": feature, "feature_text": FEATURES[feature],
                "rating": rating, "content": content, "tags": tags,
                "created_at": _ts(), "status": "new"}
    finally:
        conn.close()


def set_status(fid: int, status: str, staff_id: int | None = None) -> dict:
    """管理员标记状态：new / read / done。"""
    if status not in STATUS_TEXT:
        raise ValueError("状态只能是 new / read / done")
    conn = db.get_conn(with_db=True)
    try:
        with conn.cursor() as cur:
            cur.execute("UPDATE feedback SET status = %s, handled_by = %s, handled_at = %s "
                        "WHERE id = %s", (status, staff_id, _ts(), int(fid)))
            if cur.rowcount == 0:
                raise ValueError("这条评价不存在")
        conn.commit()
        return {"id": int(fid), "status": status}
    finally:
        conn.close()


# ---------------------------------------------------------------- 查询

_COLS = ("id, user_id, user_name, feature, rating, content, tags, session_id, "
         "status, created_at")


def _row(r) -> dict:
    return {
        "id": r[0], "user_id": r[1], "user_name": r[2] or ("用户 #%s" % r[1]),
        "feature": r[3], "feature_text": FEATURES.get(r[3], r[3]),
        "rating": int(r[4] or 0), "content": r[5] or "", "tags": r[6] or "",
        "session_id": r[7], "status": r[8] or "new",
        "status_text": STATUS_TEXT.get(r[8] or "new", r[8]),
        "created_at": str(r[9] or ""),
    }


def mine(user_id: int, limit: int = 20) -> list[dict]:
    conn = db.get_conn(with_db=True)
    try:
        with conn.cursor() as cur:
            cur.execute("SELECT " + _COLS + " FROM feedback WHERE user_id = %s "
                        "ORDER BY id DESC LIMIT %s", (int(user_id), int(limit)))
            return [_row(r) for r in cur.fetchall()]
    finally:
        conn.close()


def list_all(feature: str = "", status: str = "", limit: int = 100) -> list[dict]:
    sql = "SELECT " + _COLS + " FROM feedback WHERE 1=1"
    args = []
    if feature in FEATURES:
        sql += " AND feature = %s"
        args.append(feature)
    if status in STATUS_TEXT:
        sql += " AND status = %s"
        args.append(status)
    sql += " ORDER BY id DESC LIMIT %s"
    args.append(int(limit))
    conn = db.get_conn(with_db=True)
    try:
        with conn.cursor() as cur:
            cur.execute(sql, tuple(args))
            return [_row(r) for r in cur.fetchall()]
    finally:
        conn.close()


def stats() -> dict:
    """总览：条数 / 平均分 / 今日新增 / 待处理 / 按功能 / 星级分布。"""
    conn = db.get_conn(with_db=True)
    try:
        with conn.cursor() as cur:
            cur.execute("SELECT COUNT(*), "
                        "SUM(CASE WHEN rating > 0 THEN rating END), "
                        "SUM(CASE WHEN rating > 0 THEN 1 ELSE 0 END), "
                        "SUM(CASE WHEN status = 'new' THEN 1 ELSE 0 END), "
                        "SUM(CASE WHEN DATE(created_at) = CURDATE() THEN 1 ELSE 0 END) "
                        "FROM feedback")
            total, rsum, rcnt, unread, today = cur.fetchone()
            avg = round(float(rsum) / float(rcnt), 2) if rcnt else 0.0

            cur.execute("SELECT feature, COUNT(*), "
                        "AVG(CASE WHEN rating > 0 THEN rating END) "
                        "FROM feedback GROUP BY feature ORDER BY COUNT(*) DESC")
            by_feature = [{"feature": f, "name": FEATURES.get(f, f), "count": int(c),
                           "avg": round(float(a), 2) if a is not None else 0.0}
                          for f, c, a in cur.fetchall()]

            cur.execute("SELECT rating, COUNT(*) FROM feedback "
                        "WHERE rating > 0 GROUP BY rating ORDER BY rating")
            dist = {int(k): int(v) for k, v in cur.fetchall()}
        return {
            "total": int(total or 0), "avg": avg, "rated": int(rcnt or 0),
            "unread": int(unread or 0), "today": int(today or 0),
            "by_feature": by_feature,
            "dist": [{"star": s, "count": dist.get(s, 0)} for s in (5, 4, 3, 2, 1)],
        }
    finally:
        conn.close()


def staff_view(staff_id: int, limit: int = 40) -> dict:
    """客服端：跟本坐席相关的会话评价 + 全站最新评价（参考）。"""
    conn = db.get_conn(with_db=True)
    try:
        with conn.cursor() as cur:
            cur.execute(
                "SELECT " + ",".join("f." + c.strip() for c in _COLS.split(",")) +
                " FROM feedback f JOIN support_sessions s ON s.id = f.session_id "
                "WHERE s.staff_id = %s ORDER BY f.id DESC LIMIT %s",
                (int(staff_id), int(limit)))
            mine_rows = [_row(r) for r in cur.fetchall()]
            cur.execute("SELECT " + _COLS + " FROM feedback "
                        "WHERE feature = 'chat' ORDER BY id DESC LIMIT %s", (int(limit),))
            mine_ids = {m["id"] for m in mine_rows}
            chats = [_row(r) for r in cur.fetchall()]
            chats = [c for c in chats if c["id"] not in mine_ids]   # 我接待的那几条别重复列一遍
        return {"ok": True, "mine": mine_rows, "chats": chats,
                "mine_avg": round(sum(m["rating"] for m in mine_rows if m["rating"]) /
                                  max(1, len([m for m in mine_rows if m["rating"]])), 2)
                if any(m["rating"] for m in mine_rows) else 0.0}
    finally:
        conn.close()
