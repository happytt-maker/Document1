# -*- coding: utf-8 -*-
"""room.py —— 用户端多人聊天室（一个公共大群）

两张表：
  room_messages  每一条群消息（user 发的 / 系统提示）
  room_online    谁在群里（心跳表，超过 ONLINE_SECONDS 没动静就算离线）

为什么不建「房间 / 群成员」表：
  赛题里聊天室是「让用户互相交流」的附属功能，只有一个大群，
  多一个 talks 表就是多余的一层；以后真要分群，加 room_id 列即可。

为什么还是轮询：
  和客服模块保持一致（一个 uvicorn + 线程池，不占长连接）。群里人多时
  2.5 秒一轮、只取 after_id 之后的新消息，接口返回体只有几百字节。

防止刷屏：
  同一个人两次发送至少隔 MIN_GAP 秒；单条上限 MAX_CONTENT 字。
  撤回只允许 5 分钟内撤自己的，别人看到「已撤回」占位。
"""

import sys
import threading
import time
from datetime import datetime, timedelta
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

import db  # noqa: E402

MAX_CONTENT = 500            # 单条上限（字）
MIN_GAP = 1.0                # 同一个人两条消息的最小间隔（秒）
ONLINE_SECONDS = 45          # 心跳 5 秒一次；45 秒没心跳算掉线
HISTORY_DEFAULT = 60         # 首次进群拉多少条
HISTORY_MAX = 200
RECALL_SECONDS = 300         # 发出 5 分钟内可以撤回
ONLINE_LIST_MAX = 60

_LAST_SEND = {}              # user_id → 上次发送的单调时钟（防刷屏用，进程内）
_LAST_LOCK = threading.Lock()


def _now() -> datetime:
    return datetime.now().replace(microsecond=0)


def _ts(dt: datetime | None = None) -> str:
    return (dt or _now()).strftime("%Y-%m-%d %H:%M:%S")


def ensure_tables() -> None:
    """建聊天室两张表（幂等）。"""
    conn = db.get_conn(with_db=True)
    try:
        with conn.cursor() as cur:
            cur.execute("""
                CREATE TABLE IF NOT EXISTS room_messages (
                    id          BIGINT AUTO_INCREMENT PRIMARY KEY,
                    user_id     INT          NOT NULL COMMENT '发言人',
                    user_name   VARCHAR(50)  NOT NULL COMMENT '发言人昵称（冗余快照）',
                    avatar      VARCHAR(120) NULL     COMMENT '头像文件名快照',
                    content     TEXT         NOT NULL,
                    kind        VARCHAR(10)  NOT NULL DEFAULT 'text' COMMENT 'text / system',
                    deleted     TINYINT(1)   NOT NULL DEFAULT 0 COMMENT '1 = 已撤回',
                    created_at  DATETIME     NOT NULL,
                    INDEX idx_room (id)
                ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COMMENT='A23 聊天室消息'
            """)
            cur.execute("""
                CREATE TABLE IF NOT EXISTS room_online (
                    user_id     INT PRIMARY KEY,
                    user_name   VARCHAR(50)  NOT NULL,
                    avatar      VARCHAR(120) NULL,
                    joined_at   DATETIME     NOT NULL COMMENT '这次进群时间',
                    last_seen   DATETIME     NOT NULL COMMENT '心跳时间',
                    INDEX idx_seen (last_seen)
                ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COMMENT='A23 聊天室在线表'
            """)
        conn.commit()
    finally:
        conn.close()


# ---------------------------------------------------------------- 小工具

def _disp_name(user: dict) -> str:
    return (user.get("display_name") or user.get("username") or "用户")[:50]


def _avatar_url(name: str | None) -> str:
    return ("/api/avatar/" + name) if name else ""


def _avatar_file(user: dict) -> str | None:
    """user['avatar'] 是 /api/avatar/xxx.png 形式的 URL，这里反解出文件名存快照。"""
    a = user.get("avatar") or ""
    pre = "/api/avatar/"
    return a[len(pre):] if a.startswith(pre) and a[len(pre):] else None


def _row_msg(r) -> dict:
    return {
        "id": r[0], "user_id": r[1], "user_name": r[2] or "用户",
        "avatar": _avatar_url(r[3]), "content": r[4], "kind": r[5] or "text",
        "deleted": bool(r[6]), "created_at": str(r[7] or ""),
    }


# 昵称 / 头像每次读 users 表里最新的；用户改了资料，历史消息也跟着变
_MSG_SELECT = ("SELECT m.id, m.user_id, "
               "COALESCE(NULLIF(u.display_name, ''), m.user_name), "
               "COALESCE(NULLIF(u.avatar, ''), m.avatar), "
               "m.content, m.kind, m.deleted, m.created_at "
               "FROM room_messages m LEFT JOIN users u ON u.id = m.user_id ")


# ---------------------------------------------------------------- 在线状态

def touch(user: dict) -> None:
    """进群 / 轮询心跳：把自己标成在线（UPSERT）。"""
    conn = db.get_conn(with_db=True)
    try:
        with conn.cursor() as cur:
            now = _ts()
            cur.execute(
                "INSERT INTO room_online (user_id, user_name, avatar, joined_at, last_seen) "
                "VALUES (%s, %s, %s, %s, %s) ON DUPLICATE KEY UPDATE "
                "user_name = VALUES(user_name), avatar = VALUES(avatar), last_seen = VALUES(last_seen)",
                (int(user["id"]), _disp_name(user), _avatar_file(user), now, now))
    finally:
        conn.close()


def leave(user_id: int) -> None:
    """收起面板 = 离开（主动退出，不用等 45 秒过期）。"""
    conn = db.get_conn(with_db=True)
    try:
        with conn.cursor() as cur:
            cur.execute("DELETE FROM room_online WHERE user_id = %s", (int(user_id),))
    finally:
        conn.close()


def _online(cur, me_id: int) -> list[dict]:
    since = _ts(_now() - timedelta(seconds=ONLINE_SECONDS))
    cur.execute("SELECT user_id, user_name, avatar, joined_at, last_seen FROM room_online "
                "WHERE last_seen >= %s ORDER BY joined_at LIMIT %s",
                (since, ONLINE_LIST_MAX))
    out = []
    for r in cur.fetchall():
        out.append({"id": r[0], "name": r[1], "avatar": _avatar_url(r[2]),
                    "me": int(r[0]) == int(me_id)})
    return out


def _sweep(cur) -> None:
    """顺手把过期的人清掉，表不会越堆越大。"""
    since = _ts(_now() - timedelta(seconds=max(300, ONLINE_SECONDS * 4)))
    cur.execute("DELETE FROM room_online WHERE last_seen < %s", (since,))


# ---------------------------------------------------------------- 发消息 / 读消息

def _too_fast(uid: int) -> float:
    """防刷屏的精确计时：返回还要等多少秒（0 = 可以发）。

    为什么不只看数据库里上一条的时间：created_at 是 DATETIME（秒），
    S.9 秒发一条、S+1.1 秒再发，库里算出来差 1 秒，其实只隔了 0.2 秒。
    所以以进程内单调时钟为准，数据库那层只当兜底。
    """
    now = time.monotonic()
    with _LAST_LOCK:
        last = _LAST_SEND.get(uid)
        if last is not None and (now - last) < MIN_GAP:
            return MIN_GAP - (now - last)
        _LAST_SEND[uid] = now
        return 0.0


def send(user: dict, content: str) -> dict:
    """发一条群消息。太快 / 太长抛 ValueError（前端把话显示给用户）。"""
    text = (content or "").strip()
    if not text:
        raise ValueError("说点什么再发")
    if len(text) > MAX_CONTENT:
        raise ValueError("一条最多 %d 字，分两条发吧" % MAX_CONTENT)
    uid = int(user["id"])
    wait = _too_fast(uid)
    if wait > 0:
        raise ValueError("发得太快了，%.1f 秒后再试" % wait)
    conn = db.get_conn(with_db=True)
    try:
        with conn.cursor() as cur:
            cur.execute("SELECT created_at FROM room_messages WHERE user_id = %s "
                        "ORDER BY id DESC LIMIT 1", (int(user["id"]),))
            last = cur.fetchone()
            # 兜底：服务重启后进程内计时没了，同一秒内连发仍从这里挡住
            if last and last[0]:
                gap = (_now() - last[0]).total_seconds()
                if gap <= 0:
                    raise ValueError("发得太快了，%.1f 秒后再试" % MIN_GAP)
            now = _ts()
            cur.execute("INSERT INTO room_messages (user_id, user_name, avatar, content, kind, "
                        "created_at) VALUES (%s, %s, %s, %s, 'text', %s)",
                        (int(user["id"]), _disp_name(user), _avatar_file(user),
                         text, now))
            mid = cur.lastrowid
            # 发言算一次心跳
            cur.execute(
                "INSERT INTO room_online (user_id, user_name, avatar, joined_at, last_seen) "
                "VALUES (%s, %s, %s, %s, %s) ON DUPLICATE KEY UPDATE last_seen = VALUES(last_seen)",
                (int(user["id"]), _disp_name(user), _avatar_file(user), now, now))
            return {"id": mid, "user_id": int(user["id"]), "user_name": _disp_name(user),
                    "avatar": _avatar_url(_avatar_file(user)), "content": text,
                    "kind": "text", "deleted": False, "created_at": now}
    finally:
        conn.close()


def _rev(cur) -> int:
    """版本号 = 已撤回的条数（只增不减）。撤回不产生新 id，靠它通知客户端重拉。"""
    cur.execute("SELECT COUNT(*) FROM room_messages WHERE deleted = 1")
    return int(cur.fetchone()[0])


def poll(user: dict, after_id: int = 0, alive: bool = True) -> dict:
    """轮询：after_id 之后的新消息 + 在线名单 + 版本号。

    after_id <= 0 表示「刚进群」，返回最近 HISTORY_DEFAULT 条。
    alive=False 表示面板没开（只是想知道有没有新消息），不更新在线状态。
    """
    if alive:
        touch(user)                  # 先落心跳，名单里一定有自己的名字
    conn = db.get_conn(with_db=True)
    try:
        with conn.cursor() as cur:
            if after_id and after_id > 0:
                cur.execute(_MSG_SELECT + "WHERE m.id > %s ORDER BY m.id LIMIT %s",
                            (int(after_id), HISTORY_MAX))
                rows = cur.fetchall()
                msgs = [_row_msg(r) for r in rows]
            else:
                cur.execute(_MSG_SELECT + "ORDER BY m.id DESC LIMIT %s", (HISTORY_DEFAULT,))
                msgs = [_row_msg(r) for r in reversed(cur.fetchall())]
            online = _online(cur, int(user["id"]))
            rev = _rev(cur)
            cur.execute("SELECT COUNT(*) FROM room_messages")
            total = int(cur.fetchone()[0])
            if alive:
                _sweep(cur)
        return {"ok": True, "messages": msgs, "online": online, "online_count": len(online),
                "total": total, "rev": rev, "min_gap": MIN_GAP,
                "me_id": int(user["id"]), "server_time": _ts()}
    finally:
        conn.close()


def recall(user: dict, message_id: int) -> dict:
    """撤回自己的消息（5 分钟内）。别人只会看到「已撤回」占位。"""
    conn = db.get_conn(with_db=True)
    try:
        with conn.cursor() as cur:
            cur.execute("SELECT user_id, created_at, deleted FROM room_messages WHERE id = %s",
                        (int(message_id),))
            r = cur.fetchone()
            if not r:
                raise ValueError("这条消息已经不在了")
            if int(r[0]) != int(user["id"]):
                raise ValueError("只能撤回自己发的消息")
            if r[2]:
                return {"ok": True, "message_id": int(message_id), "rev": _rev(cur)}
            if (_now() - r[1]).total_seconds() > RECALL_SECONDS:
                raise ValueError("超过 %d 分钟了，撤不回来了" % (RECALL_SECONDS // 60))
            cur.execute("UPDATE room_messages SET deleted = 1 WHERE id = %s", (int(message_id),))
            return {"ok": True, "message_id": int(message_id), "rev": _rev(cur)}
    finally:
        conn.close()


def stats() -> dict:
    """管理端 / 自检用：群消息数、今天发言人数、当前在线。"""
    conn = db.get_conn(with_db=True)
    try:
        with conn.cursor() as cur:
            today = _now().strftime("%Y-%m-%d")
            cur.execute("SELECT COUNT(*) FROM room_messages")
            total = int(cur.fetchone()[0])
            cur.execute("SELECT COUNT(DISTINCT user_id) FROM room_messages WHERE DATE(created_at) = %s",
                        (today,))
            today_users = int(cur.fetchone()[0])
            cur.execute("SELECT COUNT(DISTINCT user_id) FROM room_messages")
            all_users = int(cur.fetchone()[0])
            since = _ts(_now() - timedelta(seconds=ONLINE_SECONDS))
            cur.execute("SELECT COUNT(*) FROM room_online WHERE last_seen >= %s", (since,))
            online = int(cur.fetchone()[0])
            return {"messages": total, "users": all_users, "users_today": today_users,
                    "online": online, "rev": _rev(cur)}
    finally:
        conn.close()


if __name__ == "__main__":
    ensure_tables()
    print("【聊天室】room_messages / room_online 已就绪")
    print(stats())
