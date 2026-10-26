# -*- coding: utf-8 -*-
"""
auth.py —— 用户账号：注册 / 登录 / 登录态（MySQL）

职责：把"谁在用这套系统"落到 MySQL —— 注册写 users 表，登录发一个随机 token 写 sessions 表，
     浏览器带 cookie 回来就能认出是谁。密码只存哈希，不存明文。

设计：
- 密码：PBKDF2-HMAC-SHA256，12 万次迭代，每个用户独立随机盐
        存法 "pbkdf2_sha256$迭代次数$盐(hex)$哈希(hex)"，校验用 hmac.compare_digest 防时序侧信道
- 登录态：token = 32 字节随机数（hex 64 位），存 sessions 表，7 天过期，退出即删
- 角色：users.role 存逗号分隔的角色集合 user / staff / admin，
        登录时按「端口」校验（用户端 / 客服端 / 管理员端），一个账号可以同时开多个端
- 表结构：users（账号密码昵称头像角色状态时间）+ sessions（token 指向 user_id、过期时间）
- 头像：文件存 agent/avatars/，users.avatar 只存文件名（不塞数据库大字段）；
        换头像时旧文件由 API 层删掉，数据库里永不残留二进制
- 数据库连不上：抛 RuntimeError，由 API 层转成"人话"提示

命令行：
    python agent/auth.py --init        建表
    python agent/auth.py --users       看已注册用户
"""

import argparse
import hashlib
import hmac
import os
import re
import secrets
from datetime import datetime, timedelta
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parent))

import db  # noqa: E402

TOKEN_COOKIE = "a23_token"
TOKEN_DAYS = 7
PBKDF2_ITER = 120_000
MIN_PASSWORD = 6
MAX_PASSWORD = 64
MAX_NICKNAME = 20

# 三个端口：用户端 / 客服端 / 管理员端。
# 账号的 role 字段是逗号分隔集合，例如 "user,staff,admin" 表示三个端都能登。
ROLES = {"user": "用户端", "staff": "客服端", "admin": "管理员端"}
ROLE_ORDER = ["user", "staff", "admin"]

# 内置超级管理员（首次建表时写入，已存在则不动它的密码）
BUILTIN_ADMIN = {"username": "2482991994", "password": "123456",
                 "display_name": "管理员", "roles": "user,staff,admin"}

AVATAR_DIR = Path(__file__).resolve().parent / "avatars"

USERNAME_RE = re.compile(r"^[A-Za-z0-9_\u4e00-\u9fa5]{3,20}$")


# ---------------------------------------------------------------- 建表

def _has_col(cur, table: str, column: str) -> bool:
    cur.execute("SELECT COUNT(*) FROM information_schema.columns "
                "WHERE table_schema = DATABASE() AND table_name = %s AND column_name = %s",
                (table, column))
    return bool(cur.fetchone()[0])


def _add_col(cur, column: str, ddl: str) -> bool:
    """补一列；专用账号没 ALTER 权限时打印人话提示并跳过（不拦启动）。"""
    if _has_col(cur, "users", column):
        return True
    try:
        cur.execute("ALTER TABLE users ADD COLUMN " + ddl)
        return True
    except Exception as e:
        code = e.args[0] if getattr(e, "args", None) else None
        if code not in (1142, 1044):
            raise
        print("【用户表】账号没有 ALTER 权限，%s 列没补上；用 root 执行一次：" % column)
        print("        ALTER TABLE users ADD COLUMN " + ddl + ";")
        return False


def ensure_tables() -> str:
    """建 users / sessions 两张表（幂等，含老库补列）。返回库名。"""
    cfg = db.load_config()
    conn = db.get_conn(with_db=True)
    try:
        with conn.cursor() as cur:
            cur.execute("""
                CREATE TABLE IF NOT EXISTS users (
                    id            INT AUTO_INCREMENT PRIMARY KEY,
                    username      VARCHAR(50)  NOT NULL COMMENT '登录账号',
                    password_hash VARCHAR(255) NOT NULL COMMENT 'PBKDF2 哈希（含盐）',
                    display_name  VARCHAR(50)  NULL     COMMENT '昵称（显示名）',
                    avatar        VARCHAR(120) NULL     COMMENT '头像文件名（agent/avatars/ 下）',
                    role          VARCHAR(40)  NOT NULL DEFAULT 'user'
                                  COMMENT '角色集合，逗号分隔：user / staff / admin',
                    status        VARCHAR(10)  NOT NULL DEFAULT 'active'
                                  COMMENT 'active 正常 / disabled 已停用',
                    created_at    DATETIME     NOT NULL COMMENT '注册时间',
                    last_login    DATETIME     NULL     COMMENT '最近登录时间',
                    last_seen     DATETIME     NULL     COMMENT '最近活跃（客服在线判定）',
                    UNIQUE KEY uk_username (username)
                ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COMMENT='A23 用户账号'
            """)
            cur.execute("""
                CREATE TABLE IF NOT EXISTS sessions (
                    token      CHAR(64) NOT NULL PRIMARY KEY COMMENT '登录令牌',
                    user_id    INT      NOT NULL COMMENT '归属用户',
                    created_at DATETIME NOT NULL COMMENT '签发时间',
                    expires_at DATETIME NOT NULL COMMENT '过期时间',
                    INDEX idx_user (user_id),
                    INDEX idx_exp (expires_at)
                ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COMMENT='A23 登录会话'
            """)
            # 老库（10-22 建的 users 没有 avatar / role / status / last_seen）：缺哪列补哪列
            _add_col(cur, "avatar", "avatar VARCHAR(120) NULL "
                                    "COMMENT '头像文件名（agent/avatars/ 下）' AFTER display_name")
            _add_col(cur, "role", "role VARCHAR(40) NOT NULL DEFAULT 'user' "
                                  "COMMENT '角色集合：user / staff / admin' AFTER avatar")
            _add_col(cur, "status", "status VARCHAR(10) NOT NULL DEFAULT 'active' "
                                    "COMMENT 'active / disabled' AFTER role")
            _add_col(cur, "last_seen", "last_seen DATETIME NULL "
                                       "COMMENT '最近活跃（客服在线判定）' AFTER last_login")
        conn.commit()
        ensure_builtin_admin()
    finally:
        conn.close()
    AVATAR_DIR.mkdir(parents=True, exist_ok=True)
    return cfg["database"]


def ensure_builtin_admin() -> None:
    """写入内置管理员账号（存在则只保证角色，不覆盖密码）。"""
    conn = db.get_conn(with_db=True)
    try:
        with conn.cursor() as cur:
            cur.execute("SELECT id FROM users WHERE username = %s", (BUILTIN_ADMIN["username"],))
            row = cur.fetchone()
            if row:
                cur.execute("UPDATE users SET role = %s WHERE id = %s",
                            (BUILTIN_ADMIN["roles"], row[0]))
                return
            cur.execute(
                "INSERT INTO users (username, password_hash, display_name, role, status, created_at) "
                "VALUES (%s, %s, %s, %s, 'active', %s)",
                (BUILTIN_ADMIN["username"], hash_password(BUILTIN_ADMIN["password"]),
                 BUILTIN_ADMIN["display_name"], BUILTIN_ADMIN["roles"], _now_str()),
            )
            conn.commit()
    finally:
        conn.close()



# ---------------------------------------------------------------- 密码

def hash_password(password: str, salt: bytes | None = None) -> str:
    salt = salt or os.urandom(16)
    dk = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt, PBKDF2_ITER)
    return "pbkdf2_sha256$%d$%s$%s" % (PBKDF2_ITER, salt.hex(), dk.hex())


def verify_password(password: str, stored: str) -> bool:
    try:
        algo, it, salt_hex, hash_hex = stored.split("$")
        if algo != "pbkdf2_sha256":
            return False
        dk = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"),
                                 bytes.fromhex(salt_hex), int(it))
        return hmac.compare_digest(dk.hex(), hash_hex)
    except Exception:
        return False


# ---------------------------------------------------------------- 用户名 / 密码校验

def check_username(username: str) -> str:
    username = (username or "").strip()
    if not USERNAME_RE.match(username):
        raise ValueError("用户名要 3–20 位，可用中文、字母、数字、下划线")
    return username


def check_password(password: str) -> str:
    password = password or ""
    if len(password) < MIN_PASSWORD:
        raise ValueError("密码至少 %d 位" % MIN_PASSWORD)
    if len(password) > MAX_PASSWORD:
        raise ValueError("密码最长 %d 位" % MAX_PASSWORD)
    return password


def _now() -> datetime:
    return datetime.now().replace(microsecond=0)


def _user_dict(uid, username, display_name, avatar=None, roles="user", status="active") -> dict:
    """对外的用户信息：avatar 给可访问的 URL；roles 给端口判断用。"""
    role_list = [r for r in (roles or "user").split(",") if r in ROLES] or ["user"]
    return {
        "id": uid,
        "username": username,
        "display_name": display_name or username,
        "avatar": ("/api/avatar/" + avatar) if avatar else "",
        "roles": role_list,
        "role_text": " / ".join(ROLES[r] for r in ROLE_ORDER if r in role_list),
        "status": status or "active",
    }


def _now_str() -> str:
    return _now().strftime("%Y-%m-%d %H:%M:%S")


# ---------------------------------------------------------------- 注册 / 登录 / 退出

def register(username: str, password: str, display_name: str = "") -> dict:
    """注册一个新账号。用户名重复报 ValueError。返回 user 字典。"""
    username = check_username(username)
    password = check_password(password)
    name = (display_name or "").strip()[:50] or username

    conn = db.get_conn(with_db=True)
    try:
        with conn.cursor() as cur:
            cur.execute("SELECT id FROM users WHERE username = %s", (username,))
            if cur.fetchone():
                raise ValueError("这个用户名已经被注册了，换一个试试")
            cur.execute(
                "INSERT INTO users (username, password_hash, display_name, role, status, created_at) "
                "VALUES (%s, %s, %s, 'user', 'active', %s)",
                (username, hash_password(password), name, _now_str()),
            )
            return _user_dict(cur.lastrowid, username, name, None)
    finally:
        conn.close()


def login(username: str, password: str, portal: str = "user") -> tuple[str, dict]:
    """校验账号密码 + 端口权限 → 签发 token。

    portal：user 用户端 / staff 客服端 / admin 管理员端。
    账号密码错、没该端权限、账号被停用，都抛 ValueError（话术分开说清）。"""
    username = (username or "").strip()
    password = password or ""
    portal = portal if portal in ROLES else "user"
    conn = db.get_conn(with_db=True)
    try:
        with conn.cursor() as cur:
            cur.execute("SELECT id, username, password_hash, display_name, avatar, role, status "
                        "FROM users WHERE username = %s", (username,))
            row = cur.fetchone()
            if not row or not verify_password(password, row[2]):
                raise ValueError("用户名或密码不对")
            if (row[6] or "active") != "active":
                raise ValueError("这个账号已被停用，请联系管理员")
            roles = [r for r in (row[5] or "user").split(",") if r in ROLES] or ["user"]
            if portal not in roles:
                raise ValueError("这个账号没有%s权限" % ROLES[portal])
            cur.execute("UPDATE users SET last_login = %s, last_seen = %s WHERE id = %s",
                        (_now_str(), _now_str(), row[0]))
            token = _new_token(cur, row[0])
            return token, _user_dict(row[0], row[1], row[3], row[4], row[5], row[6])
    finally:
        conn.close()


def _new_token(cur, user_id: int) -> str:
    """在同一个连接/游标里签发 token（顺手清掉过期会话）。"""
    token = secrets.token_hex(32)
    now = _now()
    cur.execute("DELETE FROM sessions WHERE expires_at < %s",
                (now.strftime("%Y-%m-%d %H:%M:%S"),))
    cur.execute("INSERT INTO sessions (token, user_id, created_at, expires_at) "
                "VALUES (%s, %s, %s, %s)",
                (token, user_id, now.strftime("%Y-%m-%d %H:%M:%S"),
                 (now + timedelta(days=TOKEN_DAYS)).strftime("%Y-%m-%d %H:%M:%S")))
    return token


def user_by_token(token: str | None) -> dict | None:
    """用 cookie 里的 token 换用户信息；无效/过期返回 None。"""
    if not token:
        return None
    conn = db.get_conn(with_db=True)
    try:
        with conn.cursor() as cur:
            cur.execute(
                "SELECT u.id, u.username, u.display_name, u.avatar, u.role, u.status "
                "FROM sessions s JOIN users u ON u.id = s.user_id "
                "WHERE s.token = %s AND s.expires_at > %s",
                (token, _now_str()),
            )
            row = cur.fetchone()
            if not row or (row[5] or "active") != "active":     # 停用后立即失效
                return None
            return _user_dict(row[0], row[1], row[2], row[3], row[4], row[5])
    finally:
        conn.close()


def get_user(user_id: int) -> dict | None:
    """按 id 取用户（改完资料回传最新状态用）。"""
    conn = db.get_conn(with_db=True)
    try:
        with conn.cursor() as cur:
            cur.execute("SELECT id, username, display_name, avatar, role, status "
                        "FROM users WHERE id = %s", (int(user_id),))
            row = cur.fetchone()
            return _user_dict(row[0], row[1], row[2], row[3], row[4], row[5]) if row else None
    finally:
        conn.close()


def avatar_file(user_id: int) -> str:
    """当前头像文件名（没有返回空串）——换头像时用它删旧文件。"""
    conn = db.get_conn(with_db=True)
    try:
        with conn.cursor() as cur:
            cur.execute("SELECT avatar FROM users WHERE id = %s", (int(user_id),))
            row = cur.fetchone()
            return (row[0] or "") if row else ""
    finally:
        conn.close()


def check_nickname(name: str) -> str:
    name = (name or "").strip()
    if not name:
        raise ValueError("昵称不能为空")
    if len(name) > MAX_NICKNAME:
        raise ValueError("昵称最多 %d 个字" % MAX_NICKNAME)
    return name


def update_profile(user_id: int, display_name: str | None = None,
                   avatar: str | None = None) -> dict:
    """改昵称 / 头像。两个参数都不传（None）表示不改；avatar 传空串 = 恢复默认。

    返回更新后的用户信息。"""
    sets, vals = [], []
    if display_name is not None:
        sets.append("display_name = %s")
        vals.append(check_nickname(display_name))
    if avatar is not None:
        sets.append("avatar = %s")
        vals.append(avatar or None)
    if not sets:
        raise ValueError("没有要修改的内容")

    vals.append(int(user_id))
    conn = db.get_conn(with_db=True)
    try:
        with conn.cursor() as cur:
            cur.execute("UPDATE users SET " + ", ".join(sets) + " WHERE id = %s", tuple(vals))
    finally:
        conn.close()
    user = get_user(user_id)
    if not user:
        raise ValueError("用户不存在")
    return user


def is_logged_in(token: str | None) -> bool:
    """判断有没有有效登录态。

    数据库连不上时【放行】：这套系统即便 MySQL 挂了，文档操作仍要能用
    （历史记录本来就会降级），不能因为库挂了把所有人锁在登录页外面。"""
    try:
        return user_by_token(token) is not None
    except RuntimeError:
        return True
    except Exception:
        return True


def logout(token: str | None) -> None:
    if not token:
        return
    conn = db.get_conn(with_db=True)
    try:
        with conn.cursor() as cur:
            cur.execute("DELETE FROM sessions WHERE token = %s", (token,))
    finally:
        conn.close()


def list_users(limit: int = 50) -> list[dict]:
    """已注册用户清单（给命令行看，不含密码）。"""
    conn = db.get_conn(with_db=True)
    try:
        with conn.cursor() as cur:
            cur.execute("SELECT id, username, display_name, created_at, last_login, role, status "
                        "FROM users ORDER BY id DESC LIMIT %s", (int(limit),))
            rows = cur.fetchall()
            return [{"id": r[0], "username": r[1], "display_name": r[2],
                     "created_at": str(r[3] or ""), "last_login": str(r[4] or "-"),
                     "roles": r[5] or "user", "status": r[6] or "active"} for r in rows]
    finally:
        conn.close()


# ---------------------------------------------------------------- 管理端：改角色 / 停用 / 重置密码

def normal_roles(roles) -> str:
    """把前端传来的角色数组规整成 "user,staff,admin"（去重、按固定顺序）。"""
    if isinstance(roles, str):
        roles = [x.strip() for x in roles.split(",")]
    got = {r for r in (roles or []) if r in ROLES}
    if not got:
        raise ValueError("至少要保留一个权限")
    if "admin" in got:
        got.update({"user", "staff"})      # 管理员默认三端都能进
    return ",".join(r for r in ROLE_ORDER if r in got)


def set_roles(user_id: int, roles) -> dict:
    text = normal_roles(roles)
    conn = db.get_conn(with_db=True)
    try:
        with conn.cursor() as cur:
            cur.execute("UPDATE users SET role = %s WHERE id = %s", (text, int(user_id)))
            if cur.rowcount == 0:
                raise ValueError("用户不存在")
    finally:
        conn.close()
    user = get_user(user_id)
    if not user:
        raise ValueError("用户不存在")
    return user


def set_status(user_id: int, status: str) -> dict:
    status = "disabled" if status == "disabled" else "active"
    conn = db.get_conn(with_db=True)
    try:
        with conn.cursor() as cur:
            cur.execute("UPDATE users SET status = %s WHERE id = %s", (status, int(user_id)))
            if cur.rowcount == 0:
                raise ValueError("用户不存在")
            if status == "disabled":       # 停用顺手踢掉登录态
                cur.execute("DELETE FROM sessions WHERE user_id = %s", (int(user_id),))
    finally:
        conn.close()
    user = get_user(user_id)
    if not user:
        raise ValueError("用户不存在")
    return user


def reset_password(user_id: int, password: str) -> None:
    password = check_password(password)
    conn = db.get_conn(with_db=True)
    try:
        with conn.cursor() as cur:
            cur.execute("UPDATE users SET password_hash = %s WHERE id = %s",
                        (hash_password(password), int(user_id)))
            if cur.rowcount == 0:
                raise ValueError("用户不存在")
            cur.execute("DELETE FROM sessions WHERE user_id = %s", (int(user_id),))
    finally:
        conn.close()


def create_staff(username: str, password: str, display_name: str = "") -> dict:
    """管理员开客服账号：直接建成 staff 角色。"""
    username = check_username(username)
    password = check_password(password)
    name = (display_name or "").strip()[:50] or username
    conn = db.get_conn(with_db=True)
    try:
        with conn.cursor() as cur:
            cur.execute("SELECT id FROM users WHERE username = %s", (username,))
            if cur.fetchone():
                raise ValueError("这个账号已经存在了，换一个试试")
            cur.execute(
                "INSERT INTO users (username, password_hash, display_name, role, status, created_at) "
                "VALUES (%s, %s, %s, 'user,staff', 'active', %s)",
                (username, hash_password(password), name, _now_str()),
            )
            return _user_dict(cur.lastrowid, username, name, None, "user,staff")
    finally:
        conn.close()


def main():
    ap = argparse.ArgumentParser(description="A23 用户账号表")
    ap.add_argument("--init", action="store_true", help="建 users / sessions 表（幂等）")
    ap.add_argument("--users", action="store_true", help="看已注册用户")
    args = ap.parse_args()

    if args.init:
        try:
            name = ensure_tables()
            print("【用户表】已就绪（库 %s · users + sessions）" % name)
        except RuntimeError as e:
            print("【用户表】%s" % e)
            sys.exit(1)
    if args.users:
        try:
            rows = list_users()
            print("已注册用户 %d 个：" % len(rows))
            for r in rows:
                print("  #%-3d %-16s %-12s 注册 %s ｜ 最近登录 %s"
                      % (r["id"], r["username"], r["display_name"], r["created_at"][:16],
                         r["last_login"][:16]))
        except RuntimeError as e:
            print("【用户表】%s" % e)
            sys.exit(1)
    if not (args.init or args.users):
        ap.print_help()


if __name__ == "__main__":
    main()
