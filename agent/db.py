# -*- coding: utf-8 -*-
"""
db.py —— 模块（1）· 第 5 步：操作历史入库（MySQL）

职责：把"每一次用户指令 + 执行结果"落到 MySQL——供事后查询 / 审计 / 统计，
     同时正面落实赛题"数据库存储"的要求（模块 1 的操作记录 + 模块 2 的数据底座）。

设计：
- 配置从 agent/db_config.json 读（没有就用默认值）；库名 a23_agent（独立，不碰别的项目）
- 表 commands：一条指令一行（会话号/轮次/指令/动作/文档/输出/耗时/OK/备注）
- 连接失败给"人话"提示（服务没启动 / 密码不对 / 库没建）
- 四个口：
    ensure_db()      建库建表（幂等，可反复跑）
    log_command(..)  写一条操作历史
    recent(n)        最近 n 条
    by_doc(name)     某份文档的操作史
- 命令行：
    python agent/db.py --init         建库建表
    python agent/db.py --recent 10    看最近 10 条
    python agent/db.py --doc 设备采购  按文档看操作史
    python agent/db.py --make-user    建专用账号 a23_app（可选；最小权限，跑一次即可）
"""

import json
import sys
from datetime import datetime
from pathlib import Path

import pymysql

BASE = Path(__file__).resolve().parent
CONFIG_PATH = BASE / "db_config.json"

DEFAULT_CONFIG = {
    "host": "localhost",
    "port": 3306,
    "user": "root",
    "password": "123456",
    "database": "a23_agent",
    "charset": "utf8mb4",
}


def load_config() -> dict:
    """读连接配置：db_config.json 优先，缺项用默认值补。"""
    cfg = dict(DEFAULT_CONFIG)
    if CONFIG_PATH.exists():
        try:
            cfg.update(json.loads(CONFIG_PATH.read_text(encoding="utf-8")))
        except Exception:
            pass
    return cfg


def get_conn(with_db: bool = True):
    """连接 MySQL。失败时抛 RuntimeError，并带上"人话"排查提示。"""
    cfg = load_config()
    kw = dict(
        host=cfg["host"], port=int(cfg["port"]),
        user=cfg["user"], password=cfg["password"],
        charset=cfg.get("charset", "utf8mb4"),
        autocommit=True,
    )
    if with_db:
        kw["database"] = cfg["database"]
    try:
        return pymysql.connect(**kw)
    except pymysql.err.OperationalError as e:
        code = e.args[0] if e.args else 0
        if code == 2003:
            tip = ("连不上 MySQL 服务——检查服务是否在运行。"
                   "命令行执行一次：net start MySQL80")
        elif code == 1045:
            tip = "用户名或密码不对——改 agent/db_config.json 里的 user / password"
        elif code == 1049:
            tip = "库还不存在——先跑一次初始化：python agent/db.py --init"
        else:
            tip = str(e)
        raise RuntimeError("MySQL 连接失败：" + tip) from e


def ensure_db() -> str:
    """建库（如果不存在）+ 建表（如果不存在）。幂等，可反复跑。返回库名。

    小技巧：先试着直接连目标库——连得上说明库已存在，就跳过建库语句
    （这样"专用账号没有建库权限"也不影响日常使用）。"""
    cfg = load_config()
    db = cfg["database"]

    need_create = True
    try:
        probe = get_conn(with_db=True)
        probe.close()
        need_create = False
    except RuntimeError:
        pass

    if need_create:
        conn = get_conn(with_db=False)
        try:
            with conn.cursor() as cur:
                cur.execute("CREATE DATABASE IF NOT EXISTS `%s` "
                            "DEFAULT CHARACTER SET utf8mb4 COLLATE utf8mb4_general_ci" % db)
        finally:
            conn.close()

    conn = get_conn(with_db=True)
    try:
        with conn.cursor() as cur:
            cur.execute("""
                CREATE TABLE IF NOT EXISTS commands (
                    id          INT AUTO_INCREMENT PRIMARY KEY,
                    session_id  VARCHAR(64)  NOT NULL COMMENT '会话编号',
                    turn        INT          NOT NULL COMMENT '会话内第几轮',
                    ts          DATETIME     NOT NULL COMMENT '完成时间',
                    command     TEXT         NOT NULL COMMENT '用户指令原文',
                    action      VARCHAR(32)  NULL     COMMENT '四类操作',
                    op          VARCHAR(32)  NULL     COMMENT '具体动作',
                    doc         VARCHAR(255) NULL     COMMENT '作用的文档',
                    out_file    VARCHAR(512) NULL     COMMENT '输出文件（只读操作为空）',
                    ok          TINYINT      NOT NULL DEFAULT 0 COMMENT '是否成功',
                    elapsed_sec FLOAT        NULL     COMMENT '本轮耗时（秒）',
                    note        VARCHAR(512) NULL     COMMENT '备注（反问/失败原因等）',
                    INDEX idx_ts (ts),
                    INDEX idx_doc (doc),
                    INDEX idx_session (session_id)
                ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COMMENT='A23 模块1 操作历史'
            """)
            cur.execute("""
                CREATE TABLE IF NOT EXISTS imports (
                    id          INT AUTO_INCREMENT PRIMARY KEY,
                    ts          DATETIME     NOT NULL COMMENT '导入时间',
                    filename    VARCHAR(255) NOT NULL COMMENT '用户上传时的原始文件名',
                    stored_name VARCHAR(255) NOT NULL COMMENT '实际存储文件名（重名会自动加后缀）',
                    filetype    VARCHAR(16)  NULL     COMMENT '自动识别出的格式',
                    size_bytes  INT          NULL     COMMENT '文件大小',
                    blocks      INT          NULL     COMMENT '解析出的内容块数',
                    ok          TINYINT      NOT NULL DEFAULT 0 COMMENT '是否成功',
                    elapsed_sec FLOAT        NULL     COMMENT '解析耗时（秒）',
                    note        VARCHAR(512) NULL     COMMENT '备注（失败原因等）',
                    INDEX idx_ts2 (ts),
                    INDEX idx_name (stored_name)
                ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COMMENT='A23 模块2 文档导入记录'
            """)
    finally:
        conn.close()
    return db


# ---------------- 专用账号（可选） ----------------
APP_USER = "a23_app"
APP_PASS = "A23app@2026"          # 演示口令；正式部署请换成强密码


def ensure_user() -> str:
    """建"专用应用账号"（可选，跑一次即可）：
    只对 a23_agent 库给 增删改查 + 建表 权限（不给整库管理权），
    避免程序日常操作一直裸用 root。返回账号名。"""
    cfg = load_config()
    conn = get_conn(with_db=False)
    try:
        with conn.cursor() as cur:
            cur.execute("CREATE USER IF NOT EXISTS '%s'@'localhost' "
                        "IDENTIFIED BY %%s" % APP_USER, (APP_PASS,))
            cur.execute("GRANT SELECT, INSERT, UPDATE, DELETE, CREATE ON `%s`.* "
                        "TO '%s'@'localhost'" % (cfg["database"], APP_USER))
    finally:
        conn.close()
    return APP_USER


def log_command(session_id: str, turn: int, command: str, action=None, op=None,
                doc=None, out_file=None, ok=False, elapsed_sec=None, note=None):
    """写一条操作历史。返回新行 id。"""
    conn = get_conn(with_db=True)
    try:
        with conn.cursor() as cur:
            cur.execute(
                "INSERT INTO commands "
                "(session_id, turn, ts, command, action, op, doc, out_file, ok, elapsed_sec, note) "
                "VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)",
                (session_id, int(turn),
                 datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                 command, action, op, doc, out_file, 1 if ok else 0,
                 elapsed_sec, (note or "")[:500]))
            return cur.lastrowid
    finally:
        conn.close()


def log_import(filename: str, stored_name: str, filetype=None, size_bytes=None,
               blocks=None, ok=False, elapsed_sec=None, note=None):
    """写一条"文档导入"记录（模块2 第1步）。返回新行 id。"""
    conn = get_conn(with_db=True)
    try:
        with conn.cursor() as cur:
            cur.execute(
                "INSERT INTO imports "
                "(ts, filename, stored_name, filetype, size_bytes, blocks, ok, elapsed_sec, note) "
                "VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s)",
                (datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                 filename, stored_name, filetype, size_bytes, blocks,
                 1 if ok else 0, elapsed_sec, (note or "")[:500]))
            return cur.lastrowid
    finally:
        conn.close()


def recent_imports(n: int = 10):
    """最近 n 条导入记录（倒序）。"""
    conn = get_conn(with_db=True)
    try:
        with conn.cursor(pymysql.cursors.DictCursor) as cur:
            cur.execute("SELECT * FROM imports ORDER BY id DESC LIMIT %s", (int(n),))
            return cur.fetchall()
    finally:
        conn.close()


def recent(n: int = 10):
    """最近 n 条（倒序）。"""
    conn = get_conn(with_db=True)
    try:
        with conn.cursor(pymysql.cursors.DictCursor) as cur:
            cur.execute("SELECT * FROM commands ORDER BY id DESC LIMIT %s", (int(n),))
            return cur.fetchall()
    finally:
        conn.close()


def by_doc(name: str):
    """按文档名模糊筛（含该字样的都拉出来）。"""
    conn = get_conn(with_db=True)
    try:
        with conn.cursor(pymysql.cursors.DictCursor) as cur:
            cur.execute("SELECT * FROM commands WHERE doc LIKE %s ORDER BY id DESC",
                        ("%" + name + "%",))
            return cur.fetchall()
    finally:
        conn.close()


def print_rows(rows):
    """把查询结果打成一张表。"""
    if not rows:
        print("（无记录）")
        return
    print("%-4s %-19s %-3s %-10s %-10s %-16s %-30s %-3s %s" %
          ("id", "时间", "轮", "类别", "动作", "文档", "指令", "OK", "耗时s"))
    print("-" * 122)
    for r in rows:
        print("%-4s %-19s %-3s %-10s %-10s %-16s %-30s %-3s %s" % (
            r["id"], str(r["ts"])[:19], r["turn"],
            r.get("action") or "-", r.get("op") or "-",
            (r.get("doc") or "-")[:16], (r.get("command") or "")[:28],
            "Y" if r.get("ok") else "N", r.get("elapsed_sec")))


def main():
    import argparse
    ap = argparse.ArgumentParser(description="A23 操作历史库（MySQL）")
    ap.add_argument("--init", action="store_true", help="建库建表（幂等）")
    ap.add_argument("--make-user", action="store_true", help="建专用账号 a23_app（可选；最小权限）")
    ap.add_argument("--recent", type=int, help="看最近 N 条")
    ap.add_argument("--doc", help="按文档名筛选")
    args = ap.parse_args()

    try:
        if args.init:
            db = ensure_db()
            conn = get_conn(with_db=True)
            try:
                with conn.cursor() as cur:
                    cur.execute("SELECT COUNT(*) FROM commands")
                    n = cur.fetchone()[0]
            finally:
                conn.close()
            print("库 `%s` 与表 `commands` 就绪（现有 %d 条记录）" % (db, n))
        elif args.make_user:
            user = ensure_user()
            print("专用账号 `%s`@localhost 已就绪：只对库 `a23_agent` 有 增删改查+建表 权限" % user)
            print("启用（可选）：把 db_config.json 里的 user / password 换成它")
        elif args.recent:
            print_rows(recent(args.recent))
        elif args.doc:
            print_rows(by_doc(args.doc))
        else:
            ap.print_help()
    except RuntimeError as e:
        print("【错误】%s" % e)
        sys.exit(1)


if __name__ == "__main__":
    main()
