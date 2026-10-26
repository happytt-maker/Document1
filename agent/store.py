# -*- coding: utf-8 -*-
"""
store.py —— 模块（2）· 第 4 步：提取结果入库（数据库存储）

赛题原文（模块2）："...提取文件文本关键信息、实体数据或用户指定内容，进行数据库存储操作"，
并"确保信息的保真度与入库的规范性"。

职责：把三条提取线（用户指定内容 / 关键信息画像 / 实体数据）的产物统一落进 MySQL，
     供按文档、按类型、按关键词查询。

设计：
- 库复用 a23_agent（与 commands/imports 同库，模块内管下列 5 张新表）
- 五张表：
    documents      文档登记（进过库的文档：格式/来源/首见/最近）
    entities       实体数据（7 类：类型/值/次数/出处/摘抄/核验/批次）
    profiles       文档画像主表（主题/文档类型）
    profile_items  画像明细（要点 point / 关键要素 element：正文/摘抄/出处）
    extractions    用户指定内容（字段提取明细：字段/值/原文摘抄/出处）
- 保真度：每条数据都带 quote（原文摘抄）+ loc（出处坐标），与产物文件一字不差
- 规范性：幂等入库（同文档同类型同值 → 先查后写；重复跑不膨胀、不重复）
- 批次：每次入库生成 run_id（时间戳），同批数据同一批次号，可追溯
- 两条入口：
    实时 —— edit.py 执行完自动调用 save_any(产物 json)（"提取即入库"）
    补录 —— python agent/store.py --import-all（把历史产物全部补录进库）

命令行：
    python agent/store.py --init                  建表（幂等）
    python agent/store.py --stats                 库统计
    python agent/store.py --docs                  按文档看数据分布
    python agent/store.py --entities --doc 设备采购 [--type 日期] [--q 关键词]
    python agent/store.py --fields [--doc X] [--q K]
    python agent/store.py --profile 设备采购        看某份文档的画像
    python agent/store.py --search 华景            全库搜（实体 + 字段值）
    python agent/store.py --import-all            扫描 输出/ 把三线产物全部落库（幂等）
"""

import json
import sys
from datetime import datetime
from pathlib import Path

try:
    from db import get_conn
except ImportError:                                # 允许从别的目录 import
    sys.path.insert(0, str(Path(__file__).resolve().parent))
    from db import get_conn

BASE = Path(__file__).resolve().parent
PROJECT = BASE.parent
OUT_DIR = BASE / "输出"

NOW = lambda: datetime.now().strftime("%Y-%m-%d %H:%M:%S")   # noqa: E731


# ================================================================
#  建表
# ================================================================

def ensure_tables() -> str:
    """建 5 张表（幂等，可反复跑）。返回库名。"""
    conn = get_conn(with_db=True)
    try:
        with conn.cursor() as cur:
            cur.execute("""
                CREATE TABLE IF NOT EXISTS documents (
                    id          INT AUTO_INCREMENT PRIMARY KEY,
                    doc_name    VARCHAR(255) NOT NULL COMMENT '文档名（含扩展名）',
                    format      VARCHAR(16)  NULL     COMMENT '格式：docx/xlsx/md/txt',
                    source      VARCHAR(16)  NULL     COMMENT '来源：标准库/导入',
                    first_seen  DATETIME     NOT NULL COMMENT '首次入库时间',
                    last_seen   DATETIME     NOT NULL COMMENT '最近入库时间',
                    UNIQUE KEY uk_name (doc_name(200))
                ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COMMENT='A23 模块2 入库文档登记'
            """)
            cur.execute("""
                CREATE TABLE IF NOT EXISTS entities (
                    id          INT AUTO_INCREMENT PRIMARY KEY,
                    doc_name    VARCHAR(255)  NOT NULL COMMENT '来源文档',
                    etype       VARCHAR(32)   NOT NULL COMMENT '类型（英文键）',
                    etype_cn    VARCHAR(32)   NULL     COMMENT '类型（中文）',
                    value       VARCHAR(600)  NOT NULL COMMENT '实体值（原文写法）',
                    freq        INT           NOT NULL DEFAULT 1 COMMENT '出现次数',
                    loc         VARCHAR(255)  NULL     COMMENT '出处坐标',
                    quote       VARCHAR(1000) NULL     COMMENT '原文摘抄（保真度）',
                    verify      VARCHAR(16)   NULL     COMMENT '格式核验：ok/suspect/na',
                    source      VARCHAR(16)   NULL     COMMENT '来源：model/rule',
                    run_id      VARCHAR(32)   NULL     COMMENT '入库批次',
                    ts          DATETIME      NOT NULL COMMENT '入库时间',
                    INDEX idx_doc (doc_name(120)),
                    INDEX idx_type (etype),
                    INDEX idx_value (value(120))
                ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COMMENT='A23 模块2 实体数据'
            """)
            cur.execute("""
                CREATE TABLE IF NOT EXISTS profiles (
                    id          INT AUTO_INCREMENT PRIMARY KEY,
                    doc_name    VARCHAR(255) NOT NULL,
                    topic       TEXT         NULL     COMMENT '主题（一句话）',
                    doc_type    VARCHAR(64)  NULL     COMMENT '文档类型',
                    run_id      VARCHAR(32)  NULL     COMMENT '入库批次',
                    ts          DATETIME     NOT NULL COMMENT '入库时间',
                    UNIQUE KEY uk_name (doc_name(200))
                ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COMMENT='A23 模块2 文档画像（主表）'
            """)
            cur.execute("""
                CREATE TABLE IF NOT EXISTS profile_items (
                    id          INT AUTO_INCREMENT PRIMARY KEY,
                    profile_id  INT           NOT NULL COMMENT '所属画像',
                    kind        VARCHAR(16)   NOT NULL COMMENT 'point=核心要点 / element=关键要素',
                    seq         INT           NOT NULL COMMENT '序号',
                    title       VARCHAR(128)  NULL     COMMENT '要素名称（point 为空）',
                    detail      TEXT          NULL     COMMENT '正文（point）/ 值（element）',
                    quote       VARCHAR(1000) NULL     COMMENT '原文摘抄',
                    loc         VARCHAR(255)  NULL     COMMENT '出处坐标',
                    loc_ok      TINYINT       NULL     COMMENT '出处是否定位成功',
                    INDEX idx_pid (profile_id)
                ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COMMENT='A23 模块2 画像明细（要点+要素）'
            """)
            cur.execute("""
                CREATE TABLE IF NOT EXISTS extractions (
                    id          INT AUTO_INCREMENT PRIMARY KEY,
                    doc_name    VARCHAR(255)  NOT NULL COMMENT '来源文档',
                    field       VARCHAR(128)  NOT NULL COMMENT '用户点名的字段',
                    value       VARCHAR(1000) NOT NULL COMMENT '提取值',
                    evidence    VARCHAR(1000) NULL     COMMENT '原文摘抄',
                    loc         VARCHAR(255)  NULL     COMMENT '出处坐标',
                    run_id      VARCHAR(32)   NULL     COMMENT '入库批次',
                    ts          DATETIME      NOT NULL COMMENT '入库时间',
                    INDEX idx_doc (doc_name(120)),
                    INDEX idx_field (field)
                ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COMMENT='A23 模块2 用户指定内容提取'
            """)
    finally:
        conn.close()
    from db import load_config
    return load_config()["database"]


# ================================================================
#  来源识别（标准库 / 导入）
# ================================================================

_SRC_CACHE = {}


def _source_of(doc_name: str) -> str:
    """判断文档来自'标准库'还是'导入'（按文件实际所在目录；带进程级缓存）。"""
    if doc_name in _SRC_CACHE:
        return _SRC_CACHE[doc_name]
    hit = "标准库"
    try:
        for folder, label in ((PROJECT / "导入文档", "导入"),
                              (PROJECT / "测试数据" / "导入样本", "导入"),
                              (PROJECT / "测试数据" / "源文档", "标准库")):
            if folder.exists():
                for _ in folder.rglob(doc_name):
                    hit = label
                    raise StopIteration
    except StopIteration:
        pass
    _SRC_CACHE[doc_name] = hit
    return hit


def _touch_document(cur, doc_name: str):
    """登记/更新文档（存在则更新 last_seen）。"""
    if not doc_name:
        return
    now = NOW()
    fmt = Path(doc_name).suffix.lstrip(".").lower() or None
    src = _source_of(doc_name)
    cur.execute("SELECT id FROM documents WHERE doc_name=%s", (doc_name,))
    row = cur.fetchone()
    if row:
        cur.execute("UPDATE documents SET last_seen=%s, format=%s, source=%s WHERE id=%s",
                    (now, fmt, src, row[0]))
    else:
        cur.execute("INSERT INTO documents (doc_name, format, source, first_seen, last_seen) "
                    "VALUES (%s, %s, %s, %s, %s)", (doc_name, fmt, src, now, now))


# ================================================================
#  三线落库
# ================================================================

def _save_entities(cur, data: dict, run_id: str):
    """实体产物 → entities 表（同步式：产物是整表快照，库跟它对齐）。

    返回 (新增数, 更新数, 移除数)——移除 = 上次有、这次产物里没有的旧值。
    """
    doc = data.get("doc") or ""
    ts = NOW()
    cur.execute("SELECT id, value FROM entities WHERE doc_name=%s", (doc,))
    have = {r[1]: r[0] for r in cur.fetchall()}
    want = {}
    for e in data.get("entities") or []:
        v = (e.get("value") or "").strip()
        if v:
            want[v] = e
    new = upd = removed = 0
    for v, e in want.items():
        vals = (e.get("type") or e.get("etype") or "unknown", e.get("type_cn"), v,
                int(e.get("count") or 1), e.get("loc"), e.get("quote"),
                e.get("verify"), e.get("source"), run_id, ts)
        if v in have:
            cur.execute("UPDATE entities SET etype=%s, etype_cn=%s, value=%s, freq=%s, "
                        "loc=%s, quote=%s, verify=%s, source=%s, run_id=%s, ts=%s WHERE id=%s",
                        vals + (have[v],))
            upd += 1
        else:
            cur.execute("INSERT INTO entities (etype, etype_cn, value, freq, loc, quote, "
                        "verify, source, run_id, ts, doc_name) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)",
                        vals + (doc,))
            new += 1
    for v, eid in have.items():
        if v not in want:
            cur.execute("DELETE FROM entities WHERE id=%s", (eid,))
            removed += 1
    return new, upd, removed


def _save_profile(cur, data: dict, run_id: str):
    """画像产物 → profiles + profile_items 表。返回 (新增, 更新)（口径：1=画像本体）。"""
    doc = data.get("doc") or ""
    ts = NOW()
    cur.execute("SELECT id FROM profiles WHERE doc_name=%s", (doc,))
    row = cur.fetchone()
    if row:
        pid = row[0]
        cur.execute("DELETE FROM profile_items WHERE profile_id=%s", (pid,))
        cur.execute("UPDATE profiles SET topic=%s, doc_type=%s, run_id=%s, ts=%s WHERE id=%s",
                    (data.get("topic"), data.get("doc_type"), run_id, ts, pid))
        new, upd = 0, 1
    else:
        cur.execute("INSERT INTO profiles (doc_name, topic, doc_type, run_id, ts) "
                    "VALUES (%s,%s,%s,%s,%s)",
                    (doc, data.get("topic"), data.get("doc_type"), run_id, ts))
        pid = cur.lastrowid
        new, upd = 1, 0
    seq = 0
    for p in data.get("points") or []:
        seq += 1
        cur.execute("INSERT INTO profile_items (profile_id, kind, seq, title, detail, quote, loc, loc_ok) "
                    "VALUES (%s,'point',%s,NULL,%s,%s,%s,%s)",
                    (pid, seq, p.get("text"), p.get("quote"), p.get("loc"),
                     1 if p.get("loc_ok") else 0))
    for el in data.get("elements") or []:
        seq += 1
        cur.execute("INSERT INTO profile_items (profile_id, kind, seq, title, detail, quote, loc, loc_ok) "
                    "VALUES (%s,'element',%s,%s,%s,%s,%s,%s)",
                    (pid, seq, el.get("name"), el.get("value"), el.get("quote"), el.get("loc"),
                     1 if el.get("loc_ok") else 0))
    return new, upd, 0


def _save_extraction(cur, data: dict, run_id: str):
    """字段提取产物 → extractions 表（按字段同步：每个字段以最新提取为准）。

    返回 (新增数, 更新数, 移除数)。
    """
    doc = data.get("doc") or ""
    ts = NOW()
    new = upd = removed = 0
    for f in data.get("fields") or []:
        field = (f.get("field") or f.get("orig") or "").strip()
        if not field:
            continue
        cur.execute("SELECT id, value FROM extractions WHERE doc_name=%s AND field=%s",
                    (doc, field))
        have = {r[1]: r[0] for r in cur.fetchall()}
        want = []
        for it in f.get("items") or []:
            v = (it.get("value") or "").strip()
            if v:
                want.append((v, it))
        wv = {v for v, _ in want}
        for v, it in want:
            loc = it.get("loc_text") or ""
            if v in have:
                cur.execute("UPDATE extractions SET evidence=%s, loc=%s, run_id=%s, ts=%s WHERE id=%s",
                            (it.get("evidence"), loc, run_id, ts, have[v]))
                upd += 1
            else:
                cur.execute("INSERT INTO extractions (doc_name, field, value, evidence, loc, run_id, ts) "
                            "VALUES (%s,%s,%s,%s,%s,%s,%s)",
                            (doc, field, v, it.get("evidence"), loc, run_id, ts))
                new += 1
        for v, eid in have.items():
            if v not in wv:
                cur.execute("DELETE FROM extractions WHERE id=%s", (eid,))
                removed += 1
    return new, upd, removed


def _detect_kind(json_path: Path, data: dict) -> str:
    """识别产物类型：entities / profile / extract。"""
    name = json_path.stem
    if name.endswith("_实体"):
        return "entities"
    if name.endswith("_画像"):
        return "profile"
    if name.endswith("_提取结果"):
        return "extract"
    # 按内容兜底
    if "entities" in data:
        return "entities"
    if "points" in data or "topic" in data:
        return "profile"
    return "extract"


def save_any(json_path, run_id: str = None) -> dict:
    """把一份产物 json 落库（自动识别类型）。返回 {kind, new, updated, doc}。

    幂等：同一份数据重复落库 → 只更新，不重复膨胀。
    """
    jp = Path(json_path)
    data = json.loads(jp.read_text(encoding="utf-8"))
    kind = _detect_kind(jp, data)
    run_id = run_id or ("r" + datetime.now().strftime("%Y%m%d-%H%M%S"))
    conn = get_conn(with_db=True)
    try:
        conn.begin()
        with conn.cursor() as cur:
            _touch_document(cur, data.get("doc") or "")
            if kind == "entities":
                new, upd, rm = _save_entities(cur, data, run_id)
            elif kind == "profile":
                new, upd, rm = _save_profile(cur, data, run_id)
            else:
                new, upd, rm = _save_extraction(cur, data, run_id)
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()
    return {"kind": kind, "new": new, "updated": upd, "removed": rm,
            "doc": data.get("doc") or ""}


# ================================================================
#  历史产物补录（--import-all）
# ================================================================

def import_all(root=None, verbose=True) -> dict:
    """扫描目录下三线产物 json 全部落库（幂等）。返回统计。"""
    root = Path(root) if root else OUT_DIR
    pats = ["*_实体.json", "*_画像.json", "*_提取结果.json"]
    files = []
    for pat in pats:
        files += sorted(root.rglob(pat))
    # 同源多份（如归档目录里的历史快照）按 mtime 升序：最新跑的那份最后落库
    files.sort(key=lambda p: (p.stat().st_mtime, str(p)))
    # 跳过验收归档里的"明细"等杂项（它们不匹配上面三个后缀，天然被滤掉）
    totals = {"files": 0, "new": 0, "updated": 0, "removed": 0, "fail": 0}
    for fp in files:
        try:
            r = save_any(fp)
            totals["files"] += 1
            totals["new"] += r["new"]
            totals["updated"] += r["updated"]
            totals["removed"] += r.get("removed", 0)
            if verbose:
                print("  [OK] %-46s %s ｜ 新增 %d · 更新 %d · 移除 %d"
                      % (fp.relative_to(root), r["kind"], r["new"], r["updated"],
                         r.get("removed", 0)))
        except Exception as e:
            totals["fail"] += 1
            if verbose:
                print("  [XX] %-46s %s" % (fp.name, e))
    return totals


# ================================================================
#  查询口
# ================================================================

def _q(sql, args=()):
    conn = get_conn(with_db=True)
    try:
        import pymysql
        with conn.cursor(pymysql.cursors.DictCursor) as cur:
            cur.execute(sql, args)
            return cur.fetchall()
    finally:
        conn.close()


def stats() -> dict:
    out = {}
    for t in ("documents", "entities", "profiles", "profile_items", "extractions"):
        out[t] = _q("SELECT COUNT(*) AS n FROM `%s`" % t)[0]["n"]
    rows = _q("SELECT MAX(ts) AS m FROM entities")
    out["last_ts"] = str(rows[0]["m"]) if rows and rows[0]["m"] else None
    return out


def list_docs():
    return _q("""
        SELECT d.doc_name, d.format, d.source, d.first_seen, d.last_seen,
               (SELECT COUNT(*) FROM entities e WHERE e.doc_name=d.doc_name) AS n_ent,
               (SELECT COUNT(*) FROM extractions x WHERE x.doc_name=d.doc_name) AS n_fld,
               (SELECT COUNT(*) FROM profiles p WHERE p.doc_name=d.doc_name) AS n_prof
        FROM documents d ORDER BY d.last_seen DESC, d.doc_name
    """)


def query_entities(doc=None, etype=None, q=None, limit=200):
    sql = "SELECT * FROM entities WHERE 1=1"
    args = []
    if doc:
        sql += " AND doc_name LIKE %s"; args.append("%" + doc + "%")
    if etype:
        sql += " AND (etype=%s OR etype_cn=%s)"; args += [etype, etype]
    if q:
        sql += " AND value LIKE %s"; args.append("%" + q + "%")
    sql += " ORDER BY doc_name, etype, id LIMIT %s"
    args.append(int(limit))
    return _q(sql, args)


def query_fields(doc=None, q=None, limit=200):
    sql = "SELECT * FROM extractions WHERE 1=1"
    args = []
    if doc:
        sql += " AND doc_name LIKE %s"; args.append("%" + doc + "%")
    if q:
        sql += " AND (value LIKE %s OR field LIKE %s)"; args += ["%" + q + "%", "%" + q + "%"]
    sql += " ORDER BY doc_name, id LIMIT %s"
    args.append(int(limit))
    return _q(sql, args)


def get_profile(doc: str):
    """按文档名（模糊）取画像：返回 (主表行, 明细行列表)。"""
    rows = _q("SELECT * FROM profiles WHERE doc_name LIKE %s ORDER BY ts DESC LIMIT 1",
              ("%" + doc + "%",))
    if not rows:
        return None, []
    p = rows[0]
    items = _q("SELECT * FROM profile_items WHERE profile_id=%s ORDER BY seq", (p["id"],))
    return p, items


def search(q: str, limit=100):
    """全库搜：实体值 + 字段值（带来源类型标注）。"""
    ents = _q("SELECT doc_name, etype_cn AS kind, value, freq, loc, quote FROM entities "
              "WHERE value LIKE %s ORDER BY doc_name LIMIT %s", ("%" + q + "%", int(limit)))
    flds = _q("SELECT doc_name, field AS kind, value, NULL AS freq, loc, evidence AS quote "
              "FROM extractions WHERE value LIKE %s ORDER BY doc_name LIMIT %s",
              ("%" + q + "%", int(limit)))
    return (ents + flds)[:limit]


# ================================================================
#  命令行
# ================================================================

def _print_table(rows, cols, title=None):
    if title:
        print(title)
    if not rows:
        print("（无记录）")
        return
    heads, keys = zip(*cols)
    def _fmt(v, w):
        s = "" if v is None else str(v)
        s = s.replace("\n", " ")
        return s[:w].ljust(w)
    widths = []
    for h, k in cols:
        mx = max([len(h)] + [len(str(r.get(k) or "")) for r in rows])
        widths.append(min(max(mx, 10), 26))
    print("  ".join(_fmt(h, w) for h, w in zip(heads, widths)))
    print("-" * (sum(widths) + 2 * (len(widths) - 1)))
    for r in rows:
        print("  ".join(_fmt(r.get(k), w) for k, w in zip(keys, widths)))
    print("（共 %d 条）" % len(rows))


def main():
    import argparse
    ap = argparse.ArgumentParser(description="A23 提取结果仓库（MySQL · 模块2 第4步）")
    ap.add_argument("--init", action="store_true", help="建表（幂等）")
    ap.add_argument("--stats", action="store_true", help="库统计")
    ap.add_argument("--docs", action="store_true", help="按文档看数据分布")
    ap.add_argument("--entities", action="store_true", help="查实体（可配 --doc/--type/--q）")
    ap.add_argument("--fields", action="store_true", help="查字段提取（可配 --doc/--q）")
    ap.add_argument("--profile", help="看某份文档的画像（文档名可只写一段）")
    ap.add_argument("--search", help="全库搜关键词（实体+字段值）")
    ap.add_argument("--import-all", action="store_true", help="扫描 输出/ 把历史产物全部落库")
    ap.add_argument("--doc", help="按文档名（模糊）")
    ap.add_argument("--type", dest="etype", help="实体类型（中文或英文键）")
    ap.add_argument("--q", help="值关键词")
    ap.add_argument("--limit", type=int, default=200)
    args = ap.parse_args()

    try:
        if args.init:
            db = ensure_tables()
            s = stats()
            print("库 `%s` 五张表就绪：documents=%d · entities=%d · profiles=%d · profile_items=%d · extractions=%d"
                  % (db, s["documents"], s["entities"], s["profiles"], s["profile_items"], s["extractions"]))
        elif args.stats:
            s = stats()
            print("【库统计】a23_agent")
            print("  文档 %d 份 ｜ 实体 %d 条 ｜ 画像 %d 份（明细 %d 条）｜ 字段提取 %d 条"
                  % (s["documents"], s["entities"], s["profiles"], s["profile_items"], s["extractions"]))
            print("  最近入库：%s" % (s["last_ts"] or "—"))
        elif args.docs:
            _print_table(list_docs(),
                         [("文档", "doc_name"), ("格式", "format"), ("来源", "source"),
                          ("实体", "n_ent"), ("字段", "n_fld"), ("画像", "n_prof"),
                          ("最近入库", "last_seen")],
                         title="【入库文档一览】")
        elif args.entities:
            rows = query_entities(args.doc, args.etype, args.q, args.limit)
            _print_table(rows,
                         [("文档", "doc_name"), ("类型", "etype_cn"), ("值", "value"),
                          ("次数", "freq"), ("出处", "loc"), ("核验", "verify")],
                         title="【实体数据】")
        elif args.fields:
            rows = query_fields(args.doc, args.q, args.limit)
            _print_table(rows,
                         [("文档", "doc_name"), ("字段", "field"), ("值", "value"), ("出处", "loc")],
                         title="【字段提取】")
        elif args.profile:
            p, items = get_profile(args.profile)
            if not p:
                print("（没有《%s》的画像）" % args.profile)
            else:
                print("【画像】%s（%s）" % (p["doc_name"], p.get("doc_type") or "—"))
                print("  主题：%s" % (p.get("topic") or "—"))
                for it in items:
                    tag = "要点" if it["kind"] == "point" else ("要素·%s" % (it.get("title") or ""))
                    body = it.get("detail") or ""
                    print("  [%s] %s ～(%s)" % (tag, body[:70], it.get("loc") or "—"))
        elif args.search:
            rows = search(args.search, args.limit)
            _print_table(rows,
                         [("文档", "doc_name"), ("类别", "kind"), ("值", "value"), ("出处", "loc")],
                         title="【全库搜索：%s】" % args.search)
        elif args.import_all:
            print("【补录】扫描 %s" % OUT_DIR)
            t = import_all()
            print("完成：%d 个产物文件 ｜ 新增 %d 条 · 更新 %d 条 · 移除 %d 条 ｜ 失败 %d"
                  % (t["files"], t["new"], t["updated"], t["removed"], t["fail"]))
        else:
            ap.print_help()
    except RuntimeError as e:
        print("【错误】%s" % e)
        sys.exit(1)


if __name__ == "__main__":
    main()
