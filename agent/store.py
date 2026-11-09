# -*- coding: utf-8 -*-
"""
store.py —— 模块（2）· 第 4 步：提取结果入库（数据库存储）

赛题原文（模块2）："...提取文件文本关键信息、实体数据或用户指定内容，进行数据库存储操作"，
并"确保信息的保真度与入库的规范性"。

职责：把提取产物统一落进 MySQL，供按文档、按类型、按关键词查询。

      ⚠️ 2026-11-05 起主入口是「统一知识抽取」（knowledge.py）：
      先归纳文档自身结构（doc_schemas），再按 对象/属性/关系/事件/要点 五类入库（knowledge_items）。
      下面的三条老线（用户指定内容 / 关键信息画像 / 实体数据）**保留但不再是默认** ——
      用户明确说"抽实体""出画像"时才走，老验收脚本仍靠它们。

设计：
- 库复用 a23_agent（与 commands/imports 同库，模块内管下列 7 张新表）
- 七张表：
    documents      文档登记（进过库的文档：格式/来源/首见/最近）
    entities       实体数据（7 类：类型/值/次数/出处/摘抄/核验/批次）      ← 老线
    profiles       文档画像主表（主题/文档类型）                          ← 老线
    profile_items  画像明细（要点 point / 关键要素 element：正文/摘抄/出处）  ← 老线
    extractions    用户指定内容（字段提取明细：字段/值/原文摘抄/出处）      ← 老线
    doc_schemas    文档结构（领域/用途/对象及其属性/关注点）              ← 新线
    knowledge_items 知识条目（object/attribute/relation/event/summary）   ← 新线
- 保真度：每条数据都带 quote（原文摘抄）+ loc（出处坐标），与产物文件一字不差
- 规范性：老线幂等入库（同文档同类型同值 → 先查后写）；新线按文档"先删后插"（重跑即覆盖，不会叠）
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
import re
import sys
from collections import defaultdict
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
            # ---- 统一知识抽取（2026-11-05）：文档自己长出来的结构 + 一条条知识 ----
            # 为什么不再分 entities / profiles / extractions：
            # 那三张表的列是按"系统预想的分类"固定的；换行业就装不下。
            # 这两张表只记"文档说了什么"，对象类型/属性名都是文档自己的说法。
            cur.execute("""
                CREATE TABLE IF NOT EXISTS doc_schemas (
                    id          INT AUTO_INCREMENT PRIMARY KEY,
                    doc_name    VARCHAR(255) NOT NULL COMMENT '文档名',
                    domain      VARCHAR(64)  NULL     COMMENT '业务领域（文档自述）',
                    doc_type    VARCHAR(64)  NULL     COMMENT '文档类型',
                    purpose     TEXT         NULL     COMMENT '用途/一句话说明',
                    objects     LONGTEXT     NULL     COMMENT '对象清单 JSON：类型/记录形态/属性',
                    attr_keys   LONGTEXT     NULL     COMMENT '属性名清单 JSON（这份文档长出来的词表）',
                    requirement TEXT         NULL     COMMENT '用户关注点原话',
                    source      VARCHAR(16)  NULL     COMMENT 'model=模型归纳 / structure=骨架直读',
                    run_id      VARCHAR(32)  NULL,
                    ts          DATETIME     NOT NULL,
                    UNIQUE KEY uk_doc (doc_name(200))
                ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COMMENT='A23 文档知识结构（每份文档自己的schema）'
            """)
            cur.execute("""
                CREATE TABLE IF NOT EXISTS knowledge_items (
                    id          INT AUTO_INCREMENT PRIMARY KEY,
                    doc_name    VARCHAR(255)  NOT NULL COMMENT '来源文档',
                    kind        VARCHAR(16)   NOT NULL COMMENT 'object/attribute/relation/event/summary',
                    obj_type    VARCHAR(64)   NULL     COMMENT '业务对象类型（文档自己的说法）',
                    obj_name    VARCHAR(255)  NULL     COMMENT '对象实例标识',
                    attr        VARCHAR(128)  NULL     COMMENT '属性名（文档自己的说法）',
                    value       VARCHAR(1200) NULL     COMMENT '值（原文写法）',
                    unit        VARCHAR(24)   NULL     COMMENT '单位',
                    quote       VARCHAR(1000) NULL     COMMENT '原文摘抄',
                    loc         VARCHAR(255)  NULL     COMMENT '出处坐标',
                    verified    TINYINT       NULL     COMMENT '1=值逐字可核 / 0=仅摘抄可溯',
                    focus       TINYINT       NULL     COMMENT '1=命中用户关注点',
                    source      VARCHAR(24)   NULL     COMMENT 'table / text:exact / text:quote-only / schema',
                    run_id      VARCHAR(32)   NULL,
                    ts          DATETIME      NOT NULL,
                    INDEX idx_doc (doc_name(120)),
                    INDEX idx_kind (kind),
                    INDEX idx_obj (obj_type),
                    INDEX idx_attr (attr),
                    INDEX idx_val (value(120))
                ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COMMENT='A23 文档知识条目（对象/属性/关系/事件/要点）'
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
    # 统一知识表（2026-11-05）：还没建表时不要让整个统计挂掉
    for t in ("doc_schemas", "knowledge_items"):
        try:
            out[t] = _q("SELECT COUNT(*) AS n FROM `%s`" % t)[0]["n"]
        except Exception:
            out[t] = 0
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
#  统一知识抽取（2026-11-05）：写入 / 查询
# ================================================================

def save_knowledge(doc_name: str, schema: dict, units: list,
                   run_id: str = None, requirement: str = "") -> dict:
    """把一份文档的知识结构与条目写进库。

    语义是「这份文档的最新理解」：先删旧条目再整批写入。
    不做增量合并 —— 增量会让"上次抽错的那条"永远赖在库里，
    而知识抽取本来就是可重跑的只读操作，重跑即覆盖最符合直觉。

    返回 {"db", "line", "n"}。
    """
    run_id = run_id or NOW().replace("-", "").replace(":", "").replace(" ", "")
    conn = get_conn(with_db=True)
    try:
        with conn.cursor() as cur:
            _touch_document(cur, doc_name)
            cur.execute("DELETE FROM knowledge_items WHERE doc_name=%s", (doc_name,))
            rows = []
            for u in units:
                rows.append((
                    doc_name, u.get("kind") or "attribute",
                    (u.get("obj_type") or "")[:64], (u.get("obj_name") or "")[:255],
                    (u.get("attr") or "")[:128], (u.get("value") or "")[:1200],
                    (u.get("unit") or "")[:24], (u.get("quote") or "")[:1000],
                    (u.get("loc") or "")[:255],
                    1 if u.get("verified") else 0, 1 if u.get("focus") else 0,
                    (u.get("source") or "")[:24], run_id, NOW(),
                ))
            if rows:
                cur.executemany("""
                    INSERT INTO knowledge_items
                      (doc_name, kind, obj_type, obj_name, attr, value, unit, quote, loc,
                       verified, focus, source, run_id, ts)
                    VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
                """, rows)
            objs = schema.get("objects") or []
            attrs = []
            for o in objs:
                for a in o.get("attrs") or []:
                    attrs.append(a.get("name"))
            cur.execute("""
                INSERT INTO doc_schemas
                  (doc_name, domain, doc_type, purpose, objects, attr_keys,
                   requirement, source, run_id, ts)
                VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
                ON DUPLICATE KEY UPDATE
                  domain=VALUES(domain), doc_type=VALUES(doc_type), purpose=VALUES(purpose),
                  objects=VALUES(objects), attr_keys=VALUES(attr_keys),
                  requirement=VALUES(requirement), source=VALUES(source),
                  run_id=VALUES(run_id), ts=VALUES(ts)
            """, (doc_name, (schema.get("domain") or "")[:64], (schema.get("doc_type") or "")[:64],
                  schema.get("purpose") or "",
                  json.dumps(objs, ensure_ascii=False),
                  json.dumps(list(dict.fromkeys(attrs)), ensure_ascii=False),
                  requirement or "", schema.get("source") or "", run_id, NOW()))
        conn.commit()
    finally:
        conn.close()

    cnt = {}
    for u in units:
        cnt[u.get("kind")] = cnt.get(u.get("kind"), 0) + 1
    line = "已写入 %d 条（%s）· 领域「%s」· 结构 %s" % (
        len(units),
        " / ".join("%s%d" % (k, v) for k, v in sorted(cnt.items())),
        schema.get("domain") or "未识别",
        "模型归纳" if schema.get("source") == "model" else "文档骨架直读")
    from db import load_config
    return {"db": load_config()["database"], "line": line, "n": len(units), "counts": cnt}


def query_knowledge(doc=None, kind=None, obj_type=None, q=None, limit=300):
    sql = "SELECT * FROM knowledge_items WHERE 1=1"
    args = []
    if doc:
        sql += " AND doc_name LIKE %s"; args.append("%" + doc + "%")
    if kind:
        sql += " AND kind=%s"; args.append(kind)
    if obj_type:
        sql += " AND obj_type LIKE %s"; args.append("%" + obj_type + "%")
    if q:
        sql += " AND (value LIKE %s OR attr LIKE %s OR obj_name LIKE %s)"
        args += ["%" + q + "%"] * 3
    sql += " ORDER BY doc_name, kind, obj_type, id LIMIT %s"
    args.append(int(limit))
    return _q(sql, args)


def knowledge_stats():
    """知识视图统计：按 kind、按领域、按对象类型。"""
    out = {"items": 0, "docs": 0, "by_kind": {}, "by_domain": {}, "by_obj": [], "last_ts": None}
    rows = _q("SELECT kind, COUNT(*) AS n FROM knowledge_items GROUP BY kind")
    out["by_kind"] = {r["kind"]: r["n"] for r in rows}
    out["items"] = sum(out["by_kind"].values())
    rows = _q("SELECT COUNT(*) AS n FROM doc_schemas")
    out["docs"] = rows[0]["n"] if rows else 0
    rows = _q("SELECT domain, COUNT(*) AS n FROM doc_schemas GROUP BY domain ORDER BY n DESC")
    out["by_domain"] = {r["domain"] or "未识别": r["n"] for r in rows}
    rows = _q("SELECT obj_type, COUNT(*) AS n FROM knowledge_items "
              "WHERE kind='attribute' GROUP BY obj_type ORDER BY n DESC LIMIT 12")
    out["by_obj"] = [{"obj_type": r["obj_type"], "n": r["n"]} for r in rows]
    rows = _q("SELECT MAX(ts) AS m FROM knowledge_items")
    out["last_ts"] = str(rows[0]["m"]) if rows and rows[0]["m"] else None
    return out


def get_schema(doc: str):
    rows = _q("SELECT * FROM doc_schemas WHERE doc_name LIKE %s ORDER BY ts DESC LIMIT 1",
              ("%" + doc + "%",))
    return rows[0] if rows else None


def list_schemas(limit=100):
    rows = _q("SELECT doc_name, domain, doc_type, purpose, attr_keys, source, ts "
              "FROM doc_schemas ORDER BY ts DESC LIMIT %s", (int(limit),))
    return rows


def knowledge_docs(limit=300):
    """有知识的文档清单（数据仓库左栏用）：每份带条目数 / 对象数 / 属性数 / 领域。

    2026-11-05 加：之前前端只能把所有文档的条目混在一张表里，
    用户没法"点进某个文档看它抽到了什么"。
    """
    rows = _q("SELECT doc_name, COUNT(*) AS items, "
              "COUNT(DISTINCT obj_type) AS objs, "
              "SUM(CASE WHEN kind='attribute' THEN 1 ELSE 0 END) AS attrs, "
              "SUM(CASE WHEN kind='summary' THEN 1 ELSE 0 END) AS points, "
              "MAX(ts) AS ts "
              "FROM knowledge_items GROUP BY doc_name ORDER BY ts DESC LIMIT %s",
              (int(limit),))
    dom = {r["doc_name"]: r for r in _q(
        "SELECT doc_name, domain, doc_type, requirement FROM doc_schemas")}
    for r in rows:
        s = dom.get(r["doc_name"]) or {}
        r["domain"] = s.get("domain") or ""
        r["doc_type"] = s.get("doc_type") or ""
        r["requirement"] = s.get("requirement") or ""
        r["ts"] = str(r["ts"]) if r.get("ts") else ""
    return rows


def knowledge_doc_detail(doc: str, limit=3000):
    """一份文档的完整视图：结构 + 按对象分组 + 每个属性的值清单。

    返回：
      schema   结构（解析好的 objects / attr_keys / requirement）
      counts   各类条目数
      groups   对象 → 该对象的条目（属性在前，同属性按 id 稳定）
      lists    属性 → 去重值列表（"这份文档里有哪些城市"就是看这里）
    """
    if not doc:
        return None
    rows = _q("SELECT * FROM knowledge_items WHERE doc_name=%s "
              "ORDER BY kind, obj_type, obj_name, id LIMIT %s", (doc, int(limit)))
    if not rows:
        rows = _q("SELECT * FROM knowledge_items WHERE doc_name LIKE %s "
                  "ORDER BY kind, obj_type, obj_name, id LIMIT %s",
                  ("%" + doc + "%", int(limit)))
    if not rows:
        return None
    real = rows[0]["doc_name"]
    sch = _q("SELECT * FROM doc_schemas WHERE doc_name=%s ORDER BY ts DESC LIMIT 1", (real,))
    schema = sch[0] if sch else {}
    if schema:
        for k in ("objects", "attr_keys"):
            try:
                schema[k] = json.loads(schema.get(k) or "[]")
            except Exception:
                schema[k] = []
        schema["ts"] = str(schema["ts"]) if schema.get("ts") else ""

    counts = {}
    groups, lists = {}, {}
    seen_val = {}
    for r in rows:
        counts[r["kind"]] = counts.get(r["kind"], 0) + 1
        r["ts"] = str(r["ts"]) if r.get("ts") else ""
        key = r["obj_type"] or "未分类"
        groups.setdefault(key, []).append(r)
        if r["kind"] == "attribute" and r["attr"]:
            k2 = (r["obj_type"], r["attr"])
            if k2 not in lists:
                lists[k2] = []
                seen_val[k2] = set()
            v = (r.get("value") or "").strip()
            if v and v not in seen_val[k2]:
                seen_val[k2].add(v)
                lists[k2].append({"value": v, "unit": r.get("unit") or "",
                                  "obj_name": r.get("obj_name") or "",
                                  "loc": r.get("loc") or "",
                                  "verified": bool(r.get("verified")),
                                  "focus": bool(r.get("focus"))})
    list_out = [{"obj_type": k[0], "attr": k[1], "n": len(v), "values": v}
                for k, v in lists.items()]
    list_out.sort(key=lambda x: -x["n"])
    return {"doc_name": real, "schema": schema, "counts": counts,
            "total": len(rows), "groups": groups, "lists": list_out}


# ================================================================
#  知识图谱（2026-11-07 增强：多属性推导关系 · 多视角切换）
# ================================================================
#
#  一条条目只说明"某个对象有某个属性"；真正让图连起来的是属性之间、
#  对象之间的**重合**。所以这里不只画字面写出的东西，还按可核验的重合
#  推导关系（全部有据可查，不做任何猜测）：
#
#    rel     库里明确写出的关系：主体 -[谓词]-> 客体
#    attr    对象 -[属性名]-> 属性值（只属于这一个对象的值）
#    shared  同一个值被 ≥2 个对象使用，或同一个值出现在不同属性名下
#            （"苏州精诚自动化"既是某运单的"收件人"、又是某物料的"供应商"）
#            → 这是两个对象之间真实存在的纽带，不是模型猜的
#    type    同属一个 obj_type：对象 -[属于]-> 类型中枢
#    doc     出自同一份文档：对象 -[出自]-> 文档中枢
#
#  views 决定画哪几类，默认 rel,attr,shared,type；doc 会引入文档节点，
#  由前端"同文档"视角按需请求。

_GRAPH_ORDER = {"rel": 0, "shared": 1, "link": 2, "attr": 3, "type": 4, "doc": 5}
_GRAPH_ATTR_SKIP = {"序号", "序列", "序列号", "no", "No", "NO", "编号序号", "#"}
_GRAPH_TYPE_SKIP = {"关系", "事件", "要点", "summary", "relation", "event"}
_NUM_ONLY = re.compile(r"^[+-]?\d+(\.\d+)?$")


def _grel_target(value):
    """relation 的 value 形如「主体 → 客体」；兼容 -> / —> / =>。"""
    for sep in ("→", "->", "—>", "=>"):
        if sep in value:
            return value.split(sep)[-1].strip()
    return ""


def _short_doc(name):
    """文档节点标签：去扩展名、去版本尾巴，太长截断。"""
    s = str(name or "").strip()
    for ext in (".docx", ".doc", ".xlsx", ".xls", ".pdf", ".md", ".txt"):
        if s.lower().endswith(ext):
            s = s[: -len(ext)]
            break
    return s if len(s) <= 26 else s[:24] + "…"


def knowledge_graph(doc=None, limit=6000, max_nodes=260, max_edges=900,
                    views=None, min_shared=2, collapse=False):
    """把知识条目拼成多关系图谱（2026-11-07 增强）。

    返回 {"nodes", "edges", "stats", "relation_types", "docs", "truncated"}。
    每条边都带 kind（怎么来的）与 label（哪条属性能读出来），
    前端据此做视角切换与关系筛选；不做任何模型臆造。

    collapse=True 时多做一种"关系转换"：把「A -[性别]-> 值(男) <-[性别]- B」
    这种经值节点中转的两跳关系，折叠成「A -[性别=男]-> B」的直接关系。
    用户就能一眼看到"这两个对象为什么相关"，而不是自己顺着值节点看两跳。
    用户数超过 8 的大团不折叠（会变成几十条边糊成一团），仍走值节点。
    """
    want = [x.strip() for x in str(views or "rel,shared,type").split(",") if x.strip()]
    want = [x for x in want if x in _GRAPH_ORDER] or ["rel", "shared", "type"]
    want_set = set(want)
    min_shared = max(2, int(min_shared or 2))

    sql = ("SELECT doc_name, kind, obj_type, obj_name, attr, value "
           "FROM knowledge_items")
    args = []
    if doc:
        sql += " WHERE doc_name=%s"
        args.append(doc)
    sql += " LIMIT %s"
    rows = _q(sql, args + [int(limit)])

    attrs = defaultdict(list)        # 对象 → [(属性, 值)]
    otypes = {}                      # 对象 → 类型
    odocs = defaultdict(list)        # 对象 → [文档]
    val_owner = defaultdict(set)     # 值 → {用到它的对象}
    val_attr = defaultdict(set)      # 值 → {指向它的属性名}
    rels = []

    for r in rows:
        kind = r.get("kind")
        if kind == "relation":
            subj = (r.get("obj_name") or "").strip()
            pred = (r.get("attr") or "").strip()
            obj = _grel_target(r.get("value") or "")
            if subj and pred and obj and subj != obj:
                rels.append((subj, pred, obj))
            continue
        if kind != "attribute":
            continue
        o = (r.get("obj_name") or "").strip()
        a = (r.get("attr") or "").strip()
        v = (r.get("value") or "").strip()
        if not (o and a and v):
            continue
        if len(v) > 26 or "\n" in v:      # 长段落画不成节点
            continue
        if len(v) == 1 and _NUM_ONLY.match(v):
            continue                      # "1""2" 这种当枢纽只是噪音
        attrs[o].append((a, v))
        otypes.setdefault(o, (r.get("obj_type") or "").strip())
        d = (r.get("doc_name") or "").strip()
        if d and d not in odocs[o]:
            odocs[o].append(d)
        val_owner[v].add(o)
        val_attr[v].add(a)

    edges, seen = [], set()

    def add(s, t, label, kind, w=1, **extra):
        key = (s, t, label, kind)
        if key in seen:
            return
        seen.add(key)
        e = {"s": s, "t": t, "label": label, "kind": kind, "w": w}
        e.update(extra)
        edges.append(e)

    # ① 字面写出的关系
    if "rel" in want_set:
        for subj, pred, obj in rels:
            add("o:" + subj, "o:" + obj, pred, "rel", 3)

    # ② 属性边；值被别人用过 → 升格成 shared（这就是"多出来的关系"）
    #    collapse 时 shared 不画在值节点上，改由下面的 link 边直接连对象
    if "attr" in want_set or ("shared" in want_set and not collapse):
        for o, pairs in attrs.items():
            for a, v in pairs:
                if a in _GRAPH_ATTR_SKIP:
                    continue
                users = val_owner.get(v) or set()
                names = val_attr.get(v) or set()
                cross = (len(names) >= 2 and len(users) >= 2
                         and not _NUM_ONLY.match(v))
                shared = len(users) >= min_shared or cross
                kind = "shared" if shared else "attr"
                if shared and collapse:
                    continue
                if kind not in want_set:
                    continue
                add("o:" + o, "v:" + v, a, kind,
                    len(users) if shared else 1,
                    cross=bool(cross), users=len(users))

    # ②b 关系转换：A -[性别]-> 男 <-[性别]- B   ⇒   A -[性别=男]-> B
    #     对象太多了会糊成一团，所以只用 2~8 个对象共用的小团
    if collapse and ("shared" in want_set or "link" in want_set):
        for v, users in val_owner.items():
            if not (min_shared <= len(users) <= 8):
                continue
            names = sorted(val_attr.get(v) or [])
            if not names:
                continue
            lab = "%s=%s" % (names[0], v)
            ul = sorted(users)
            for i in range(len(ul)):
                for j in range(i + 1, len(ul)):
                    add("o:" + ul[i], "o:" + ul[j], lab, "link", len(ul))

    # ③ 类型中枢
    if "type" in want_set:
        for o, t in otypes.items():
            if not t or t in _GRAPH_TYPE_SKIP:
                continue
            add("o:" + o, "t:" + t, "属于", "type", 1)

    # ④ 文档中枢
    if "doc" in want_set:
        for o, ds in odocs.items():
            for d in ds:
                add("o:" + o, "d:" + d, "出自", "doc", 1)

    total_edges = len(edges)
    total_nodes = len(attrs) + len(val_owner)
    # 先保显式关系，再保共享度高的（w 大的），最后才是普通属性
    edges.sort(key=lambda e: (_GRAPH_ORDER[e["kind"]], -int(e.get("w") or 1), e["label"]))
    if len(edges) > max_edges:
        edges = edges[:int(max_edges)]

    # ---- 由边反推节点（保证不出现孤立点）----
    nodes, order = {}, []
    for e in edges:
        for nid in (e["s"], e["t"]):
            if nid in nodes:
                continue
            if nid.startswith("o:"):
                nm = nid[2:]
                nodes[nid] = {"id": nid, "label": nm[:40], "type": "object",
                              "group": otypes.get(nm) or "关系对象",
                              "docs": odocs.get(nm, [])[:3],
                              "n_attr": len(attrs.get(nm, []))}
            elif nid.startswith("v:"):
                v = nid[2:]
                names = sorted(val_attr.get(v) or [])
                users = len(val_owner.get(v) or ())
                nodes[nid] = {"id": nid, "label": v[:24], "type": "value",
                              "group": names[0] if names else "",
                              "n_obj": users, "attrs": names[:5],
                              "cross": len(names) >= 2 and not _NUM_ONLY.match(v)}
            elif nid.startswith("t:"):
                nodes[nid] = {"id": nid, "label": nid[2:][:24], "type": "type",
                              "group": "类型", "docs": [], "n_attr": 0}
            else:
                nm = nid[2:]
                nodes[nid] = {"id": nid, "label": _short_doc(nm), "type": "doc",
                              "group": "文档", "docs": [], "n_attr": 0}
            order.append(nid)

    for n in nodes.values():
        n["deg"] = 0
    for e in edges:
        nodes[e["s"]]["deg"] += 1
        nodes[e["t"]]["deg"] += 1

    # 节点仍超限：对象全保，其余按度数留（值节点是最容易砍的）
    if len(nodes) > int(max_nodes):
        ranked = sorted(nodes.values(),
                        key=lambda n: (0 if n["type"] == "object" else 1, -n["deg"]))
        keep = {n["id"] for n in ranked[:int(max_nodes)]}
        nodes = {k: v for k, v in nodes.items() if k in keep}
        edges = [e for e in edges if e["s"] in nodes and e["t"] in nodes]
        for n in nodes.values():          # 度数按最终保留的边重算
            n["deg"] = 0
        for e in edges:
            nodes[e["s"]]["deg"] += 1
            nodes[e["t"]]["deg"] += 1

    # ---- 关系类型清单（前端图例 / 筛选 / 侧栏用）----
    rt = defaultdict(int)
    for e in edges:
        rt[(e["kind"], e["label"])] += 1
    relation_types = [{"kind": k, "label": l, "n": n}
                      for (k, l), n in sorted(rt.items(), key=lambda x: (-x[1], x[0][0]))]

    by_kind = defaultdict(int)
    for e in edges:
        by_kind[e["kind"]] += 1
    shared_vals = [v for v in val_owner if len(val_owner[v]) >= min_shared]
    cross_vals = [v for v in val_attr
                  if len(val_attr[v]) >= 2 and len(val_owner.get(v) or ()) >= 2
                  and not _NUM_ONLY.match(v)]
    stats = {
        "rows": len(rows),
        "objects": len(attrs),
        "values": len(val_owner),
        "shared_values": len(shared_vals),
        "cross_values": len(cross_vals),
        "docs": len({d for ds in odocs.values() for d in ds}),
        "edges": {k: by_kind.get(k, 0) for k in ("rel", "shared", "link", "attr", "type", "doc")},
        "nodes": len(nodes),
        "edge_total": total_edges,
        "node_total": total_nodes,
    }
    return {"nodes": list(nodes.values()), "edges": edges,
            "stats": stats, "relation_types": relation_types,
            "views": want,
            "docs": sorted({d for ds in odocs.values() for d in ds}),
            "total_rows": len(rows),
            "truncated": total_edges > len(edges) or total_nodes > len(nodes)}


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
