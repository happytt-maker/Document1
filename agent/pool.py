# -*- coding: utf-8 -*-
"""
pool.py —— 模块（2）· 第 1 步：检索池管理（文档导入的"入池"环节）

检索池 = doc_reader/output.json —— 全量"文档块数据"，matcher / extract_ops
检索时的底料。上传一份新文档后，把它解析出的块并进这个池子，后续所有
指令（编辑 / 提取）立刻就能检索到它。

四个口：
  add(path)      单文件解析并入池（上传链路用；同名替换 = 重传即更新）
  rebuild()      全量重建 = 标准 18 份 + 导入文档（想恢复"含上传"的全景时用）
  baseline()     只扫标准 18 份 → 池子回到比赛基线（演示前复位用）
  status()       池子现状（共几份 / 各格式 / 哪些来自导入）

命令行（在项目根目录跑）：
  python agent/pool.py --status          看池子现状
  python agent/pool.py --rebuild         重扫 源文档 + 导入文档
  python agent/pool.py --baseline        复位：只留标准 18 份
  python agent/pool.py --add 某文件       单文件入池（等同于"上传"的解析部分）

设计说明：
- 写池用"临时文件 + 原子替换"，防止"上传正在写、指令同时在读"读到半截文件
- 上传文件与标准文档重名时，存储名自动加"（导入）"后缀，绝不覆盖基线
- 本模块不碰 doc_reader 原有代码（main.py / run_all.py 照常可用于恢复基线）
"""
import argparse
import json
import os
import sys
import time
from datetime import datetime
from pathlib import Path

BASE = Path(__file__).resolve().parents[1]           # = 研一项目/
DOC_READER_DIR = BASE / "doc_reader"
sys.path.insert(0, str(DOC_READER_DIR))              # 复用读取器（自动识别格式）

from readers import read_one, is_supported, SUPPORTED   # noqa: E402

POOL_JSON = DOC_READER_DIR / "output.json"           # 检索池文件
SRC_DIR = BASE / "测试数据" / "源文档"                 # 标准文档库（18 份 · 比赛基线）
IMPORT_DIR = BASE / "导入文档"                         # 导入文档库（用户上传）


# ==================== 池的读 / 写 ====================

def load_pool() -> dict:
    """读池（不存在或坏了就给一个空壳，不崩）。"""
    if POOL_JSON.exists():
        try:
            return json.loads(POOL_JSON.read_text(encoding="utf-8"))
        except Exception:
            pass
    return {
        "source_dir": str(SRC_DIR),
        "supported_types": sorted(SUPPORTED),
        "file_count": 0, "ok_count": 0, "failed_count": 0, "skipped_count": 0,
        "documents": [], "failed": [],
    }


def _save_pool(data: dict):
    """原子写：先写 .tmp 再替换——防止并发读到半截文件。"""
    tmp = POOL_JSON.with_name(POOL_JSON.name + ".tmp")
    tmp.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    os.replace(tmp, POOL_JSON)


def _sync_counts(data: dict):
    """重算计数（增量模式下 file_count 没有再扫的意义，按实际重新数）。"""
    docs = data.get("documents", [])
    fails = data.get("failed", [])
    data["ok_count"] = len(docs)
    data["failed_count"] = len(fails)
    data["file_count"] = len(docs) + len(fails)


# ==================== 存储名规划（上传落盘前用） ====================

def plan_stored_name(raw_name: str) -> Path:
    """给上传文件规划"不冲突的存储名"。

    规则（按顺序）：
      1. 与标准文档库重名 → 加"（导入）"后缀（绝不顶掉基线；
         后续再传同名的，走第 2 条递增）
      2. 导入文档库里已存在 → 加 "(2)/(3)/..." 递增
    返回：应保存到的完整路径（仅规划，不落盘）。
    """
    raw_name = Path(raw_name).name                     # 防路径注入
    stem, ext = Path(raw_name).stem, Path(raw_name).suffix
    name = raw_name

    # 1) 与标准库重名 → 后缀标注
    if SRC_DIR.exists():
        std = {p.name for p in SRC_DIR.rglob("*") if p.is_file()}
        if name in std:
            name = "%s（导入）%s" % (stem, ext)

    # 2) 导入库内递增
    IMPORT_DIR.mkdir(parents=True, exist_ok=True)
    cand = IMPORT_DIR / name
    n = 1
    while cand.exists():
        n += 1
        cand = IMPORT_DIR / ("%s(%d)%s" % (Path(name).stem, n, ext))
    return cand


# ==================== 核心：单文件入池 ====================

def add(path) -> dict:
    """把一个文件解析并入池（同名文档替换 = 重传更新）。返回摘要 dict。"""
    path = Path(path)
    if not path.exists():
        return {"ok": False, "error": "文件不存在：%s" % path}
    if not is_supported(path):
        return {"ok": False,
                "error": "不支持的格式「%s」。当前支持：%s"
                         % (path.suffix or "未知", " / ".join(SUPPORTED))}

    t0 = time.time()
    try:
        doc = read_one(path)                            # ← 自动识别格式并解析
    except Exception as e:
        return {"ok": False, "error": "解析失败：%s: %s" % (type(e).__name__, e)}
    if doc is None:
        return {"ok": False, "error": "读取器不认识这个文件格式"}

    doc["_imported"] = True                             # 标记：来自导入（供展示区分）
    data = load_pool()
    docs = [d for d in data.get("documents", []) if d.get("file") != doc["file"]]
    replaced = len(docs) != len(data.get("documents", []))
    docs.append(doc)
    data["documents"] = docs
    data["failed"] = [f for f in data.get("failed", [])
                      if Path(str(f.get("file", ""))).name != doc["file"]]
    _sync_counts(data)
    _save_pool(data)

    return {"ok": True, "file": doc["file"], "type": doc.get("type"),
            "blocks": len(doc.get("blocks", [])),
            "elapsed": round(time.time() - t0, 2),
            "replaced": replaced}


# ==================== 全量重建 / 复位 ====================

def _scan_and_read(dirs, with_import_flag: bool):
    """扫若干目录 → 逐个解析。返回 (documents, failed)。"""
    documents, failed = [], []
    for d in dirs:
        if not d.exists():
            continue
        for p in sorted(d.rglob("*")):
            if not p.is_file() or not is_supported(p):
                continue
            try:
                doc = read_one(p)
                if doc is None:
                    continue
                if with_import_flag and p.is_relative_to(IMPORT_DIR):
                    doc["_imported"] = True
                documents.append(doc)
            except Exception as e:
                failed.append({"file": str(p), "error": repr(e)})
    return documents, failed


def rebuild() -> dict:
    """全量重建：标准 18 份 + 导入文档 → 池。"""
    t0 = time.time()
    documents, failed = _scan_and_read([SRC_DIR, IMPORT_DIR], True)
    data = {
        "source_dir": "%s ｜ %s" % (SRC_DIR, IMPORT_DIR),
        "supported_types": sorted(SUPPORTED),
        "file_count": len(documents) + len(failed),
        "ok_count": len(documents),
        "failed_count": len(failed),
        "skipped_count": 0,
        "documents": documents,
        "failed": failed,
    }
    _save_pool(data)
    imported = sum(1 for d in documents if d.get("_imported"))
    return {"ok": True, "total": len(documents), "imported": imported,
            "failed": len(failed), "elapsed": round(time.time() - t0, 2)}


def baseline() -> dict:
    """复位到比赛基线：只扫标准 18 份。（导入文件仍在磁盘，随时可 rebuild 回来。）"""
    t0 = time.time()
    documents, failed = _scan_and_read([SRC_DIR], False)
    data = {
        "source_dir": str(SRC_DIR),
        "supported_types": sorted(SUPPORTED),
        "file_count": len(documents) + len(failed),
        "ok_count": len(documents),
        "failed_count": len(failed),
        "skipped_count": 0,
        "documents": documents,
        "failed": failed,
    }
    _save_pool(data)
    return {"ok": True, "total": len(documents), "failed": len(failed),
            "elapsed": round(time.time() - t0, 2)}


def purge_imports() -> dict:
    """演示复位：清空导入文档库（删除全部上传文件）+ 池子复位基线。

    只在"导入文档/"目录内操作，不碰标准文档库和任何其它目录。
    返回 {"ok", "removed": [文件名...], "total": 复位后的份数}。
    """
    removed = []
    if IMPORT_DIR.exists():
        for p in sorted(IMPORT_DIR.rglob("*")):
            if p.is_file():
                try:
                    p.unlink()
                    removed.append(p.name)
                except OSError:
                    pass
    r = baseline()
    return {"ok": True, "removed": removed, "total": r["total"]}


# ==================== 现状 ====================

def status() -> dict:
    """池子现状：总数 / 各格式 / 导入文档清单。"""
    data = load_pool()
    docs = data.get("documents", [])
    by_type = {}
    imported = []
    for d in docs:
        t = d.get("type", "?")
        by_type[t] = by_type.get(t, 0) + 1
        if d.get("_imported"):
            imported.append({"file": d.get("file"),
                             "blocks": len(d.get("blocks", []))})
    return {"total": len(docs), "by_type": by_type,
            "imported": imported,
            "failed": len(data.get("failed", [])),
            "pool_file": str(POOL_JSON)}


# ==================== 命令行 ====================

def main():
    ap = argparse.ArgumentParser(description="A23 检索池管理（模块 2 · 文档导入）")
    ap.add_argument("--status", action="store_true", help="看池子现状")
    ap.add_argument("--rebuild", action="store_true", help="全量重建：源文档 + 导入文档")
    ap.add_argument("--baseline", action="store_true", help="复位：只留标准 18 份")
    ap.add_argument("--purge", action="store_true",
                    help="演示复位：清空「导入文档/」（上传件全删）+ 池子复位基线")
    ap.add_argument("--add", help="单文件入池（等于「上传」的解析部分）")
    args = ap.parse_args()

    if args.status:
        st = status()
        print("检索池：%s" % st["pool_file"])
        print("总数：%d 份（失败 %d）" % (st["total"], st["failed"]))
        print("按格式：%s" % "、".join("%s×%d" % (k, v) for k, v in sorted(st["by_type"].items())))
        if st["imported"]:
            print("导入文档（%d 份）：" % len(st["imported"]))
            for it in st["imported"]:
                print("  - %s（%d 块）" % (it["file"], it["blocks"]))
        else:
            print("导入文档：无（当前为纯标准库状态）")
    elif args.rebuild:
        r = rebuild()
        print("重建完成：共 %d 份（其中导入 %d 份）｜失败 %d ｜%.2fs"
              % (r["total"], r["imported"], r["failed"], r["elapsed"]))
    elif args.baseline:
        r = baseline()
        print("已复位到基线：共 %d 份 ｜失败 %d ｜%.2fs" % (r["total"], r["failed"], r["elapsed"]))
    elif args.purge:
        r = purge_imports()
        if r["removed"]:
            print("已清空导入库：删除 %d 个上传件 —— %s"
                  % (len(r["removed"]), "、".join(r["removed"][:8])
                     + ("…" if len(r["removed"]) > 8 else "")))
        else:
            print("导入库本来就是空的（无需清理）")
        print("池子已复位基线：共 %d 份" % r["total"])
    elif args.add:
        r = add(args.add)
        if r.get("ok"):
            print("入池成功：%s ｜格式 %s ｜%d 块 ｜%.2fs%s"
                  % (r["file"], r["type"], r["blocks"], r["elapsed"],
                     "（同名替换）" if r.get("replaced") else ""))
        else:
            print("入池失败：%s" % r.get("error"))
            sys.exit(1)
    else:
        ap.print_help()


if __name__ == "__main__":
    main()
