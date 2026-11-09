# -*- coding: utf-8 -*-
"""OCR 结果接入下游的连通性验证

验的不是"能认字"，而是"认出来的字真的能当文档用"：
  上传扫描版 PDF → 进检索池 → 池子里的块带页码和来源 → 能在全文里搜到关键实体
  → 图片 OCR 后同样能入池。
（不做模型抽取：那需要填大模型密钥，跟 OCR 本身无关。）
"""
import json
import os
import sys
import urllib.request

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, "agent"))
sys.path.insert(0, os.path.join(ROOT, "doc_reader"))

BASE = "http://127.0.0.1:8000"
SAMPLE = os.path.join(ROOT, "测试数据", "OCR样本")
IMPORT_DIR = os.path.join(ROOT, "导入文档")

PASS, FAIL, LOG = [], [], []


def ok(name, cond, extra=""):
    (PASS if cond else FAIL).append(name)
    LOG.append("%s %s%s" % ("[通过]" if cond else "[失败]", name,
                            ("  —— " + str(extra)) if extra else ""))
    print(LOG[-1])


def _opener():
    return urllib.request.build_opener(urllib.request.ProxyHandler({}))


def upload(path, name):
    import uuid
    boundary = "----a23" + uuid.uuid4().hex
    with open(path, "rb") as fh:
        data = fh.read()
    body = (
        ("--%s\r\nContent-Disposition: form-data; name=\"file\"; filename=\"%s\"\r\n"
         "Content-Type: application/octet-stream\r\n\r\n" % (boundary, name)).encode("utf-8")
        + data + ("\r\n--%s--\r\n" % boundary).encode("utf-8"))
    req = urllib.request.Request(BASE + "/api/upload", data=body, method="POST")
    req.add_header("Content-Type", "multipart/form-data; boundary=" + boundary)
    try:
        with _opener().open(req, timeout=300) as r:
            return r.status, json.loads(r.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        return e.code, json.loads(e.read().decode("utf-8") or "{}")


def main():
    LOG.append("=" * 62)
    LOG.append("OCR → 下游链路连通性验证（%s）" % BASE)
    LOG.append("=" * 62)

    # 1) 上传扫描版 PDF
    code, j = upload(os.path.join(SAMPLE, "扫描版-采购合同.pdf"), "OCR联通-扫描合同.pdf")
    ok("扫描版 PDF 上传成功", code == 200 and j.get("ok"), j.get("error") or code)
    ok("入池块数 > 0", (j.get("blocks") or 0) > 0, j.get("blocks"))
    ok("标识为扫描件（走了 OCR）", j.get("scanned") is True, j.get("scanned"))
    stored = j.get("stored_name") or ""

    # 2) 池子里这份文档长什么样
    import pool
    data = pool.load_pool()
    doc = next((d for d in data.get("documents", []) if d.get("file") == stored), None)
    ok("检索池里能按文件名找到它", doc is not None, stored)
    if doc:
        blocks = doc.get("blocks") or []
        ok("块都带页码定位",
           all((b.get("locator") or {}).get("page") for b in blocks),
           [(b.get("locator") or {}).get("kind") for b in blocks][:4])
        ok("块标了来源 = ocr",
           all(b.get("source") == "ocr" for b in blocks),
           set(b.get("source") for b in blocks))
        full = "\n".join(b.get("text") or "" for b in blocks)
        for key in ("华东理工大学", "苏州精工机械有限公司", "HD-2026-0317",
                    "384000", "1102023109005566778"):
            ok("全文里搜得到「%s」" % key, key in full)

        # 3) 模拟检索：按关键词找块（下游 matcher 就是这个思路）
        hits = [b for b in blocks if "华东理工大学" in (b.get("text") or "")]
        ok("按实体名能定位到具体块", len(hits) >= 1,
           (hits[0].get("text") or "")[:40] if hits else "没命中")
        if hits:
            loc = hits[0].get("locator") or {}
            ok("定位信息够用来回溯原文", bool(loc.get("page")), loc)

    # 4) 图片 OCR 后同样入池
    code, j2 = upload(os.path.join(SAMPLE, "扫描件-设备验收单.png"), "OCR联通-验收单.png")
    ok("图片上传入池成功", code == 200 and j2.get("ok"), j2.get("error") or code)
    ok("图片也是 OCR 出来的块", (j2.get("blocks") or 0) > 0, j2.get("blocks"))
    d2 = pool.load_pool()
    doc2 = next((d for d in d2.get("documents", []) if d.get("file") == j2.get("stored_name")), None)
    ok("图片在池子里", doc2 is not None, j2.get("stored_name"))
    if doc2:
        t = "\n".join(b.get("text") or "" for b in doc2.get("blocks") or [])
        ok("图片全文里有「数控车床」", "数控车床" in t)
        ok("图片全文里有「CK6140」", "CK6140" in t)

    # 5) 清场
    removed = []
    for f in list(os.listdir(IMPORT_DIR)):
        if f.startswith("OCR联通"):
            try:
                os.remove(os.path.join(IMPORT_DIR, f))
                removed.append(f)
            except OSError:
                pass
    r = pool.rebuild()
    ok("清场完成，池子回到 %d 份" % r.get("total", 0), r.get("ok") is True,
       "删了：" + "、".join(removed))

    LOG.append("")
    LOG.append("=" * 62)
    LOG.append("通过 %d 项 ／ 失败 %d 项" % (len(PASS), len(FAIL)))
    if FAIL:
        LOG.append("失败项：" + "；".join(FAIL))
    LOG.append("=" * 62)
    print("\n".join(LOG[-6:]))

    out = os.path.join(ROOT, "收尾对照", "OCR-下游联通-1026.txt")
    with open(out, "w", encoding="utf-8") as fh:
        fh.write("\n".join(LOG))
    print("写出：" + out)
    return 0 if not FAIL else 1


if __name__ == "__main__":
    sys.exit(main())
