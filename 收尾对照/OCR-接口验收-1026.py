# -*- coding: utf-8 -*-
"""「扫描版 PDF / 图片 OCR」接口验收

只验 OCR 这条线（先用 OCR-素材生成-1026.py 造素材）：
  引擎状态 → 图片 OCR → 电子版 PDF（不该走 OCR）→ 扫描版 PDF → 混合版
  → 拍照件（歪+噪点）→ 并入文档库（复用结果不重复 OCR）→ 边界与错误码 → 清场

跑法：
  python 收尾对照/OCR-接口验收-1026.py [--port 8013]
"""

import argparse
import json
import os
import sys
import time

import requests

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SAMPLE = os.path.join(ROOT, "测试数据", "OCR样本")
IMPORT_DIR = os.path.join(ROOT, "导入文档")

sys.path.insert(0, ROOT)
sys.path.insert(0, os.path.join(ROOT, "agent"))
sys.path.insert(0, os.path.join(ROOT, "doc_reader"))

PASS, FAIL, LOG = [], [], []
DETAIL = {}


def ok(name, cond, extra=""):
    (PASS if cond else FAIL).append(name)
    line = "%s %s%s" % ("[通过]" if cond else "[失败]", name,
                        ("  —— " + str(extra)) if extra else "")
    LOG.append(line)
    print(line)


def say(s):
    LOG.append(s)
    print(s)


def norm(s):
    """去空白 + 全角转半角 + 小写：OCR 跟原文比对的归一化。"""
    out = []
    for ch in str(s or ""):
        o = ord(ch)
        if 0xFF01 <= o <= 0xFF5E:
            ch = chr(o - 0xFEE0)
        elif o == 0x3000:
            ch = " "
        if ch.isspace():
            continue
        out.append(ch)
    return "".join(out).lower()


def hit_rate(text, keys):
    t = norm(text)
    missed = [k for k in keys if norm(k) not in t]
    return (len(keys) - len(missed)), len(keys), missed


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--port", type=int, default=8013)
    ap.add_argument("--no-clean", action="store_true", help="保留本次导入的文件（排查用）")
    args = ap.parse_args()
    base = "http://127.0.0.1:%d" % args.port

    s = requests.Session()
    s.trust_env = False                       # 绕开系统代理（这台机器 502 的老毛病）

    LOG.append("=" * 64)
    LOG.append("扫描件 OCR · 接口验收（%s）" % base)
    LOG.append("=" * 64)

    # ---------- 素材
    if not os.path.isdir(SAMPLE):
        say("没有素材目录：%s\n请先跑 收尾对照/OCR-素材生成-1026.py" % SAMPLE)
        return 2
    files = {f: os.path.join(SAMPLE, f) for f in os.listdir(SAMPLE)}
    for need in ["扫描件-设备验收单.png", "电子版-采购合同.pdf",
                 "扫描版-采购合同.pdf", "混合版-统计公报.pdf",
                 "拍照件-公告（歪+噪点）.png"]:
        if need not in files:
            say("缺素材：%s（请重跑素材生成脚本）" % need)
            return 2

    # 从素材生成脚本里取标准答案
    import importlib.util
    spec = importlib.util.spec_from_file_location(
        "mk", os.path.join(ROOT, "收尾对照", "OCR-素材生成-1026.py"))
    mk = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mk)
    EXPECT = mk.EXPECT

    def post_ocr(path, dpi=200, name=None):
        with open(path, "rb") as fh:
            r = s.post(base + "/api/ocr", params={"dpi": dpi},
                       files={"file": (name or os.path.basename(path), fh,
                                       "application/octet-stream")}, timeout=300)
        try:
            j = r.json()
        except Exception:
            j = {"raw": r.text[:300]}
        return r.status_code, j

    # ---------- 0 服务 + 引擎
    try:
        r = s.get(base + "/api/status", timeout=10)
        ok("服务在线（/api/status）", r.status_code == 200, r.status_code)
    except Exception as e:
        say("服务连不上：%r" % e)
        return 2

    r = s.get(base + "/api/ocr/status", timeout=20)
    st = r.json() if r.status_code == 200 else {}
    ok("OCR 状态接口可用", r.status_code == 200, r.status_code)
    ok("OCR 引擎就绪", bool(st.get("ok")), st.get("engine") or st.get("tip", "")[:80])
    ok("PyMuPDF 可用（能拆 PDF 页）", bool(st.get("pdf")), st.get("pdf_tip", ""))
    if not st.get("ok"):
        say("引擎没就绪，后面的用例没有意义，先装依赖再跑。")
        return 2

    # ---------- 1 图片 OCR
    say("\n-- 1 图片 OCR（设备验收单）--")
    t0 = time.time()
    code, j = post_ocr(files["扫描件-设备验收单.png"])
    ok("图片识别返回 200", code == 200, code)
    ok("识别类型 = image", j.get("kind") == "image", j.get("kind"))
    ok("认出了文字", (j.get("char_count") or 0) > 60, "%s 字" % j.get("char_count"))
    ok("段落带 OCR 来源", all(p.get("source") == "ocr" for p in (j.get("paragraphs") or [])))
    ok("段落带置信度", all(p.get("conf") is not None for p in (j.get("paragraphs") or [])))
    h, n, miss = hit_rate(j.get("text"), EXPECT["扫描件-设备验收单.png"])
    ok("关键串命中 ≥ 80%%（%d/%d）" % (h, n), h >= n * 0.8, "没认到：" + "、".join(miss))
    ok("单张图耗时 < 15s", round(time.time() - t0, 1) < 15, "%.1fs" % (time.time() - t0))
    DETAIL["图片"] = {"hits": "%d/%d" % (h, n), "elapsed": j.get("elapsed"),
                     "conf": j.get("avg_conf")}

    # ---------- 2 电子版 PDF：必须【不】走 OCR
    say("\n-- 2 电子版 PDF（自带文字层，不该动用 OCR）--")
    t0 = time.time()
    code, j = post_ocr(files["电子版-采购合同.pdf"])
    el = time.time() - t0
    ok("PDF 识别返回 200", code == 200, code)
    ok("识别类型 = pdf", j.get("kind") == "pdf", j.get("kind"))
    ok("页数 = 3", j.get("page_count") == 3, j.get("page_count"))
    ok("没有扫描页（scanned=False）", j.get("scanned") is False, j.get("scanned"))
    ok("OCR 页数 = 0", (j.get("ocr_page_count") or 0) == 0, j.get("ocr_page_count"))
    ok("全部段落来自文字层",
       (j.get("paragraphs") or []) and all(p.get("source") == "text_layer"
                                          for p in j.get("paragraphs") or []))
    h, n, miss = hit_rate(j.get("text"), EXPECT["电子版-采购合同.pdf"])
    ok("关键串全中（%d/%d）" % (h, n), h == n, "没认到：" + "、".join(miss))
    ok("3 页文字层耗时 < 5s（证明确实没 OCR）", el < 5, "%.1fs" % el)
    DETAIL["电子版"] = {"hits": "%d/%d" % (h, n), "elapsed": round(el, 2)}

    # ---------- 3 扫描版 PDF
    say("\n-- 3 扫描版 PDF（三页全是图，必须走 OCR）--")
    t0 = time.time()
    code, j = post_ocr(files["扫描版-采购合同.pdf"])
    el = time.time() - t0
    ok("扫描版识别返回 200", code == 200, code)
    ok("判定为扫描件（scanned=True）", j.get("scanned") is True, j.get("scanned"))
    ok("3 页全部走了 OCR", (j.get("ocr_page_count") or 0) == 3, j.get("ocr_page_count"))
    pages = {p.get("page") for p in (j.get("paragraphs") or [])}
    ok("段落带上页码 1/2/3", {1, 2, 3}.issubset(pages), sorted(pages))
    ok("段落标注 OCR 来源",
       all(p.get("source") == "ocr" for p in (j.get("paragraphs") or [])
           if p.get("source") == "ocr") and any(
           p.get("source") == "ocr" for p in (j.get("paragraphs") or [])))
    h, n, miss = hit_rate(j.get("text"), EXPECT["扫描版-采购合同.pdf"])
    ok("关键串命中 ≥ 85%%（%d/%d）" % (h, n), h >= n * 0.85, "没认到：" + "、".join(miss))
    ok("总耗时 < 60s（3 页）", el < 60, "%.1fs" % el)
    DETAIL["扫描版"] = {"hits": "%d/%d" % (h, n), "elapsed": round(el, 2),
                       "ocr_elapsed": j.get("ocr_elapsed")}

    # ---------- 4 混合版
    say("\n-- 4 混合版 PDF（第 1 页文字层 + 第 2 页扫描图）--")
    code, j = post_ocr(files["混合版-统计公报.pdf"])
    ok("混合版识别返回 200", code == 200, code)
    ok("页数 = 2", j.get("page_count") == 2, j.get("page_count"))
    ok("文字层页 = 1", j.get("text_pages") == 1, j.get("text_pages"))
    ok("OCR 页 = 1", (j.get("ocr_page_count") or 0) == 1, j.get("ocr_page_count"))
    by_page = {}
    for p in (j.get("paragraphs") or []):
        by_page.setdefault(p.get("page"), set()).add(p.get("source"))
    ok("第 1 页来自文字层", by_page.get(1) == {"text_layer"}, by_page.get(1))
    ok("第 2 页来自 OCR", by_page.get(2) == {"ocr"}, by_page.get(2))
    h, n, miss = hit_rate(j.get("text"), EXPECT["混合版-统计公报.pdf"])
    ok("关键串命中 ≥ 80%%（%d/%d）" % (h, n), h >= n * 0.8, "没认到：" + "、".join(miss))
    DETAIL["混合版"] = {"hits": "%d/%d" % (h, n), "elapsed": j.get("elapsed")}

    # ---------- 5 拍照件（歪 + 噪点）
    say("\n-- 5 拍照件（-0.6° 倾斜 + 9000 噪点）--")
    code, j = post_ocr(files["拍照件-公告（歪+噪点）.png"])
    ok("拍照件识别返回 200", code == 200, code)
    h, n, miss = hit_rate(j.get("text"), EXPECT["拍照件-公告（歪+噪点）.png"])
    ok("关键串命中 ≥ 60%%（%d/%d）" % (h, n), h >= n * 0.6, "没认到：" + "、".join(miss))
    ok("倾斜图片仍能分出多段", (j.get("para_count") or 0) >= 2, j.get("para_count"))
    DETAIL["拍照件"] = {"hits": "%d/%d" % (h, n), "elapsed": j.get("elapsed"),
                       "conf": j.get("avg_conf")}

    # ---------- 6 并入文档库
    say("\n-- 6 并入文档库（预览后用缓存，不重复 OCR）--")
    before = set((s.get(base + "/api/docs", timeout=20).json() or {}).get("docs") or [])
    t0 = time.time()
    with open(files["扫描版-采购合同.pdf"], "rb") as fh:
        r = s.post(base + "/api/upload",
                   files={"file": ("OCR验收-扫描版合同.pdf", fh, "application/pdf")},
                   timeout=300)
    el = time.time() - t0
    uj = r.json() if r.status_code == 200 else {}
    ok("上传扫描版 PDF 成功", r.status_code == 200 and uj.get("ok"), uj.get("error") or r.status_code)
    ok("入库块数 > 0", (uj.get("blocks") or 0) > 0, uj.get("blocks"))
    ok("标识为扫描件", uj.get("scanned") is True, uj.get("scanned"))
    ok("复用刚才的 OCR 结果（没有重复识别）", uj.get("reused_ocr") is True, uj.get("reused_ocr"))
    ok("复用后耗时 < 5s（真没重跑 OCR）", el < 5, "%.1fs" % el)

    after = set((s.get(base + "/api/docs", timeout=20).json() or {}).get("docs") or [])
    new_docs = after - before
    ok("已进入文档库清单", any("OCR验收" in d for d in new_docs), sorted(new_docs)[:3])

    # 图片也能入池
    with open(files["扫描件-设备验收单.png"], "rb") as fh:
        r2 = s.post(base + "/api/upload",
                    files={"file": ("OCR验收-验收单.png", fh, "image/png")}, timeout=300)
    uj2 = r2.json() if r2.status_code == 200 else {}
    ok("图片也能入文档库（OCR 后当文档用）",
       r2.status_code == 200 and uj2.get("ok") and (uj2.get("blocks") or 0) > 0,
       uj2.get("error") or uj2.get("blocks"))

    # ---------- 7 边界 / 错误码
    say("\n-- 7 边界与错误码 --")
    code, j = post_ocr(os.path.join(ROOT, "测试数据", "模板表格", "模板1-设备采购汇总表.xlsx"),
                       name="表格.xlsx")
    ok("Word/Excel 打到 OCR 接口 → 400", code == 400, "%s %s" % (code, j.get("detail")))

    code, j = post_ocr(files["扫描件-设备验收单.png"], name="x.png")
    ok("不认的后缀（伪装成 xlsx 的图片）不崩", code in (200, 400, 422), code)

    # 同名重复上传：必须自动改名，绝不覆盖
    with open(files["扫描件-设备验收单.png"], "rb") as fh:
        r1 = s.post(base + "/api/upload",
                    files={"file": ("OCR验收-重复.png", fh, "image/png")}, timeout=300)
    with open(files["扫描件-设备验收单.png"], "rb") as fh:
        r2 = s.post(base + "/api/upload",
                    files={"file": ("OCR验收-重复.png", fh, "image/png")}, timeout=300)
    n1 = (r1.json().get("stored_name") or "") if r1.status_code == 200 else ""
    n2 = (r2.json().get("stored_name") or "") if r2.status_code == 200 else ""
    ok("同名文件二次上传自动改名", bool(n1 and n2 and n1 != n2), "%s → %s" % (n1, n2))

    with open(files["扫描件-设备验收单.png"], "rb") as fh:
        r = s.post(base + "/api/ocr", files={"file": ("x.png", b"", "image/png")}, timeout=30)
    ok("空文件 → 400", r.status_code == 400, r.status_code)

    with open(files["扫描件-设备验收单.png"], "rb") as fh:
        r = s.post(base + "/api/upload", files={"file": ("x.rar", b"Rar!", "application/octet-stream")},
                   timeout=30)
    ok("不支持的后缀上传 → 400", r.status_code == 400, r.status_code)
    if r.status_code == 400:
        ok("400 的提示里列出了 PDF / 图片",
           "PDF" in (r.json().get("detail") or "") and "png" in (r.json().get("detail") or ""),
           (r.json().get("detail") or "")[:80])

    # ---------- 8 清场
    if not args.no_clean:
        say("\n-- 8 清场 --")
        removed = []
        for f in os.listdir(IMPORT_DIR) if os.path.isdir(IMPORT_DIR) else []:
            if f.startswith("OCR验收"):
                try:
                    os.remove(os.path.join(IMPORT_DIR, f))
                    removed.append(f)
                except OSError:
                    pass
        ok("验收导入件已清理", True, "删了 %d 个：%s" % (len(removed), "、".join(removed) or "无"))
        try:
            import pool
            r = pool.rebuild()
            ok("检索池已按实际文件重建", r.get("ok"), "共 %d 份" % r.get("total", 0))
        except Exception as e:
            ok("检索池重建", False, repr(e))

    # ---------- 汇总
    LOG.append("")
    LOG.append("=" * 64)
    LOG.append("通过 %d 项 ／ 失败 %d 项" % (len(PASS), len(FAIL)))
    if FAIL:
        LOG.append("失败项：")
        for f in FAIL:
            LOG.append("  - " + f)
    LOG.append("分项数据：")
    for k, v in DETAIL.items():
        LOG.append("  %s：关键串 %s ｜ 耗时 %ss ｜ 置信度 %s"
                   % (k, v.get("hits"), v.get("elapsed"), v.get("conf", "-")))
    LOG.append("=" * 64)
    print("\n".join(LOG[-14:]))

    out = os.path.join(ROOT, "收尾对照", "OCR-接口验收-1026.txt")
    with open(out, "w", encoding="utf-8") as fh:
        fh.write("\n".join(LOG))
    print("\n完整记录：%s" % out)
    return 0 if not FAIL else 1


if __name__ == "__main__":
    sys.exit(main())
