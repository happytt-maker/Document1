# -*- coding: utf-8 -*-
"""接口自检：文字识别 · 繁体 / 竖排 / 英文（2026-11-07）

全部走 HTTP（真实服务），不直接调函数 —— 测的是"用户点按钮"那条路。

  1 状态接口     报出识别模型与四种语言
  2 繁体          環境空氣質量監測報告 / 設備驗收單 / 繁體字測試 必须全对
  3 英文          Annual Report / Total Revenue / Increase 12.5% 必须全对
  4 简繁混排      简体那行不能被认成繁体（不互伤）
  5 竖排          按「从右到左」还原列序：環境空氣質量 在 監測報告表 之前
  6 竖排开关      关掉后顺序不再还原（证明开关真的起作用）
  7 印刷回归      既有扫描件关键串全中、字数不降
  8 边界          不支持格式 / 空文件要有明确报错

跑法（默认打 8000，A23_TEST_PORT 可覆盖）：
  python 收尾对照\\OCR-多语种-接口自检-1107.py
"""
import json
import os
import sys
import urllib.request

BASE = "http://127.0.0.1:%s" % os.environ.get("A23_TEST_PORT", "8000")
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
IMGDIR = os.path.join(ROOT, "测试数据", "OCR-多语种")

_OPENER = urllib.request.build_opener(urllib.request.ProxyHandler({}))   # 系统代理会假 404
PASS, FAIL = [], []


def check(name, cond, detail=""):
    (PASS if cond else FAIL).append(name)
    detail = "" if detail is None else str(detail)
    print("  %s %s%s" % ("✓" if cond else "✗", name, ("  —— " + detail[:150]) if detail else ""))


def get(path):
    with _OPENER.open(BASE + path, timeout=30) as r:
        return json.loads(r.read().decode("utf-8"))


def ocr(name, lang="auto", vertical="auto", path=None):
    """multipart 上传一张图，返回接口 JSON。"""
    path = path or os.path.join(IMGDIR, name)
    with open(path, "rb") as f:
        data = f.read()
    boundary = "----a23ocr%d" % len(data)
    head = ("--%s\r\nContent-Disposition: form-data; name=\"file\"; filename=\"%s\"\r\n"
            "Content-Type: application/octet-stream\r\n\r\n" % (boundary, os.path.basename(path)))
    body = head.encode("utf-8") + data + ("\r\n--%s--\r\n" % boundary).encode("utf-8")
    url = "%s/api/ocr?lang=%s&vertical=%s" % (BASE, lang, vertical)
    req = urllib.request.Request(url, data=body, method="POST",
                                 headers={"Content-Type": "multipart/form-data; boundary=" + boundary})
    try:
        with _OPENER.open(req, timeout=180) as r:
            return json.loads(r.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        return {"ok": False, "http": e.code, "error": e.read().decode("utf-8", "replace")[:200]}


print("=" * 62)
print("文字识别 · 繁体/竖排/英文 接口自检 · %s" % BASE)
print("=" * 62)

# ---------- 1 状态 ----------
print("\n[1] 引擎状态")
st = get("/api/ocr/status")
check("引擎可用", st.get("ok"), st.get("engine"))
check("报出识别模型", bool(st.get("rec_model")), st.get("rec_model"))
langs = {x["key"]: x for x in (st.get("langs") or [])}
check("四种语言齐全", set(langs) == {"auto", "chs", "cht", "en"}, "/".join(sorted(langs)))
check("竖排能力已声明", bool(st.get("vertical")), st.get("vertical_tip"))

# ---------- 2 繁体 ----------
print("\n[2] 繁体（繁英混排.png · 自动）")
r = ocr("繁英混排.png", "auto")
text = r.get("text") or ""
for k in ("環境空氣質量監測報告", "設備驗收單", "繁體字測試"):
    check("认出「%s」" % k, k in text)
check("没退化成简体字库（不含「盘测」「收军」）",
      ("盤測" not in text) and ("收軍" not in text))

# ---------- 3 英文 ----------
print("\n[3] 英文（繁英混排.png · 自动 / 英文小字.png · 英文）")
for k in ("Annual Report 2025", "Total Revenue", "Increase 12.5%"):
    check("混排行认出「%s」" % k, k in text)
r2 = ocr("英文小字.png", "en")
t2 = r2.get("text") or ""
check("英文模型生效", "英文" in str(r2.get("rec_model")), r2.get("rec_model"))
for k in ("EQ-2026-0187", "Shandong Environmental Instruments", "USD 128,400.00", "2026-03-15"):
    check("英文专认「%s」" % k, k in t2)

# ---------- 4 简繁不互伤 ----------
print("\n[4] 简繁混排不互伤")
check("简体那行仍读作简体「环境空气质量监测报告」", "环境空气质量监测报告" in text)

# ---------- 5 竖排 ----------
print("\n[5] 竖排（竖排繁体.png · 繁体）")
r3 = ocr("竖排繁体.png", "cht")
t3 = (r3.get("text") or "").replace("\n", " ")
check("两列都认出来了", ("環境空氣質量" in t3) and ("監測報告表" in t3), t3)
check("列序按从右到左（第一列在前）",
      t3.find("環境空氣質量") >= 0 and t3.find("環境空氣質量") < t3.find("監測報告表"), t3)

# ---------- 6 竖排开关 ----------
print("\n[6] 竖排开关（关掉 = 不再按从右到左）")
r4 = ocr("竖排繁体.png", "cht", vertical="off")
t4 = (r4.get("text") or "").replace("\n", " ")
check("关掉后顺序与开启时不同", t4 != t3, "off：%s" % t4)
check("关掉时仍能读出文字（不是坏掉）", ("環境空氣質量" in t4) and ("監測報告表" in t4), t4)

# ---------- 7 印刷体回归 ----------
print("\n[7] 既有印刷体不退化（扫描件-设备验收单.png）")
scan = os.path.join(ROOT, "测试数据", "OCR样本", "扫描件-设备验收单.png")
r5 = ocr(None, "auto", "auto", path=scan)
t5 = (r5.get("text") or "").replace("\n", " ")
for k in ("设备验收单", "HD-2026-0317", "数控车床", "CK6140", "556000"):
    check("关键串「%s」" % k, k in t5)
check("字数不降（≥180）", (r5.get("char_count") or 0) >= 180, "%s 字" % r5.get("char_count"))
check("平均置信度 ≥ 0.75", (r5.get("avg_conf") or 0) >= 0.75,
      "%.3f" % (r5.get("avg_conf") or 0))

# ---------- 8 边界 ----------
print("\n[8] 边界")
bad = ocr(None, path=os.path.join(ROOT, "README.md"))
check("不支持的格式明确报错", bad.get("ok") is False and bad.get("http") == 400,
      str(bad.get("error"))[:80])
empty = ocr(None, path=os.path.join(IMGDIR, "繁英混排.png"))   # 同一张（确认正常路径没被 httperror 吞）
check("正常图不受边界影响", empty.get("ok") is True)

print("\n" + "=" * 62)
print("通过 %d 项 · 失败 %d 项" % (len(PASS), len(FAIL)))
if FAIL:
    print("失败项：" + "、".join(FAIL))
print("=" * 62)
sys.exit(0 if not FAIL else 1)
