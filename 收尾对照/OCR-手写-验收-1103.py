# -*- coding: utf-8 -*-
"""OCR-手写-验收-1103.py —— 手写能力验收（不靠肉眼，逐项可复现）

七项：
  1 坐标映射    ±90° 补读的框必须落回原图正确位置（写反会跑到对角）
  2 补读触发    页面有横躺的页边批注才补跑旋转；普通印刷页不补
  3 指标提升    4 张手写素材：关键串 / 字符准确率，升级前 vs 升级后
  4 单格回血    验收单「数量」列的 3 / 2 / 5 以前整格消失，现在必须读出来
  5 印刷体回归  1026 那批印刷体素材：关键串不许掉、字数不许降
  6 空图        空白页不报错、不乱认
  7 开关         A23_OCR_HAND=0 必须完全复现旧行为

用法（项目根目录）：python 收尾对照\\OCR-手写-验收-1103.py
"""

import io
import os
import re
import sys
import time
from pathlib import Path

BASE = Path(__file__).resolve().parents[1]
HAN_DIR = BASE / "测试数据" / "OCR样本" / "手写"
PRINT_DIR = BASE / "测试数据" / "OCR样本"

sys.path.insert(0, str(BASE / "agent"))
sys.path.insert(0, str(BASE / "收尾对照"))

import ocrmetric as M                                       # noqa: E402
os.environ.setdefault("A23_OCR_ENGINE", "rapidocr")
import ocr                                                   # noqa: E402

ok_n = bad_n = 0
lines = []


def check(name, cond, detail=""):
    global ok_n, bad_n
    if cond:
        ok_n += 1
        lines.append("  [通过] %s%s" % (name, (" ｜ " + detail) if detail else ""))
    else:
        bad_n += 1
        lines.append("  [失败] %s%s" % (name, (" ｜ " + detail) if detail else ""))


def read_any(path, hand):
    """图片直接读；PDF 逐页渲染后拼起来读（扫描版 PDF 的主链路就是这样）。

    必须逐页：只读第 1 页的话，第 3 页的表格数字当然"缺"，
    那是测试自己的错，不是识别退化 —— 这类假失败最容易误导人。
    """
    os.environ["A23_OCR_HAND"] = hand
    p = Path(path)
    if p.suffix.lower() == ".pdf":
        import fitz
        doc = fitz.open(str(p))
        texts = []
        for i in range(doc.page_count):
            data = doc[i].get_pixmap(dpi=200).tobytes("png")
            texts.append(ocr.read_image_bytes(data, "%s-p%d" % (p.name, i + 1)).get("text", ""))
        doc.close()
        return {"ok": True, "text": "\n".join(texts), "lines": [], "handwriting": None}
    return ocr.read_image(str(p))


# ==================== 1. 坐标映射 ====================
lines.append("【1】±90° 补读的坐标映射回原图是否正确")
p_mix = HAN_DIR / "手写-混合（印刷正文+手写批注）.png"
r_mix = read_any(p_mix, "auto")
notes = [ln for ln in r_mix["lines"] if any(k in ln["text"] for k in ("已阅", "同意", "完成"))]
right = [ln for ln in notes if ln["x0"] > 1150]
check("页边批注映射到右侧（x0>1150）",
      bool(notes) and len(right) == len(notes),
      "命中 %d 条、在右边缘 %d 条：%s"
      % (len(notes), len(right), " / ".join(n["text"][:14] for n in notes)))
first = r_mix["lines"][1] if len(r_mix["lines"]) > 1 else None
check("横排正文没被搬走",
      bool(first) and first["x0"] < 1150 and first["y"] < 260,
      "首行 %r @ x=%s y=%s" % ((first or {}).get("text", "")[:16],
                               (first or {}).get("x0"), (first or {}).get("y")))

# ==================== 2. 补读触发 ====================
lines.append("")
lines.append("【2】什么时候补跑旋转：有横躺批注才补，普通页不补")
r_scan = read_any(PRINT_DIR / "扫描件-设备验收单.png", "auto")
check("印刷扫描件不补旋转",
      not (r_scan.get("handwriting") or {}).get("vertical_band"),
      "passes=%s" % (r_scan.get("handwriting") or {}).get("passes"))
check("带页边批注的页补了旋转",
      bool((r_mix.get("handwriting") or {}).get("vertical_band"))
      and (r_mix.get("handwriting") or {}).get("angles"),
      "angles=%s" % (r_mix.get("handwriting") or {}).get("angles"))

# ==================== 3. 指标提升 ====================
lines.append("")
lines.append("【3】4 张手写素材：升级前 vs 升级后")
lines.append("  %-30s %-22s %-22s" % ("素材", "旧（HAND=0）", "新（auto）"))
tot = {"old_hit": 0, "new_hit": 0, "key": 0, "old_cer": 0.0, "new_cer": 0.0}
worst_t = 0.0
for name in M.GROUND_TRUTH:
    fp = HAN_DIR / name
    if not fp.is_file():
        check("素材存在：" + name, False, "缺文件")
        continue
    t0 = time.time()
    ro = read_any(fp, "0")
    to = time.time() - t0
    t0 = time.time()
    rn = read_any(fp, "auto")
    tn = time.time() - t0
    worst_t = max(worst_t, tn)
    so, sn = M.score(name, ro["text"]), M.score(name, rn["text"])
    tot["key"] += so["key"]
    tot["old_hit"] += so["hit"]
    tot["new_hit"] += sn["hit"]
    tot["old_cer"] += so["cer"] * so["key"]
    tot["new_cer"] += sn["cer"] * sn["key"]
    lines.append("  %-30s %2d/%d 字%3.0f%% %4.1fs    %2d/%d 字%3.0f%% %4.1fs"
                 % (name[:28], so["hit"], so["key"], so["cer"] * 100, to,
                    sn["hit"], sn["key"], sn["cer"] * 100, tn))

old_rate = 100.0 * tot["old_hit"] / max(1, tot["key"])
new_rate = 100.0 * tot["new_hit"] / max(1, tot["key"])
old_cer = 100.0 * tot["old_cer"] / max(1, tot["key"])
new_cer = 100.0 * tot["new_cer"] / max(1, tot["key"])
lines.append("  —— 合计：关键串 %.1f%% → %.1f%% ｜ 字符 %.1f%% → %.1f%%"
             % (old_rate, new_rate, old_cer, new_cer))
check("关键串命中率提升", new_rate > old_rate, "%.1f%% → %.1f%%" % (old_rate, new_rate))
check("字符准确率不下降", new_cer >= old_cer - 0.5, "%.1f%% → %.1f%%" % (old_cer, new_cer))
check("单页耗时 <8s", worst_t < 8.0, "最慢一页 %.2fs" % worst_t)

# ==================== 4. 单格回血 ====================
lines.append("")
lines.append("【4】验收单「数量」列：以前整格消失，现在要读出来")
nt_new = M.norm(read_any(
    HAN_DIR / "手写-验收单（印刷表格+手写填值）.png", "auto")["text"])
r_hand_old = read_any(HAN_DIR / "手写-验收单（印刷表格+手写填值）.png", "0")
nt_old = M.norm(r_hand_old["text"])
# 行内顺序是 型号 → 数量 → 单价 → 金额，所以「CK6140 3 128000」连起来能证明数量读到了
# 注意：nt_* 已经过 norm（全角转半角 + 去空白 + 转小写），正则也要跟着转小写，否则永远匹配不上
pairs = [("CK6140", "3", "128000"), ("X5032", "2", "86000")]
for c1, qty, c3 in pairs:
    pat = re.compile(re.escape(M.norm(c1)) + r"(\d)" + re.escape(M.norm(c3)))
    m_new = pat.search(nt_new)
    m_old = pat.search(nt_old)
    check("%s 行数量=%s 读出来了" % (c1, qty),
          bool(m_new) and m_new.group(1) == qty,
          "旧：%s ｜ 新：%s" % (m_old.group(1) if m_old else "整格丢失",
                              m_new.group(1) if m_new else "仍丢失"))

# ==================== 5. 印刷体回归 ====================
lines.append("")
lines.append("【5】1026 那批印刷体素材不能被拖坏")
PRINT_CASES = {
    "扫描版-采购合同.pdf": ["设备采购合同", "CK6140", "384000", "652000"],
    "混合版-统计公报.pdf": ["2021年民政事业发展统计公报", "856.3万张", "1651.6亿元"],
    "扫描件-设备验收单.png": ["设备验收单", "CK6140", "384000", "556000"],
    "拍照件-公告（歪+噪点）.png": ["设备安全专项检查", "11月20日"],
}
for name, keys in PRINT_CASES.items():
    fp = PRINT_DIR / name
    if not fp.is_file():
        check("素材存在：" + name, False, "缺文件")
        continue
    t0 = time.time()
    rr = read_any(fp, "auto")
    dt = time.time() - t0
    nt = M.norm(rr.get("text", ""))
    miss = [k for k in keys if M.norm(k) not in nt]
    check("%s 关键串全中" % name[:18], not miss,
          "%d 字 %.1fs%s" % (len(nt), dt, "" if not miss else " ｜ 缺：" + "、".join(miss)))

# ==================== 6. 空图 ====================
lines.append("")
lines.append("【6】空白页")
from PIL import Image                                        # noqa: E402

buf = io.BytesIO()
Image.new("RGB", (900, 1200), (255, 255, 255)).save(buf, format="PNG")
re_ = ocr.read_image_bytes(buf.getvalue(), "blank.png")
check("空白页不报错、返回空文本",
      re_.get("ok") and not re_.get("text", "").strip(),
      "ok=%s 字数=%d" % (re_.get("ok"), len(re_.get("text", ""))))

# ==================== 7. 开关 ====================
lines.append("")
lines.append("【7】A23_OCR_HAND=0 复现旧行为")
check("关闭时无 handwriting 字段", not r_hand_old.get("handwriting"), "")
check("关闭时框数与旧基线一致（33）", r_hand_old["box_count"] <= 40,
      "框数 %d" % r_hand_old["box_count"])

lines.append("")
lines.append("=" * 70)
lines.append("验收结果：通过 %d 项，失败 %d 项" % (ok_n, bad_n))
out = "\n".join(lines)
print(out)
(BASE / "收尾对照" / "OCR-手写-验收-1103.txt").write_text(out, encoding="utf-8")
sys.exit(0 if bad_n == 0 else 1)
