# -*- coding: utf-8 -*-
"""实验：怎么判断"这页有手写字被阈值吞了"，又不用多跑一遍。

关键发现（这决定了实现方式）：
  快速那一遍用 text_score=0.5，低分结果在引擎内部就被丢了，外面看不见，
  所以"置信度分布"这个信号自己消失了，没法当触发条件。

  但引擎的过滤是**最后一步**（filter_boxes_rec_by_score）。
  所以把 text_score 设 0，一次就能拿到"全部框 + 各自的分数"，
  然后自己数"有多少低于 0.5" —— 这就是被吞掉的比例，零额外开销：
    · 印刷体：低于 0.5 的极少（识别都很确定）
    · 手写  ：一批单字停在 0.2~0.5（模型认得出来，但不够自信）

  而且这一步对印刷体是**行为等价**的：拿到全部后再按 0.5 筛一遍，
  结果和原来一模一样（同样的 det 框、同样的 rec 结果），只是多了一次列表过滤。

用法：python 收尾对照\OCR-手写-墨迹实验-1103.py
"""
import shutil
import sys
from pathlib import Path

BASE = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(BASE / "agent"))
import ocr  # noqa: E402

HAND = BASE / "测试数据" / "OCR样本" / "手写"
PRINT = BASE / "测试数据" / "OCR样本"

cases = [("手写", p) for p in sorted(HAND.glob("*.png"))]
cases += [("印刷", PRINT / nm) for nm in
          ("扫描件-设备验收单.png", "拍照件-公告（歪+噪点）.png")]

ocr.engine_now()          # 先把引擎定下来（_run_rapidocr 直接用 _ENGINE_OBJ）

lines = []
for tag, p in cases:
    png, w, h, tmpdir = ocr._prepare(src_path=p)
    try:
        raw = ocr._run_rapidocr(png, {"text_score": 0.0, "box_thresh": 0.30})
        keep = [it for it in raw if (it[2] or 0) >= 0.5]
        low = len(raw) - len(keep)
        share = low / len(raw) if raw else 0.0
        lines.append("%-4s %-34s 全框%3d 低于0.5 %2d 占 %4.0f%% ｜ 保留 %3d（应与旧版一致）"
                     % (tag, p.stem[:32], len(raw), low, share * 100, len(keep)))
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)

print("\n".join(lines))
(BASE / "收尾对照" / "_手写阈值实验.txt").write_text("\n".join(lines), encoding="utf-8")
