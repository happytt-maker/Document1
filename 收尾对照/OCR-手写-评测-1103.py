# -*- coding: utf-8 -*-
"""
OCR-手写-评测-1103.py —— 同一个标准答案，量不同 OCR 引擎在手写上的差距

为什么要单独一套评测：
  印刷体那份（1026）只验"认不认得出来"，手写要验的是"差多少"。
  所以这里用两个互补的指标：
    ① 关键串命中率 —— 标准答案里的每个词，在识别结果里能不能原样找到
       （去空白、全角转半角后再比；这是业务最关心的：关键信息有没有丢）
    ② 字符准确率 —— 1 - 编辑距离 / 答案长度（看整体认对了几成字）
  两个都算，是因为"关键串全中但整篇乱"和"整篇像但关键数字错"都出现过。

引擎用 --engine 选：
  rapidocr   现有 PP-OCRv4（基线）
  ppocr5     RapidOCR 2.x / PP-OCRv5（如果装了）
  vlm        qwen-vl-ocr（走 agent/llm_config.py 的 key）

跑法（项目根目录）：
  python 收尾对照/OCR-手写-评测-1103.py --engine rapidocr
  python 收尾对照/OCR-手写-评测-1103.py --engine rapidocr --out 收尾对照/手写评测-基线.txt
"""

import argparse
import sys
import time
import unicodedata
from pathlib import Path

BASE = Path(__file__).resolve().parents[1]
IMG_DIR = BASE / "测试数据" / "OCR样本" / "手写"

# 标准答案（与素材生成脚本保持一致；改素材要同步改这里）
GROUND_TRUTH = {
    "手写-验收单（印刷表格+手写填值）.png": [
        "设备验收单", "机械加工实训中心设备采购", "HD-2026-0317",
        "2026年10月18日", "张建国", "机械工程学院", "运行正常",
        "符合合同约定，同意验收",
        "数控车床", "CK6140", "384000", "立式铣床", "X5032", "172000", "556000",
    ],
    "手写-会议记录（整篇手写）.png": [
        "设备安全专项检查会议记录", "2026年10月24日", "实验楼B座102室",
        "张建国", "李红梅", "王海涛",
        "本月完成设备巡检126台", "发现隐患3处",
        "立式铣床X5032主轴异响", "需更换轴承",
        "下周组织安全培训",
    ],
    "手写-报销单（数字+涂改）.png": [
        "差旅费", "386.50", "设备配件采购", "1280.00", "快递费", "45.00", "1711.50",
    ],
    "手写-混合（印刷正文+手写批注）.png": [
        "关于开展设备安全专项检查的通知", "实验室设备安全管理",
        "2026年11月1日至11月20日", "自查自纠", "集中检查",
        "已阅", "同意", "11月20日前必须完成",
    ],
}

# 每个文件"应该认出多少字"（用来算字符准确率的分母），由生成脚本文本累加而来
FULL_TEXT = {
    "手写-验收单（印刷表格+手写填值）.png":
        "设备验收单 项目名称：机械加工实训中心设备采购 合同编号：HD-2026-0317 "
        "验收日期：2026年10月18日 验收人：张建国 使用部门：机械工程学院 "
        "设备状态：运行正常 验收结论：符合合同约定，同意验收 "
        "序号 设备名称 型号 数量 单价（元） 金额（元） "
        "1 数控车床 CK6140 3 128000 384000 2 立式铣床 X5032 2 86000 172000 "
        "合计 — — 5 — 556000",
    "手写-会议记录（整篇手写）.png":
        "设备安全专项检查会议记录 时间：2026年10月24日 下午三点 "
        "地点：实验楼B座102室 参会人：张建国、李红梅、王海涛 "
        "一、本月完成设备巡检126台，发现隐患3处 "
        "二、立式铣床X5032主轴异响，需更换轴承 "
        "三、下周组织安全培训，全员参加 记录人：李红梅",
    "手写-报销单（数字+涂改）.png":
        "费用报销单 报销人：王海涛 部门：机械工程学院 "
        "日期 事由 金额（元） 2026-10-08 差旅费（苏州—上海） 386.50 "
        "2026-10-12 设备配件采购 1280.00 2026-10-18 快递费 45.00 "
        "合计 1711.50 审批意见：同意报销",
    "手写-混合（印刷正文+手写批注）.png":
        "关于开展设备安全专项检查的通知 各部门、各学院： "
        "为切实加强实验室设备安全管理，消除安全隐患，经研究决定， "
        "于2026年11月1日至11月20日开展设备安全专项检查。 "
        "一、检查范围：全校单价10万元以上的教学科研仪器设备。 "
        "二、工作安排：分自查自纠和集中检查两个阶段。 "
        "三、有关要求：请各部门按时报送自查报告。 "
        "已阅 同意 张建国 10.25 11月20日前必须完成",
}


def norm(s: str) -> str:
    """归一化：全角→半角、去所有空白、统一大小写。"""
    if not s:
        return ""
    s = unicodedata.normalize("NFKC", s)
    return "".join(ch for ch in s if not ch.isspace()).lower()


def lev(a: str, b: str) -> int:
    """编辑距离（滚动数组，样例长度只有几百，够快）。"""
    if a == b:
        return 0
    if not a:
        return len(b)
    if not b:
        return len(a)
    prev = list(range(len(b) + 1))
    for i, ca in enumerate(a, 1):
        cur = [i]
        for j, cb in enumerate(b, 1):
            cur.append(min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + (ca != cb)))
        prev = cur
    return prev[-1]


# ---------------- 引擎适配 ----------------

def run_rapidocr(path):
    """现有 ocr.py（PP-OCRv4）—— 走项目自己的入口，保证和线上一致。"""
    sys.path.insert(0, str(BASE / "agent"))
    import ocr
    r = ocr.read_image(str(path))
    if not r.get("ok"):
        return "", r.get("error") or "失败"
    return r["text"], None


def run_ppocr5(path):
    """RapidOCR 2.x（默认模型是 PP-OCRv5）。没装就如实报错。"""
    try:
        from rapidocr import RapidOCR                      # type: ignore
    except Exception as e:
        return "", "没装 rapidocr 2.x：%s" % e
    eng = RapidOCR()
    res = eng(str(path))
    txts = getattr(res, "txts", None)
    if txts is None and isinstance(res, (list, tuple)):
        txts = res
    return "\n".join(str(t) for t in (txts or [])), None


def run_vlm(path):
    """qwen-vl-ocr：整图送模型，只取文字。"""
    sys.path.insert(0, str(BASE / "agent"))
    from llm_config import chat
    import base64
    b64 = base64.b64encode(Path(path).read_bytes()).decode()
    msgs = [{
        "role": "user",
        "content": [
            {"type": "image_url",
             "image_url": {"url": "data:image/png;base64," + b64}},
            {"type": "text",
             "text": "把这张图片里的所有文字按阅读顺序原样输出，"
                     "保留数字和标点，不要解释、不要加标题。"},
        ],
    }]
    text, usage, cost, err = chat(msgs, timeout=120, max_tokens=2000)
    if err:
        return "", err
    return text or "", None


ENGINES = {"rapidocr": run_rapidocr, "ppocr5": run_ppocr5, "vlm": run_vlm}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--engine", default="rapidocr")
    ap.add_argument("--out", default="")
    ap.add_argument("--only", default="")     # 只跑文件名包含这个串的
    a = ap.parse_args()

    fn = ENGINES.get(a.engine)
    if fn is None:
        print("没有这个引擎：%s（可选 %s）" % (a.engine, "、".join(ENGINES)))
        return 1

    lines = []
    tot_hit = tot_key = 0
    tot_err = tot_len = 0
    t_all = time.time()

    for name, keys in GROUND_TRUTH.items():
        if a.only and a.only not in name:
            continue
        p = IMG_DIR / name
        if not p.is_file():
            lines.append("【跳过】找不到素材：%s" % name)
            continue
        t0 = time.time()
        text, err = fn(p)
        cost = time.time() - t0
        if err:
            lines.append("【失败】%s → %s" % (name, err))
            continue

        nt = norm(text)
        hits = [k for k in keys if norm(k) in nt]
        miss = [k for k in keys if norm(k) not in nt]
        exp = norm(FULL_TEXT.get(name, ""))
        d = lev(nt, exp) if exp else 0
        cer = (1 - d / len(exp)) if exp else 0.0

        tot_hit += len(hits)
        tot_key += len(keys)
        tot_err += d
        tot_len += len(exp)

        lines.append("")
        lines.append("=" * 70)
        lines.append("%s ｜ %s ｜ %.2fs ｜ 认出 %d 字"
                     % (name, a.engine, cost, len(nt)))
        lines.append("-" * 70)
        lines.append("关键串 %d/%d（%.0f%%）｜ 字符准确率 %.1f%%"
                     % (len(hits), len(keys), 100.0 * len(hits) / len(keys), cer * 100))
        if miss:
            lines.append("没认出来：" + "  ".join(miss))
        lines.append("--- 识别原文 ---")
        lines.append(text)

    lines.append("")
    lines.append("=" * 70)
    lines.append("【汇总】引擎 %s ｜ 关键串 %d/%d = %.1f%% ｜ 字符准确率 %.1f%% ｜ 总耗时 %.1fs"
                 % (a.engine, tot_hit, tot_key,
                    100.0 * tot_hit / max(1, tot_key),
                    (1 - tot_err / max(1, tot_len)) * 100, time.time() - t_all))

    out = "\n".join(lines)
    print(out)
    if a.out:
        (BASE / a.out).write_text(out, encoding="utf-8")
        print("\n已写入：%s" % (BASE / a.out))
    return 0


if __name__ == "__main__":
    sys.exit(main())
