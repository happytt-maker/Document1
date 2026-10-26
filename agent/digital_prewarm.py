# -*- coding: utf-8 -*-
"""把小橘最常说的几句先合成好（欢迎语 + 6 条常见问题答案）。

改过 support.AI_FAQ 的文案之后跑一次，现场点快捷指令就是「立即出声」，
不用等 1~2 秒的联网合成。
"""

import time

import digital_tts
import support

LINES = ["你好呀，我是小橘，有什么想问的直接说。"]
LINES += [f["a"] for f in support.AI_FAQ]

t0 = time.time()
hit = 0
for i, line in enumerate(LINES, 1):
    t = time.time()
    p = digital_tts.speak(line)
    if p:
        hit += 1
        print("%d/%d  %5.2fs  %6d B  %s" % (
            i, len(LINES), time.time() - t, p.stat().st_size, line[:18].replace("\n", " ")), flush=True)
    else:
        print("%d/%d  合成失败：%s" % (i, len(LINES), line[:24]), flush=True)

print("\n共 %d 句，成功 %d，用时 %.1fs" % (len(LINES), hit, time.time() - t0))
print("缓存：%s" % digital_tts.stats())
