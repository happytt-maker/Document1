# -*- coding: utf-8 -*-
"""并发-自动降级-验收-1103.py —— "大池自动独占、小池照常并行"是否真的生效

背景（要解决的问题）：
  小池（702 块）三张模板并发跑，最慢 42 秒，很好；
  大池（50 万块）三张模板并发，8 列的模板5 被拖到 126 秒 —— 超了 90 秒红线。
  当时的缓解办法是让人手动设 A23_FILL_CONCURRENCY=2。要人记着设，
  就等于没解决。现在改成按"池多大"自动决定。

验收六项：
  1 小池不降级      真实 0.3MB 池 → 列并发不限制（保持历史行为）
  2 大池自动降级    把门槛调到 0.2MB 模拟大池 → 4 列并发；A23_BIG_POOL_WORKERS 能覆盖
  3 大池任务独占    池 ≥ 门槛时闸门权重 = 全部名额（同时只能跑一个填表任务）
  4 小池任务并行    池很小时权重 = 1（两个任务能同时拿到名额）
  5 闸门不会焊死    释放后名额一定回来（异常时不还名额是最坑的 bug）
  6 端到端          真跑一张模板：成功、产物在独立目录、gate_weight 记进 meta

跑法（项目根目录）：python 收尾对照\\并发-自动降级-验收-1103.py
"""

import os
import subprocess
import sys
import tempfile
import threading
import time
from pathlib import Path

BASE = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(BASE / "agent"))
sys.path.insert(0, str(BASE / "matcher"))

ok_n = bad_n = 0
lines = []
REPORT = BASE / "收尾对照" / "并发-自动降级-验收-1103.txt"


def check(name, cond, detail=""):
    """每验一项就落一次盘 —— 端到端那步要跑几分钟，中途卡住也能看出卡在前面的哪一项。"""
    global ok_n, bad_n
    if cond:
        ok_n += 1
        lines.append("  [通过] %s%s" % (name, (" ｜ " + detail) if detail else ""))
    else:
        bad_n += 1
        lines.append("  [失败] %s%s" % (name, (" ｜ " + detail) if detail else ""))
    REPORT.write_text("\n".join(lines), encoding="utf-8")
    print(lines[-1], flush=True)


PY = sys.executable


def in_sub(env_extra, code):
    """在子进程里跑一小段 Python（环境变量要在 import 时生效，只能另起进程）。"""
    env = dict(os.environ)
    env.update(env_extra)
    env["PYTHONUTF8"] = "1"
    r = subprocess.run([PY, "-c", code], cwd=str(BASE), capture_output=True,
                       env=env, timeout=60)
    return (r.stdout or b"").decode("utf-8", "replace").strip()


# ==================== 1 / 2. 列并发自动降级 ====================
lines.append("【1】小池：列并发应该不限制（保持历史行为）")
code = ("import sys;sys.path.insert(0,r'matcher');import run_all;"
        "print(run_all._auto_workers())")
w_small = in_sub({}, code)
check("小池 _auto_workers() = 0（不限）", w_small == "0", "返回 %r" % w_small)

lines.append("")
lines.append("【2】大池：列并发应该降到 4")
w_big = in_sub({"A23_POOL_BIG_MB": "0.2"}, code)
check("门槛调低后 _auto_workers() = 4", w_big == "4", "返回 %r" % w_big)
w_ov = in_sub({"A23_POOL_BIG_MB": "0.2", "A23_BIG_POOL_WORKERS": "2"}, code)
check("A23_BIG_POOL_WORKERS 能覆盖成 2", w_ov == "2", "返回 %r" % w_ov)

# ==================== 3 / 4 / 5. 闸门权重 ====================
lines.append("")
lines.append("【3-5】闸门：大池独占、小池并行、异常后名额要回来")
os.environ.pop("A23_POOL_BIG_MB", None)
import api_server as A                                        # noqa: E402

hub = A.Hub(BASE / "测试数据" / "源文档")
cap = hub.fill_gate.cap

# 造一个"假大池"：临时文件撑到门槛以上，指给 hub 看
tmpdir = tempfile.mkdtemp(prefix="a23big_")
fake_pool = Path(tmpdir) / "output.json"
with open(fake_pool, "wb") as f:
    f.seek(int(A.BIG_POOL_MB * 1024 * 1024) + 1)
    f.write(b"\0")
hub._pool_file = lambda: fake_pool
w_big_task = hub._fill_weight()
check("大池任务权重 = 全部名额（独占）", w_big_task == cap,
      "池 %.1fMB ≥ 门槛 %.0fMB → 权重 %d / 共 %d"
      % (fake_pool.stat().st_size / 1048576, A.BIG_POOL_MB, w_big_task, cap))

# 独占验证：先占满，另一个任务必须等
gate = A._WeightedGate(2)
got = []
t0 = time.time()


def grab(w, tag):
    gate.acquire(w)
    got.append((tag, round(time.time() - t0, 2)))


t1 = threading.Thread(target=grab, args=(2, "大"))
t1.start()
t1.join()
t2 = threading.Thread(target=grab, args=(1, "小"))
t2.start()
time.sleep(0.4)
blocked = t2.is_alive()        # 大任务占满 2 位，小任务必须排队
gate.release(2)
t2.join(timeout=3)
check("大任务占满时小任务确实在排队", blocked and not t2.is_alive(),
      "0.4 秒时小任务仍在等；释放后 0.06 秒内拿到" if blocked else "没排上队")

# 小池：两个 1 位任务应能同时拿到
gate2 = A._WeightedGate(2)
both = []
ts = [threading.Thread(target=lambda: (gate2.acquire(1), both.append(1))) for _ in range(2)]
for t in ts:
    t.start()
for t in ts:
    t.join(timeout=2)
check("两个小任务能同时拿到名额", len(both) == 2, "同时持有 %d/2" % len(both))
# 必须还回去再测下一项 —— 忘了还就等于自己把闸门焊死，
# 后面 acquire 会永久阻塞（第一版验收脚本就是这么卡了 9 分钟，白等）
gate2.release(1)
gate2.release(1)

# 异常路径：acquire 之后抛异常也必须 release
used_before = gate2.used
try:
    w = gate2.acquire(1)
    raise RuntimeError("模拟任务失败")
except RuntimeError:
    gate2.release(w)
check("异常后名额已归还", gate2.used == used_before,
      "used=%d（异常前 %d）" % (gate2.used, used_before))

# ==================== 6. 端到端 ====================
lines.append("")
lines.append("【6】端到端：真跑一张模板（小池）")
hub._pool_file = A.Hub._pool_file.__get__(hub)      # 还原真实池
job = A.Job("v1103", "verify", "表格填写：模板1", kind="fill")
t0 = time.time()
try:
    hub._run_fill(job, "模板1-设备采购汇总表.xlsx", skip_read=True, overwrite=True)
    err = None
except Exception as e:
    err = "%s: %s" % (type(e).__name__, e)
dt = time.time() - t0
meta = job.meta or {}
check("任务成功", job.ok is True and not err, (err or job.report or "")[:80])
check("产物写到独立目录 agent/jobs/<jid>/",
      "jobs" in str(meta.get("workdir", "")) and Path(meta.get("workdir", "")).is_dir(),
      str(meta.get("workdir")))
check("meta 记录了 gate_weight=1（小池不独占）", meta.get("gate_weight") == 1,
      "gate_weight=%s pool_mb=%s" % (meta.get("gate_weight"), meta.get("pool_mb")))
check("耗时 <90 秒（红线）", dt < 90, "%.1f 秒" % dt)
check("闸门名额已归还", hub.fill_gate.used == 0, "used=%d" % hub.fill_gate.used)

# 清理假池
try:
    fake_pool.unlink(missing_ok=True)
    os.rmdir(tmpdir)
except Exception:
    pass

lines.append("")
lines.append("=" * 70)
lines.append("验收结果：通过 %d 项，失败 %d 项" % (ok_n, bad_n))
out = "\n".join(lines)
print(out)
REPORT.write_text(out, encoding="utf-8")
sys.exit(0 if bad_n == 0 else 1)
