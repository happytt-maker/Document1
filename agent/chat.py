# -*- coding: utf-8 -*-
"""
chat.py —— 模块（1）· 第 5 步：交互式会话（多轮连续指令 + 历史入库）

"不用每句都提文档名"：会话记住上一句操作的文档，
本句说"把它的标题居中"也能正确执行（指代消解 = 模型上下文 + 规则兜底 双保险）。

两种模式：
  交互模式（真人用）：
    python agent/chat.py
      直接聊天式输入指令；特殊命令：
        :docs     看文档库清单
        :current  看会话当前锁定的文档
        :history  看最近的数据库记录
        :q        退出
  脚本模式（自动验收 / 演示）：
    python agent/chat.py --script agent/会话测试.json
      顺序执行一组指令（模拟连续多轮），全程自动

每一轮做的事：
  ① 组装会话上下文（current_doc）→ ② 解析+执行（复用 edit.run_edit）
  → ③ 更新会话（本轮锁定的文档 → 成为新的 current_doc）→ ④ 写 MySQL 历史

数据库：agent/db.py（库 a23_agent · 表 commands）；配置：agent/db_config.json
"""

import argparse
import json
import sys
import time
from datetime import datetime
from pathlib import Path

BASE = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(BASE / "agent"))

import db                                                        # noqa: E402
from edit import run_edit, scan_docs, DEFAULT_DOCS_DIRS          # noqa: E402

# 输出兜底：终端表示不了的字符就替换，不让程序崩
try:
    sys.stdout.reconfigure(errors="replace")
except Exception:
    pass


class Session:
    """一个会话的上下文：连续几条指令共享的"记忆"。"""

    def __init__(self, docs_dir: Path, session_id: str = None):
        self.docs_dir = docs_dir
        self.session_id = session_id or ("s" + datetime.now().strftime("%m%d-%H%M%S"))
        self.turn = 0                 # 会话内轮次
        self.current_doc = None       # 当前锁定的文档（上一句操作过的）
        self.work_copy = None         # 工作副本路径（多轮累积修改的"当前版"）
        self.work_doc = None          # 工作副本归属的文档名
        self.last_hint = None         # 上一轮的目标位置（"把它的加粗去掉"类指代句要用）

    def context(self):
        """给解析器的会话上下文（没有锁定文档时给 None）。"""
        if not self.current_doc:
            return None
        ctx = {"current_doc": self.current_doc}
        if self.last_hint:
            ctx["last_hint"] = self.last_hint
        return ctx

    def run_turn(self, text: str, image: str = None):
        """跑一轮：解析 → 执行 → 更新上下文 → 入库。返回 (ok, report, meta, 耗时s)。

        image（可选）＝ 本轮要插入的图片引用（上传后的存储名/文件名），
        只在指令是"插入图片"时用到；不传则执行层自动取最近上传的一张。
        """
        self.turn += 1
        meta = {}
        # 工作副本（多轮累积）：本轮如果接着同一份文档改，就基于上一轮产物出发
        workcopy = None
        if self.work_copy is not None and self.work_doc is not None:
            workcopy = {"doc": self.work_doc, "path": self.work_copy}

        t0 = time.time()
        ok, report = run_edit(text, self.docs_dir, context=self.context(),
                              meta=meta, workcopy=workcopy, image=image)
        elapsed = time.time() - t0

        # 更新会话：只要本轮"锁定到了某文档"，它就是新的 current_doc
        # （哪怕执行失败——用户仍然是在弄这份文档，下一句的"它"还指它）
        if meta.get("doc"):
            if meta["doc"] != self.work_doc:
                # 本轮操作的是别的文档 → 工作副本重新起点（回到该文档的原件）
                self.work_doc = meta["doc"]
                self.work_copy = None
            self.current_doc = meta["doc"]

        # 工作副本推进：本轮产出的文件 = 下一轮的起点（修改逐轮累积）
        if meta.get("out_path"):
            self.work_copy = Path(meta["out_path"])

        # 记录本轮目标位置（下一轮"把它的加粗去掉"类指代句要用；多操作取第一个任务）
        intent = meta.get("intent") or {}
        h = ""
        loc0 = intent.get("locate")
        if isinstance(loc0, dict):
            h = (loc0.get("hint") or "").strip()
        if not h:
            tasks = intent.get("tasks")
            if isinstance(tasks, list) and tasks and isinstance(tasks[0], dict):
                tl = tasks[0].get("locate")
                if isinstance(tl, dict):
                    h = (tl.get("hint") or "").strip()
        if h:
            self.last_hint = h

        # 写数据库
        note = ""
        if intent.get("action") == "不明确":
            note = "反问：" + (intent.get("clarify") or "")
        elif not ok:
            note = "未完成"
        try:
            db.log_command(self.session_id, self.turn, text,
                           action=intent.get("action"), op=intent.get("op"),
                           doc=meta.get("doc"), out_file=meta.get("out_path"),
                           ok=ok, elapsed_sec=round(elapsed, 2), note=note)
        except RuntimeError as e:
            report += "\n【数据库】写入失败：%s" % e

        return ok, report, meta, elapsed

    def summary_line(self, ok: bool, elapsed: float) -> str:
        return (">> %s ｜ 耗时 %.1fs ｜ 当前会话文档 = %s" %
                ("成功" if ok else "未完成", elapsed,
                 self.current_doc or "(未锁定)"))


def run_script(sess: Session, script_path: Path):
    """脚本模式：顺序跑一组指令（自动验收用）。"""
    data = json.loads(script_path.read_text(encoding="utf-8"))
    steps = data.get("steps") if isinstance(data, dict) else data
    name = data.get("name", script_path.name) if isinstance(data, dict) else script_path.name

    print("=" * 64)
    print("会话脚本：%s" % name)
    print("共 %d 步 ｜ 会话号 %s ｜ 文档库 %s" % (len(steps), sess.session_id, sess.docs_dir))
    print("=" * 64)

    ok_n = 0
    for i, text in enumerate(steps, 1):
        print()
        print("-" * 64)
        print("### 第 %d 轮：%s" % (i, text))
        print("-" * 64)
        ok, report, meta, elapsed = sess.run_turn(text)
        print(report)
        print()
        print(sess.summary_line(ok, elapsed))
        if ok:
            ok_n += 1

    print()
    print("=" * 64)
    print("会话结束：共 %d 轮 ｜ 成功 %d 轮 ｜ 当前文档 = %s" %
          (sess.turn, ok_n, sess.current_doc or "(未锁定)"))
    print("历史已入库 MySQL（库 a23_agent）｜ 查看：python agent/db.py --recent %d" % (sess.turn + 2))
    print("=" * 64)


def run_interactive(sess: Session):
    """交互模式：真人聊天式下指令。"""
    print("=" * 64)
    print("A23 智能文档助手 · 交互模式（会话号 %s）" % sess.session_id)
    print("直接输入指令，例如：把设备采购合同的标题居中")
    print("然后可以只说过：把它的金额加粗（'它'= 上一句操作的文档）")
    print("特殊命令：:docs 文档库 ｜ :current 当前文档 ｜ :history 最近记录 ｜ :q 退出")
    print("=" * 64)

    while True:
        try:
            text = input("\n指令> ").strip()
        except (EOFError, KeyboardInterrupt):
            print()
            break
        if not text:
            continue
        if text in (":q", ":quit", "exit", "quit"):
            break
        if text == ":docs":
            for n in scan_docs(sess.docs_dir):
                print("  " + n)
            continue
        if text == ":current":
            print("  当前会话文档：" + (sess.current_doc or "(还没锁定)"))
            continue
        if text == ":history":
            try:
                db.print_rows(db.recent(10))
            except RuntimeError as e:
                print("【错误】%s" % e)
            continue

        ok, report, meta, elapsed = sess.run_turn(text)
        print()
        print(report)
        print()
        print(sess.summary_line(ok, elapsed))

    print()
    print("会话结束（共 %d 轮）。历史已存入 MySQL（库 a23_agent）。" % sess.turn)


def main():
    ap = argparse.ArgumentParser(description="A23 智能文档助手 · 多轮会话")
    ap.add_argument("--docs", default=None,
                    help="文档库目录（默认：源文档 + 导入文档 双目录）")
    ap.add_argument("--script", help="脚本模式：顺序跑 JSON 里的多条指令（自动验收用）")
    ap.add_argument("--session", help="会话号（默认按时间生成）")
    args = ap.parse_args()

    # 数据库先备好（连不上也不拦路——指令照常执行，只是历史功能不可用）
    try:
        db.ensure_db()
    except RuntimeError as e:
        print("【数据库】%s" % e)
        print("（提示：历史记录功能不可用，但指令照常执行）")

    docs = Path(args.docs) if args.docs else DEFAULT_DOCS_DIRS
    sess = Session(docs, session_id=args.session)
    if args.script:
        run_script(sess, Path(args.script))
    else:
        run_interactive(sess)


if __name__ == "__main__":
    main()
