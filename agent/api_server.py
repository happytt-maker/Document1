# -*- coding: utf-8 -*-
"""
api_server.py —— HTTP API + 异步调用 + 网页操作台（三模块服务端）

赛题对应：
  - "智能识别分类模块提供 API 接口，支持异步调用" → POST /api/recognize
  - 模块（3）表格填写 → POST /api/fill（异步跑 模板→找数→填写→自检）
  - 模块（1）自然语言操作 → POST /api/command

形态：
  - FastAPI 提供 HTTP 接口（默认本机 127.0.0.1:8000）
  - 每个任务提交后【立即返回 job_id】（不阻塞），客户端轮询任务状态拿结果
    ——长任务（提取 10~15 秒 / 填表 20~60 秒）不卡界面
  - 复用 chat.Session：多轮上下文 / 工作副本 / MySQL 历史 全部保持一致
  - 浏览器打开 http://127.0.0.1:8000/ 即为网页操作台（agent/web/index.html）

接口一览：
  GET  /                       网页操作台（未登录 → 302 到 /login）
  GET  /login                  登录 / 注册页
  POST /api/auth/register      注册 {username, password, display_name?} → 自动登录
  POST /api/auth/login         登录 {username, password} → 写 cookie
  POST /api/auth/logout        退出登录（删 token）
  GET  /api/auth/me            当前登录用户（顶栏显示）
  POST /api/auth/profile       改昵称 {display_name}
  POST /api/auth/avatar        换头像（上传图片，自动裁方缩到 256px）
  POST /api/auth/avatar/remove 恢复默认头像
  GET  /api/avatar/{文件名}     头像图片
  GET  /api/status             服务状态（含数据库连通、文档数）
  GET  /api/docs               文档库清单（标准库 + 导入库）
  POST /api/upload             上传文档 → 自动识别格式 → 解析并入检索池（模块2）
  POST /api/upload_image       上传图片 → {id}；随后指令说"把这张图插到××"（模块1）
  GET  /api/images             最近上传的图片清单
  GET  /api/images/{name}      图片预览
  GET  /api/imports            导入记录（MySQL imports 表）
  POST /api/command            提交指令 {text, session_id?, image_id?} → {job_id, status}
  POST /api/recognize          智能识别分类（异步）：文档 × 画像/实体/字段 → 自动入库
  POST /api/fill               表格填写（异步）：模板 → 读懂/找数/填写/自检全流程
  GET  /api/fill/templates     可填模板清单（Excel + Word 双形态）
  GET  /api/fill/download      下载已填写文件（?template=模板名）
  GET  /api/job/{job_id}       查任务状态/结果（轮询用，三种任务通用）
  GET  /api/history            操作历史（MySQL，?limit=15）
  GET  /api/session/{sid}      会话状态（当前锁定文档、已执行轮数）
  GET  /api/db/stats           数据仓库统计（模块2 第4步：提取结果入库）
  GET  /api/db/entities        查实体（?doc=&type=&q=&limit=）
  GET  /api/db/fields          查字段提取（?doc=&q=&limit=）
  GET  /api/db/profile         看某份文档的画像（?doc=）
  GET  /api/db/search          全库搜（?q=，实体+字段值）
  GET  /api/db/audit           质量体检（?doc=，反向审计 + 三线查全率；模块2 第5步）
  GET  /docs                   接口文档（FastAPI 自带）

并发模型：
  - 线程池 4 个工人：不同任务可并行；【同一会话内部串行】（一把会话锁），
    保证多轮"它/这份文档"的上下文不乱
  - 表格填写任务全局串行（一把填表锁）：③ 抽取的中间产物是共享文件名，互斥最稳
  - 任务状态在内存（服务重启后旧任务丢失，属预期）

启动：
  python agent/api_server.py                # 默认 127.0.0.1:8000
  python agent/api_server.py --port 8100    # 换端口
"""

import argparse
import hashlib
import json
import os
import re
import subprocess
import sys
import tempfile
import threading
import time
import itertools
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime
from pathlib import Path

BASE = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(BASE / "agent"))
sys.path.insert(0, str(BASE / "doc_reader"))       # 复用文档读取器（含 PDF / 图片 OCR）

# 输出兜底：终端表示不了的字符就替换，不让程序崩
try:
    sys.stdout.reconfigure(errors="replace")
except Exception:
    pass

from fastapi import FastAPI, HTTPException, UploadFile, File, Request, Response   # noqa: E402
from fastapi.responses import FileResponse, RedirectResponse         # noqa: E402
from pydantic import BaseModel                        # noqa: E402
import uvicorn                                        # noqa: E402

import db                                             # noqa: E402
import llm_config as llm_model                        # noqa: E402  全项目模型配置（阿里云百炼）
import auth                                           # noqa: E402  用户账号：注册 / 登录 / 登录态
import support                                        # noqa: E402  在线客服：AI + 人工（会话/消息）
import room                                           # noqa: E402  用户端多人聊天室（公共大群）
import feedback                                       # noqa: E402  用户评价 / 意见反馈
import digital_tts                                    # noqa: E402  数字人小橘的嗓子（edge-tts + 缓存）
import asr                                            # noqa: E402  语音输入的本机识别兜底（faster-whisper）
import pool as doc_pool                               # noqa: E402  模块2：检索池（导入入池）
import store as db_store                              # noqa: E402  模块2 第4步：提取结果仓库
import quality as quality_engine                      # noqa: E402  模块2 第5步：质量验证引擎
import image_ops                                      # noqa: E402  模块1：图片插入（上传/自动定位）
import ocr as ocr_engine                              # noqa: E402  扫描版 PDF / 图片的文字识别
import semantic as sem_engine                         # noqa: E402  字段语义匹配（本地向量）
from chat import Session                              # noqa: E402
from edit import scan_docs, DEFAULT_DOCS_DIRS, IMPORT_DIR     # noqa: E402
from edit import find_doc_file, run_edit                      # noqa: E402  识别分类 API 用

WEB_DIR = Path(__file__).resolve().parent / "web"

# 池多大算"大池"（MB）：超过它就只让一个填表任务独占跑（见 _WeightedGate）。
# 依据：日常 18 份文档的池 0.3MB；压测 16 份大文档的池 194.8MB。
# 30MB 是日常池的 100 倍，中间几乎没有样本 —— 宁可门槛放高，也别误伤正常文档。
BIG_POOL_MB = float(os.environ.get("A23_POOL_BIG_MB") or 30)

app = FastAPI(title="A23 智能文档助手 API", version="1.1",
              description="三模块服务端：自然语言操作文档（编辑 / 排版 / 格式 / 提取）+ "
                          "智能识别分类（画像 / 实体 / 字段 → 自动入库，异步）+ "
                          "表格填写（读懂模板 → 找数 → 填写 → 自检，异步）")


# ---------------------------------------------------------------- 工具

def jsonable(v):
    """把结果里的 Path / 集合等转成可 JSON 化的类型（递归）。"""
    if v is None or isinstance(v, (int, float, bool, str)):
        return v
    if isinstance(v, Path):
        return str(v)
    if isinstance(v, dict):
        return {str(k): jsonable(x) for k, x in v.items()}
    if isinstance(v, (list, tuple, set)):
        return [jsonable(x) for x in v]
    return str(v)


# ---------------------------------------------------------------- 任务 & 会话中枢

class Job:
    """一个任务：指令（command）/ 识别分类（recognize）/ 表格填写（fill）共用。"""

    def __init__(self, jid: str, session_id: str, text: str,
                 kind: str = "command", runner=None, image: str = None):
        self.id = jid
        self.session_id = session_id
        self.text = text
        self.kind = kind              # command / recognize / fill
        self.runner = runner          # 自定义执行体（None = 走会话指令链路）
        self.image = image            # 本轮要插入的图片（指令类任务才用）
        self.status = "queued"        # queued → running → done / failed
        self.created = time.time()
        self.started = None
        self.finished = None
        self.elapsed = None
        self.ok = None
        self.report = None
        self.meta = None
        self.error = None

    def to_dict(self) -> dict:
        d = {
            "job_id": self.id, "session_id": self.session_id, "text": self.text,
            "kind": self.kind,
            "status": self.status, "created": self.created,
            "started": self.started, "finished": self.finished,
        }
        if self.elapsed is not None:
            d["elapsed"] = self.elapsed
        if self.status in ("done", "failed"):
            d["result"] = {
                "ok": bool(self.ok),
                "report": self.report,
                "meta": self.meta,
                "error": self.error,
            }
        return d


class _WeightedGate:
    """按"重量"放行的闸门（2026-11-03）。

    为什么不用普通的 Semaphore：
      小池上一个任务很快（20~40 秒），并发 2~3 个是对的；
      但 50 万块的大池上，一个任务自己就要跑 70~90 秒，同时跑 3 个会把
      8 列的模板5 拖到 126 秒（超红线）。同样是"并发"，两种池要的答案相反。

      所以闸门按"这个任务占几个位"来分配：
        小任务 = 1 位 → 小池可以 2~3 个并行
        大任务 = 全部位 → 大池独占，保证每个任务都在 90 秒内
    """

    def __init__(self, cap: int):
        self.cap = max(1, int(cap))
        self.used = 0
        self.cond = threading.Condition()

    def _w(self, weight):
        return max(1, min(int(weight or 1), self.cap))

    def acquire(self, weight=1):
        w = self._w(weight)
        with self.cond:
            while self.used + w > self.cap:
                self.cond.wait()
            self.used += w
        return w

    def release(self, weight=1):
        w = self._w(weight)
        with self.cond:
            self.used = max(0, self.used - w)
            self.cond.notify_all()

    # 兼容 `with hub.fill_gate:` 的写法（按 1 位算）
    def __enter__(self):
        self.acquire(1)
        return self

    def __exit__(self, *a):
        self.release(1)
        return False


# ---------------------------------------------------------------- 能力问答（秒回，不走模型）
_HELP_TEXT = """我可以帮你做这些事（点左边的大卡片也能进）：

1) 改文档 · 说人话就行
   「把合同里的乙方改成华景公司」「标题居中、加粗」「第 3 段删掉」
   支持 docx / xlsx / md / txt / pdf 和图片；改完另存，原件不动。
   不想打字就点左边的「🤖 智能操作」：加粗 / 居中 / 下划线 这些点一下就行，
   选不选文件都能用（选了就作用在它上面，没选就在话里点名）。

2) 取数据 · 读懂文档再抽
   「提取这份合同的关键信息」「把文档里的城市都列出来」
   自动认文档自己的结构（对象 / 属性 / 关系 / 事件 / 要点），抽完直接入库。

3) 查数据仓库
   按文档分组，点进去能看这份文档抽到了什么；还能看知识图谱、做质量体检。

4) 填表格
   选模板 → 自动找数填进去 → 自检 → 下载（模板不会被改动）。

5) 认字（OCR）
   扫描版 PDF、照片、截图 → 可编辑文字，认完能并入文档库继续用。

6) 语音输入 / 聊天室 / 意见反馈
   输入框左边的 🎤 说话就能输入；有问题也可以找 AI 客服或转人工。

想直接开始：说一句「把 XX 文档的标题改成 XX」，或者从左边「识别入库」挑一份文档试试。
"""

_HELP_PAT = re.compile(
    r"(你有什么(功能|用|本事)|你能(做|干)(什么|啥)|你(会|能)(做)?什么|"
    r"有(哪些|什么)功能|功能(有|是)(哪些|什么)|能帮我(做|干)(什么|啥)|"
    r"怎么(用|使用)|使用(说明|方法|帮助)|help|^\?+$)", re.I)


def help_answer(text: str):
    """判断是不是"问能力/问怎么用"，是就回一份清单（前端直接显示，不折叠）。"""
    t = (text or "").strip()
    if not t or len(t) > 40:
        return None
    if _HELP_PAT.search(t):
        return _HELP_TEXT
    return None


class Hub:
    """全局中枢：会话表 + 任务表 + 线程池。"""

    # 识别线别名（英文键 + 中文说法都能提交）
    TASK_ALIASES = {
        "knowledge": "knowledge", "知识": "knowledge", "自动": "knowledge",
        "自动抽取": "knowledge", "统一抽取": "knowledge",
        "profile": "profile", "画像": "profile", "关键信息": "profile",
        "entities": "entities", "实体": "entities", "实体数据": "entities",
        "fields": "fields", "字段": "fields", "指定字段": "fields",
        "用户指定内容": "fields", "指定内容": "fields",
    }

    def __init__(self, docs_dir):
        # docs_dir 可以是单个目录，也可以是目录列表（模块2 起默认：源文档 + 导入文档）
        self.docs_dir = docs_dir
        self.sessions = {}                 # sid → Session（内存对象，带多轮上下文）
        self.session_locks = {}            # sid → 锁（同一会话串行）
        self.sessions_lock = threading.Lock()

        self.jobs = {}                     # jid → Job
        self.jobs_lock = threading.Lock()
        self._job_seq = itertools.count(1)

        # ---- 并发改造（2026-11-03）----
        # 以前这里是 self.fill_lock，把"读取 → 抽取 → 写回"整条流水线锁成串行，
        # 理由是抽取中间产物共享文件名。问题是：赛题要求 16 份文档阅读一次、
        # 5 张模板各跑一遍，本来可以并行，却被这把锁压成排队。
        # 现在拆成两件事：
        #   ① 读取阶段全局串行（read_lock）—— 它会整体重写 output.json，不能让两个任务同时写；
        #   ② 抽取/写回阶段各自独立目录（jobs/<jid>/），可以真并发，用 fill_gate 限流。
        self.read_lock = threading.Lock()
        self.fill_gate = _WeightedGate(
            max(1, int(os.environ.get("A23_FILL_CONCURRENCY") or 2)))
        self.jobs_root = BASE / "agent" / "jobs"

        # 池里有哪些文档（文件名集合）。按 mtime 缓存 —— 每来一个填表任务
        # 都重读一次 87MB 的池子只为拿文件名，太亏（2026-11-07）。
        self._pool_idx = {}
        self.pool_idx_lock = threading.Lock()

        self.pool = ThreadPoolExecutor(max_workers=4, thread_name_prefix="a23-job")

    # ---- 会话
    def get_session(self, sid: str) -> Session:
        with self.sessions_lock:
            if sid not in self.sessions:
                self.sessions[sid] = Session(self.docs_dir, session_id=sid)
            return self.sessions[sid]

    def lock_for(self, sid: str) -> threading.Lock:
        with self.sessions_lock:
            if sid not in self.session_locks:
                self.session_locks[sid] = threading.Lock()
            return self.session_locks[sid]

    # ---- 任务（提交立即返回；执行在线程池里跑，客户端轮询 /api/job/{id}）
    def _new_job(self, job: Job) -> Job:
        with self.jobs_lock:
            self.jobs[job.id] = job
        self.pool.submit(self._run, job)
        return job

    def submit(self, text: str, sid: str = None, image: str = None) -> Job:
        if not sid:
            sid = "api" + datetime.now().strftime("%m%d-%H%M%S")
        jid = "j%04d" % next(self._job_seq)
        job = Job(jid, sid, text, image=image)
        ans = help_answer(text)
        if ans:
            # 「你有什么功能」这类问题不该走大模型：直接给一份能力清单，
            # 秒回、也不会被解析成"不明确"后把答案折叠起来（2026-11-07）
            job.status = "done"
            job.ok = True
            job.report = ans
            job.meta = {"action": "帮助", "op": "help"}
            job.elapsed = 0.1
            job.finished = time.time()
            with self.jobs_lock:
                self.jobs[job.id] = job
            return job
        return self._new_job(job)

    def submit_recognize(self, docs: list, tasks: list, fields: list,
                         requirement: str = "") -> Job:
        """提交「智能识别分类」任务（异步）：文档 × 提取线 → 自动入库。"""
        jid = "r%04d" % next(self._job_seq)
        label = "识别分类：%s" % "、".join(docs[:6])
        job = Job(jid, "recognize", label, kind="recognize")
        job.runner = lambda j: self._run_recognize(j, docs, tasks, fields, requirement)
        return self._new_job(job)

    def submit_fill(self, template: str, skip_read: bool = False,
                    overwrite: bool = False, require: str = "",
                    docs: list = None) -> Job:
        """提交「表格填写」任务（异步）：跑 pipeline 全流程。

        require（2026-11-03）：用户要求原话。带行条件（如"日期一列从 2020/7/1 到 2020/8/31"）时，
        matcher 会走结构化直通并按条件筛行；不传就等于没有条件，会走原来的大模型路径。

        docs（2026-11-07）：源文档名单。给了就只在选中的这几份里找数（见 _scoped_pool）。
        """
        jid = "f%04d" % next(self._job_seq)
        job = Job(jid, "fill", "表格填写：%s" % template, kind="fill")
        job.runner = lambda j: self._run_fill(j, template, skip_read, overwrite,
                                              require, docs or [])
        return self._new_job(job)

    # ---- 智能操作（模块①：编辑 / 排版 / 格式，2026-11-07 独立入口）----
    # 「智能操作」弹窗和首页对话框走的是同一条执行链（edit.run_edit），
    # 区别只有一个：弹窗里点过文件，就把那份文件当好 doc_override 传进去；
    # 点的是快捷按钮，就带一份 preset 意图（跳过模型，0.3 秒出结果）。
    def submit_op(self, text: str = "", doc: str = None, preset: dict = None,
                  label: str = "") -> Job:
        jid = "o%04d" % next(self._job_seq)
        job = Job(jid, "op", label or text or "智能操作", kind="op")
        job.runner = lambda j: self._run_op(j, text, doc, preset)
        return self._new_job(job)

    def _run_op(self, job: Job, text: str, doc: str, preset: dict):
        with self.lock_for("op"):        # 同一份"智能操作"会话串行（工作副本要连着走）
            sess = self.get_session("op")
            ok, report, meta, elapsed = sess.run_turn(
                text or "", preset=preset, doc_override=doc)
            job.ok = bool(ok)
            job.report = report
            job.meta = jsonable(meta)
            job.elapsed = round(elapsed, 2)
            job.status = "done"

    def _run(self, job: Job):
        job.status = "running"
        job.started = time.time()
        if job.runner is not None:                   # 识别 / 填写：自定义执行体
            try:
                job.runner(job)
            except Exception as e:                   # 任何异常不让服务倒
                job.ok = False
                job.error = "%s: %s" % (type(e).__name__, e)
                job.status = "failed"
            job.finished = time.time()
            return
        lock = self.lock_for(job.session_id)
        with lock:                                   # 同一会话串行执行
            try:
                sess = self.get_session(job.session_id)
                ok, report, meta, elapsed = sess.run_turn(job.text, image=job.image)
                job.ok = bool(ok)
                job.report = report
                job.meta = jsonable(meta)
                job.elapsed = round(elapsed, 2)
                job.status = "done"
            except Exception as e:                   # 任何异常不让服务倒
                job.ok = False
                job.error = "%s: %s" % (type(e).__name__, e)
                job.status = "failed"
        job.finished = time.time()

    # ---- 识别分类执行体（模块2 能力 · 专属异步 API 的核心） ----
    def _recognize_cmd(self, line: str, doc: str, fields: list):
        """把（识别线, 文档, 字段）拼成一条走既定链路的指令。"""
        if line == "profile":
            return "提取%s的关键信息" % doc
        if line == "entities":
            return "识别%s里的实体数据" % doc
        if line == "fields":
            return "从%s里提取%s" % (doc, "和".join(fields))
        return None

    def _run_recognize(self, job, docs, tasks, fields, requirement=""):
        LINE_CN = {"knowledge": "知识", "profile": "画像", "entities": "实体", "fields": "字段"}

        def fail_line(rep):
            for l in rep.splitlines():
                s = l.strip()
                if s.startswith(("【失败】", "【提示】", "【反问】")):
                    return s
            return "（详情见任务报告）"

        t0 = time.time()
        results = []
        n_fail = 0
        for doc in docs:
            entry = {"doc": doc, "lines": {}}
            for line in tasks:
                if line == "knowledge":
                    # 统一知识抽取（2026-11-05）：不进 run_edit 的意图解析，
                    # 直接走引擎 —— 这里已经是"确定要抽"，再解析一遍纯属浪费一次模型调用。
                    from knowledge import run_knowledge
                    p = find_doc_file(doc, self.docs_dir)
                    t1 = time.time()
                    ok, report, _files = run_knowledge(
                        "自动读懂%s并抽取关键信息入库" % doc,
                        {"params": {"requirement": requirement}},
                        Path(p) if p else Path(doc),
                        requirement=requirement)
                    entry["lines"]["knowledge"] = {
                        "ok": bool(ok), "cmd": "knowledge",
                        "elapsed": round(time.time() - t1, 2),
                        "error": None if ok else fail_line(report),
                        "report": report,
                    }
                    if not ok:
                        n_fail += 1
                    continue
                cmd = self._recognize_cmd(line, doc, fields)
                if cmd is None:
                    continue
                t1 = time.time()
                meta = {}
                ok, report = run_edit(cmd, self.docs_dir, meta=meta)
                entry["lines"][line] = {
                    "ok": bool(ok),
                    "cmd": cmd,
                    "elapsed": round(time.time() - t1, 2),
                    "error": None if ok else fail_line(report),
                }
                if not ok:
                    n_fail += 1
            results.append(entry)

        # 提取即入库：跑完即可查——附库内分布快照
        snap = []
        try:
            for d in db_store.list_docs():
                snap.append({"doc_name": d["doc_name"], "n_ent": d["n_ent"],
                             "n_fld": d["n_fld"], "n_prof": d["n_prof"]})
        except Exception:
            pass
        k_snap = {}
        try:
            k_snap = db_store.knowledge_stats()
        except Exception:
            pass

        lines_out = ["智能识别分类 —— %d 份文档 × 识别线：%s"
                     % (len(docs), " / ".join(LINE_CN.get(k, k) for k in tasks))]
        for e in results:
            bits = []
            for k in tasks:
                r = e["lines"].get(k)
                if not r:
                    continue
                bits.append("%s %s %.1fs" % (LINE_CN.get(k, k),
                                             "✓" if r["ok"] else "✗", r["elapsed"]))
            lines_out.append("◆《%s》%s" % (e["doc"], " ｜ ".join(bits)))
            for k in tasks:
                r = e["lines"].get(k)
                if r and not r["ok"]:
                    lines_out.append("    ✗ %s：%s" % (LINE_CN.get(k, k), r["error"]))
                elif r and k == "knowledge" and r.get("report"):
                    # 知识线给一句话摘要（领域/对象/条目），细节在数据仓库
                    for l in (r["report"] or "").splitlines():
                        s = l.strip()
                        if s.startswith(("◆ 领域", "◆ 对象", "◆ 入库条目")):
                            lines_out.append("    " + s)
        lines_out.append("")
        if k_snap:
            lines_out.append("知识库：%d 条（对象 %d / 属性 %d / 关系 %d / 事件 %d / 要点 %d）"
                             % (k_snap.get("items", 0),
                                k_snap.get("by_kind", {}).get("object", 0),
                                k_snap.get("by_kind", {}).get("attribute", 0),
                                k_snap.get("by_kind", {}).get("relation", 0),
                                k_snap.get("by_kind", {}).get("event", 0),
                                k_snap.get("by_kind", {}).get("summary", 0)))
        lines_out.append("提取结果已自动入库（a23_agent）——可在「数据仓库」查询")

        job.ok = (n_fail == 0)
        job.report = "\n".join(lines_out)
        job.meta = jsonable({"docs": results, "store": snap, "knowledge": k_snap,
                             "elapsed": round(time.time() - t0, 2)})
        job.elapsed = round(time.time() - t0, 2)
        job.status = "done"

    # ---- 共享池：多任务共用一份，只有"生成"这一步串行 ----
    def _pool_file(self):
        return BASE / "doc_reader" / "output.json"

    def _reader_stamp(self):
        """读取器代码的最后修改时间。

        为什么光比源文档不够：2026-11-03 给 xlsx_reader 补了 SUM 求值，
        源文档一个字节没动，旧池看起来"还很新"，于是被继续复用 ——
        池里存的还是 "=SUM(F3:F12)" 这串字符，补丁等于没生效。
        读取逻辑一变，池就必须重建，所以把代码时间也算进"新鲜度"。
        """
        newest = 0.0
        for p in (BASE / "doc_reader").rglob("*.py"):
            try:
                newest = max(newest, p.stat().st_mtime)
            except OSError:
                pass
        return newest

    def _ensure_pool(self):
        """确保共享池存在、且不比源文档（和读取器代码）旧；返回池路径。

        为什么池共享：赛题场景是"16 份文档上传一次 → 5 张模板分别填"，
        池的内容跟模板无关，没必要每个任务各读一遍（大文档要 20 秒）。
        为什么读取必须串行：读取会整体重写 output.json，
        两个任务同时写就是互相覆盖 —— 这是原来那把全局锁唯一真正的理由。
        """
        pool = self._pool_file()
        dirs = self.docs_dir if isinstance(self.docs_dir, (list, tuple)) else [self.docs_dir]
        newest = self._reader_stamp()
        for d in dirs:
            try:
                for f in Path(d).rglob("*"):
                    if f.is_file():
                        newest = max(newest, f.stat().st_mtime)
            except OSError:
                pass

        def _fresh():
            return pool.exists() and pool.stat().st_mtime >= newest

        if _fresh():
            return pool
        with self.read_lock:
            # 双检：等锁这一会儿，别的任务可能已经读完了
            if _fresh():
                return pool
            proc = subprocess.run(
                [sys.executable, "run_all.py"],
                cwd=str(BASE / "doc_reader"),
                capture_output=True, timeout=900,
                env={**os.environ, "PYTHONIOENCODING": "utf-8"})
            if not _fresh():
                msg = (proc.stderr or proc.stdout or b"").decode("utf-8", "replace")
                raise RuntimeError("读源文档失败：" + (msg.strip()[-300:] or "池文件没生成"))
        return pool

    def _pool_mb(self):
        """池有多大（MB）。拿不到就返回 0 —— 宁可当小池，也不要因为文件没读成拒服务。"""
        try:
            return os.path.getsize(str(self._pool_file())) / 1024.0 / 1024.0
        except OSError:
            return 0.0

    def _fill_weight(self):
        """这个填表任务占几个闸门位：大池占用全部（独占），小池占 1 位。"""
        if self._pool_mb() >= BIG_POOL_MB:
            return self.fill_gate.cap
        return 1

    # ---- 表格填写执行体（模块3 · 跑 pipeline 全流程） ----
    def _pool_files(self):
        """池子里已经有哪些文档（文件名集合）。按 mtime 缓存，不反复读大盘。"""
        pool = self._pool_file()
        try:
            st = pool.stat()
        except OSError:
            return set()
        key = (st.st_mtime_ns, st.st_size)
        with self.pool_idx_lock:
            if self._pool_idx.get("key") == key:
                return self._pool_idx["files"]
        try:
            data = json.loads(pool.read_text(encoding="utf-8"))
            files = {d.get("file") for d in (data.get("documents") or [])}
        except Exception:
            files = set()
        with self.pool_idx_lock:
            self._pool_idx = {"key": key, "files": files}
        return files

    def _read_doc_cached(self, p: Path):
        """读一份文档，按「文件名 + 文件 mtime + 读取器代码版本」缓存。

        为什么要有这层：共享池是 `doc_reader/main.py` 按「测试数据/源文档」建的，
        `导入文档/` 里的文件不在池里（池重建也会把它们冲掉）。指定源文档时
        `_scoped_pool` 得现场把这些文件读一遍 —— 2 万行的 xlsx 要 6~7 秒，
        而同一份源文档往往要喂给好几张模板，每回都重读等于白等。

        缓存命中就只花一次 JSON 反序列化；文件改过或读取器改过，key 就变了，
        自然失效（不做"猜文件没变"这种判断）。
        """
        try:
            f = Path(p)
            stamp = int(f.stat().st_mtime)
        except OSError:
            return None
        cache_dir = BASE / "agent" / ".cache" / "docs"
        cache_dir.mkdir(parents=True, exist_ok=True)
        key = "%s@%d@%d" % (f.stem, stamp, int(self._reader_stamp()))
        safe = re.sub(r'[\\/:*?"<>|]', "_", key)[:150]
        cache = cache_dir / (safe + ".json")
        if cache.exists():
            try:
                return json.loads(cache.read_text(encoding="utf-8"))
            except Exception:
                pass                                  # 缓存坏了就当没命中，别让它挡住任务
        from readers import read_one
        d = read_one(f)
        if d:
            try:
                cache.write_text(json.dumps(d, ensure_ascii=False), encoding="utf-8")
                # 同一份文件的旧版本缓存留着没用，顺手清掉（只清同名的、别的 mtime 的）
                for old in cache_dir.glob("%s@*.json" % re.sub(r'[\\/:*?"<>|]', "_", f.stem)[:120]):
                    if old.name != cache.name:
                        try:
                            old.unlink()
                        except OSError:
                            pass
            except Exception:
                pass
        return d

    def _scoped_pool(self, work: Path, names: list):
        """把共享池裁成「这一单只看这几份文档」的小池（2026-11-07）。

        为什么必须裁：模板的列名（金额 / 日期 / 负责人）在文档库里常常好几份都有。
        全库一起检索时，抽到的值可能来自另一份合同 —— 值本身没错，但不是这张表要的。
        用户勾了源文档就等于给了范围，这里只在这个范围里找数。

        池里没有的（刚上传还没来得及入池 / 池重建漏了）现场读一遍 ——
        不能让用户"明明选了这份文档，却因为它不在池里被静默跳过"。
        小池写在任务目录里，绝不改共享池（别的任务还要用全库）。

        返回 (小池路径, 实际纳入的文件名列表)。
        """
        want = [str(n).strip() for n in (names or []) if str(n).strip()]
        pool = self._pool_file()
        try:
            data = json.loads(pool.read_text(encoding="utf-8"))
        except Exception:
            data = {"documents": []}
        docs = data.get("documents") or []
        wset = set(want)
        picked = [d for d in docs if d.get("file") in wset]
        have = {d.get("file") for d in picked}
        for n in want:
            if n in have:
                continue
            p = find_doc_file(n, self.docs_dir)
            if p is None:
                continue
            try:
                d = self._read_doc_cached(p)
            except Exception:
                d = None
            if d:
                picked.append(d)

        out = dict(data)
        out["documents"] = picked
        out["file_count"] = len(picked)
        out["ok_count"] = len(picked)
        out["failed"] = []
        out["source_dir"] = "、".join(want)
        dst = work / "pool.json"
        dst.write_text(json.dumps(out, ensure_ascii=False), encoding="utf-8")
        return dst, [d.get("file") for d in picked]

    def _run_fill(self, job, template: str, skip_read: bool, overwrite: bool,
                  require: str = "", docs: list = None):
        docs = [str(d).strip() for d in (docs or []) if str(d).strip()]
        t0 = time.time()
        # 按文件名定位实际所在目录（系统自带两批 + 用户上传那批都认），
        # 该目录再通过 A23_TEMPLATE_DIR 传给 pipeline 当「模板目录」。
        tpl_path = _find_template(template)
        if tpl_path is None:
            raise ValueError("模板目录里找不到：%s" % template)
        tpl_dir = tpl_path.parent
        tpl_dir_name = tpl_dir.name

        # 池先备好（读取阶段全局串行），之后各任务只做抽取+写回。
        # 必须在进闸门**之前**做完：读取本身是串行的，占着闸门等读取纯属浪费名额。
        if docs:
            # 指定了源文档：小池由 _scoped_pool 现裁，池里没有的那几份它自己会读。
            # 这里绝不重建全库池 —— 重建会把导入文档冲掉，反而帮倒忙。
            if not skip_read and not self._pool_file().exists():
                self._ensure_pool()
        elif not skip_read:
            self._ensure_pool()

        # 并发闸门：小池按 1 位并行，大池占满全部位（=独占）。
        # 以前这里是 fill_lock（整条流水线串行），并发能力等于 0；
        # 后来改成固定 2 路，大池上又会互相拖到超时 —— 见 _WeightedGate 的说明。
        weight = self._fill_weight()
        self.fill_gate.acquire(weight)
        try:
            # 每个任务一套独立产物目录：抽取结果、成品、模板结构都不共享，
            # 这样两个任务并行时不会互相覆盖同名文件（以前正是这个问题逼出了全局锁）。
            work = self.jobs_root / job.id
            res_dir = work / "matcher"
            out_dir = work / "out"
            res_dir.mkdir(parents=True, exist_ok=True)
            out_dir.mkdir(parents=True, exist_ok=True)

            cmd = [sys.executable, str(BASE / "pipeline.py"), "--template", template]
            cmd.append("--skip-read")          # 池已经备好了
            if overwrite:
                cmd.append("--overwrite")
            env = os.environ.copy()
            env["PYTHONIOENCODING"] = "utf-8"
            env["A23_TEMPLATE_DIR"] = str(tpl_dir)
            env["A23_MATCHER_DIR"] = str(res_dir)
            env["A23_OUT_DIR"] = str(out_dir)
            env["A23_STRUCTURE_JSON"] = str(work / "structure.json")
            # 只看用户勾的那几份文档（2026-11-07）：共享池裁成这一单专用的小池，
            # 用 A23_OUTPUT_JSON 指给 matcher（extract / retrieve 都认这个变量）。
            scope = []
            if docs:
                have = self._pool_files()
                if all(d in have for d in docs):
                    # 快路：池里都有 → 只读时过滤，不拷贝（87MB 池实测省 9 秒）
                    env["A23_DOC_SCOPE"] = "\n".join(docs)
                    scope = docs
                else:
                    # 有刚上传还没进池的：现裁一份小池，缺的那几份顺手读进来
                    scoped, scope = self._scoped_pool(work, docs)
                    env["A23_OUTPUT_JSON"] = str(scoped)
            # 用户要求原话透传：matcher 从里面解析行条件（日期区间等）。
            # 不传 = 没有条件 = 整列全取（老行为）。
            if require.strip():
                env["A23_USER_REQ"] = require.strip()

            proc = subprocess.run(cmd, cwd=str(BASE), capture_output=True,
                                  env=env, timeout=420)
            out = proc.stdout.decode("utf-8", errors="replace")
            err = proc.stderr.decode("utf-8", errors="replace")
            elapsed = round(time.time() - t0, 2)

            stopped = re.search(r"流程在中途停下：(.+)", out)

            # 关键证据行（从流水线日志里抓）
            stages = {}
            for m in re.finditer(r"【(.+?)】完成 ｜ ([\d.]+) 秒", out):
                stages[m.group(1)] = float(m.group(2))
            bits = []
            m = re.search(r"总耗时：([\d.]+) 秒", out)
            if m:
                bits.append("总耗时 %s 秒" % m.group(1))
            m = re.search(r"写入 (\d+) 格", out)
            if m:
                bits.append("写入 %s 格" % m.group(1))
            m = re.search(r"值一致 ([\d/]+) 格", out)
            if m:
                bits.append("值一致 %s 格" % m.group(1))
            m = re.search(r"格式差异 (\d+) 处", out)
            if m:
                bits.append("格式差异 %s 处" % m.group(1))
            # 行条件与走法（2026-11-03）：填得少不代表填对了 —— 界面必须显示"按什么条件筛的行"
            m = re.search(r"【③ 行条件】(.+)", out)
            if m:
                bits.append("行条件 %s" % m.group(1).strip())
            if "结构化直通" in out:
                bits.append("走法 结构化直通（0 次模型调用）")
            if "结论：可以交卷" in out:
                bits.append("结论：可以交卷")

            m = re.search(r"交卷文件 → (.+)", out)
            out_file = m.group(1).strip() if m else None

            ok = (proc.returncode == 0) and (stopped is None) and (out_file is not None)
            rep = ["【表格填写】模板：%s（%s）" % (template, tpl_dir_name)]
            if docs:
                shown = "、".join(scope[:6]) + ("…" if len(scope) > 6 else "")
                rep.append("源文档：%s（只在这 %d 份里找数）" % (shown or "（没找到匹配的文件）",
                                                         len(scope)))
            else:
                rep.append("源文档：未指定 —— 整个文档库都参与检索")
            if stopped:
                rep.append("流程在中途停下：%s" % stopped.group(1).strip())
                rep.append("（提示：待填区已有内容时默认保护性不写；"
                           "确认要覆盖时，重提任务并勾选「允许覆盖旧值」）")
            if bits:
                rep.append("  ｜ ".join(bits))
            # 直通过程中的告警（2026-11-07）：最典型的是「这块没有任何筛选条件，
            # 只取了源表前 N 行」—— 这种情况值看起来填满了、其实未必是用户要的那段，
            # 必须让用户在结果卡片上一眼看到，而不是埋在日志里。
            _warns = re.findall(r"⚠ ([^\n]+)", out)
            for _w in _warns[:3]:
                rep.append("⚠ " + _w.strip())
            if out_file:
                rep.append("交卷文件 → %s" % out_file)
            if not out_file and err.strip():
                rep.append("stderr：%s" % err.strip()[-400:])

            # 字段语义匹配的对照（2026-10-26）：模板列名不是源文档里的写法时，
            # 把"它被认成了什么"贴到报告里 —— 用户能直接从界面看出映射关系。
            _ol = out.splitlines()
            for _i, _ln in enumerate(_ol):
                if "【③ 字段语义匹配】" in _ln:
                    for _nxt in _ol[_i + 1:_i + 9]:
                        if _nxt.startswith("  ") and "→" in _nxt:
                            rep.append("  " + _nxt.strip())
                        else:
                            break
                    break

            job.ok = bool(ok)
            job.report = "\n".join(rep)
            job.meta = jsonable({"template": template, "dir": tpl_dir_name,
                                 "elapsed": elapsed, "stages": stages,
                                 "output_file": out_file,
                                 "workdir": str(work),
                                 "pool_mb": round(self._pool_mb(), 1),
                                 "gate_weight": weight,
                                 "log_tail": "\n".join(out.splitlines()[-24:])})
            job.elapsed = elapsed
            job.status = "done"
        finally:
            # 无论成功失败都要还名额，否则一次异常就把闸门焊死，后面所有任务永久排队
            self.fill_gate.release(weight)

    def get_job(self, jid: str) -> Job:
        with self.jobs_lock:
            return self.jobs.get(jid)


hub: "Hub | None" = None


# ---------------------------------------------------------------- 路由

@app.get("/login")
def login_page(request: Request, portal: str = ""):
    """登录 / 注册页；已经登录的人直接回对应端口。"""
    try:
        user = auth.user_by_token(request.cookies.get(auth.TOKEN_COOKIE))
    except RuntimeError:
        user = None
    if user:
        homes = {"user": "/", "staff": "/staff", "admin": "/admin"}
        if portal in homes and portal in user.get("roles", []):
            return RedirectResponse(homes[portal], status_code=302)
        for p in ("admin", "staff", "user"):
            if p in user.get("roles", []):
                return RedirectResponse(homes[p], status_code=302)
        return RedirectResponse("/", status_code=302)
    page = WEB_DIR / "login.html"
    if not page.exists():
        raise HTTPException(status_code=500, detail="网页文件缺失：agent/web/login.html")
    return FileResponse(page, media_type="text/html; charset=utf-8")


@app.get("/admin")
def admin_page(request: Request):
    """管理员端（需 admin 角色）。"""
    user = _portal_user(request, "admin")
    if not user:
        return RedirectResponse("/login?portal=admin", status_code=302)
    page = WEB_DIR / "admin.html"
    if not page.exists():
        raise HTTPException(status_code=500, detail="网页文件缺失：agent/web/admin.html")
    return FileResponse(page, media_type="text/html; charset=utf-8")


@app.get("/staff")
def staff_page(request: Request):
    """客服端工作台（需 staff 角色）。"""
    user = _portal_user(request, "staff")
    if not user:
        return RedirectResponse("/login?portal=staff", status_code=302)
    page = WEB_DIR / "staff.html"
    if not page.exists():
        raise HTTPException(status_code=500, detail="网页文件缺失：agent/web/staff.html")
    return FileResponse(page, media_type="text/html; charset=utf-8")


def _portal_user(request: Request, portal: str) -> dict | None:
    """页面级校验：登录了 + 有该端权限才放行；数据库连不上时放行（与后端降级一致）。"""
    try:
        user = auth.user_by_token(request.cookies.get(auth.TOKEN_COOKIE))
    except RuntimeError:
        return {"id": 0, "username": "offline", "display_name": "离线模式",
                "roles": ["user", "staff", "admin"], "status": "active"}
    if not user:
        return None
    if portal not in user.get("roles", []):
        return None
    return user


@app.get("/")
def index(request: Request):
    """网页操作台（需登录；没登录先去登录页）。"""
    if not auth.is_logged_in(request.cookies.get(auth.TOKEN_COOKIE)):
        return RedirectResponse("/login", status_code=302)
    page = WEB_DIR / "index.html"
    if not page.exists():
        raise HTTPException(status_code=500, detail="网页文件缺失：agent/web/index.html")
    return FileResponse(page, media_type="text/html; charset=utf-8")


# ---------------------------------------------------------------- 账号（注册 / 登录 / 退出）

class RegisterIn(BaseModel):
    username: str
    password: str
    display_name: str | None = None


class LoginIn(BaseModel):
    username: str
    password: str
    portal: str = "user"          # user 用户端 / staff 客服端 / admin 管理员端


class ProfileIn(BaseModel):
    display_name: str | None = None


MAX_AVATAR_BYTES = 5 * 1024 * 1024          # 原图上限 5MB
AVATAR_SIZE = 256                           # 统一裁成 256×256 方图
AVATAR_EXTS = {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".webp"}


def _set_token_cookie(resp: Response, token: str):
    resp.set_cookie(auth.TOKEN_COOKIE, token, max_age=auth.TOKEN_DAYS * 86400,
                    httponly=True, samesite="lax", path="/")


@app.post("/api/auth/register")
def api_register(body: RegisterIn, response: Response):
    """注册 → 直接自动登录（省得再输一遍）。"""
    try:
        user = auth.register(body.username, body.password, body.display_name or "")
        token, _ = auth.login(body.username, body.password)
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))
    except RuntimeError as e:
        raise HTTPException(status_code=503, detail=str(e))
    _set_token_cookie(response, token)
    return {"ok": True, "user": user}


@app.post("/api/auth/login")
def api_login(body: LoginIn, response: Response):
    try:
        token, user = auth.login(body.username, body.password, body.portal or "user")
    except ValueError as e:
        raise HTTPException(status_code=401, detail=str(e))
    except RuntimeError as e:
        raise HTTPException(status_code=503, detail=str(e))
    _set_token_cookie(response, token)
    return {"ok": True, "user": user, "portal": body.portal or "user"}


@app.post("/api/auth/logout")
def api_logout(request: Request, response: Response):
    try:
        auth.logout(request.cookies.get(auth.TOKEN_COOKIE))
    except Exception:
        pass
    response.delete_cookie(auth.TOKEN_COOKIE, path="/")
    return {"ok": True}


@app.get("/api/auth/me")
def api_me(request: Request):
    """当前登录用户（顶栏显示用）；没登录返回 user=None。"""
    try:
        user = auth.user_by_token(request.cookies.get(auth.TOKEN_COOKIE))
    except RuntimeError:
        return {"ok": False, "user": None, "warning": "数据库未连接，登录态无法校验"}
    return {"ok": bool(user), "user": user}


# ---------------------------------------------------------------- 三个端口：权限守卫

def _require_role(request: Request, portal: str) -> dict:
    """要登录 + 要有该端权限，否则 401 / 403。"""
    user = _require_user(request)
    if portal not in user.get("roles", []):
        raise HTTPException(status_code=403,
                            detail="这个账号没有%s权限" % auth.ROLES.get(portal, portal))
    return user


# ---------------------------------------------------------------- 用户端：AI / 人工客服

class SupportStartIn(BaseModel):
    mode: str = "ai"              # ai 智能客服 / human 转人工


class SupportSendIn(BaseModel):
    session_id: int
    content: str


class SupportIdIn(BaseModel):
    """只需要会话 id 的操作（结束会话）。"""
    session_id: int


@app.get("/api/support/faq")
def api_support_faq(request: Request):
    """AI 客服的常见问题快捷指令（编号 + 短标题）。用户回数字就直接命中，不调模型。"""
    _require_role(request, "user")
    return {"ok": True, "items": support.faq_list()}


# ---------------------------------------------------------------- 语音输入：本机离线识别
#
# 背景：浏览器自带的 SpeechRecognition 会把麦克风音频送到谷歌 / 微软的服务器，
# 国内网络经常连不上（前端只拿到 network 错误）。这两条接口让录音留在本机识别，
# 是那条通道挂掉之后的兜底——不联网也能把话变成字。

MAX_ASR_BYTES = 20 * 1024 * 1024          # 20MB，约三分钟 opus，一句话绰绰有余


@app.get("/api/asr/status")
def api_asr_status(request: Request):
    """离线语音自检：模型在不在、加载没加载（前端提示 / 验收脚本用）。"""
    _require_user(request)
    return {"ok": True, **asr.info()}


@app.post("/api/asr")
def api_asr(request: Request, file: UploadFile = File(...)):
    """录音 → 文字。全部在本机完成，不出网。

    用普通 def（不是 async def）：识别是 CPU 密集的同步活儿，交给线程池跑，
    否则会卡住整个事件循环，别的请求都得排队等它。
    """
    _require_user(request)
    raw = file.file.read()
    if not raw:
        raise HTTPException(status_code=400, detail="录音是空的，再说一次")
    if len(raw) > MAX_ASR_BYTES:
        raise HTTPException(status_code=413, detail="这段录音太长了，一句一句说")

    suffix = {"audio/webm": ".webm", "audio/ogg": ".ogg",
              "audio/mp4": ".m4a", "audio/wav": ".wav",
              "audio/x-wav": ".wav", "audio/mpeg": ".mp3"}.get(
                  (file.content_type or "").split(";")[0].strip(), ".webm")
    tmp = None
    try:
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
        tmp.write(raw)
        tmp.close()
        result = asr.transcribe(tmp.name)
    except RuntimeError as e:                 # 模型没下载 / 引擎起不来
        raise HTTPException(status_code=503, detail=str(e))
    except Exception as e:
        raise HTTPException(status_code=500, detail="识别失败：%s" % e)
    finally:
        if tmp is not None:
            try:
                os.unlink(tmp.name)
            except OSError:
                pass

    if not result.get("text"):
        raise HTTPException(status_code=422, detail="没听清，再说一次")
    return {"ok": True, **result}


@app.get("/api/tts")
def api_tts(request: Request, text: str = ""):
    """数字人「小橘」的嗓子：文字 → mp3（本地缓存，同句第二次不再联网）。"""
    _require_user(request)
    p = digital_tts.speak(text)
    if not p:
        raise HTTPException(status_code=503, detail="语音暂时不可用")
    return FileResponse(str(p), media_type="audio/mpeg",
                        headers={"Cache-Control": "public, max-age=604800"})


@app.get("/api/digital/status")
def api_digital_status(request: Request):
    """数字人自检：声音是否就绪、缓存了多少句（验收脚本用）。"""
    _require_user(request)
    return {"ok": True, "name": "小橘", **digital_tts.stats()}


@app.post("/api/support/start")
def api_support_start(body: SupportStartIn, request: Request):
    """开一次咨询（AI 或人工）。已有进行中的会话就复用。"""
    user = _require_role(request, "user")
    try:
        sess = support.start(user, body.mode)
        return {"ok": True, "session": sess, "messages": support.history(sess["id"], 0, 100),
                "queue_ahead": support.staff_ahead(sess["id"]) if sess["status"] == "waiting" else 0}
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))
    except RuntimeError as e:
        raise HTTPException(status_code=503, detail=str(e))


@app.get("/api/support/mine")
def api_support_mine(request: Request):
    """当前进行中的咨询（刷新页面后恢复现场）。"""
    user = _require_role(request, "user")
    try:
        sess = support.my_session(user["id"])
    except RuntimeError as e:
        raise HTTPException(status_code=503, detail=str(e))
    if not sess:
        return {"ok": True, "session": None, "messages": [], "queue_ahead": 0}
    return {"ok": True, "session": sess, "messages": support.history(sess["id"], 0, 100),
            "queue_ahead": support.staff_ahead(sess["id"]) if sess["status"] == "waiting" else 0}


@app.post("/api/support/send")
def api_support_send(body: SupportSendIn, request: Request):
    """用户发消息；AI 会话当场带一条模型回复回来。"""
    user = _require_role(request, "user")
    try:
        return support.user_send(user, body.session_id, body.content)
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))
    except RuntimeError as e:
        raise HTTPException(status_code=503, detail=str(e))


@app.get("/api/support/poll")
def api_support_poll(session_id: int, after: int = 0, request: Request = None):
    """轮询新消息（用户侧）。"""
    user = _require_role(request, "user")
    try:
        sess = support.poll(session_id, after)
        if sess["session"]["user_id"] != int(user["id"]):
            raise HTTPException(status_code=403, detail="这不是你的会话")
        return sess
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))


@app.post("/api/support/close")
def api_support_close(body: SupportIdIn, request: Request):
    user = _require_role(request, "user")
    try:
        return {"ok": True, "session": support.close_by_user(user["id"], body.session_id)}
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))


# ---------------------------------------------------------------- 客服端：工作台

class StaffSessionIn(BaseModel):
    session_id: int


class StaffMsgIn(BaseModel):
    session_id: int
    content: str


@app.post("/api/staff/heartbeat")
def api_staff_heartbeat(request: Request):
    """客服在线心跳（前端 15 秒一次）。"""
    user = _require_role(request, "staff")
    try:
        support.touch(user["id"])
    except Exception:
        pass
    return {"ok": True}


@app.get("/api/staff/sessions")
def api_staff_sessions(request: Request):
    """待接入队列 + 我正在接待的会话。"""
    user = _require_role(request, "staff")
    support.touch(user["id"])
    try:
        return support.staff_sessions(user["id"])
    except RuntimeError as e:
        raise HTTPException(status_code=503, detail=str(e))


@app.post("/api/staff/accept")
def api_staff_accept(body: StaffSessionIn, request: Request):
    user = _require_role(request, "staff")
    try:
        return {"ok": True, "session": support.staff_accept(user, body.session_id)}
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))


@app.post("/api/staff/send")
def api_staff_send(body: StaffMsgIn, request: Request):
    user = _require_role(request, "staff")
    try:
        return support.staff_send(user, body.session_id, body.content)
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))


@app.post("/api/staff/close")
def api_staff_close(body: StaffSessionIn, request: Request):
    user = _require_role(request, "staff")
    try:
        return {"ok": True, "session": support.staff_close(user, body.session_id)}
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))


@app.get("/api/staff/poll")
def api_staff_poll(request: Request, session_id: int, after: int = 0):
    """客服轮询某个会话（只有自己接待的才能看）。"""
    user = _require_role(request, "staff")
    try:
        sess = support.poll(session_id, after)
        s = sess["session"]
        if s["staff_id"] != int(user["id"]) and s["status"] != "waiting":
            raise HTTPException(status_code=403, detail="这不是你接待的会话")
        return sess
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))


# ---------------------------------------------------------------- 用户端：多人聊天室

class RoomSendIn(BaseModel):
    content: str


class RoomRecallIn(BaseModel):
    message_id: int


@app.get("/api/room/poll")
def api_room_poll(request: Request, after: int = 0, alive: int = 1):
    """群里拉消息：after 之后的新消息 + 在线名单。

    after=0 表示刚进群（返回最近 60 条）；alive=0 表示面板没开，
    只想知道有没有新消息，不把自己标成在线。
    """
    user = _require_role(request, "user")
    try:
        return room.poll(user, after, bool(alive))
    except RuntimeError as e:
        raise HTTPException(status_code=503, detail=str(e))


@app.post("/api/room/send")
def api_room_send(body: RoomSendIn, request: Request):
    user = _require_role(request, "user")
    try:
        return {"ok": True, "message": room.send(user, body.content)}
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))
    except RuntimeError as e:
        raise HTTPException(status_code=503, detail=str(e))


@app.post("/api/room/recall")
def api_room_recall(body: RoomRecallIn, request: Request):
    """撤回自己 5 分钟内发的消息。"""
    user = _require_role(request, "user")
    try:
        return room.recall(user, body.message_id)
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))
    except RuntimeError as e:
        raise HTTPException(status_code=503, detail=str(e))


@app.post("/api/room/leave")
def api_room_leave(request: Request):
    """收起聊天室 = 离开在线名单。"""
    user = _require_role(request, "user")
    try:
        room.leave(user["id"])
        return {"ok": True}
    except RuntimeError as e:
        raise HTTPException(status_code=503, detail=str(e))


# ---------------------------------------------------------------- 用户评价 / 意见反馈

class FeedbackIn(BaseModel):
    feature: str = "other"        # doc_edit / extract / fill / voice / chat / other
    rating: int = 0               # 1–5 星；0 = 只写意见
    content: str = ""
    tags: str = ""
    session_id: int | None = None


class FeedbackStatusIn(BaseModel):
    status: str = "read"          # new / read / done


@app.post("/api/feedback")
def api_feedback_add(body: FeedbackIn, request: Request):
    """用户提交评价（登录即可，不要求 user 角色，客服自己也能提）。"""
    user = _require_user(request)
    try:
        return {"ok": True, "item": feedback.add(user, body.feature, body.rating,
                                                 body.content, body.session_id, body.tags)}
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))
    except RuntimeError as e:
        raise HTTPException(status_code=503, detail=str(e))


@app.get("/api/feedback/mine")
def api_feedback_mine(request: Request, limit: int = 20):
    user = _require_user(request)
    try:
        return {"ok": True, "items": feedback.mine(user["id"], limit)}
    except RuntimeError as e:
        raise HTTPException(status_code=503, detail=str(e))


@app.get("/api/staff/feedback")
def api_staff_feedback(request: Request, limit: int = 40):
    """客服端：我接待会话收到的评价 + 全站客服评价。"""
    user = _require_role(request, "staff")
    try:
        return feedback.staff_view(user["id"], limit)
    except RuntimeError as e:
        raise HTTPException(status_code=503, detail=str(e))


@app.get("/api/admin/feedback")
def api_admin_feedback(request: Request, feature: str = "", status: str = "", limit: int = 100):
    """管理员端：全量评价 + 总览统计。"""
    _require_role(request, "admin")
    try:
        return {"ok": True, "items": feedback.list_all(feature, status, limit),
                "stats": feedback.stats()}
    except RuntimeError as e:
        raise HTTPException(status_code=503, detail=str(e))


@app.post("/api/admin/feedback/{fid}/status")
def api_admin_feedback_status(fid: int, body: FeedbackStatusIn, request: Request):
    me = _require_role(request, "admin")
    try:
        return {"ok": True, **feedback.set_status(fid, body.status, me["id"])}
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))


# ---------------------------------------------------------------- 管理端

class RolesIn(BaseModel):
    roles: list[str] = []


class StatusIn(BaseModel):
    status: str = "active"


class PwdIn(BaseModel):
    password: str


class NewUserIn(BaseModel):
    username: str
    password: str
    display_name: str | None = None


@app.get("/api/admin/overview")
def api_admin_overview(request: Request):
    _require_role(request, "admin")
    try:
        return {"ok": True, "stats": support.admin_overview()}
    except RuntimeError as e:
        raise HTTPException(status_code=503, detail=str(e))


@app.get("/api/admin/users")
def api_admin_users(request: Request, q: str = "", limit: int = 200):
    _require_role(request, "admin")
    try:
        return {"ok": True, "users": support.admin_users(q, limit)}
    except RuntimeError as e:
        raise HTTPException(status_code=503, detail=str(e))


@app.post("/api/admin/users/{uid}/roles")
def api_admin_roles(uid: int, body: RolesIn, request: Request):
    me = _require_role(request, "admin")
    if int(uid) == int(me["id"]) and "admin" not in body.roles:
        raise HTTPException(status_code=400, detail="不能取消自己的管理员权限")
    try:
        return {"ok": True, "user": auth.set_roles(uid, body.roles)}
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))


@app.post("/api/admin/users/{uid}/status")
def api_admin_status(uid: int, body: StatusIn, request: Request):
    me = _require_role(request, "admin")
    if int(uid) == int(me["id"]) and body.status == "disabled":
        raise HTTPException(status_code=400, detail="不能停用自己")
    try:
        return {"ok": True, "user": auth.set_status(uid, body.status)}
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))


@app.post("/api/admin/users/{uid}/password")
def api_admin_password(uid: int, body: PwdIn, request: Request):
    _require_role(request, "admin")
    try:
        auth.reset_password(uid, body.password)
        return {"ok": True}
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))


@app.post("/api/admin/users")
def api_admin_create_user(body: NewUserIn, request: Request):
    """管理员开客服账号。"""
    _require_role(request, "admin")
    try:
        return {"ok": True, "user": auth.create_staff(body.username, body.password,
                                                      body.display_name or "")}
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))


@app.get("/api/admin/staff")
def api_admin_staff(request: Request):
    _require_role(request, "admin")
    try:
        return {"ok": True, "staff": support.admin_staff()}
    except RuntimeError as e:
        raise HTTPException(status_code=503, detail=str(e))


@app.get("/api/admin/sessions")
def api_admin_sessions(request: Request, status: str = "", limit: int = 100):
    _require_role(request, "admin")
    try:
        return {"ok": True, "sessions": support.admin_sessions(status, limit)}
    except RuntimeError as e:
        raise HTTPException(status_code=503, detail=str(e))


@app.get("/api/admin/session/{sid}")
def api_admin_session(sid: int, request: Request):
    _require_role(request, "admin")
    try:
        return {"ok": True, **support.session_transcript(sid)}
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))


# ---------------------------------------------------------------- 个人资料（昵称 / 头像）

def _require_user(request: Request) -> dict:
    """从 cookie 认人；没登录 / 过期一律 401。"""
    try:
        user = auth.user_by_token(request.cookies.get(auth.TOKEN_COOKIE))
    except RuntimeError as e:
        raise HTTPException(status_code=503, detail=str(e))
    if not user:
        raise HTTPException(status_code=401, detail="登录已过期，请重新登录")
    return user


def _square_png(data: bytes) -> bytes:
    """任意图片 → 居中裁成方图 → 256×256 PNG。不是图片抛 ValueError。"""
    import io as _io
    try:
        from PIL import Image
    except ImportError as e:
        raise ValueError("服务器缺少 Pillow，暂时不能处理头像") from e
    try:
        with Image.open(_io.BytesIO(data)) as im:
            im.load()
            w, h = im.size
            side = min(w, h)
            left, top = (w - side) // 2, (h - side) // 2
            im = im.crop((left, top, left + side, top + side)).convert("RGBA")
            resample = getattr(Image, "Resampling", Image).LANCZOS
            im = im.resize((AVATAR_SIZE, AVATAR_SIZE), resample)
            buf = _io.BytesIO()
            im.save(buf, "PNG", optimize=True)
            return buf.getvalue()
    except Exception as e:
        raise ValueError("这张图片读不出来，换一张试试（jpg / png / webp）") from e


def _drop_avatar_file(name: str) -> None:
    """删旧头像文件（只认文件名，防止路径串出去）。"""
    if not name or name != Path(name).name:
        return
    try:
        p = auth.AVATAR_DIR / name
        if p.is_file():
            p.unlink()
    except Exception:
        pass


@app.post("/api/auth/profile")
def api_profile(body: ProfileIn, request: Request):
    """改昵称（显示名）。账号本身不能改，避免登录名混乱。"""
    user = _require_user(request)
    if body.display_name is None:
        raise HTTPException(status_code=400, detail="没有要修改的内容")
    try:
        fresh = auth.update_profile(user["id"], display_name=body.display_name)
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))
    except RuntimeError as e:
        raise HTTPException(status_code=503, detail=str(e))
    return {"ok": True, "user": fresh}


@app.post("/api/auth/avatar")
async def api_avatar_upload(request: Request, file: UploadFile = File(...)):
    """换头像：校验 → 裁方缩到 256 → 落盘 → 库里的 avatar 换成新文件名 → 删旧文件。"""
    user = _require_user(request)
    raw = await file.read()
    if not raw:
        raise HTTPException(status_code=400, detail="图片是空的，重新选一张")
    if len(raw) > MAX_AVATAR_BYTES:
        raise HTTPException(status_code=400, detail="图片最大 5 MB，压缩一下再传")
    ext = Path(file.filename or "").suffix.lower()
    if ext and ext not in AVATAR_EXTS:
        raise HTTPException(status_code=400, detail="只支持 jpg / png / gif / bmp / webp")
    try:
        png = _square_png(raw)
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))

    import secrets as _secrets
    name = "u%d_%s.png" % (user["id"], _secrets.token_hex(6))
    auth.AVATAR_DIR.mkdir(parents=True, exist_ok=True)
    (auth.AVATAR_DIR / name).write_bytes(png)

    old = auth.avatar_file(user["id"])
    try:
        fresh = auth.update_profile(user["id"], avatar=name)
    except Exception:
        _drop_avatar_file(name)          # 写库失败别留垃圾文件
        raise
    if old and old != name:
        _drop_avatar_file(old)
    return {"ok": True, "user": fresh}


@app.post("/api/auth/avatar/remove")
def api_avatar_remove(request: Request):
    """恢复默认头像（字母头像）。"""
    user = _require_user(request)
    old = auth.avatar_file(user["id"])
    try:
        fresh = auth.update_profile(user["id"], avatar="")
    except RuntimeError as e:
        raise HTTPException(status_code=503, detail=str(e))
    _drop_avatar_file(old)
    return {"ok": True, "user": fresh}


@app.get("/api/avatar/{name}")
def api_avatar_get(name: str):
    """头像图片。文件名每次换头像都变，所以可以放心长缓存。"""
    if not name or name != Path(name).name or name.startswith("."):
        raise HTTPException(status_code=404, detail="头像不存在")
    p = auth.AVATAR_DIR / name
    if not p.is_file():
        raise HTTPException(status_code=404, detail="头像不存在")
    return FileResponse(p, headers={"Cache-Control": "public, max-age=31536000"})


@app.get("/api/status")
def api_status():
    """服务状态：服务本身 + 数据库 + 文档库。"""
    dirs = hub.docs_dir if isinstance(hub.docs_dir, (list, tuple)) else [hub.docs_dir]
    st = {
        "service": "ok",
        "time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "docs_dir": " ｜ ".join(str(d) for d in dirs),
        "jobs_total": len(hub.jobs),
        "sessions": len(hub.sessions),
    }
    try:
        st["docs_count"] = len(scan_docs(hub.docs_dir))
    except Exception:
        st["docs_count"] = None
    try:
        db.recent(1)
        st["database"] = "ok"
    except RuntimeError as e:
        st["database"] = "unavailable"
        st["database_tip"] = str(e)
    return st


def _docs_detail(names: list) -> list:
    """给文档库清单补上"看得见"的信息：格式 / 大小 / 修改时间 / 来源。

    scan_docs 只回文件名（给模型当清单刚好）；但「数据仓库 → 文档库」那块是给人看的，
    得知道多大、什么时候进来的、是自己传的还是标准库自带的，否则同名不同版分不出来。
    以 scan_docs 的结果为准——它认为在库里的才列，顺序也照它，绝不多列。
    """
    dirs = hub.docs_dir if isinstance(hub.docs_dir, (list, tuple)) else [hub.docs_dir]
    want = set(names)
    try:
        import_root = Path(IMPORT_DIR).resolve()
    except OSError:
        import_root = Path(IMPORT_DIR)
    meta = {}
    for d in dirs:
        root = Path(d)
        try:
            src = "import" if root.resolve() == import_root else "standard"
            files = sorted(root.rglob("*"))
        except OSError:
            continue
        for p in files:
            if p.name not in want or p.name in meta:
                continue
            try:
                if not p.is_file():
                    continue
                st = p.stat()
            except OSError:
                continue
            meta[p.name] = {
                "name": p.name,
                "ext": p.suffix.lower().lstrip("."),
                "size": st.st_size,
                "mtime": time.strftime("%Y-%m-%d %H:%M", time.localtime(st.st_mtime)),
                "src": src,
            }
    return [meta[n] for n in names if n in meta]


@app.get("/api/docs")
def api_docs():
    """文档库清单（标准库 + 导入库 合并）。

    docs   —— 文件名数组（老格式，前端/模型都按字符串用，不变）
    detail —— 明细（格式 / 大小 / 修改时间 / 来源），给「数据仓库 → 文档库」展示
    counts —— 总数与我上传的份数
    """
    names = scan_docs(hub.docs_dir)
    detail = _docs_detail(names)
    return {
        "docs": names,
        "detail": detail,
        "counts": {
            "total": len(names),
            "import": sum(1 for x in detail if x.get("src") == "import"),
        },
        "import_dir": str(IMPORT_DIR),
    }


@app.delete("/api/docs")
def api_docs_delete(name: str):
    """删掉一份【我上传的】文档（2026-11-07）。

    标准库文档不许删（比赛基线）；上传的文件不真删，挪到
    导入文档/_已删除/ 下带时间戳，删错了能捞回来。
    """
    r = doc_pool.remove(name)
    if not r.get("ok"):
        raise HTTPException(status_code=403 if "标准文档" in str(r.get("error")) else 400,
                            detail=r.get("error") or "删除失败")
    try:
        total = len(scan_docs(hub.docs_dir))
    except Exception:
        total = None
    return {"ok": True, "name": r.get("name"), "archived": r.get("archived"),
            "docs_total": total}


# ---------------------------------------------------------------- 文档导入（模块2 · 第1步）

# 能导入的格式 = 读取器支持的格式（pdf / 图片也在这里，扫描件走 OCR）
SUPPORTED_EXTS = set(doc_pool.SUPPORTED)


def _ext_desc() -> str:
    return ("Word(.docx) / Excel(.xlsx) / PDF(.pdf) / Markdown(.md) / 纯文本(.txt) / "
            "图片(.png .jpg .jpeg .bmp .webp .tif)")


@app.post("/api/upload")
def api_upload(file: UploadFile = File(...)):
    """上传一份非结构化文档 → 自动识别格式并解析 → 并入检索池 → 记录入库。

    上传后这份文件【立刻】可被指令操作（编辑 / 排版 / 格式 / 提取），无需重启或重扫。
    重名自动加后缀，绝不覆盖标准文档或此前上传的文件。

    PDF / 图片：自带文字层的 PDF 直接抽；扫描页和图片走 OCR（见 agent/ocr.py），
    认出来的文字同样带页码定位进池子，后面的检索 / 抽取 / 填表照常能用。
    """
    raw_name = Path(file.filename or "").name            # 防路径注入
    if not raw_name:
        raise HTTPException(status_code=400, detail="没收到文件名")
    ext = Path(raw_name).suffix.lower()
    if ext not in SUPPORTED_EXTS:
        note = "不支持的格式「%s」" % (ext or "无扩展名")
        try:
            db.log_import(raw_name, raw_name, None, None, None, False, None, note)
        except RuntimeError:
            pass
        raise HTTPException(
            status_code=400,
            detail="不支持的格式「%s」——当前支持：%s"
                   % (ext or "无扩展名", _ext_desc()))

    # 1) 落盘（规划不冲突的存储名）
    stored = doc_pool.plan_stored_name(raw_name)
    data = file.file.read()
    try:
        stored.write_bytes(data)
    except OSError as e:
        raise HTTPException(status_code=500, detail="文件保存失败：%s" % e)

    # 2) 自动识别 + 解析 + 并入检索池
    #    PDF / 图片：如果刚在「文字识别」里认出过，直接用那份结果入池，不重复 OCR
    pre = None
    if ext == ".pdf" or ext in OCR_IMAGE_EXTS:
        pre = _ocr_cache_get(data)
    res = doc_pool.add(stored, doc_override=pre)

    # 3) 导入记录入库（数据库不可用不拦上传）
    db_tip = None
    try:
        db.log_import(raw_name, stored.name, res.get("type"),
                      len(data), res.get("blocks"), bool(res.get("ok")),
                      res.get("elapsed"), res.get("error") or "")
    except RuntimeError as e:
        db_tip = str(e)

    if not res.get("ok"):
        return {"ok": False, "filename": raw_name, "stored_name": stored.name,
                "error": res.get("error"), "db_tip": db_tip}

    try:
        docs_total = len(scan_docs(hub.docs_dir))
    except Exception:
        docs_total = None
    return {
        "ok": True, "filename": raw_name, "stored_name": stored.name,
        "filetype": res.get("type"), "blocks": res.get("blocks"),
        "size": len(data), "elapsed": res.get("elapsed"),
        "docs_total": docs_total, "db_tip": db_tip,
        # PDF / 图片专有：几页、几页是扫描页走了 OCR、平均置信度
        "page_count": res.get("page_count"),
        "ocr_page_count": res.get("ocr_page_count"),
        "scanned": res.get("scanned"),
        "avg_conf": res.get("avg_conf"),
        "notes": res.get("notes"),
        "reused_ocr": pre is not None,
    }


# ---------------------------------------------------------------- OCR（扫描版 PDF / 图片 → 文字）

OCR_IMAGE_EXTS = {".png", ".jpg", ".jpeg", ".bmp", ".webp", ".tif", ".tiff"}
OCR_MAX_BYTES = 40 * 1024 * 1024            # 单文件 40MB（扫描件常见 10~20MB）

# 刚识别过的文件先留在内存里：「预览」之后点「并入文档库」直接复用，
# 不再 OCR 第二遍（一份 20 页扫描件能省十几秒）。只留最近 8 份，不占内存。
_OCR_CACHE = {}
_OCR_CACHE_MAX = 8


def _ocr_cache_put(data: bytes, doc: dict):
    key = hashlib.sha1(data).hexdigest()
    _OCR_CACHE[key] = doc
    while len(_OCR_CACHE) > _OCR_CACHE_MAX:
        _OCR_CACHE.pop(next(iter(_OCR_CACHE)))


def _ocr_cache_get(data: bytes):
    return _OCR_CACHE.get(hashlib.sha1(data).hexdigest())


def _ocr_paragraphs(blocks: list) -> list:
    """块 → 前端展示用的扁平段落（带页码 / 来源 / 置信度）。"""
    out = []
    for b in blocks or []:
        loc = b.get("locator") or {}
        kind = loc.get("kind") or ""
        out.append({
            "text": b.get("text") or "",
            "page": loc.get("page"),
            "source": b.get("source") or ("ocr" if kind.endswith("ocr") else "text"),
            "conf": b.get("ocr_conf", loc.get("conf")),
        })
    return out


@app.get("/api/ocr/status")
def api_ocr_status():
    """OCR 引擎状态（前端据此提示"能不能用 / 该装什么"）。"""
    return ocr_engine.status()


# ---------------------------------------------------------------- 字段语义匹配
# 2026-10-26：字段匹配从"别名表 + 字面"改成"本地向量语义 + 别名兜底"。
# 这两个接口是给前端和排查用的：能把"凭什么认为『合同总价』就是『合同总金额』"说清楚。

class SemanticMatchIn(BaseModel):
    query: str
    candidates: list[str] | None = None
    doc: str | None = None
    topk: int = 3


@app.get("/api/llm/status")
def api_llm_status():
    """当前模型线路状态：用哪家、哪家还配着、哪家在冷却。

    2026-11-02 加：免费额度用完会自动换下一家，但"换到哪家了"得能看见，
    否则出了问题只能靠猜。前端和排障都看这个。
    """
    st = llm_model.status()
    st["describe"] = llm_model.describe()
    st["has_key"] = llm_model.has_key()
    return st


@app.get("/api/semantic/status")
def api_semantic_status():
    """语义模型状态：装没装、多大、多少维、阈值多少。

    模型缺失时 available=false 并给出原因，但**不影响其他功能** ——
    匹配会自动降级成"别名表 + 字面"。
    """
    return sem_engine.status()


@app.post("/api/semantic/match")
def api_semantic_match(body: SemanticMatchIn):
    """把一句话对到候选字段上，返回 Top-K 和每一名的依据。

    body.candidates 不传时，用 body.doc 那份文档里的真实说法（表格列头 + "XX："标签）。
    """
    cands = body.candidates
    if not cands:
        if not body.doc:
            raise HTTPException(status_code=400, detail="要么给 candidates，要么给 doc")
        import extract_ops
        cands = extract_ops.doc_candidates(body.doc)
    if not cands:
        return {"query": body.query, "candidates": [], "matches": [],
                "note": "这份文档里没找到可用的字段说法（没有表格列头，也没有「XX：」标签）"}

    res = sem_engine.match(body.query, cands, topk=max(1, min(10, body.topk)),
                           aliases=_load_aliases_safe(), min_score=0.0)
    return {
        "query": body.query,
        "candidate_count": len(cands),
        "matches": res,
        "note": "score 是三路（别名表 / 字面 / 语义向量）里最强的一路；via 说明靠哪一路命中",
    }


def _load_aliases_safe() -> dict:
    """读别名表（坏了就当空表 —— 别名表只是兜底，不能成为故障点）。"""
    try:
        import extract_ops
        return extract_ops.load_aliases()
    except Exception:
        return {}


@app.post("/api/ocr")
def api_ocr(file: UploadFile = File(...), dpi: int = 200,
            lang: str = "auto", vertical: str = "auto"):
    """上传一张图或一份 PDF → 立刻识别出文字（预览用，不写进检索池）。

    想让它变成"能下指令的文档"，识别完点「并入文档库」即可（走 /api/upload）。

    lang      auto / chs / cht / en —— 识别模型（2026-11-07 加：繁体、英文）
    vertical  auto / on / off —— 竖排处理（竖着的字按从上到下、从右到左还原）

    返回：
      text        全文
      paragraphs  逐段（带页码 / 是否 OCR / 置信度）—— 前端按这个渲染
      scanned     PDF 里有没有扫描页
      elapsed     耗时（秒）
    """
    raw_name = Path(file.filename or "").name
    if not raw_name:
        raise HTTPException(status_code=400, detail="没收到文件名")
    ext = Path(raw_name).suffix.lower()
    if ext != ".pdf" and ext not in OCR_IMAGE_EXTS:
        raise HTTPException(
            status_code=400,
            detail="这个接口只认 PDF 和图片（%s）——Word / Excel 请直接用「📥 导入文档」"
                   % " / ".join(sorted(OCR_IMAGE_EXTS)))

    data = file.file.read()
    if not data:
        raise HTTPException(status_code=400, detail="文件是空的")
    if len(data) > OCR_MAX_BYTES:
        raise HTTPException(status_code=413,
                            detail="文件太大（%.1f MB），上限 %d MB"
                                   % (len(data) / 1024 / 1024, OCR_MAX_BYTES // 1024 // 1024))

    st = ocr_engine.status()
    if not st.get("ok"):
        raise HTTPException(status_code=503, detail=st.get("tip") or "OCR 引擎不可用")
    if ext == ".pdf" and not st.get("pdf"):
        raise HTTPException(status_code=503,
                            detail=st.get("pdf_tip") or "读 PDF 需要 PyMuPDF")

    import time as _t
    t0 = _t.time()

    # ---- 图片：直接认 ----
    if ext in OCR_IMAGE_EXTS:
        r = ocr_engine.read_image_bytes(data, raw_name, lang=lang, vertical=vertical)
        if not r.get("ok"):
            raise HTTPException(status_code=500, detail=r.get("error") or "OCR 失败")
        blocks = []
        for p in r.get("paragraphs") or []:
            blocks.append({"text": p.get("text"), "source": "ocr",
                           "locator": {"kind": "image_para", "conf": p.get("conf")},
                           "ocr_conf": p.get("conf")})
        paras = _ocr_paragraphs(blocks)
        _ocr_cache_put(data, {"file": raw_name, "type": "image", "blocks": blocks,
                              "ocr": True, "avg_conf": r.get("avg_conf"),
                              "char_count": r.get("char_count")})
        return {
            "ok": True, "filename": raw_name, "kind": "image",
            "elapsed": round(_t.time() - t0, 2),
            "engine": r.get("engine"),
            "width": r.get("width"), "height": r.get("height"),
            "char_count": r.get("char_count"), "para_count": len(paras),
            "avg_conf": r.get("avg_conf"),
            "low_conf_lines": r.get("low_conf_lines") or 0,
            "scanned": False, "page_count": None, "ocr_page_count": None,
            "text": r.get("text") or "",
            "paragraphs": paras,
            "lang": r.get("lang"), "lang_label": r.get("lang_label"),
            "rec_model": r.get("rec_model"), "vertical": r.get("vertical"),
            "notes": [x for x in [r.get("vertical"), r.get("tip"), r.get("low_conf_tip")] if x],
        }

    # ---- PDF：落临时文件 → 复用 pdf_reader（逐页判断文字层 / 扫描页） ----
    from readers import read_pdf                       # noqa: E402  延迟导入
    tmp = Path(tempfile.mkdtemp(prefix="a23ocr_")) / raw_name
    try:
        tmp.write_bytes(data)
        try:
            doc = read_pdf(tmp, dpi=max(120, min(int(dpi), 400)),
                           lang=lang, vertical=vertical)
        except ValueError as e:
            raise HTTPException(status_code=422, detail=str(e))
        except RuntimeError as e:
            raise HTTPException(status_code=500, detail=str(e))
    finally:
        try:
            tmp.unlink()
            tmp.parent.rmdir()
        except OSError:
            pass

    blocks = doc.get("blocks") or []
    _ocr_cache_put(data, doc)               # 供紧接着的「并入文档库」复用
    paras = _ocr_paragraphs(blocks)
    text = "\n\n".join(b.get("text") or "" for b in blocks)
    notes = list(doc.get("notes") or [])
    if doc.get("ocr_failed"):
        bad = "、".join("第 %d 页" % x["page"] for x in doc["ocr_failed"][:5])
        notes.append("有 %d 页没认出来（%s）——多半是空白页或整页图表"
                     % (len(doc["ocr_failed"]), bad))
    return {
        "ok": True, "filename": raw_name, "kind": "pdf",
        "elapsed": round(_t.time() - t0, 2),
        "page_count": doc.get("page_count"),
        "text_pages": len(doc.get("text_pages") or []),
        "ocr_page_count": doc.get("ocr_page_count"),
        "scanned": bool(doc.get("scanned")),
        "ocr_elapsed": doc.get("ocr_elapsed"),
        "dpi": doc.get("dpi"),
        "char_count": len("".join(text.split())),
        "para_count": len(paras),
        "text": text,
        "paragraphs": paras,
        "lang": ocr_engine._norm_lang(lang),
        "lang_label": ocr_engine.LANG_LABEL.get(ocr_engine._norm_lang(lang)),
        "rec_model": ocr_engine._rec_model(lang)[1],
        "notes": notes,
    }


# ---------------------------------------------------------------- 图片上传（模块1 · 插图）

@app.post("/api/upload_image")
def api_upload_image(file: UploadFile = File(...)):
    """上传一张图片 → 返回 image_id；随后发一句"把这张图插到××"即可。

    图片只落盘到 导入图片/，不进检索池（它不是文档，不该混进 blocks），
    也不参与模块2 的抽取。
    """
    raw_name = Path(file.filename or "").name
    if not raw_name:
        raise HTTPException(status_code=400, detail="没收到文件名")
    data = file.file.read()
    try:
        info = image_ops.save_image(data, raw_name)
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))
    except OSError as e:
        raise HTTPException(status_code=500, detail="图片保存失败：%s" % e)
    return {"ok": True, **info}


@app.get("/api/images")
def api_images(limit: int = 50):
    """最近上传的图片清单（新的在前）。"""
    limit = max(1, min(int(limit), 200))
    return {"images": image_ops.list_images(limit)}


@app.get("/api/images/{name}")
def api_image_file(name: str):
    """图片预览（卡片缩略图直接用 <img src>）。"""
    path = image_ops.resolve_image(name)
    if path is None or not path.is_file():
        raise HTTPException(status_code=404, detail="图片不存在")
    return FileResponse(str(path))


@app.get("/api/imports")
def api_imports(limit: int = 20):
    """导入记录（MySQL imports 表，最新在前）。"""
    limit = max(1, min(int(limit), 100))
    try:
        rows = db.recent_imports(limit)
    except RuntimeError as e:
        return {"rows": [], "error": str(e)}
    out = []
    for r in rows:
        out.append({
            "id": r["id"], "ts": str(r["ts"])[:19],
            "filename": r["filename"], "stored_name": r["stored_name"],
            "filetype": r.get("filetype"), "size_bytes": r.get("size_bytes"),
            "blocks": r.get("blocks"), "ok": bool(r.get("ok")),
            "elapsed_sec": r.get("elapsed_sec"), "note": r.get("note"),
        })
    return {"rows": out}


class CommandIn(BaseModel):
    text: str
    session_id: str | None = None
    image_id: str | None = None          # 本轮要插入的图片（上传后返回的 id）


@app.post("/api/command")
def api_command(body: CommandIn):
    """提交指令：立即返回 job_id（异步），结果用 /api/job/{id} 轮询。"""
    text = (body.text or "").strip()
    if not text:
        raise HTTPException(status_code=400, detail="指令不能为空")
    job = hub.submit(text, (body.session_id or "").strip() or None,
                     image=(body.image_id or "").strip() or None)
    return {"job_id": job.id, "session_id": job.session_id, "status": job.status}


# ---------------------------------------------------------------- 智能操作（模块① 独立入口）
# 两条路走同一套执行器（edit.run_edit）：
#   ① 快捷按钮 → 前端直接说 op + params，这里拼成意图（preset）→ **不调模型**
#   ② 下面的自然语言框 → 照常走解析（和首页对话框一模一样），只是带上"点选的文件"
_OP_ACTION = {
    "set_align": "排版", "set_spacing": "排版", "set_indent": "排版",
    "set_bold": "格式调整", "set_italic": "格式调整", "set_underline": "格式调整",
    "set_strike": "格式调整", "set_highlight": "格式调整",
    "set_color": "格式调整", "set_font": "格式调整", "set_size": "格式调整",
}
_OP_CN = {
    "replace": "替换", "insert": "插入", "delete": "删除",
    "set_align": "对齐", "set_spacing": "行距", "set_indent": "缩进",
    "set_bold": "加粗", "set_italic": "斜体", "set_underline": "下划线",
    "set_strike": "删除线", "set_highlight": "高亮", "set_color": "颜色",
    "set_font": "字体", "set_size": "字号",
}


def _op_phrase(op: str, params: dict) -> str:
    """把 op + params 说成一句人话（报告标题 / 操作历史里显示的文案）。"""
    p = params or {}
    if op == "set_align":
        return {"center": "居中", "left": "左对齐", "right": "右对齐",
                "justify": "两端对齐"}.get(str(p.get("align")), "对齐")
    if op == "set_spacing":
        return "行距调成 %s" % (p.get("spacing") or "")
    if op == "set_indent":
        return str(p.get("indent") or "首行缩进")
    if op == "set_font":
        return "字体改成 %s" % (p.get("font") or "")
    if op == "set_size":
        return "字号调成 %s" % (p.get("size") or "")
    if op == "set_color":
        return "标成%s" % (p.get("color") or "红色")
    if op == "set_highlight":
        return "%s高亮" % (p.get("color") or "黄色")
    if p.get("off"):
        return "取消%s" % _OP_CN.get(op, op)
    return _OP_CN.get(op, op)


class OpIn(BaseModel):
    """智能操作的一次提交。

    text  ｜ 自然语言那句（快捷按钮可不填）
    doc   ｜ 用户点选的文件名（可以没有：那就从 text 里点名）
    op / params / hint / by ｜ 快捷按钮用；给了 op 就不调模型
    """
    text: str = ""
    doc: str | None = None
    op: str | None = None
    params: dict | None = None
    hint: str = ""
    by: str | None = None


@app.post("/api/op")
def api_op(body: OpIn):
    """提交一次「智能操作」：快捷按钮（op）或自然语言（text）都走这里。"""
    doc = (body.doc or "").strip() or None
    text = (body.text or "").strip()
    op = (body.op or "").strip() or None

    if op:
        if op not in _OP_ACTION:
            raise HTTPException(status_code=400,
                                detail="不支持的快捷操作：%s（可用：%s）"
                                       % (op, "、".join(_OP_ACTION)))
        hint = (body.hint or "").strip()
        by = (body.by or ("keyword" if hint else "whole")).strip()
        if by not in ("whole", "keyword", "table"):
            by = "keyword" if hint else "whole"
        intent = {
            "action": _OP_ACTION[op],
            "op": op,
            "doc": doc,
            "doc_hint": doc or (hint or ""),
            "locate": {"by": by, "hint": hint},
            "params": dict(body.params or {}),
            "clarify": "",
        }
        label = "快捷·%s" % _OP_CN.get(op, op)
        if not text:
            text = "把%s%s%s" % ("《%s》的" % doc if doc else "",
                                 (hint if hint and hint != "全文" else ""),
                                 _op_phrase(op, body.params or {}))
        job = hub.submit_op(text=text, doc=doc, preset=intent, label=label)
        return {"job_id": job.id, "status": job.status, "mode": "quick", "op": op}

    if not text:
        raise HTTPException(status_code=400, detail="说一句要做什么，或者点上面的快捷按钮")
    job = hub.submit_op(text=text, doc=doc, preset=None, label=text[:40])
    return {"job_id": job.id, "status": job.status, "mode": "nl"}


class OpOpenIn(BaseModel):
    file: str                      # 产物文件名（不带目录）
    mode: str = "file"             # file = 打开；folder = 在资源管理器里定位


def _find_edit_output(name: str):
    """按文件名找「智能操作」的产物，只认 agent/输出/ 下真实存在的文件。

    和 /api/fill/open 同一套谨慎：路径由服务端自己拼，前端只给文件名，
    免得这个接口变成"能在本机打开任意文件"的口子。
    """
    name = Path(name or "").name
    if not name:
        return None
    d = (BASE / "agent" / "输出").resolve()
    p = (d / name).resolve()
    try:
        if p.parent != d or not p.is_file():
            return None
    except OSError:
        return None
    return p


@app.post("/api/op/open")
def api_op_open(body: OpOpenIn):
    """在**这台机器**上打开智能操作产物 / 打开它所在目录。"""
    p = _find_edit_output(body.file)
    if p is None:
        raise HTTPException(status_code=404,
                            detail="没找到这份产物：%s（跑一次操作后再点）" % body.file)
    try:
        if body.mode == "folder":
            subprocess.Popen(["explorer", "/select,", str(p)])
        elif os.name == "nt":
            os.startfile(str(p))                       # noqa: S606（本机应用，路径由服务端算）
        else:
            subprocess.Popen(["xdg-open", str(p)])
    except Exception as e:
        raise HTTPException(status_code=500, detail="打开失败：%s" % e)
    return {"ok": True, "path": str(p), "mode": body.mode}


@app.get("/api/job/{job_id}")
def api_job(job_id: str):
    """查任务状态/结果。"""
    job = hub.get_job(job_id)
    if job is None:
        raise HTTPException(status_code=404, detail="任务不存在（服务重启后旧任务会丢失）")
    return job.to_dict()


@app.get("/api/history")
def api_history(limit: int = 15):
    """操作历史（MySQL）。"""
    limit = max(1, min(int(limit), 100))
    try:
        rows = db.recent(limit)
    except RuntimeError as e:
        return {"rows": [], "error": str(e)}
    out = []
    for r in rows:
        out.append({
            "id": r["id"], "ts": str(r["ts"])[:19], "session_id": r["session_id"],
            "turn": r["turn"], "command": r["command"],
            "action": r.get("action"), "op": r.get("op"), "doc": r.get("doc"),
            "out_file": r.get("out_file"), "ok": bool(r.get("ok")),
            "elapsed_sec": r.get("elapsed_sec"), "note": r.get("note"),
        })
    return {"rows": out}


@app.get("/api/session/{sid}")
def api_session(sid: str):
    """会话状态：当前锁定文档 / 已跑轮数。"""
    with hub.sessions_lock:
        sess = hub.sessions.get(sid)
    if sess is None:
        return {"session_id": sid, "exists": False}
    return {
        "session_id": sid, "exists": True,
        "turn": sess.turn,
        "current_doc": sess.current_doc,
        "work_copy": str(sess.work_copy) if sess.work_copy else None,
    }


# ---------------------------------------------------------------- 数据仓库（模块2 第4步：提取结果入库）

@app.get("/api/db/stats")
def api_db_stats():
    """数据仓库统计：文档 / 实体 / 画像 / 字段 各多少条。"""
    try:
        return {"ok": True, "stats": db_store.stats()}
    except RuntimeError as e:
        return {"ok": False, "error": str(e)}


@app.get("/api/db/entities")
def api_db_entities(doc: str = None, type: str = None, q: str = None, limit: int = 200):
    """按文档 / 类型 / 关键词查实体数据。"""
    limit = max(1, min(int(limit), 500))
    try:
        rows = db_store.query_entities(doc, type, q, limit)
    except RuntimeError as e:
        return {"ok": False, "error": str(e), "rows": []}
    return {"ok": True, "rows": jsonable(rows)}


@app.get("/api/db/fields")
def api_db_fields(doc: str = None, q: str = None, limit: int = 200):
    """查"用户指定内容"（字段提取）结果。"""
    limit = max(1, min(int(limit), 500))
    try:
        rows = db_store.query_fields(doc, q, limit)
    except RuntimeError as e:
        return {"ok": False, "error": str(e), "rows": []}
    return {"ok": True, "rows": jsonable(rows)}


@app.get("/api/db/profile")
def api_db_profile(doc: str):
    """看某份文档的画像（主题 + 要点 + 要素）。"""
    try:
        p, items = db_store.get_profile(doc)
    except RuntimeError as e:
        return {"ok": False, "error": str(e)}
    if not p:
        return {"ok": True, "found": False}
    return {"ok": True, "found": True, "profile": jsonable(p), "items": jsonable(items)}


# ---------------------------------------------------------------- 统一知识（2026-11-05）
# 数据仓库的主视图：不再分"实体/画像/字段"三张表去查，
# 而是「这份文档有哪些对象、各有什么属性、值是多少、出自哪里」。

class KnowledgeIn(BaseModel):
    docs: list[str]
    requirement: str = ""


@app.get("/api/knowledge/stats")
def api_knowledge_stats():
    try:
        return {"ok": True, "stats": jsonable(db_store.knowledge_stats())}
    except RuntimeError as e:
        return {"ok": False, "error": str(e)}


@app.get("/api/knowledge/items")
def api_knowledge_items(doc: str = None, kind: str = None, obj: str = None,
                        q: str = None, limit: int = 300):
    """查知识条目。kind ∈ object / attribute / relation / event / summary。"""
    limit = max(1, min(int(limit), 1000))
    try:
        rows = db_store.query_knowledge(doc, kind, obj, q, limit)
    except RuntimeError as e:
        return {"ok": False, "error": str(e), "rows": []}
    return {"ok": True, "rows": jsonable(rows)}


@app.get("/api/knowledge/schema")
def api_knowledge_schema(doc: str):
    """某份文档归纳出的结构（对象/属性/领域/用途）。"""
    try:
        row = db_store.get_schema(doc)
    except RuntimeError as e:
        return {"ok": False, "error": str(e)}
    if not row:
        return {"ok": True, "found": False}
    item = jsonable(row)
    for k in ("objects", "attr_keys"):
        try:
            item[k] = json.loads(item.get(k) or "[]")
        except Exception:
            item[k] = []
    return {"ok": True, "found": True, "schema": item}


@app.get("/api/knowledge/schemas")
def api_knowledge_schemas(limit: int = 100):
    try:
        rows = db_store.list_schemas(limit)
    except RuntimeError as e:
        return {"ok": False, "error": str(e), "rows": []}
    out = []
    for r in rows:
        d = jsonable(r)
        try:
            d["attr_keys"] = json.loads(d.get("attr_keys") or "[]")
            d["attr_count"] = len(d["attr_keys"])
        except Exception:
            d["attr_count"] = 0
        out.append(d)
    return {"ok": True, "rows": out}


@app.get("/api/knowledge/docs")
def api_knowledge_docs(limit: int = 300):
    """数据仓库左栏：有知识的文档清单（带条目数 / 对象数 / 属性数 / 领域）。

    2026-11-05 加 —— 用户要"按文档分类"，不再把所有文档的条目混成一张表。
    """
    try:
        rows = db_store.knowledge_docs(limit)
    except RuntimeError as e:
        return {"ok": False, "error": str(e), "rows": []}
    return {"ok": True, "rows": jsonable(rows)}


@app.get("/api/knowledge/doc")
def api_knowledge_doc(doc: str, limit: int = 3000):
    """点进某份文档：结构 + 按对象分组的条目 + 每个属性的值清单。

    lists 里就是"这份文档里有哪些城市/哪些学生/哪些单号"的答案。
    """
    limit = max(1, min(int(limit), 5000))
    try:
        d = db_store.knowledge_doc_detail(doc, limit)
    except RuntimeError as e:
        return {"ok": False, "error": str(e)}
    if not d:
        return {"ok": True, "found": False, "doc": doc}
    return {"ok": True, "found": True,
            "doc_name": d["doc_name"], "schema": d["schema"],
            "counts": d["counts"], "total": d["total"],
            "groups": d["groups"], "lists": d["lists"]}


@app.get("/api/knowledge/graph")
def api_knowledge_graph(doc: str = "", limit: int = 6000,
                        views: str = "rel,shared,type",
                        min_shared: int = 2,
                        collapse: int = 0,
                        max_nodes: int = 260, max_edges: int = 900):
    """知识图谱数据：节点 + 边 + 关系类型清单。

    doc  为空 = 全库；给了就只看这一份（前端"点进文档看图谱"用它）。
    views 选画哪几类关系（逗号分隔）：rel 显式关系 / attr 属性 / shared 共享属性
          / type 同类型 / doc 同文档 —— 对应前端"关系视角"。
    min_shared   一个值要被几个对象共用才算"共享"（默认 2，即同值即关联）
    collapse=1   关系转换：把「A -属性-> 值 <-属性- B」折叠成「A -[属性=值]-> B」
    max_nodes / max_edges  画布上限（超了先丢值节点，对象优先保留）

    前端只负责画，关系怎么来的全在这里定，保证图上每条边都能回到原文条目。
    """
    limit = max(50, min(int(limit), 20000))
    try:
        g = db_store.knowledge_graph(
            doc=(doc or "").strip() or None, limit=limit,
            views=views, min_shared=max(2, min(int(min_shared or 2), 12)),
            collapse=bool(int(collapse or 0)),
            max_nodes=max(20, min(int(max_nodes or 260), 600)),
            max_edges=max(40, min(int(max_edges or 900), 4000)))
    except RuntimeError as e:
        return {"ok": False, "error": str(e), "nodes": [], "edges": [],
                "stats": {}, "relation_types": []}
    return {"ok": True, "doc": doc or "", **jsonable(g)}


@app.post("/api/knowledge/run")
def api_knowledge_run(body: KnowledgeIn):
    """提交统一抽取任务（异步）。语义与 /api/recognize 的 knowledge 线一致。"""
    docs_in = [d.strip() for d in (body.docs or []) if d and d.strip()]
    if not docs_in:
        raise HTTPException(status_code=400, detail="请至少给一个文档名（或名字片段）")
    resolved, missing = [], []
    for d in docs_in:
        p = find_doc_file(d, hub.docs_dir)
        if p is None:
            missing.append(d)
        elif p.name not in resolved:
            resolved.append(p.name)
    if missing:
        raise HTTPException(status_code=404,
                            detail="文档库里找不到：%s（可先用 GET /api/docs 看清单）" % "、".join(missing))
    job = hub.submit_recognize(resolved, ["knowledge"], [], requirement=(body.requirement or "").strip())
    return {"job_id": job.id, "status": job.status, "kind": job.kind,
            "docs": resolved, "requirement": (body.requirement or "").strip()}


@app.get("/api/db/search")
def api_db_search(q: str, limit: int = 100):
    """全库搜关键词（实体值 + 字段值）。"""
    limit = max(1, min(int(limit), 300))
    try:
        rows = db_store.search(q, limit)
    except RuntimeError as e:
        return {"ok": False, "error": str(e), "rows": []}
    return {"ok": True, "rows": jsonable(rows)}


@app.get("/api/db/audit")
def api_db_audit(doc: str = None):
    """质量体检（模块2 第5步）：反向审计（库→原文）+ 三线查全率（零模型、秒级）。"""
    try:
        audit = quality_engine.audit_all(doc=doc)
        scores = quality_engine.score_all(doc=doc)
        summ = quality_engine.summarize(audit, scores)
    except RuntimeError as e:
        return {"ok": False, "error": str(e)}
    except Exception as e:
        return {"ok": False, "error": "%s: %s" % (type(e).__name__, e)}
    return {"ok": True, "audit": summ["audit"], "scores": summ["scores"],
            "elapsed": audit["elapsed"]}


# ---------------------------------------------------------------- 智能识别分类（异步 API · 赛题技术要求）

class RecognizeIn(BaseModel):
    docs: list[str]                       # 文档名（或名字片段），1 个或多个
    tasks: list[str] = ["knowledge"]      # 默认走统一知识抽取；老三条线仍可显式指定
    fields: list[str] = []                # tasks 含 fields 时必填：要提取的字段名（用户指定内容）
    requirement: str = ""                 # 用户关注点（原话）：决定抽什么优先、哪些条目重点标记


@app.post("/api/recognize")
def api_recognize(body: RecognizeIn):
    """智能识别：提交文档 → 异步跑「统一知识抽取」并自动入库（默认）。

    老三条线（画像 / 实体 / 字段）仍可用 tasks 显式指定，兼容旧脚本。
    提交立即返回 job_id（不阻塞）；用 GET /api/job/{id} 轮询进度与结果。
    """
    docs_in = [d.strip() for d in (body.docs or []) if d and d.strip()]
    if not docs_in:
        raise HTTPException(status_code=400, detail="请至少给一个文档名（或名字片段）")

    tasks = []
    for t in (body.tasks or []):
        key = Hub.TASK_ALIASES.get(str(t).strip())
        if key and key not in tasks:
            tasks.append(key)
    # 什么都不传 / 传了空数组 / 传了不认识的名字 —— 一律按统一知识抽取跑。
    # 面向用户的那句话是「点一下就把文档读懂入库」，不该逼他先选抽哪条线。
    if not tasks:
        tasks = ["knowledge"]

    fields = [f.strip() for f in (body.fields or []) if f and f.strip()]
    if ("fields" in tasks) and not fields:
        raise HTTPException(status_code=400,
                            detail="选择了「字段」线时，请在 fields 里给出要提取的字段名（如 「甲方名称」「合同总金额」）")

    # 解析文档（精确 → 包含；找不到就报错，绝不悄悄换文档）
    resolved, missing = [], []
    for d in docs_in:
        p = find_doc_file(d, hub.docs_dir)
        if p is None:
            missing.append(d)
        elif p.name not in resolved:
            resolved.append(p.name)
    if missing:
        raise HTTPException(status_code=404,
                            detail="文档库里找不到：%s（可先用 GET /api/docs 看清单）" % "、".join(missing))

    job = hub.submit_recognize(resolved, tasks, fields, requirement=(body.requirement or "").strip())
    return {"job_id": job.id, "status": job.status, "kind": job.kind,
            "docs": resolved, "tasks": tasks, "fields": fields,
            "requirement": (body.requirement or "").strip()}


# ---------------------------------------------------------------- 表格填写（异步 API · 模块3）

class FillIn(BaseModel):
    template: str
    # 源文档（2026-11-07）：指定"这张表的数据从哪几份文档里找"。
    # 空 = 老行为（整个文档库都参与检索）。给了名字就只在这几份里找数，
    # 避免同名列（金额 / 日期 / 负责人）从别的文档里取到"对但不该用"的值。
    docs: list[str] = []
    skip_read: bool = False               # 跳过「重读源文档」（提速；文档没变时可用）
    overwrite: bool = False               # 待填区已有内容时是否覆盖（默认保护性拒写）
    # 用户要求原话（2026-11-03）。例如「将内容中日期一列从 2020/7/1 到 2020/8/31 的数据填入模板」。
    # 带行条件时会走结构化直通 + 按条件筛行；不传就等于没有条件（老行为）。
    require: str = ""


# 用户自己上传的模板单独放一个目录：不跟系统自带的两批混在一起，
# 好处是「删除」只敢动这一批，系统模板删不掉。
FILL_UPLOAD_DIR = BASE / "测试数据" / "模板上传"
FILL_TPL_EXTS = {".xlsx", ".xlsm", ".docx"}
FILL_UPLOAD_MAX = 12 * 1024 * 1024          # 12MB，模板一般几百 KB，给足余量


def _fill_template_dirs():
    """可填模板目录：系统自带两批（Excel / Word 分目录）+ 用户上传（混放）。

    第三项的 kind 给 None，表示按文件扩展名逐个判断（上传目录里 xlsx 和
    docx 是混着的，而 template_reader 的 read_folder 两种都能读）。
    """
    return [(BASE / "测试数据" / "模板表格", "xlsx"),
            (BASE / "测试数据" / "模板表格-word", "docx"),
            (FILL_UPLOAD_DIR, None)]


def _fill_templates():
    out = []
    for d, kind in _fill_template_dirs():
        if not d.exists():
            continue
        for f in sorted(d.iterdir()):
            if not f.is_file() or f.suffix.lower() not in FILL_TPL_EXTS:
                continue
            if f.name.startswith(".uploading-"):      # 正在上传的临时文件，别当成模板
                continue
            k = kind or ("docx" if f.suffix.lower() == ".docx" else "xlsx")
            st = f.stat()
            out.append({"name": f.name, "kind": k, "dir": d.name,
                        "upload": (d == FILL_UPLOAD_DIR),
                        "size": st.st_size,
                        "mtime": datetime.fromtimestamp(st.st_mtime).strftime("%Y-%m-%d %H:%M")})
    return out


def _find_template(name: str):
    """按文件名在所有可填模板目录里定位实际文件（先系统、后上传）。"""
    target = Path(name).name
    for d, _kind in _fill_template_dirs():
        p = d / target
        if p.is_file():
            return p
    return None


def _clean_tpl_name(raw: str) -> str:
    """把上传的文件名洗成一个安全的模板名：去路径、去特殊字符、限长。"""
    name = Path(str(raw or "")).name.strip()
    name = re.sub(r'[\\/:*?"<>|\r\n\t]', "_", name)
    name = name.strip(" .")
    stem, ext = os.path.splitext(name)
    stem = stem.strip(" .")[:80] or "未命名模板"
    return stem + ext.lower()


def _validate_tpl_file(path: Path) -> None:
    """确认这真是个能打开的 Office 文件，不合格抛 ValueError（直接给人话）。"""
    with path.open("rb") as fh:
        head = fh.read(4)
    if len(head) < 4 or not head.startswith(b"PK"):
        raise ValueError("这不是有效的 Office 文件（.xlsx / .docx 都是压缩包格式，"
                         "文件头对不上）。如果你传的是 .xls 或 .doc 老格式，"
                         "请用 Excel / Word 另存为 .xlsx / .docx 再上传。")
    try:
        if path.suffix.lower() == ".docx":
            from docx import Document
            Document(str(path))
        else:
            from openpyxl import load_workbook
            wb = load_workbook(str(path), read_only=True)
            wb.close()
    except Exception as e:
        raise ValueError("这个文件打不开，可能已损坏或加了密码（%s）" % str(e)[:120])


@app.get("/api/fill/templates")
def api_fill_templates():
    """可填模板清单（系统自带 + 用户上传，upload 字段区分）。"""
    return {"templates": _fill_templates()}


@app.get("/api/fill/history")
def api_fill_history(template: str = ""):
    """历史填写参照（2026-11-02）：这张表（或同类表）上次是怎么填的。

    两个用法：
      不带 template → 返回最近 20 次填写的摘要（"填过哪些表"）
      带 template   → 拿这张模板的列名去查历史，返回可复用的列映射
                      （哪一列上次是从源文档的哪个列头抽的）
    """
    mdir = BASE / "matcher"
    if str(mdir) not in sys.path:
        sys.path.insert(0, str(mdir))
    import history

    cols = []
    _err = ""
    if template:
        try:
            data = json.loads((BASE / "template_reader" / "模板表结构.json")
                              .read_text(encoding="utf-8"))
            for t in data:
                if t.get("file") == template and t.get("sheets"):
                    cols = [c["title"] for c in t["sheets"][0].get("columns", [])
                            if c.get("title") and c["title"] != "序号"]
                    break
        except Exception as e:
            # 不静默吞：查不到历史时，"结构文件没读到"和"这张表没填过"
            # 是两件完全不同的事，得能看出来
            cols = []
            _err = "%s: %s" % (e.__class__.__name__, e)

    if template and cols:
        hit = history.recall(template, cols)
        if not hit:
            return {"template": template, "columns": cols, "hit": None,
                    "lines": ["这张模板还没有历史填写记录"],
                    "cols": {}}
        return {"template": template, "columns": cols, "hit": hit["hit"],
                "from": hit.get("from"), "lines": hit["lines"],
                "cols": hit["cols"]}

    d = history._load()
    runs = [{"ts": r.get("ts"), "template": r.get("template"),
             "n_cols": r.get("n_cols"), "n_rows": r.get("n_rows"),
             "times": r.get("times", 1),
             "source_files": r.get("source_files"), "cols": r.get("cols")}
            for r in d["runs"][-20:]]
    # template / columns 一起回传：查不到时能一眼看出是"模板名没传对"
    # 还是"结构文件里没这张表"，不用再去猜
    out = {"runs": runs, "total": len(d["runs"]),
           "template": template, "columns": cols}
    if _err:
        out["error"] = _err
    return out


@app.post("/api/fill/templates/upload")
def api_fill_template_upload(file: UploadFile = File(...)):
    """上传一张自己的模板（.xlsx / .xlsm / .docx）→ 立刻出现在模板下拉框里。

    三件保护性的规矩：
      1. 只收 Office 新格式，且必须能真的被 openpyxl / python-docx 打开；
      2. 跟任何已有模板重名就拒绝（改个名再传），避免"到底填的是哪张"；
      3. 先写临时文件、验过再落正式名，不会留下半截坏文件。
    """
    raw = file.filename or ""
    if not raw.strip():
        raise HTTPException(status_code=400, detail="没收到文件名")

    name = _clean_tpl_name(raw)
    ext = Path(name).suffix.lower()
    if ext not in FILL_TPL_EXTS:
        raise HTTPException(
            status_code=400,
            detail="只支持 .xlsx / .xlsm / .docx —— 你传的是「%s」。"
                   ".xls / .doc 是旧格式，请另存为 .xlsx / .docx 再上传。"
                   % (ext or "无扩展名"))

    data = file.file.read()
    if not data:
        raise HTTPException(status_code=400, detail="这个文件是空的（0 字节）")
    if len(data) > FILL_UPLOAD_MAX:
        raise HTTPException(
            status_code=400,
            detail="文件 %.1fMB，超过 %dMB 上限。模板一般只有几百 KB，"
                   "是不是传错了文件？" % (len(data) / 1048576.0,
                                          FILL_UPLOAD_MAX // 1048576))

    if _find_template(name) is not None:
        raise HTTPException(
            status_code=409,
            detail="已经有同名模板「%s」了（系统自带或之前传过）。"
                   "换一个名字再上传，避免分不清填的是哪一张。" % name)

    FILL_UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
    final = FILL_UPLOAD_DIR / name
    tmp = FILL_UPLOAD_DIR / (".uploading-" + name)
    try:
        tmp.write_bytes(data)
        _validate_tpl_file(tmp)
        if final.exists():
            raise HTTPException(status_code=409, detail="目标文件已存在，换个名字再传")
        tmp.replace(final)
    except HTTPException:
        try:
            tmp.unlink()
        except OSError:
            pass
        raise
    except ValueError as e:
        try:
            tmp.unlink()
        except OSError:
            pass
        raise HTTPException(status_code=400, detail=str(e))
    except OSError as e:
        try:
            tmp.unlink()
        except OSError:
            pass
        raise HTTPException(status_code=500, detail="保存失败：%s" % e)

    return {"ok": True, "name": name, "kind": "docx" if ext == ".docx" else "xlsx",
            "size": len(data), "templates": _fill_templates()}


@app.delete("/api/fill/templates")
def api_fill_template_delete(name: str):
    """删掉一张模板（2026-11-07 起：系统自带的也能删）。

    不真删 —— 统一挪到 测试数据/模板归档/ 下带时间戳，删错了把文件挪回去就行。
    这样既满足"系统模板也能删"，又不会把比赛用的模板永久弄丢。
    """
    if not str(name or "").strip():
        raise HTTPException(status_code=400, detail="没给模板名")
    target = _clean_tpl_name(name)

    p = _find_template(target)
    if p is None:
        raise HTTPException(status_code=404, detail="没有叫「%s」的模板" % target)
    # 只允许删模板目录里的文件，防止被喂一个 ../.. 的路径
    try:
        ok_dir = any(p.resolve().parent == d.resolve()
                     for d, _k in _fill_template_dirs())
    except OSError:
        ok_dir = False
    if not ok_dir:
        raise HTTPException(status_code=400, detail="这个模板名不合法")

    archive = BASE / "测试数据" / "模板归档"
    archive.mkdir(parents=True, exist_ok=True)
    dest = archive / ("%s.%s%s" % (p.stem, datetime.now().strftime("%Y%m%d-%H%M%S"),
                                   p.suffix))
    try:
        import shutil
        shutil.move(str(p), str(dest))
    except OSError as e:
        # WinError 32/33、EACCES、以及「另一个程序正在使用」都等于文件被占着
        msg = str(e)
        locked = (isinstance(e, PermissionError)
                  or getattr(e, "winerror", None) in (32, 33)
                  or getattr(e, "errno", None) == 13
                  or "another process" in msg or "另一个程序" in msg
                  or "正在使用" in msg or "safe-delete" in msg)
        if locked:
            raise HTTPException(
                status_code=423,
                detail="删不掉：「%s」现在被 WPS / Excel 之类占着，先把它关掉再删。"
                       % target)
        raise HTTPException(status_code=500, detail="删除失败：%s" % msg[:200])
    left = _fill_templates()
    return {"ok": True, "deleted": target, "archived": str(dest),
            "uploaded_left": sum(1 for t in left if t["upload"]),
            "templates": left}


@app.post("/api/fill")
def api_fill(body: FillIn):
    """表格填写：模板 + 源文档（可选）→ 异步跑「读懂模板 → 找数 → 填写 → 自检」。

    提交立即返回 job_id（不阻塞）；用 GET /api/job/{id} 轮询进度与结果。
    完成后可用 GET /api/fill/download?template=模板名 下载已填写文件。

    docs（2026-11-07）：源文档名单。给了就只在这些文档里找数；不给 = 全库检索。
    """
    name = (body.template or "").strip()
    if not name:
        raise HTTPException(status_code=400, detail="请给出模板文件名")
    allowed = {t["name"] for t in _fill_templates()}
    if name not in allowed:
        raise HTTPException(status_code=404,
                            detail="模板目录里没有「%s」——可用：%s"
                                   % (name, "、".join(sorted(allowed)[:12])))

    docs = [str(d).strip() for d in (body.docs or []) if str(d).strip()]
    if docs:
        try:
            in_lib = set(scan_docs(hub.docs_dir))
        except Exception:
            in_lib = set()
        miss = [d for d in docs if d not in in_lib]
        if miss:
            raise HTTPException(
                status_code=404,
                detail="文档库里没有：%s。先在「⬆ 上传源文档」把它传进来，"
                       "或者取消勾选它再用。" % "、".join(miss[:6]))

    job = hub.submit_fill(name, skip_read=bool(body.skip_read),
                          overwrite=bool(body.overwrite),
                          require=(body.require or "").strip(),
                          docs=docs)
    return {"job_id": job.id, "status": job.status, "kind": job.kind,
            "template": name, "docs": docs}


def _find_fill_output(stem: str):
    """按模板名找「已填写」产物，返回最新的那个路径；找不到返回 None。

    找两处（2026-11-07 补）：
      ① writer/输出/ —— 手工跑 fill.py 的产物
      ② agent/jobs/<任务号>/out/ —— 网页提交的任务，产物按任务隔离，
         以前这里只认 ①，于是网页上点「下载」会 404（文件明明已经生成了）。
    """
    stem = Path(stem).stem
    if (not stem) or ("/" in stem) or ("\\" in stem):
        return None
    cands = []
    for name in ("%s-已填写.docx" % stem, "%s-已填写.xlsx" % stem):
        p = BASE / "writer" / "输出" / name
        if p.exists():
            cands.append(p)
    for p in (BASE / "agent" / "jobs").glob("*/out/%s-已填写.*" % stem):
        if p.is_file():
            cands.append(p)
    if not cands:
        return None
    return max(cands, key=lambda p: p.stat().st_mtime)


class FillOpenIn(BaseModel):
    template: str
    mode: str = "file"            # file = 打开这份文件；folder = 在资源管理器里定位它


@app.post("/api/fill/open")
def api_fill_open(body: FillOpenIn):
    """在**这台机器**上打开已填写文件 / 打开它所在的文件夹。

    只认「已填写」产物，路径由服务端自己算（不接收前端传来的任意路径），
    免得变成一个"能打开任意文件"的口子。
    """
    stem = (body.template or "").strip()
    if not stem:
        raise HTTPException(status_code=400, detail="请给出模板名")
    p = _find_fill_output(stem)
    if p is None:
        raise HTTPException(status_code=404,
                            detail="还没找到「%s」的已填写产物 —— 先跑一次表格填写" % stem)
    try:
        if body.mode == "folder":
            # /select, 会打开资源管理器并选中这个文件
            subprocess.Popen(["explorer", "/select,", str(p)])
        elif os.name == "nt":
            os.startfile(str(p))                       # noqa: S606（本机应用，路径由服务端算）
        else:
            subprocess.Popen(["xdg-open", str(p)])
    except Exception as e:
        raise HTTPException(status_code=500, detail="打开失败：%s" % e)
    return {"ok": True, "path": str(p), "mode": body.mode}


@app.get("/api/fill/download")
def api_fill_download(template: str):
    """下载已填写文件（按模板名取最新产物，xlsx / docx 自动识别）。"""
    hit = _find_fill_output(template)
    if hit is None:
        raise HTTPException(status_code=404,
                            detail="还没有「%s」的已填写产物——先提交 /api/fill 任务"
                                   % Path(template).stem)
    return FileResponse(hit, filename=hit.name)


# ---------------------------------------------------------------- 启动

def main():
    global hub
    ap = argparse.ArgumentParser(description="A23 智能文档助手 · HTTP API 服务")
    ap.add_argument("--host", default="127.0.0.1", help="监听地址（默认本机）")
    ap.add_argument("--port", type=int, default=8000, help="端口（默认 8000）")
    ap.add_argument("--docs", default=None,
                    help="文档库目录（默认：源文档 + 导入文档 双目录）")
    args = ap.parse_args()

    hub = Hub(Path(args.docs) if args.docs else DEFAULT_DOCS_DIRS)

    # 数据库先备好（连不上不拦路——指令照常执行，只是历史 / 登录功能降级）
    try:
        db.ensure_db()
        auth.ensure_tables()
        support.ensure_tables()
        room.ensure_tables()
        feedback.ensure_table()
        print("【数据库】MySQL 就绪（库 a23_agent · 表 commands / imports / users / sessions"
              " / support_sessions / support_messages / feedback / room_messages / room_online）")
    except RuntimeError as e:
        print("【数据库】%s" % e)
        print("（提示：历史记录与登录功能不可用，其余功能正常）")

    try:
        n_docs = len(scan_docs(hub.docs_dir))
    except Exception:
        n_docs = 0

    dirs = hub.docs_dir if isinstance(hub.docs_dir, (list, tuple)) else [hub.docs_dir]
    print("=" * 64)
    print("A23 智能文档助手 · API 服务已启动")
    print("网页操作台：http://127.0.0.1:%d/" % args.port)
    print("接口文档：  http://127.0.0.1:%d/docs" % args.port)
    print("文档库：    %s（%d 份）" % (" ｜ ".join(str(d) for d in dirs), n_docs))
    print("模型：      %s%s" % (llm_model.describe(),
                                "" if llm_model.has_key() else "  ← 还没填密钥，AI 功能不可用"))
    print("语音输入：  浏览器语音服务不通时，自动改用本机离线识别（%s）"
          % ("模型就绪" if asr.ready() else "模型未下载，见 download_asr_model.py"))
    _ocr_st = ocr_engine.status()
    print("扫描件 OCR：%s" % (
        "就绪（%s，PDF %s）" % (_ocr_st.get("engine"),
                              "可拆页" if _ocr_st.get("pdf") else "未装 PyMuPDF")
        if _ocr_st.get("ok")
        else "不可用 ← pip install rapidocr_onnxruntime pymupdf"))
    print("=" * 64)

    # 后台把识别模型热起来：不拖慢启动，用户第一次按麦克风就是热的
    if asr.ready():
        threading.Thread(target=asr.warm_up, name="asr-warmup", daemon=True).start()
    # OCR 引擎同理：第一次点「文字识别」就不用等模型加载
    if _ocr_st.get("ok"):
        threading.Thread(target=ocr_engine.warmup, name="ocr-warmup", daemon=True).start()

    uvicorn.run(app, host=args.host, port=args.port, log_level="info")


if __name__ == "__main__":
    main()
