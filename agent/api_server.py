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
from chat import Session                              # noqa: E402
from edit import scan_docs, DEFAULT_DOCS_DIRS, IMPORT_DIR     # noqa: E402
from edit import find_doc_file, run_edit                      # noqa: E402  识别分类 API 用

WEB_DIR = Path(__file__).resolve().parent / "web"

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


class Hub:
    """全局中枢：会话表 + 任务表 + 线程池。"""

    # 识别线别名（英文键 + 中文说法都能提交）
    TASK_ALIASES = {
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
        self.fill_lock = threading.Lock()  # 表格填写全局串行（抽取中间产物共享文件名）

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
        return self._new_job(Job(jid, sid, text, image=image))

    def submit_recognize(self, docs: list, tasks: list, fields: list) -> Job:
        """提交「智能识别分类」任务（异步）：文档 × 三线提取 → 自动入库。"""
        jid = "r%04d" % next(self._job_seq)
        label = "识别分类：%s" % "、".join(docs[:6])
        job = Job(jid, "recognize", label, kind="recognize")
        job.runner = lambda j: self._run_recognize(j, docs, tasks, fields)
        return self._new_job(job)

    def submit_fill(self, template: str, skip_read: bool = False,
                    overwrite: bool = False) -> Job:
        """提交「表格填写」任务（异步）：跑 pipeline 全流程。"""
        jid = "f%04d" % next(self._job_seq)
        job = Job(jid, "fill", "表格填写：%s" % template, kind="fill")
        job.runner = lambda j: self._run_fill(j, template, skip_read, overwrite)
        return self._new_job(job)

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

    def _run_recognize(self, job, docs, tasks, fields):
        LINE_CN = {"profile": "画像", "entities": "实体", "fields": "字段"}

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
        lines_out.append("")
        lines_out.append("提取结果已自动入库（a23_agent）——可在「数据仓库」查询")

        job.ok = (n_fail == 0)
        job.report = "\n".join(lines_out)
        job.meta = jsonable({"docs": results, "store": snap,
                             "elapsed": round(time.time() - t0, 2)})
        job.elapsed = round(time.time() - t0, 2)
        job.status = "done"

    # ---- 表格填写执行体（模块3 · 跑 pipeline 全流程） ----
    def _run_fill(self, job, template: str, skip_read: bool, overwrite: bool):
        with self.fill_lock:                         # 全局串行：互斥共享中间产物
            t0 = time.time()
            # 按文件名定位实际所在目录（系统自带两批 + 用户上传那批都认），
            # 该目录再通过 A23_TEMPLATE_DIR 传给 pipeline 当「模板目录」。
            tpl_path = _find_template(template)
            if tpl_path is None:
                raise ValueError("模板目录里找不到：%s" % template)
            tpl_dir = tpl_path.parent
            tpl_dir_name = tpl_dir.name

            cmd = [sys.executable, str(BASE / "pipeline.py"), "--template", template]
            if skip_read:
                cmd.append("--skip-read")
            if overwrite:
                cmd.append("--overwrite")
            env = os.environ.copy()
            env["PYTHONIOENCODING"] = "utf-8"
            env["A23_TEMPLATE_DIR"] = str(tpl_dir)

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
            if "结论：可以交卷" in out:
                bits.append("结论：可以交卷")

            m = re.search(r"交卷文件 → (.+)", out)
            out_file = m.group(1).strip() if m else None

            ok = (proc.returncode == 0) and (stopped is None) and (out_file is not None)
            rep = ["【表格填写】模板：%s（%s）" % (template, tpl_dir_name)]
            if stopped:
                rep.append("流程在中途停下：%s" % stopped.group(1).strip())
                rep.append("（提示：待填区已有内容时默认保护性不写；"
                           "确认要覆盖时，重提任务并勾选「允许覆盖旧值」）")
            if bits:
                rep.append("  ｜ ".join(bits))
            if out_file:
                rep.append("交卷文件 → %s" % out_file)
            if not out_file and err.strip():
                rep.append("stderr：%s" % err.strip()[-400:])

            job.ok = bool(ok)
            job.report = "\n".join(rep)
            job.meta = jsonable({"template": template, "dir": tpl_dir_name,
                                 "elapsed": elapsed, "stages": stages,
                                 "output_file": out_file,
                                 "log_tail": "\n".join(out.splitlines()[-24:])})
            job.elapsed = elapsed
            job.status = "done"

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


@app.get("/api/docs")
def api_docs():
    """文档库清单（标准库 + 导入库 合并）。"""
    return {"docs": scan_docs(hub.docs_dir)}


# ---------------------------------------------------------------- 文档导入（模块2 · 第1步）

SUPPORTED_EXTS = {".docx", ".md", ".markdown", ".xlsx", ".txt"}


@app.post("/api/upload")
def api_upload(file: UploadFile = File(...)):
    """上传一份非结构化文档 → 自动识别格式并解析 → 并入检索池 → 记录入库。

    上传后这份文档【立刻】可被指令操作（编辑 / 排版 / 格式 / 提取），无需重启或重扫。
    重名自动加后缀，绝不覆盖标准文档或此前上传的文件。
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
            detail="不支持的格式「%s」——当前支持：Word(.docx) / Excel(.xlsx) / "
                   "Markdown(.md) / 纯文本(.txt)" % (ext or "无扩展名"))

    # 1) 落盘（规划不冲突的存储名）
    stored = doc_pool.plan_stored_name(raw_name)
    data = file.file.read()
    try:
        stored.write_bytes(data)
    except OSError as e:
        raise HTTPException(status_code=500, detail="文件保存失败：%s" % e)

    # 2) 自动识别 + 解析 + 并入检索池
    res = doc_pool.add(stored)

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
    tasks: list[str] = ["profile", "entities"]   # 默认两条全自动线；fields 需另给字段名再开
    fields: list[str] = []                # tasks 含 fields 时必填：要提取的字段名（用户指定内容）


@app.post("/api/recognize")
def api_recognize(body: RecognizeIn):
    """智能识别分类：提交文档 → 异步跑 画像 / 实体 / 字段 提取 → 自动入库。

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
    if not tasks:
        raise HTTPException(status_code=400,
                            detail="tasks 至少包含其一：profile / entities / fields（画像 / 实体 / 字段）")

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

    job = hub.submit_recognize(resolved, tasks, fields)
    return {"job_id": job.id, "status": job.status, "kind": job.kind,
            "docs": resolved, "tasks": tasks, "fields": fields}


# ---------------------------------------------------------------- 表格填写（异步 API · 模块3）

class FillIn(BaseModel):
    template: str
    skip_read: bool = False               # 跳过「重读源文档」（提速；文档没变时可用）
    overwrite: bool = False               # 待填区已有内容时是否覆盖（默认保护性拒写）


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
    """删掉一张【自己上传的】模板。系统自带的删不了（这是故意的）。"""
    target = _clean_tpl_name(name)
    if not str(name or "").strip():
        raise HTTPException(status_code=400, detail="没给模板名")
    p = FILL_UPLOAD_DIR / target
    # 再多一道保险：洗过的名字也只允许落在「模板上传」这一层目录里
    try:
        if p.resolve().parent != FILL_UPLOAD_DIR.resolve():
            raise HTTPException(status_code=400, detail="这个模板名不合法")
    except OSError:
        raise HTTPException(status_code=400, detail="这个模板名不合法")
    if not p.is_file():
        if _find_template(target) is not None:
            raise HTTPException(
                status_code=403,
                detail="「%s」是系统自带的模板，不能删除（可以自己改个名字再上传一份）"
                       % target)
        raise HTTPException(status_code=404, detail="没有叫「%s」的上传模板" % target)
    try:
        p.unlink()
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
    return {"ok": True, "deleted": target,
            "uploaded_left": sum(1 for t in left if t["upload"]),
            "templates": left}


@app.post("/api/fill")
def api_fill(body: FillIn):
    """表格填写：提交模板名 → 异步跑「读懂模板 → 找数 → 填写 → 自检」全流程。

    提交立即返回 job_id（不阻塞）；用 GET /api/job/{id} 轮询进度与结果。
    完成后可用 GET /api/fill/download?template=模板名 下载已填写文件。
    """
    name = (body.template or "").strip()
    if not name:
        raise HTTPException(status_code=400, detail="请给出模板文件名")
    allowed = {t["name"] for t in _fill_templates()}
    if name not in allowed:
        raise HTTPException(status_code=404,
                            detail="模板目录里没有「%s」——可用：%s"
                                   % (name, "、".join(sorted(allowed)[:12])))
    job = hub.submit_fill(name, skip_read=bool(body.skip_read),
                          overwrite=bool(body.overwrite))
    return {"job_id": job.id, "status": job.status, "kind": job.kind,
            "template": name}


@app.get("/api/fill/download")
def api_fill_download(template: str):
    """下载已填写文件（按模板名取最新产物，xlsx / docx 自动识别）。"""
    stem = Path(template).stem
    if (not stem) or ("/" in stem) or ("\\" in stem):
        raise HTTPException(status_code=400, detail="模板名不合法")
    out_dir = BASE / "writer" / "输出"
    candidates = [out_dir / ("%s-已填写.xlsx" % stem),
                  out_dir / ("%s-已填写.docx" % stem)]
    hit = next((p for p in candidates if p.exists()), None)
    if hit is None:
        raise HTTPException(status_code=404,
                            detail="还没有「%s」的已填写产物——先提交 /api/fill 任务" % stem)
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
    print("=" * 64)

    # 后台把识别模型热起来：不拖慢启动，用户第一次按麦克风就是热的
    if asr.ready():
        threading.Thread(target=asr.warm_up, name="asr-warmup", daemon=True).start()

    uvicorn.run(app, host=args.host, port=args.port, log_level="info")


if __name__ == "__main__":
    main()
