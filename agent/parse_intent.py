# -*- coding: utf-8 -*-
"""
parse_intent.py —— 模块（1）· 第 1 步：指令解析器（自然语言 → 意图 JSON）

任务：用户说一句话（比如"把合同里的甲方改成采购方"），
     把它翻译成一条结构化的 JSON 操作描述，交给后面的"定位 + 执行"去干活。

一句话的旅程（模块（1）五段链）：
  ① 接指令 → ② 解析（本文件）→ ③ 定位 → ④ 执行 → ⑤ 输出新文件
  本文件只负责 ②：听懂，不动手。

四类操作（对照赛题原文）：
  编辑     —— 改文字内容（替换/增加/删除）
  排版     —— 改段落/页面布局（对齐/行距/缩进）
  格式调整 —— 改文字外观（加粗/字体/字号/颜色）
  内容提取 —— 从文档里读取、整理信息（不改文档）

跑法（VSCode 打开项目根目录，终端里）：
  单条：python agent/parse_intent.py "把设备采购合同里的甲方改成采购方"
  批量：python agent/parse_intent.py --batch agent/测试指令集.json --out agent/测试记录-第1轮.json

输出：
  屏幕：解析结果 JSON（单条）/ 逐条对比 + 统计（批量）
  文件：--out 指定时，把每条的完整解析结果存成 JSON
"""
import argparse
import json
import re
import sys
import time
import threading
from pathlib import Path

import requests

# 输出兜底：遇到终端表示不了的字符就替换，不让程序崩
try:
    sys.stdout.reconfigure(errors="replace")
except Exception:
    pass

# ================== 配置区（只改这里） ==================
# 2026-10-25：模型换成「阿里云百炼 MaaS」的 OpenAI 兼容接口。
# 密钥 / 地址 / 模型名统一在 agent/llm_config.py 里配（只填那一处），这里只引用。
sys.path.insert(0, str(Path(__file__).resolve().parent))
from llm_config import API_KEY, API_URL, MODEL, chat as _llm_chat  # noqa: E402
# 速度优化（2026-10-12）：解析任务只做"听懂人话→填 JSON"，不需要深度推理。
# 关思考的开关也挪到了 llm_config.py 的 THINKING_OFF_STYLE（默认不发，最兼容）。
# ========================================================

BASE = Path(__file__).resolve().parents[1]                  # = 研一项目/
DEFAULT_DOCS_DIR = BASE / "测试数据" / "源文档"               # 默认文档库（可 --docs 换）

SYSTEM_PROMPT = ("你是一个文档操作指令解析器。用户给出一句自然语言指令，"
                 "你把它翻译成一条结构化的 JSON 操作描述。你只负责听懂，不负责执行。"
                 "严格只输出 JSON，不要任何解释。")


def as_dir_list(docs_dir) -> list:
    """目录参数归一化：单个目录 / 目录列表 → [Path, ...]（只保留存在的）。

    模块（2）起文档库 = "测试数据/源文档" + "导入文档" 双目录，
    所以库参数统一支持"一个目录，或一串目录"。
    """
    if not docs_dir:
        return []
    items = docs_dir if isinstance(docs_dir, (list, tuple, set)) else [docs_dir]
    out = []
    for d in items:
        if not d:
            continue
        p = Path(d)
        if p.exists():
            out.append(p)
    return out


def scan_docs(docs_dir):
    """扫描文档库（一个目录或多个目录），返回文件名清单（去重、保序）。"""
    names = []
    seen = set()
    for d in as_dir_list(docs_dir):
        for p in sorted(d.rglob("*")):
            if p.is_file() and p.suffix.lower() in (".docx", ".doc", ".xlsx", ".xls", ".md", ".txt"):
                if p.name not in seen:
                    seen.add(p.name)
                    names.append(p.name)
    return names


def build_prompt(text: str, doc_names: list, context: dict = None) -> str:
    """把指令 + 文档库清单 + 输出规范拼成给模型的一段话。

    context（可选）＝ 会话上下文（多轮连续指令的指代消解），形如
      {"current_doc": "项目推进会会议纪要.docx"}
    带上它，用户说"把它的标题居中"这类指代句时，模型才知道"它"是谁。
    """
    if doc_names:
        docs_block = "【当前文档库】（用户提到的文档，要尽量匹配到这个清单里的文件名，填进 doc 字段；匹配不上时 doc 填 null）\n" \
                     + "\n".join("- " + n for n in doc_names)
    else:
        docs_block = "【当前文档库】未知（doc 字段填 null，把用户话里的称呼原样填进 doc_hint）"

    ctx_block = ""
    if context and context.get("current_doc"):
        cd = context["current_doc"]
        ctx_block = ("\n【会话上下文】用户正在连续操作文档：上一句操作的文档是《%s》。\n"
                     "如果本句用「它 / 这个文件 / 刚才那份」等指代词称呼文档、且没有说别的文档名，"
                     "请把 doc 字段指向《%s》；如果本句明确提到了别的文档，以本句为准。\n"
                     % (cd, cd))
        lh = str(context.get("last_hint") or "").strip()
        if lh:
            ctx_block += ("上一句操作的目标位置是「%s」——如果本句说「把它的加粗去掉」"
                          "「把它改成红色」这类没提具体目标的话，locate.hint 填「%s」。\n"
                          % (lh, lh))

    return f"""你是文档系统的指令解析器。请把下面这句用户指令解析成 JSON。

【四类操作，只选一类】
- 编辑：改变文字内容（替换、添加、删除文字），或把一张图片插进文档
- 排版：改变段落/页面的布局（对齐、行距、缩进等）
- 格式调整：改变文字外观（字体、字号、加粗、颜色等）
- 内容提取：从文档里读取/整理信息（不改动文档）
若指令信息不足以判断要做什么，action 用「不明确」，clarify 写一句反问用户的话。两种常见情况：
① 说不清要做什么（如"处理一下那个文件"）→ 反问要做什么；若句中的模糊称呼（如"那几个表格""那份纪要"）能唯一对应文档库里的某一份，doc 字段要填上这份文件名，不要留 null（若有多个候选则留 null，反问是哪一份）。
② 操作内容没说全——让"加上/补上"某个信息、但没给出要加的具体内容（如"把联系电话补上"但没说号码是多少）→ 反问要补的具体内容；绝不能拿"电话""地址"这类信息名称本身当内容去插入。

{docs_block}{ctx_block}

【输出格式】严格按这个 JSON 结构（不要多字段、不要少字段）：
{{
  "action": "编辑 | 排版 | 格式调整 | 内容提取 | 多操作 | 不明确",
  "op": "replace | insert | delete | insert_image | set_align | set_spacing | set_indent | set_bold | set_italic | set_underline | set_strike | set_highlight | set_font | set_size | set_color | extract | profile | entities | multi | none",
  "doc": "文档文件名（来自文档库清单）；匹配不上写 null",
  "doc_hint": "用户话里对文档的称呼（原话照抄）",
  "locate": {{"by": "whole | keyword | table", "hint": "定位线索（关键词），by=whole 时可留空"}},
  "params": {{}},
  "clarify": "仅 action=不明确 时填写反问；否则空字符串"
}}

【多操作（一句话里有多个操作）】
一句话里用"并且 / 同时 / 再 / 还要 / 顺便"等词连接了多个操作时（如"标题右对齐，并且加粗"）：action 用「多操作」、op 用「multi」，并加一个 tasks 字段——按顺序列出每个操作，每个操作是一张完整的"op + locate + params"卡片（文档名用顶层的 doc，不用重复）：
{{
  "action": "多操作", "op": "multi",
  "doc": "设备采购合同.docx", "doc_hint": "设备采购合同",
  "locate": {{"by": "keyword", "hint": "标题"}},
  "params": {{}},
  "tasks": [
    {{"op": "set_align", "locate": {{"by": "keyword", "hint": "标题"}}, "params": {{"align": "right"}}}},
    {{"op": "set_bold", "locate": {{"by": "keyword", "hint": "标题"}}, "params": {{}}}}
  ],
  "clarify": ""
}}
- 多个操作必须属于同一份文档；若用户要求操作两份不同的文档，只解析第一份的操作。
- 单个操作时不要输出 tasks 字段（用上面的单操作结构）。
- 提取（extract）不和其它操作混在同一句里。

【各 op 的 params 写法】
- replace: {{"find": "要改的词", "replace_with": "改成什么"}}
- insert:  {{"text": "要加的内容"}}
- insert_image: {{"position": "start | end | auto", "width_ratio": 0.7}}（插入图片；position 只有明确说"开头"才填 start、明确说"最后/文末"才填 end，其余填 auto；width_ratio 是图片占正文宽度的比例，没说就 0.7）
- delete:  {{"find": "要删的词"}}
- set_bold: {{}}
- set_italic: {{}}（斜体）
- set_underline: {{}}（下划线）
- set_strike: {{}}（删除线）
- set_highlight: {{"color": "颜色词（如 黄色/绿色/青色），没说就留空"}}（高亮 / 突出显示）
- set_font: {{"font": "字体名（黑体 / 宋体 / 楷体 / 微软雅黑等）"}}
- set_size: {{"size": 数字}}
- set_color: {{"color": "颜色"}}
- set_align: {{"align": "center | left | right"}}
- set_spacing: {{"spacing": "行距描述（如 1.5 倍）"}}
- set_indent: {{"indent": "缩进描述（如 首行缩进2字符）"}}
- extract: {{"fields": ["要提取的字段1", "字段2"]}}——字段名保留关键限定语、尽量对齐文档里的实际叫法（如"每种物料的名称"→"物料名称"，别只写"名称"）；按用户说的顺序列出
- profile: {{}}（关键信息 / 文档画像——用户说"提取这份文档的关键信息""总结一下这份文档""这份文档讲了什么""生成文档画像"时用 profile；用户点名要具体字段（如"姓名和工号"）时用 extract）
- entities: {{}}（实体数据识别——用户说"识别/提取这份文档里的实体数据""把实体都列出来""把里面的人名/日期/金额这些抽出来"时用 entities；点名具体字段用 extract，要梗概用 profile）
- 取消类：用户说"取消 / 去掉 / 去除 某种格式"时（如"把标题的加粗去掉"），op 仍用对应的 set_bold / set_italic / set_underline / set_strike / set_highlight，并加上 {{"off": true}}。
- 注意：斜体、下划线、删除线、高亮是**独立的 op**（set_italic / set_underline / set_strike / set_highlight），绝不要塞进 set_font 的 font 参数里。

【图片插入（insert_image）】
- 只有用户要把「图片 / 图 / 照片 / 截图 / 配图」放进文档时，op 才用 insert_image（action 仍填「编辑」）。
- insert 是"插入一段文字"，insert_image 是"插入一张图片"，绝不能混用；用户没说图片时不要用 insert_image。
- 插入位置写在 locate.hint 里，照抄用户的位置说法：如"标题下面""第二部分""图 3 后面""文末"；完全没说位置就填 ""。
- 用户若说"插到开头/最前面"，params.position 填 start、locate.hint 可留空；说"插到最后/文末"则 position 填 end。
- 用户没说要插到哪份文档、但只有一份正在操作的文档时，doc 填 null，由系统沿用上一轮的文档。

【locate.by 的用法】
- whole：整个文档范围（如"合同里所有金额"）
- keyword：靠一个定位词找到位置（如"含'待办'的段落"→hint=待办；"标题"→hint=标题；"落款处"→hint=落款）
- table：明确指表格区域（如"员工信息表的表头"→hint=表头）

【示例】
用户：把《研发项目立项管理办法》里的"项目经理"改成"项目负责人"
输出：{{"action": "编辑", "op": "replace", "doc": "研发项目立项管理办法.docx", "doc_hint": "研发项目立项管理办法", "locate": {{"by": "whole", "hint": ""}}, "params": {{"find": "项目经理", "replace_with": "项目负责人"}}, "clarify": ""}}

用户：给销售台账的表头加粗
输出：{{"action": "格式调整", "op": "set_bold", "doc": "销售台账.xlsx", "doc_hint": "销售台账", "locate": {{"by": "table", "hint": "表头"}}, "params": {{}}, "clarify": ""}}

用户：把工作总结的标题居中
输出：{{"action": "排版", "op": "set_align", "doc": "年度数字化转型工作总结.docx", "doc_hint": "工作总结", "locate": {{"by": "keyword", "hint": "标题"}}, "params": {{"align": "center"}}, "clarify": ""}}

用户：把设备验收记录里的设备名称和验收结论抽出来
输出：{{"action": "内容提取", "op": "extract", "doc": "设备验收记录.txt", "doc_hint": "设备验收记录", "locate": {{"by": "whole", "hint": ""}}, "params": {{"fields": ["设备名称", "验收结论"]}}, "clarify": ""}}

用户：把设备采购合同的关键信息整理出来
输出：{{"action": "内容提取", "op": "profile", "doc": "设备采购合同.docx", "doc_hint": "设备采购合同", "locate": {{"by": "whole", "hint": ""}}, "params": {{}}, "clarify": ""}}

用户：这份文档讲了什么
输出：{{"action": "内容提取", "op": "profile", "doc": null, "doc_hint": "", "locate": {{"by": "whole", "hint": ""}}, "params": {{}}, "clarify": ""}}

用户：识别设备采购合同里的实体数据
输出：{{"action": "内容提取", "op": "entities", "doc": "设备采购合同.docx", "doc_hint": "设备采购合同", "locate": {{"by": "whole", "hint": ""}}, "params": {{}}, "clarify": ""}}

用户：把这份文档里的实体都列出来
输出：{{"action": "内容提取", "op": "entities", "doc": null, "doc_hint": "", "locate": {{"by": "whole", "hint": ""}}, "params": {{}}, "clarify": ""}}

用户：那个文件帮我处理一下
输出：{{"action": "不明确", "op": "none", "doc": null, "doc_hint": "那个文件", "locate": {{"by": "whole", "hint": ""}}, "params": {{}}, "clarify": "请问是哪份文档？想对它做什么操作（改文字 / 调格式 / 抽信息）？"}}

用户：处理一下那份纪要
输出：{{"action": "不明确", "op": "none", "doc": "项目推进会会议纪要.docx", "doc_hint": "那份纪要", "locate": {{"by": "whole", "hint": ""}}, "params": {{}}, "clarify": "请问要对《项目推进会会议纪要.docx》做什么操作（改文字 / 调格式 / 抽信息）？"}}

用户：帮我把合同里乙方的地址加上
输出：{{"action": "不明确", "op": "none", "doc": "设备采购合同.docx", "doc_hint": "合同", "locate": {{"by": "keyword", "hint": "乙方"}}, "params": {{}}, "clarify": "请问要加上的地址具体是什么？请给出内容。"}}

用户：把年度数字化转型工作总结的标题居中，并且加粗
输出：{{"action": "多操作", "op": "multi", "doc": "年度数字化转型工作总结.docx", "doc_hint": "工作总结", "locate": {{"by": "keyword", "hint": "标题"}}, "params": {{}}, "tasks": [{{"op": "set_align", "locate": {{"by": "keyword", "hint": "标题"}}, "params": {{"align": "center"}}}}, {{"op": "set_bold", "locate": {{"by": "keyword", "hint": "标题"}}, "params": {{}}}}], "clarify": ""}}

用户：把项目周报的标题设成斜体
输出：{{"action": "格式调整", "op": "set_italic", "doc": "项目周报-第12周.md", "doc_hint": "项目周报", "locate": {{"by": "keyword", "hint": "标题"}}, "params": {{}}, "clarify": ""}}

用户：给员工信息表里的"王振国"加上下划线
输出：{{"action": "格式调整", "op": "set_underline", "doc": "员工信息表.xlsx", "doc_hint": "员工信息表", "locate": {{"by": "keyword", "hint": "王振国"}}, "params": {{}}, "clarify": ""}}

用户：把设备采购合同标题的加粗去掉
输出：{{"action": "格式调整", "op": "set_bold", "doc": "设备采购合同.docx", "doc_hint": "设备采购合同", "locate": {{"by": "keyword", "hint": "标题"}}, "params": {{"off": true}}, "clarify": ""}}

用户：把这张图片插到设备采购合同的标题下面
输出：{{"action": "编辑", "op": "insert_image", "doc": "设备采购合同.docx", "doc_hint": "设备采购合同", "locate": {{"by": "keyword", "hint": "标题下面"}}, "params": {{"position": "auto", "width_ratio": 0.7}}, "clarify": ""}}

用户：在项目周报的最后加一张截图
输出：{{"action": "编辑", "op": "insert_image", "doc": "项目周报-第12周.md", "doc_hint": "项目周报", "locate": {{"by": "whole", "hint": "文末"}}, "params": {{"position": "end", "width_ratio": 0.7}}, "clarify": ""}}

【现在解析这句】
{text}
"""


# 连接复用（2026-10-12 速度优化）：每次 requests.post 都要重新做一遍
# DNS+TLS 握手（实测多花 0.2~0.4s）。按线程各存一个 Session 复用连接——
# API 服务是多线程跑任务的，Session 不适合跨线程共用，所以用线程本地存储。
_thread_local = threading.local()


def _http():
    """取当前线程的 requests.Session（没有就建一个）。"""
    s = getattr(_thread_local, "session", None)
    if s is None:
        s = requests.Session()
        _thread_local.session = s
    return s


def ask_model(user_prompt: str, system: str = None):
    """调用大模型一次，返回 (回答文本, 用量, 耗时秒数)。

    system 不传时用解析器的 SYSTEM_PROMPT；其他模块（如模块2 文档画像）
    可以传入自己的 system prompt 复用这套调用逻辑（连接复用 + JSON + 参数降级）。
    """
    # 要 JSON 就明确要 JSON（配合 System Prompt 的双保险，减少"话痨"）；
    # 地址 / 模型 / 关思考 / 参数降级都交给 llm_config.chat() 统一处理。
    content, usage, cost, err = _llm_chat(
        [{"role": "system", "content": system or SYSTEM_PROMPT},
         {"role": "user", "content": user_prompt}],
        json_mode=True, timeout=90, session=_http())
    if err:
        print("请求失败！", err)
        return None, None, cost
    return content, usage, cost


def parse_json_reply(text: str):
    """宽容地解析模型的回复：允许 ```json 围栏、前后多余文字。"""
    t = (text or "").strip()
    if t.startswith("```"):
        t = re.sub(r"^```[A-Za-z]*\s*", "", t)
        t = re.sub(r"```\s*$", "", t)
    i, j = t.find("{"), t.rfind("}")
    if i == -1 or j == -1:
        return None
    try:
        return json.loads(t[i:j + 1])
    except Exception:
        return None


def parse_one(text: str, doc_names: list, retry: int = 1, context: dict = None):
    """解析一条指令。返回 (结果dict, 耗时秒, 是否成功)。

    失败（网络/解析）自动重试 retry 次。
    context 见 build_prompt（多轮指代消解用，单条模式传 None 即可）。
    """
    prompt = build_prompt(text, doc_names, context)
    total = 0.0
    for attempt in range(retry + 1):
        content, usage, cost = ask_model(prompt)
        total += cost
        if content:
            result = parse_json_reply(content)
            if result:
                return result, total, True
    return None, total, False


def cmd_single(text: str, docs_dir: Path):
    """单条模式：解析一条指令，打印 JSON。"""
    doc_names = scan_docs(docs_dir)
    print("指令：%s" % text)
    print("文档库：%d 个文件" % len(doc_names))
    print("-" * 50)
    result, cost, ok = parse_one(text, doc_names)
    if not ok:
        print("解析失败（网络或格式问题），请重跑一次。耗时 %.1fs" % cost)
        return
    print(json.dumps(result, ensure_ascii=False, indent=2))
    print("-" * 50)
    print("耗时 %.2f 秒" % cost)


def cmd_batch(batch_file: Path, docs_dir: Path, out_file: Path = None):
    """批量模式：跑整个测试指令集，逐条对比 + 统计准确率。"""
    cases = json.loads(batch_file.read_text(encoding="utf-8"))
    doc_names = scan_docs(docs_dir)

    print("=" * 62)
    print("指令解析器 · 批量测试")
    print("文档库：%d 个文件 ｜ 测试指令：%d 条" % (len(doc_names), len(cases)))
    print("=" * 62)
    print("")

    records = []
    n_action = n_op = n_doc = 0
    total_cost = 0.0

    for c in cases:
        text = c["text"]
        result, cost, ok = parse_one(text, doc_names)
        total_cost += cost
        rec = {"id": c["id"], "text": text, "ok": ok, "cost": round(cost, 2),
               "expect": {"action": c.get("expect_action"), "op": c.get("expect_op"),
                          "doc": c.get("expect_doc")},
               "got": result}
        records.append(rec)

        if not ok:
            print("[%2s] ✗ 解析失败（%.1fs）" % (c["id"], cost))
            print("      指令：%s" % text)
            print("")
            continue

        a, o, d = result.get("action"), result.get("op"), result.get("doc")
        ea, eo, ed = c.get("expect_action"), c.get("expect_op"), c.get("expect_doc")
        ma = (a == ea)
        mo = (o == eo)
        md = (d == ed)
        n_action += ma
        n_op += mo
        n_doc += md

        mark = ("✓" if ma else "✗") + ("✓" if mo else "✗") + ("✓" if md else "✗")
        print("[%2s] %s  %s/%s ｜ doc=%s" % (c["id"], mark, a, o, d))
        print("      指令：%s" % text)
        if not ma or not mo or not md:
            print("      期望：%s/%s ｜ doc=%s" % (ea, eo, ed))
        print("      全文：%s" % json.dumps(result, ensure_ascii=False))
        print("")

    n = len(cases)
    print("=" * 62)
    print("统计（共 %d 条，总耗时 %.1fs，平均 %.2fs/条）" % (n, total_cost, total_cost / max(n, 1)))
    print("  action 准确：%d/%d = %.1f%%" % (n_action, n, 100.0 * n_action / n))
    print("  op     准确：%d/%d = %.1f%%" % (n_op, n, 100.0 * n_op / n))
    print("  doc    准确：%d/%d = %.1f%%" % (n_doc, n, 100.0 * n_doc / n))
    print("=" * 62)

    if out_file:
        payload = {
            "文档库": [str(docs_dir)],
            "统计": {"总数": n, "action_对": n_action, "op_对": n_op, "doc_对": n_doc,
                     "总耗时秒": round(total_cost, 2)},
            "明细": records,
        }
        out_file.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
        print("完整结果已存：%s" % out_file)


def main():
    ap = argparse.ArgumentParser(description="模块（1）· 指令解析器（自然语言 → 意图 JSON）")
    ap.add_argument("text", nargs="?", help="一条指令（直接写）")
    ap.add_argument("--batch", help="批量测试：测试指令集 JSON 文件路径")
    ap.add_argument("--out", help="批量模式：把完整结果存到这个文件")
    ap.add_argument("--docs", default=str(DEFAULT_DOCS_DIR), help="文档库目录（默认 测试数据/源文档）")
    args = ap.parse_args()

    docs_dir = Path(args.docs)
    if args.batch:
        cmd_batch(Path(args.batch), docs_dir, Path(args.out) if args.out else None)
    elif args.text:
        cmd_single(args.text, docs_dir)
    else:
        ap.print_help()
        print("\n示例：python agent/parse_intent.py \"把设备采购合同里的甲方改成采购方\"")


if __name__ == "__main__":
    main()
