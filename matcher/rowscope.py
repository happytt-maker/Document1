# -*- coding: utf-8 -*-
"""
rowscope.py —— 用户要求里的「行条件」：解析 + 判行
================================================================

为什么单独一个模块（2026-11-03，COVID-19 真实数据踩出来的）：

  用户要求原话：「智能填表，将两个文件的内容中**日期一列从 2020/7/1 到 2020/8/31**
  的数据填入模板中。」

  以前系统完全不看这句话 —— 它把源表每一列都当"要填的值"，
  15088 行的源表只抽了 15 行，而且那 15 行是 2020-02 的数据（条件区间之外）。
  日志上写着「成功：6/6 列」，一条都不对。**机械验收 ≠ 值对**，这就是典型。

  这里只做三件事，不碰抽取：
    ① 把要求文本解析成 {"col": "日期", "lo": date, "hi": date}
    ② 判断某个日期值在不在区间里（兼容 datetime / 字符串 / 中文日期）
    ③ 把条件说成人话，写进日志和报告（"凭什么只填这些行"要能回答）

  抽取怎么做得看源是表格还是叙述体 —— 这一步的产物交给 direct.py / extract.py。
"""

import datetime
import os
import re

# ---------------------------------------------------------------- 日期解析
# 「2020/7/1」「2020-07-01」「2020年7月1日」「2020.7.1」都要认。
_DATE = r"(\d{4})\s*[/\-年.]\s*(\d{1,2})\s*[/\-月.]\s*(\d{1,2})"
# 区间连接词：到 / 至 / ~ / ～ / 破折号 / 以及
_RANGE = re.compile(_DATE + r"\s*日?\s*(?:到|至|~|～|—|－|–|-|以及)\s*" + _DATE)
_ONE_DATE = re.compile(_DATE + r"\s*日?")

# 列名：中文「日期/时间/年月日」，英文 date/day
_COL_WORDS = ("日期", "时间", "年月日", "date", "day", "Date")


def parse_date(v):
    """把一个值解析成 date；认不出来返回 None。

    认得的形态（都是真实数据里出现过的）：
      datetime.datetime(2020,7,1)   ← openpyxl 读出来的日期格
      datetime.date
      "2020-07-01 00:00:00"        ← 日期列被存成文本
      "2020/7/1"                   ← 用户要求里的写法
      "2020年7月1日"                ← 中文公文写法
    不认：Excel 序列号（45000 这种）。原因是没有基准日期就猜不出是 1900 还是 1904
    体系，猜错一天都会让边界行漏掉/多填 —— 宁可报"没认出来"，让人补。
    """
    if v is None:
        return None
    if isinstance(v, datetime.datetime):
        return v.date()
    if isinstance(v, datetime.date):
        return v
    s = str(v).strip()
    if not s:
        return None
    m = _ONE_DATE.search(s)
    if m:
        try:
            return datetime.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        except ValueError:
            return None
    return None


def parse_requirement(text):
    """从用户要求里解析行条件。没有条件 → None。

    返回 {"col": 列名, "lo": date, "hi": date, "raw": 原话}

    解析顺序（先精确后兜底）：
      ① 环境变量 A23_ROW_FILTER="日期=2020-07-01..2020-08-31" 优先
         —— 要求写得含糊、或界面里单独填了条件时，用它把话说死
      ② 文本里找「列名 + 从 A 到 B」：先找日期区间，再在区间前面找列名；
         找不到列名就用「日期」
    """
    env = (os.environ.get("A23_ROW_FILTER") or "").strip()
    if env:
        m = re.match(r"\s*([^=＝]+)\s*[=＝]\s*(\S+?)\s*(?:\.\.|~|～|—|至|到)\s*(\S+)\s*$", env)
        if m:
            lo, hi = parse_date(m.group(2)), parse_date(m.group(3))
            if lo and hi:
                return {"col": m.group(1).strip(), "lo": lo, "hi": hi, "raw": env}

    text = (text or "").strip()
    if not text:
        return None

    m = _RANGE.search(text)
    if not m:
        return None
    try:
        lo = datetime.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        hi = datetime.date(int(m.group(4)), int(m.group(5)), int(m.group(6)))
    except ValueError:
        return None
    if hi < lo:
        lo, hi = hi, lo

    # 列名：在区间前半段里找「日期一列」「日期列」「按日期」「date 列」
    head = text[:m.start()]
    col = None
    for w in _COL_WORDS:
        mm = re.search(re.escape(w) + r"\s*(?:这一?列|列|栏|字段)?", head, re.I)
        if mm:
            col = mm.group(0)
            col = re.sub(r"\s*(?:这一?列|列|栏|字段)$", "", col).strip()
            break
    if not col:
        col = "日期"
    return {"col": col, "lo": lo, "hi": hi, "raw": text}


def describe(spec):
    """把条件说成人话（写日志/报告用）。"""
    if not spec:
        return "无行条件（整列全取）"
    return "%s ∈ [%s, %s]" % (spec["col"], spec["lo"], spec["hi"])


def in_range(v, spec):
    """值 v 是否落在条件区间内。v 认不出来 → 不算（宁缺不滥）。"""
    if not spec:
        return True
    d = parse_date(v)
    if d is None:
        return False
    return spec["lo"] <= d <= spec["hi"]


def days(spec):
    """区间天数（含两端），报告里用。"""
    if not spec:
        return 0
    return (spec["hi"] - spec["lo"]).days + 1


# =================================================================
# 通用行条件（2026-11-07）：不止日期区间，还有「城市 = 德州市」这类等值
# =================================================================
#
# 为什么必须扩（山东空气质量真实数据踩出来的）：
#
#   用户要求原文是这样的：
#       表一：监测时间：2025-11-25 09:00:00.0  城市：德州市
#       表二：监测时间：2025-11-25 09:00:00.0  城市：潍坊市
#       表三：监测时间：2025-11-25 09:00:00.0  城市：临沂市
#   源表 20000 行、16 个城市、59 个时刻。旧解析器只认「从 A 到 B」的日期区间，
#   这条要求里一个区间都没有 → 解析结果 None → 整条链路当"没有条件"处理，
#   于是把源表最前面十几行（枣庄/东营）填进了「德州市」那张表。
#   三张表还只有第一张被写，后两张全空。
#
#   所以这里补三件事：
#     ① 等值条件：列 = 值 / 列：值 / 列为值 —— 文本、日期时间都能比
#     ② 按「表一 / 表二 / 第三张表」把要求原文切段，每段一组条件
#     ③ 模板自带的表前描述（"本表记录 2025年11月25日09:00 时刻德州市各监测站点…"）
#        也能当条件用 —— 用户什么都不写时，模板自己已经说了要哪个城市

_CN_NUM = {"一": 1, "二": 2, "三": 3, "四": 4, "五": 5,
           "六": 6, "七": 7, "八": 8, "九": 9, "十": 10}

# 「表一」「表1」「第三张表」「第 2 个表」
_SEC_RE = re.compile(
    r"(?:表|表格)\s*([一二三四五六七八九十]|\d{1,2})(?![\u4e00-\u9fa5])"
    r"|第\s*([一二三四五六七八九十]|\d{1,2})\s*(?:张|个|块)?\s*表")

# 一行 key:value（冒号全角半角、等号都认）
_KV_RE = re.compile(r"^\s*[-•·*—\t ]*\s*([^\s:：=，,。；;、]{2,24})\s*[:：=]\s*(.+?)\s*$")

_FW = {i: i - 0xFEE0 for i in range(0xFF01, 0xFF5F)}     # 全角 → 半角


def norm_value(v):
    """值归一：全角转半角、去首尾空白、datetime 转字符串。

    只做"看起来一样就算一样"的归一，不猜同义（那是语义层的事）。
    """
    if v is None:
        return ""
    if isinstance(v, datetime.datetime):
        return v.strftime("%Y-%m-%d %H:%M:%S")
    if isinstance(v, datetime.date):
        return v.strftime("%Y-%m-%d")
    return str(v).translate(_FW).replace("\u3000", " ").strip()


_DT_ANY = re.compile(r"(\d{4})\s*[-/年.]\s*(\d{1,2})\s*[-/月.]\s*(\d{1,2})")
_TIME_ANY = re.compile(r"(?<!\d)(\d{1,2})\s*[:：]\s*(\d{2})(?!\d)")


def _dt_key(s):
    """'2025-11-25 09:00:00.0' → (2025,11,25,9,0)；没有时间部分 → 时分为 None。

    日期和时间**分开匹配**，不写成一条正则。写成一条会踩这个坑：
        r"(\\d{1,2})\\s*日?(?:[ T](\\d{1,2})...)?"
      "2025-11-25 10:00" 里 `\\s*` 先把空格吃掉，后面那个"可选"的时分组
      于是匹配空串也能成功 —— 整条正则当场就成立，引擎根本不会回溯去试时分。
      结果就是"日期对上了、时分是 None"，所有时刻被判成同一个值：
      德州 25 行变成 370 行（15 个时刻全进来了）。这类"看起来成功"的正则
      最难查，所以日期和时间各用一条，谁缺了都看得见。
    """
    t = norm_value(s)
    m = _DT_ANY.search(t)
    if not m:
        return None
    tm = _TIME_ANY.search(t)
    hh = int(tm.group(1)) if tm else None
    mi = int(tm.group(2)) if tm else None
    return (int(m.group(1)), int(m.group(2)), int(m.group(3)), hh, mi)


def _num(s):
    t = norm_value(s).replace(",", "")
    try:
        return float(t)
    except ValueError:
        return None


def same_value(a, b):
    """两个值算不算"同一个值"。

    层次（从严到宽，能在上一层判定就绝不往下走）：
      ① 原文相同
      ② 都是日期时间 → 有时间的比到分钟（'09:00:00.0' 与 '09:00' 相等）；
         只有一方带时间 → 只比日期（"2025-11-25" = 那天所有时刻）
      ③ 都是数字 → 按数值比（'12' = '12.0'）
      ④ 去掉内部空白后相同
    """
    na, nb = norm_value(a), norm_value(b)
    if not na or not nb:
        return False
    if na == nb:
        return True
    ka, kb = _dt_key(na), _dt_key(nb)
    if ka and kb:
        if ka[3] is None or kb[3] is None:
            return ka[:3] == kb[:3]
        return ka[:5] == kb[:5]
    fa, fb = _num(na), _num(nb)
    if fa is not None and fb is not None:
        return abs(fa - fb) < 1e-9
    return re.sub(r"\s+", "", na) == re.sub(r"\s+", "", nb)


def _sec_no(m):
    s = m.group(1) or m.group(2)
    if not s:
        return 0
    return int(s) if s.isdigit() else _CN_NUM.get(s, 0)


def parse_block_sections(text):
    """把要求原文按「表一 / 表二 / 第三张表」切段 → {表号: 该段文字}。

    切不出标记就返回 {}（= 整份要求对所有表生效）。
    以第一次出现的顺序为准，同一表号出现多次就拼起来。
    """
    text = text or ""
    marks = []
    for m in _SEC_RE.finditer(text):
        no = _sec_no(m)
        if no:
            marks.append((m.start(), m.end(), no))
    if not marks:
        return {}
    out = {}
    for i, (_s, e, no) in enumerate(marks):
        end = marks[i + 1][0] if i + 1 < len(marks) else len(text)
        out.setdefault(no, "")
        out[no] += text[e:end]
    return out


def parse_conditions(text):
    """从一段文字里抠出所有「列 = 值 / 列：区间」条件。

    只认"短键 + 冒号/等号 + 值"的行 —— 散文句子（"完成填表工作"）自然被丢掉。
    返回 [{"col","op","val"} 或 {"col","op","lo","hi"}]，按列名去重（先出现的算数）。
    """
    conds, seen = [], set()
    for line in (text or "").splitlines():
        line = line.strip().lstrip("\u3000")
        if not line:
            continue
        m = _KV_RE.match(line)
        if not m:
            continue
        col, val = m.group(1).strip(), m.group(2).strip()
        if not col or not val or col in seen:
            continue

        rng = _RANGE.search(val)
        if rng:
            try:
                lo = datetime.date(int(rng.group(1)), int(rng.group(2)), int(rng.group(3)))
                hi = datetime.date(int(rng.group(4)), int(rng.group(5)), int(rng.group(6)))
            except ValueError:
                continue
            if hi < lo:
                lo, hi = hi, lo
            conds.append({"col": col, "op": "range", "lo": lo, "hi": hi,
                          "val": val, "raw": line})
            seen.add(col)
            continue

        v = val.rstrip("。；;，,、 ").strip()
        if not v:
            continue
        conds.append({"col": col, "op": "eq", "val": v, "raw": line})
        seen.add(col)
    return conds


# 模板表前描述里的「时刻」和「城市」：靠句式抠，不靠列名
_CAP_TIME = re.compile(
    r"(\d{4})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日\s*"
    r"(\d{1,2})\s*[:：]\s*(\d{2})")
# 「德州市各监测站点」「潍坊市各区域」—— 城市名后面紧跟「各」
_CAP_CITY = re.compile(r"([\u4e00-\u9fa5]{2,8}?(?:市|州|县|区|盟|地区))各")
# 抠出来常带上前面的时间词（"09:00时刻德州市各" → "时刻德州市"），统一削掉
_CAP_NOISE = re.compile(r"^(?:时刻|时间|当日|当天|上午|下午|凌晨|晚上|早上|夜间|中午|"
                        r"本表|该表|表内|全市)")


def _clean_city(name):
    prev = None
    while prev != name:
        prev = name
        name = _CAP_NOISE.sub("", name).strip()
    return name


def parse_caption(text):
    """从模板自带的表前描述里抠条件。

    例：「本表记录2025年11月25日09:00时刻德州市各监测站点的空气质量检测数据。」
        → [时间 = 2025-11-25 09:00, 城市 = 德州市]

    为什么这么做：用户常常什么都不填就点"开始填写"，而模板作者其实已经把
    "这张表要哪个城市、哪个时刻"写在表前那句话里了 —— 那是白纸黑字的依据，
    比让模型猜可靠得多。
    """
    text = text or ""
    conds = []
    m = _CAP_TIME.search(text)
    if m:
        conds.append({"col": "监测时间", "op": "eq",
                      "val": "%s-%02d-%02d %02d:%s" % (m.group(1), int(m.group(2)),
                                                       int(m.group(3)),
                                                       int(m.group(4)), m.group(5)),
                      "raw": m.group(0), "from": "caption"})
    c = _CAP_CITY.search(text)
    if c:
        city = _clean_city(c.group(1))
        if city:
            conds.append({"col": "城市", "op": "eq", "val": city,
                          "raw": c.group(0), "from": "caption"})
    return conds


def match_value(v, cond):
    """单元格值 v 满不满足这条条件。"""
    if cond.get("op") == "range":
        d = parse_date(v)
        return d is not None and cond["lo"] <= d <= cond["hi"]
    return same_value(v, cond.get("val"))


def describe_conditions(conds):
    """把一组条件说成人话（日志/报告用）。"""
    if not conds:
        return "无条件（整列全取）"
    parts = []
    for c in conds:
        if c.get("op") == "range":
            parts.append("%s ∈ [%s, %s]" % (c["col"], c["lo"], c["hi"]))
        else:
            parts.append("%s = %s" % (c["col"], c["val"]))
    return " 且 ".join(parts)


def block_conditions(req_text, block_no, caption=""):
    """综合三层依据，算出第 block_no 张表该用的行条件。

    优先级（越具体的越优先）：
      ① 用户要求里明确写给「表X」的那一段
      ② 模板自带的表前描述（"本表记录…德州市各监测站点"）
      ③ 用户要求里没有分表标记的全局条件
    """
    sec = parse_block_sections(req_text)
    if sec:
        own = parse_conditions(sec.get(block_no, ""))
        if own:
            return own, "用户要求·表%d" % block_no
        # 有分表标记但这一段没写条件：继续往下找，不要凭空用别的表的条件
    cap = parse_caption(caption)
    if cap:
        return cap, "模板表前描述"
    if not sec:
        glob = parse_conditions(req_text)
        if glob:
            return glob, "用户要求·全局"
    return [], "无（整列全取）"
