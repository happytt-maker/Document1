# -*- coding: utf-8 -*-
"""知识图谱增强 · 接口自检（打 8000 真服务）

验后端「关系怎么算出来」：
  ① 默认视角：对象 / 共享值 / 各类边都在，且每条边都带 kind + label
  ② 五种视角各自返回对应类型的边（rel / shared / link / type / doc）
  ③ 关系转换：collapse=1 时出 link 边，标签是「属性=值」，值节点被收掉
  ④ 共享阈值 min_shared 生效
  ⑤ 限定文档 / 关系类型清单 / 越界参数不崩
  ⑥ 响应时间
"""
import json
import sys
import time
import urllib.parse
import urllib.request

sys.stdout.reconfigure(encoding="utf-8")
BASE = "http://127.0.0.1:8000"
OPENER = urllib.request.build_opener(urllib.request.ProxyHandler({}))
PASS = []


def check(name, cond, extra=""):
    PASS.append((name, bool(cond)))
    print(("  ✓ " if cond else "  ✗ ") + name + (("  " + str(extra)) if extra != "" else ""))


def get(qs):
    t = time.time()
    with OPENER.open(BASE + "/api/knowledge/graph?" + qs, timeout=30) as f:
        d = json.loads(f.read().decode())
    return d, time.time() - t


print("① 默认视角")
d, dt = get("")
st = d.get("stats") or {}
check("ok", d.get("ok") is True)
check("有节点和边", len(d.get("nodes") or []) > 50 and len(d.get("edges") or []) > 100,
      "%d 点 / %d 边" % (len(d.get("nodes") or []), len(d.get("edges") or [])))
check("统计含对象数 / 共享值 / 跨表同值",
      st.get("objects", 0) > 50 and st.get("shared_values", 0) > 10 and st.get("cross_values", 0) > 0, st)
check("默认含 显式关系+共享属性+同类型",
      st["edges"]["rel"] > 0 and st["edges"]["shared"] > 0 and st["edges"]["type"] > 0,
      st.get("edges"))
check("每条边都带 kind / label",
      all(e.get("kind") and e.get("label") for e in (d.get("edges") or [])))
check("节点都带 id / label / type",
      all(n.get("id") and n.get("label") and n.get("type") for n in (d.get("nodes") or [])))
check("对象节点带类型与来源文档",
      any(n.get("docs") and n.get("group") for n in d["nodes"] if n["type"] == "object"))
check("值节点带「几个对象共用」", any(n.get("n_obj") for n in d["nodes"] if n["type"] == "value"))
check("响应够快（<1.5s）", dt < 1.5, "%.2fs" % dt)

print("\n② 五种视角")
for v, kinds, need in (("rel", ("rel",), 3), ("shared", ("shared",), 20),
                       ("type", ("type",), 20), ("doc", ("doc",), 20),
                       ("attr,shared", ("attr", "shared"), 20)):
    g, _t = get("views=" + v)
    e = (g.get("stats") or {}).get("edges") or {}
    only = [k for k, n in e.items() if n]
    check("视角 %-11s 只出 %s" % (v, "/".join(kinds)),
          set(only) <= set(kinds) and len(g.get("edges") or []) >= need,
          "%d 边 %s" % (len(g.get("edges") or []), e))

print("\n③ 关系转换（值节点折叠成对象直连）")
g1, _ = get("views=shared")
g2, _ = get("views=shared&collapse=1")
l2 = [e for e in (g2.get("edges") or []) if e["kind"] == "link"]
check("折叠后出现 link 边", len(l2) > 20, "%d 条" % len(l2))
check("link 标签是「属性=值」", all("=" in e["label"] for e in l2),
      [e["label"] for e in l2[:4]])
check("折叠后值节点被收掉（节点变少）",
      len(g2.get("nodes") or []) < len(g1.get("nodes") or []),
      "%d → %d" % (len(g1.get("nodes") or []), len(g2.get("nodes") or [])))
check("折叠后不再有 shared 边", (g2.get("stats") or {})["edges"]["shared"] == 0,
      (g2.get("stats") or {}).get("edges"))
pair = l2[0]
check("link 两端都是对象节点",
      all(x["id"].startswith("o:") for x in g2["nodes"] if x["id"] in (pair["s"], pair["t"])))

print("\n④ 共享阈值 min_shared")
a, _ = get("views=shared&min_shared=2")
b, _ = get("views=shared&min_shared=6")
sa = (a.get("stats") or {}).get("shared_values", 0)
sb = (b.get("stats") or {}).get("shared_values", 0)
check("阈值调高后共享值变少", sb < sa, "%d → %d" % (sa, sb))
check("阈值调高后边也变少", len(b.get("edges") or []) < len(a.get("edges") or []),
      "%d → %d" % (len(a.get("edges") or []), len(b.get("edges") or [])))

print("\n⑤ 文档限定 / 关系清单 / 边界")
doc = "员工信息表.xlsx"
dd, _ = get("views=rel,shared,attr&doc=" + urllib.parse.quote(doc))
check("限定文档后只看这一份",
      len(dd.get("nodes") or []) > 3 and (dd.get("stats") or {}).get("objects", 0) <= 20,
      (dd.get("stats") or {}).get("objects"))
check("限定文档后 docs 只有一份", dd.get("docs") == [doc], dd.get("docs"))
rt = d.get("relation_types") or []
check("关系类型清单按条数倒序", [x["n"] for x in rt] == sorted([x["n"] for x in rt], reverse=True),
      rt[:3])
check("关系类型带 kind + 中文标签", all(x.get("kind") and x.get("label") for x in rt))
e1, _ = get("views=bogus")
check("不认识的 views 兜底成默认", (e1.get("stats") or {}).get("edges", {}).get("shared", 0) > 0)
e2, _ = get("max_nodes=5&max_edges=8")
check("上限压到极小也不崩（服务端保底 20 个）",
      e2.get("ok") is True and len(e2.get("nodes") or []) <= 20,
      len(e2.get("nodes") or []))
e3, _ = get("max_nodes=99999&max_edges=99999&min_shared=999")
check("上限调超大也不崩", e3.get("ok") is True)

print("\n⑥ 数据可核验")
allkinds = {e["kind"] for e in (d.get("edges") or [])}
check("边类型都在约定的五种之内",
      allkinds <= {"rel", "shared", "link", "attr", "type", "doc"}, allkinds)
shared_edges = [e for e in (d.get("edges") or []) if e["kind"] == "shared"]
check("共享边都标了「被几个对象用」", all(e.get("users", 0) >= 2 for e in shared_edges),
      [e.get("users") for e in shared_edges[:6]])
cross = [e for e in shared_edges if e.get("cross")]
check("识别出跨属性同值的边（同一实体出现在不同属性下）", len(cross) > 0, "%d 条" % len(cross))

ok = sum(1 for _n, o in PASS if o)
print("\n通过 %d/%d" % (ok, len(PASS)))
for n, o in PASS:
    if not o:
        print("  未过：", n)
sys.exit(0 if ok == len(PASS) else 1)
