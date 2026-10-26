# -*- coding: utf-8 -*-
"""_恢复日常状态.py —— 修复"跑压测崩溃未自动还原"的残留
================================================================
背景：2026-10-14 20:51 那次跑批在写报告时崩溃（by_type KeyError），
      restore_all() 没执行 → matcher / writer 残留压测产物。
      本脚本把两者恢复为「备份-压测前-1014-205120」的状态
      （= 那次跑批前 = 日常小池状态）。

逻辑与 跑压测.py 的 restore_all() 一致：
  1) 删除 matcher 里压测期间新生成的 抽取结果-*.json / 评测报告-*.txt
  2) 删除 writer/输出 里的文件
  3) 从备份恢复：模板表结构.json、运行摘要.txt、抽取结果、评测报告、writer输出
"""
import shutil
from pathlib import Path

ROOT = Path(r"C:\Users\覃敬涛\Desktop\研一项目\压测-大文档")
BASE = ROOT.parent
TR = BASE / "template_reader"
MATCHER = BASE / "matcher"
WOUT = BASE / "writer" / "输出"
BACKUP = ROOT / "备份-压测前-1014-205120"


def main():
    assert BACKUP.exists(), "备份目录不存在：" + str(BACKUP)
    n_del, n_res = 0, 0

    # 1) 删压测期间新产物
    for f in MATCHER.glob("抽取结果-*.json"):
        f.unlink()
        n_del += 1
    for f in MATCHER.glob("评测报告-*.txt"):
        f.unlink()
        n_del += 1
    for f in WOUT.glob("*"):
        if f.is_file():
            f.unlink()
            n_del += 1

    # 2) 从备份恢复
    for name, dst_dir in [("模板表结构.json", TR), ("运行摘要.txt", MATCHER)]:
        src = BACKUP / name
        if src.exists():
            shutil.copy2(src, dst_dir / name)
            n_res += 1
    for f in BACKUP.glob("抽取结果-*.json"):
        shutil.copy2(f, MATCHER / f.name)
        n_res += 1
    for f in BACKUP.glob("评测报告-*.txt"):
        shutil.copy2(f, MATCHER / f.name)
        n_res += 1
    for f in (BACKUP / "writer输出").glob("*"):
        shutil.copy2(f, WOUT / f.name)
        n_res += 1

    print("删除残留 %d 个；从备份恢复 %d 个文件" % (n_del, n_res))

    # 3) 落盘校验
    print()
    print("=== matcher 现状 ===")
    for f in sorted(MATCHER.glob("抽取结果-*.json")):
        print("  ", f.name)
    for f in sorted(MATCHER.glob("评测报告-*.txt")):
        print("  ", f.name)
    print("   运行摘要.txt", "存在" if (MATCHER / "运行摘要.txt").exists() else "缺失")
    print("=== writer/输出 现状（%d 个文件）===" % len([f for f in WOUT.glob("*") if f.is_file()]))
    for f in sorted(WOUT.glob("*")):
        if f.is_file():
            print("  ", f.name)


if __name__ == "__main__":
    main()
