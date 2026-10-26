# 文档读取器

给一个文件夹，把里面的文档读出来，输出成带**位置信息**的 JSON。

这是 A-23 赛题的第一步，也是唯一必须自己写的一步。
后面的模板理解、信息抽取、填表，都建立在「能从文档里精确定位内容」这个能力上。

## 四种格式，全部打通

- `.txt` — 纯文本（自动识别 UTF-8 / GBK 编码）
- `.md` — Markdown（额外记录每个段落所属的**标题路径**）
- `.docx` — Word（段落 + 表格，含合并单元格处理）
- `.xlsx` — Excel（多工作表 + 公式识别，含合并区域）

## 实测结果

| 格式 | 文件数 | 内容块 | 公式 | 定位方式 |
|---|---|---|---|---|
| docx | 4 | 156 | — | 段落序号 / 表-行-列 |
| md | 4 | 126 | — | 行号区间 + 标题路径 |
| txt | 6 | 35 | — | 行号区间 |
| xlsx | 4 | 384 | **38** | 工作表名 + 单元格坐标 |
| **合计** | **18** | **701** | **38** | |

验收脚本抽查 54 个块，**全部能按位置回到原文**。

## 怎么跑

### 一键全流程（最省事）

在 VSCode 里按 `F5`，选「一键全流程（读取+验收+汇总）」。
它会把读取、验收、汇总跑一遍，最后生成 `汇总报告.txt`。

### 分步跑

在 VSCode 里打开某个脚本，点右上角 ▷ 运行。可用的有：
`main.py` / `verify.py` / `check_docx.py` / `check_xlsx.py` / `demo.py`

### 在终端里

```bash
python main.py                          # 读默认的测试数据文件夹
python main.py "某个文件夹路径"           # 读指定文件夹
python main.py "路径" -o 结果.json        # 指定输出文件名
python verify.py                        # 跑验收检查
python run_all.py                       # 一键全流程
```

> 记得先在 VSCode 里选好解释器：
> `Ctrl+Shift+P` → `Python: Select Interpreter` → 输入
> `C:\Users\覃敬涛\.workbuddy\binaries\python\envs\default\Scripts\python.exe`
> 并且要用「打开文件夹」打开本目录，**不要双击 .py 文件**（会 import 失败）。

## 输出长什么样

```json
{
  "source_dir": "...",
  "supported_types": [".docx", ".md", ".txt", ".xlsx"],
  "file_count": 18,
  "ok_count": 18,
  "documents": [
    {
      "file": "员工信息表.xlsx",
      "type": "xlsx",
      "sheet_count": 1,
      "sheets": [
        { "name": "员工信息", "max_row": 10, "max_column": 12,
          "block_count": 109, "formula_count": 16 }
      ],
      "block_count": 109,
      "blocks": [
        {
          "file": "员工信息表.xlsx",
          "type": "xlsx",
          "text": "=G3+H3",
          "locator": { "kind": "cell", "sheet": "员工信息",
                       "cell": "I3", "row": 3, "col": 3 },
          "value_type": "formula_uncached",
          "formula": "=G3+H3",
          "is_formula": true
        }
      ]
    }
  ]
}
```

**每个块都带 `locator`**。拿着它回到原文，就能找到那句话 —— 这是整个系统的地基。

不同格式的块还多带各字段：

| 格式 | 额外字段 |
|---|---|
| txt | 无（只有 `locator`） |
| md | `section`（标题路径）、`is_heading`、`is_code` |
| docx | `style`、`is_heading`、`grid_span`、`v_merge` |
| xlsx | `value_type`、`formula`、`is_formula`、`merged_range` |

## 文件结构

```
doc_reader/
├── main.py                 主程序：扫描文件夹 → 读取 → 输出 JSON
├── verify.py               验收：抽查位置信息能否回到原文
├── run_all.py              一键全流程（读取 → 验收 → 汇总）
├── check_docx.py           docx 专项检查（唯一单元格数、合并标记等）
├── check_xlsx.py           xlsx 专项检查（公式、多工作表、空值）
├── demo.py                 演示脚本：展示读取器到底做了什么
├── readers/
│   ├── __init__.py         按扩展名分派给对应读取器
│   ├── base.py             公共工具（编码识别、分块规则）
│   ├── txt_reader.py       txt 读取器
│   ├── md_reader.py        md 读取器（含代码块排除逻辑）
│   ├── docx_reader.py      docx 读取器（含合并单元格去重）
│   └── xlsx_reader.py      xlsx 读取器（含公式双读）
├── output.json             运行结果（每次运行覆盖）
├── 验收报告.txt              verify.py 的输出
├── 汇总报告.txt              run_all.py 的输出
├── docx检查报告.txt          check_docx.py 的输出
├── xlsx检查报告.txt          check_xlsx.py 的输出
└── .vscode/                VSCode 配置（解释器路径、调试配置）
```

## docx 读取器的三个关键处理

**1. 合并单元格去重（最重要）**

```python
seen = set()
for r_idx, row in enumerate(table.rows, start=1):
    for c_idx, cell in enumerate(row.cells, start=1):
        key = cell._tc          # ← 底层 XML 元素，不是 id(cell)
        if key in seen:
            continue
        seen.add(key)
```

实测对比：会议纪要 6 行 4 列的表，正确值是 23 个唯一单元格。
`id(cell)` 去重得到 24（错），`cell._tc` 去重得到 23（对）。

**2. 段落和表格要分别读**

`doc.paragraphs` 不含表格里的段落，`doc.tables` 是独立列表。
只读一个会丢掉一半内容。

**3. 合并后的内容会叠加**

`cell.text` 可能是 `'吴晓敏\n吴晓敏'`（纵向合并导致两行拼一起）。
`_split_merged()` 会处理：内容全同就去重留一份，不同就用空格连起来。

## xlsx 读取器的四个关键处理

**1. 公式单元格要读两遍工作簿**（最容易踩）

openpyxl 用 `data_only=True` 读的是「**上次被 Excel 计算并保存时的缓存值**」。
程序生成的、从未被 Excel 打开过的文件，缓存槽位是空的 → 读出 `None`。

```python
wb_formulas = openpyxl.load_workbook(path, data_only=False)  # 拿公式字符串
wb_values   = openpyxl.load_workbook(path, data_only=True)   # 拿缓存值（可能 None）
```

实测：员工信息表 I3 = `=G3+H3`
- `data_only=False` → `'=G3+H3'`
- `data_only=True` → `None`

所以值类型分三种：`formula_uncached`（有公式没缓存，**本次实测 38 处全是这种**）、
`formula_cached`（有公式且有缓存）、`text/number/datetime/bool`（普通值）。

**2. 必须遍历所有工作表**

只读 `wb.active` 会漏 sheet。实测「项目进度跟踪表.xlsx」有
`项目总览`（43 块）和 `工时登记`（37 块）两个工作表。

**3. 空值和 0 是两回事**

库存盘点表里「本期入库」为空是真没数据（入库单未到），不是 0。
判断用 `is None`，不要用 `if not value`。

**4. 合并区域只有左上角有值**

其余格返回 `None`。用 `sheet.merged_cells.ranges` 拿到所有合并区，
把区域内的坐标记下来跳过，块里附带 `merged_range` 字段。

## 定位方式对照

| 格式 | locator 长什么样 | 怎么回到原文 |
|---|---|---|
| txt / md | `{"line_start": 5, "line_end": 13}` | 按行号取第 5-13 行 |
| docx 段落 | `{"kind": "paragraph", "para_index": 7}` | 取 `doc.paragraphs[6]` |
| docx 单元格 | `{"kind": "table_cell", "table_index": 1, "row": 2, "col": 3}` | 取 `doc.tables[0].rows[1].cells[2]` |
| xlsx 单元格 | `{"kind": "cell", "sheet": "员工信息", "cell": "B3", "row": 3, "col": 2}` | 取 `wb["员工信息"]["B3"]` |

`verify.py` 里就是照这张表实现回溯的（`BACKTRACKERS` 字典按扩展名分派）。

## 设计约定

**为什么每个读取器都输出同样的结构？**
这样后面加 docx、xlsx 时，主程序和下游流程一行都不用改。
本次加 xlsx 就是只写了 `readers/xlsx_reader.py` 加一行注册，`main.py` 完全没动 —— 约定生效了。

**为什么位置信息要单独放在 `locator` 里，不直接拼进 text？**
位置和内容是两回事。内容是给模型看的，位置是给程序回溯用的。混在一起以后没法拆。

**为什么分块规则是「空行分隔」？**
因为这是最通用、最少假设的规则。不同格式的文档结构差异很大，用规则去猜结构容易出错，
不如先老老实实按空行切，把判断留给后面的处理环节。

## 已知限制

- 只读内容，不写内容（写回表格是后面的事）
- 分块只按空行切（txt/md），不识别表格结构
- docx 的段落和表格是分两块存的，没有按原始顺序交错排列
- xlsx 读到的是「单元格粒度」，还没做表头识别 / 行记录切分
- 一次读一个文件夹，没有并发
- 扫描件 PDF 还没支持（需要 OCR，优先级低）

## 下一步要加什么

1. **读一张空白模板表** —— 识别表头语义，理解「这一列要填什么」
2. **把内容块喂给模型** —— 给它字段名，让它在块里找证据
3. **把结果写回模板副本** —— 且不破坏原格式（这是真正的难点）

---

创建日期：2026-09-28
最近更新：2026-09-28（加上 xlsx 支持，四种格式全部打通）
