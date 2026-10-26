# STORY.md · A-23 组会汇报

## ① 用户意图对齐

- **目标受众**：导师 + 课题组同门（研一为主），周五/周四例行组会，10~15 分钟汇报 + 提问。
- **核心目标**：让导师在 15 分钟内确认三件事——这个赛题"考什么"我们已经拆清楚；三个必须模块"都能跑、有实测数字"；剩下的是材料汇编而不是功能返工。让同门记住一句话：**从一堆乱七八糟的文档，到一张格式严整、可直接交卷的表，全链路已经打通。**
- **PPT 长度**：22 页（封面 1 + 目录 1 + 3 个章节扉页 + 15 内容页 + 结束 1 + 1 页补充）。Hero 5 页 = 22.7%。
- **视觉调性**：学术严谨、证据驱动、克制、蓝白、无营销感。
- **内容边界**：
  - 必讲：赛题三项硬要求、五层管线、三个模块各自的关键机制、准确率/响应时间/保真度三组硬数字、当前进度与缺口、下一步时间表。
  - 不讲：代码级实现细节（run 切分算法、SQL 表结构）、单个 bug 修复过程、历史踩坑。
  - 禁碰：把"未成稿的提交材料"说成已完成；使用没有实测依据的数字。

## ② 页面布局骨架

**页面总数 22 页，分 3 章。**

| 章 | 目录条目 | 页码区间 | 章节扉页 |
|---|---|---|---|
| 01 | 赛题解析：它到底考什么 | P03–P07 | P03 |
| 02 | 系统方案：我们怎么答卷 | P08–P15 | P08 |
| 03 | 验证与计划：做到哪一步 | P16–P21 | P16 |

- 目录声明 3 章 → 全篇恰好 3 个 `section` 扉页（P03 / P08 / P16），编号 01/02/03 连续，标题与目录逐字一致。
- **Hero 页**：P01 封面、P09 五层管线、P13 保真度数字、P17 硬指标实测、P22 结束页 = 5/22 = 22.7%；两两之间均至少隔 1 页 supporting。
- **rhythm 曲线**：
  - P01 peak → P02 valley → P03 transition → P04 valley → P05 valley → P06 peak → P07 valley → P08 transition
  - P09 peak → P10 valley → P11 peak → P12 valley → P13 peak → P14 valley → P15 valley → P16 transition
  - P17 peak → P18 valley → P19 valley → P20 valley(转) → P21 transition → P22 peak
  - 无连续 3 页 valley（P19/P20 后接 P21 transition；P04/P05 后接 P06 peak）。
- **非对称版式预算**：P02、P04、P06、P07、P09、P10、P11、P12、P13、P14、P15、P17、P19、P20、P21 = 15/22 = 68% ≥ 40%。
- **`左大图+右侧文字` + `非对称双栏` 合计**：P09、P11、P12、P15、P20 = 5/22 = 22.7% ≤ 40%。
- **`N卡片横排` 使用**：0 页（以非对称矩阵 / 时间轴 / 数据表替代）。

## ③ 页面大纲

### P01 封面
- type: cover ｜ role: hero ｜ rhythm: peak
- layout: 全屏视觉+大标题
- visual: SVG 五层管线抽象线稿（右侧 45%，opacity 0.16）｜ visual_role: atmosphere
- density: 字数约 60 / 图 1 组 / 留白约 30%
- anti_pattern: 禁止红金装饰、印章、烟花；禁止把标题压到画面下 1/3
- description: 赛题全称 + 副标"从非结构化文档到结构化表格的端到端系统" + 汇报人/日期。定性：这是一次"赛题拆解 + 阶段成果"汇报，不是结题答辩。

### P02 目录
- type: catalog ｜ role: supporting ｜ rhythm: valley
- layout: 左标题+右内容
- visual: 编号 01/02/03 深蓝方块｜ visual_role: anchor
- density: 字数约 130 / 图 0 / 留白约 18%
- anti_pattern: 禁止六个条目平铺；禁止圆球编号
- description: 三章结构——赛题解析 / 系统方案 / 验证与计划。判断：三章各自回答导师最关心的一句问话（考什么、怎么做、做到哪）。

### P03 章节扉页 01
- type: section ｜ role: transition ｜ rhythm: transition
- layout: 全屏视觉+大标题
- visual: 巨型数字 "01"（≥160px）｜ visual_role: anchor
- density: 字数约 70 / 图 0 / 留白约 30%
- anti_pattern: 禁止四卡片预览本章内容；禁止铺满正文
- description: 赛题解析——先把"比赛怎么判"讲清楚，后面所有数字才有意义。

### P04 赛题速览
- type: content ｜ role: supporting ｜ rhythm: valley
- layout: 巨型数字+洞察
- visual: 巨型数字 "3"（三个必须模块）+ "4"（四项技术指标）｜ visual_role: anchor
- density: 字数约 150 / 图 0 / 留白约 22%
- anti_pattern: 禁止把赛道信息塞进等宽卡片横排；禁止大段引用赛题原文
- description: 第 17 届服创大赛 A 类应用赛题，金陵科技学院出题，方向"企业服务 + 智能计算"。**3 个必须模块 + 4 项技术指标**——判断：这不是自由发挥题，是逐条打钩的验收题，所以我们的开发从一开始就按"可对照、可复跑"的方式做。

### P05 真实痛点
- type: content ｜ role: supporting ｜ rhythm: valley
- layout: 上下分栏
- visual: 左"人工流程"三段式示意（SVG 箭头）｜ visual_role: evidence
- density: 字数约 260 / 图 1 组 / 留白约 15%
- anti_pattern: 禁止拆成四张小卡；禁止用大图压缩文字
- description: 工作人员面对的是格式不一、结构松散的一堆文档：逐份读、手工摘、再敲进 Excel。重复、易错、挤占真正的工作时间。判断：痛点的本质不是"读得慢"，而是"读完之后没有结构化的落点"——所以方案必须同时解决"提取"和"落库/落表"。

### P06 赛题考什么
- type: content ｜ role: supporting ｜ rhythm: peak
- layout: 关键词矩阵（2×3）
- visual: 六格矩阵，其中"准确率 ≥80%"格用浅黄高亮｜ visual_role: anchor
- density: 字数约 210 / 图 0 / 留白约 12%
- anti_pattern: 禁止六格等宽等高无重点；禁止把"用户期望"和"技术指标"混成一坨
- description: 左列三个必须模块（文档智能操作 / 非结构化提取 / 表格填写），右列四项指标（形态不限、准确率 >80%、样本人工标注多轮审核、识别分类 API 异步）。判断：**"准确率"和"响应时间"是唯一能一票否决的两条**，其余是形态与合规。

### P07 比赛当天怎么测
- type: content ｜ role: supporting ｜ rhythm: valley
- layout: 流程时间轴（5 步）
- visual: 五段胶囊时间轴 + 右阈值标注｜ visual_role: anchor
- density: 字数约 200 / 图 1 组 / 留白约 15%
- anti_pattern: 禁止纵向长列表；禁止把 16 份样本拆成 4 张卡
- description: 16 份规格样本（5 docx / 5 xlsx / 3 md / 3 txt）一起上传 → 依次给 5 张模板 → 每张跑一次并记录响应时间 → 与样例表比对算准确率。单表 ≥80%、单文档 ≤90 秒，最后取平均。判断：**这是"批量 + 计时 + 对分"的考试，所以我们的验收脚本完全照这个流程仿真。**

### P08 章节扉页 02
- type: section ｜ role: transition ｜ rhythm: transition
- layout: 全屏视觉+大标题
- visual: 巨型数字 "02"｜ visual_role: anchor
- density: 字数约 70 / 图 0 / 留白约 30%
- anti_pattern: 禁止四卡片预览；禁止与封面雷同
- description: 系统方案——一条五层管线，三个模块分别挂在哪一层。

### P09 五层管线总架构
- type: content ｜ role: hero ｜ rhythm: peak
- layout: 非对称双栏（65:35）
- visual: L1 内联 SVG 五层管线图（左 65%，节点串 + 箭头）｜ visual_role: anchor
- density: 字数约 180 / 图 1 组 / 留白约 15%
- anti_pattern: 禁止用文字列表代替架构图；禁止把五层压成三个大块
- description: ①四格式读取 → ②内容块 + 位置坐标 → ③模板结构理解 → ④证据匹配（检索 + 大模型判定 + 程序定位）→ ⑤保格式写回。判断：**每一层都可单独验收，出问题时能定位到层，而不是"模型不行"一句话背锅。**

### P10 三条设计原则
- type: content ｜ role: supporting ｜ rhythm: valley
- layout: 左标题+右内容
- visual: 三个编号方块 01/02/03｜ visual_role: anchor
- density: 字数约 240 / 图 0 / 留白约 15%
- anti_pattern: 禁止三张等宽卡片横排；禁止把原则写成口号
- description: ①不迷信模型：模型只负责"判值 + 给摘抄"，定位与校验交给程序；②不丢出处：每条入库数据必须能回原文；③不动格式：写回只碰值、不碰样式。判断：这三条决定了系统"错得可见、错得可查"，也是后面 100% 保真度的来源。

### P11 模块一 · 文档智能操作
- type: content ｜ role: supporting ｜ rhythm: peak
- layout: 非对称双栏（60:40）
- visual: 左"一句话 → 意图 JSON → 执行 → 新文件"对话流示意（SVG）｜ visual_role: anchor
- density: 字数约 200 / 图 1 组 / 留白约 15%
- anti_pattern: 禁止只放一张功能清单；禁止把"多轮指代"藏进脚注
- description: 一句话完成改字、增删、排版、字体字号、颜色高亮、内容提取；支持"它"指代与一句话多操作；docx/xlsx/md/txt 四格式全覆盖。判断：**入口是自然语言，但执行是确定性代码**——所以同一句话重复执行结果一致，不会今天改得对、明天改错。

### P12 模块二 · 三线提取
- type: content ｜ role: supporting ｜ rhythm: valley
- layout: 非对称双栏（60:40）
- visual: 左"一份文档 → 画像 / 实体 / 字段三条线 → MySQL 五表"分支图｜ visual_role: anchor
- density: 字数约 210 / 图 1 组 / 留白约 15%
- anti_pattern: 禁止三条线画成并列卡片；禁止不提"自动入库"
- description: 画像线（主题 + 要点 + 要素）、实体线（7 类实体）、字段线（用户点名要的字段）；三线跑完自动落库五张表，支持网页 / API / 命令行三处查询。判断：提取不是终点，**进库才是**——否则前端再漂亮，数据仍然散在文件里。

### P13 保真度数字
- type: content ｜ role: hero ｜ rhythm: peak
- layout: 巨型数字+洞察
- visual: 巨型数字 "266" + "100%"｜ visual_role: anchor
- density: 字数约 150 / 图 0 / 留白约 22%
- anti_pattern: 禁止把 266 和 100% 塞进表角落；禁止只给数字不给口径
- description: 9 份文档 · 137 实体 · 4 画像 · 75 字段 = 266 条，**100% 可回溯原文**，缺摘抄/缺出处 0 条。判断：这个 100% 不是模型自评，是**零模型反向审计**——把库里的每条数据拉回原文核对，重跑结果一致。

### P14 模块三 · 表格填写
- type: content ｜ role: supporting ｜ rhythm: valley
- layout: 上大图+下方卡片
- visual: 上半"模板理解 → 检索 → 判值 → 写回 → 验收"五段流程（SVG）｜ visual_role: anchor
- density: 字数约 220 / 图 1 组 / 留白约 15%
- anti_pattern: 禁止把三铁律写进图里当小字；禁止只讲"能填"不讲"格式"
- description: 支持 Excel / Word 双形态模板，从 18 份日常文档或 16 份赛题规格文档中检索证据并写回。三铁律：永不改原模板、先查后写、只填值不碰格式。判断：**填表题的胜负不在"填进去"，在"交卷时格式还对不对"**——10 份验收报告格式差异 0 处。

### P15 系统形态与异步
- type: content ｜ role: supporting ｜ rhythm: valley
- layout: 左大图+右侧文字
- visual: L1 操作台真实截图（左 58%）｜ visual_role: evidence
- density: 字数约 200 / 图 1 张 / 留白约 12%
- anti_pattern: 禁止 50:50 等分；禁止把截图缩成 200×70 角标
- description: Web 网站形态：FastAPI + 单文件浏览器操作台，双击即启动；19 个 API 端点；识别与填写均为异步任务——提交立即返回 job_id，轮询取结果。判断：**异步不是技术炫技，是因为一次三线提取要 19 秒，界面不能干等。**

### P16 章节扉页 03
- type: section ｜ role: transition ｜ rhythm: transition
- layout: 全屏视觉+大标题
- visual: 巨型数字 "03"｜ visual_role: anchor
- density: 字数约 70 / 图 0 / 留白约 30%
- anti_pattern: 禁止四卡片预览；禁止铺满正文
- description: 验证与计划——用比赛当天的口径，验收我们自己的系统。

### P17 硬指标实测
- type: content ｜ role: hero ｜ rhythm: peak
- layout: 图表+洞察
- visual: 左表格（赛题阈值 vs 实测），右响应时间条形图｜ visual_role: anchor
- density: 字数约 190 / 图 1 张 / 留白约 12%
- anti_pattern: 禁止只报表不报阈值；禁止把 5 个时间写成一段文字
- description: 准确率：泛化 63 格严格 96.8% / 宽松 100%，模板 4 对分严格 85.7% / 宽松 100%（阈值 80%）；单表响应 10.19~43.68 秒（阈值 90 秒）；平均 30.34 秒（大文档池）/ 16.30 秒（日常池）。判断：**四项硬指标全部达标，且最差情况仍有约 2 倍余量。**

### P18 自动化验收体系
- type: content ｜ role: supporting ｜ rhythm: valley
- layout: 关键词矩阵（2×2）
- visual: 四张总检成绩单（17/17、22/22、26/26、8/8）｜ visual_role: anchor
- density: 字数约 200 / 图 0 / 留白约 15%
- anti_pattern: 禁止只写"测试通过"；禁止把 73 项写成一句结论
- description: 四份零模型、可复跑的对照总检：作品要求 17/17、模块 2 22/22、模块 3 26/26、模块 1 覆盖 8/8，合计 73 项现场复核全过。判断：**验收件本身是资产**——导师或评委随时可以重跑，不用相信我们的口头结论。

### P19 三个创新点
- type: content ｜ role: supporting ｜ rhythm: valley
- layout: 非对称双栏（65:35）
- visual: 三个编号 + 对应机制小图｜ visual_role: anchor
- density: 字数约 230 / 图 0 / 留白约 15%
- anti_pattern: 禁止写"使用了先进的大模型"这类空话；禁止超过 3 个
- description: ①自然语言文档操作：意图解析 + run 级切分，改字不破格式；②保真提取：出处定位 + 一致性重审 + 诚实留空；③零格式损失填表：模板理解 → 证据检索 → 程序判值 → 只写值。判断：**创新点不在"用了大模型"，在"给大模型套上了可验收的工程约束"。**

### P20 进度盘点
- type: content ｜ role: supporting ｜ rhythm: valley
- layout: 非对称双栏（55:45）
- visual: 左"已完成"清单 / 右"待完成"清单，中间分隔线｜ visual_role: anchor
- density: 字数约 200 / 图 0 / 留白约 15%
- anti_pattern: 禁止用对勾叉号堆砌；禁止把"未完成"淡化
- description: 已完成：三模块 + 7 条要求 + 73 项复核 + 可运行 Demo。待完成：概要介绍、简介 PPT、详细方案、演示视频、关键模块说明（≤3 创新点）、素材来源一页。判断：**剩余工作全是"汇编成稿"，不需要新功能开发**——这是当前最重要的判断。

### P21 下一步计划
- type: content ｜ role: supporting ｜ rhythm: transition
- layout: 流程时间轴（4 步）
- visual: 四段时间轴 + 里程碑标注｜ visual_role: anchor
- density: 字数约 180 / 图 1 组 / 留白约 18%
- anti_pattern: 禁止写成"继续优化"的空话；禁止给没有依据的日期承诺
- description: ①详细方案（素材最全，先写）→ ②概要介绍 → ③关键模块说明 → ④简介 PPT → ⑤演示视频（演示链路现成，最后录）。判断：**顺序按"素材密度"排，不按文档编号排**，先把最厚的写完，后面都是截取。

### P22 结束页
- type: ending ｜ role: hero ｜ rhythm: peak
- layout: 全屏视觉+大标题
- visual: 一句话结论 + 三个关键数字（266 / 100% / 73 项）｜ visual_role: anchor
- density: 字数约 60 / 图 0 / 留白约 30%
- anti_pattern: 禁止"谢谢观看"单独占一页；禁止烟花/金装饰
- description: 一句话收束：文档进、表格出，全链路已跑通，且每一步都可复跑。请导师指正。
