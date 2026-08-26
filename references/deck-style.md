# vivo 2026 内部 deck 设计系统（deck-style）

> 本文件属于 brand-marketing-hub。提炼自 8 份真实 vivo/关联 deck：泡泡玛特进度更新、S 系列试点新玩法、X500 迪士尼、漫展材料、iQOO16 名人备选、2026 FIRST 传播、X 系列艺人合作分析、品牌与 IP 营销趋势分享。实现脚本：`assets/md2pptx_vivo.py`。权威层级：VI 3.1 原件 > 本文件 > 脚本默认值。

## 1. 设计令牌

### 1.1 画布与网格

- 画布：16:9，13.333 × 7.5in。
- 安全区：左右 0.55-0.9in，上 0.42-0.55in，下 0.45in。内容不得压到底部页脚区域。
- 内容页主内容区：x=0.55-0.62in，y=1.35-1.45in，w≈12.1in，h≤5.4in。
- 页脚：左下 9pt 灰字，右下页码 9pt，页码贴右下，不居中，不做装饰条。

### 1.2 色盘

- vivo 品牌蓝：`#1E46E6`，所有 title label、编号、关键结论、表头、卡片标题统一使用。
- 明亮蓝：`#3458F6`，用于封面/尾页蓝色渐变的亮端、局部高亮和大面积蓝场过渡。
- 深蓝：`#06175E`，用于正文引导词、重要判断和深色蓝字。
- 浅蓝：`#D1EBFE`，用于封面副标题、蓝场上的次级文字。
- 卡片浅蓝：`#EAF2FF`，用于 RTB/WHAT/WHY/HOW、矩阵、轻量卡片。
- 表格隔行：`#F2F7FF`。
- 正文黑：`#111111`；正文灰：`#565656`；页脚/图注灰：`#9AA3B2`。
- 黄色点缀：克制使用，接近 `#F6C84C`。只用于关键时间、风险提示、KV 局部光点或矩阵重点格，每页≤1处，不替代 vivo 蓝。
- 警示红：`#E6001E`，仅用于高风险/不可延误节点，每页≤1处。

### 1.3 字体与字号

- 字体：微软雅黑为企业 PPT 模板基准；脚本回退链为 微软雅黑 → PingFang SC → Arial → Helvetica。macOS 中文输出优先声明 PingFang SC，不回退到无 CJK 字形的 Arial。
- 封面标题：40pt，白色，bold，行距 1.15。
- 封面副标题：20pt，浅蓝。
- 内容页视觉标题：20pt。现有 renderer 的蓝色标题条使用 14pt 白字承载页面标题，模板写作仍以 20pt 信息层级组织标题。
- 内容页正文：12pt，行距 1.25。
- 卡片/表格：10.5-11pt。
- 图注/页脚/密级：9pt。
- 标题层级只靠字号、bold、蓝色 title label 拉开，不使用花哨字体、不混用 campaign 主题字体。

## 2. Renderer 支持的 Markdown 方言

`assets/md2pptx_vivo.py` 当前支持以下指令。模板只能使用这些结构；若新增语义，必须写成备注或普通文本，不能假装 renderer 会识别。

| 写法 | 渲染结果 | 使用边界 |
|---|---|---|
| `# 文档标题` | 封面主标题 | 只写一次 |
| `<!-- deck: 副标题 -->` | 封面副标题 | 放在正文第一页前 |
| `<!-- deck: meta: 部门｜日期｜密级 -->` | 封面左下 9pt 小字 | 建议所有模板使用 |
| `## P{n}｜标题` | 新建 slide | P 编号必须连续 |
| `@layout bullets` | 蓝标题条 + bullet 内容页 | 默认布局，可省略但建议显式写出 |
| `@layout cards N` | N 列浅蓝卡片 | `N` 为列数，不是卡片总数 |
| `@layout split left/right` | 图片 + bullets 图文页 | `left/right` 表示图片位置 |
| `@layout table` | Markdown 表格页 | 适合权益、预算、排期、对比 |
| `@layout part` 或 `@part` | 章节分隔页 | 白底 Part 样式，不是蓝底 |
| `@layout end` | 蓝色尾页 | 用于 Thank You/谢谢 |
| `@sub 文案` | 标题条右侧结论句 | 12pt 黑字，≤22 字 |
| `@img 路径 | 图注` | split 页图片 cover-crop + 9pt 图注 | 显式 @img 的图片缺失时保留可见图槽占位；没有图片需求时不要写 @img，改用 bullets/cards |
| `### 卡片标题` + `- 要点` | cards 页卡片 | 每卡 2-4 条 |
| Markdown 表格 | table 页 | 表头蓝底白字，隔行浅蓝 |
| `>` 备注 | speaker notes | 可写 `> Archetype：...` 供人工识别 |

### 2.1 Archetype 标注规则

- Renderer 不识别 `@archetype`。模板中的 archetype 一律写在备注里：`> Archetype：RTB/WHAT/WHY/HOW；Layout：@layout cards 4`。
- 备注是给人和后续 parser/migration 使用，不影响现有 pptx 渲染。
- 不得发明 `@blue-cover`、`@timeline`、`@matrix` 等未实现指令。蓝封面/尾页、timeline、matrix、budget 通过现有布局组合表达。

## 3. 2026 vivo house style 总原则

- 大结构：蓝色封面/尾页 + 白色内容页 + 蓝色 title label + 右下页码。
- 风格关键词：清爽、商务、可汇报、强层级、低装饰。背景大面积留白，卡片用浅蓝，重点用 vivo 蓝。
- campaign-specific style isolation：迪士尼、FIRST、漫展、IP 角色等专属视觉只进入 KV/图片/campaign 创意页，不污染全 deck 的通用标题、页脚、表格、卡片令牌。
- 每页只讲一个判断。标题是主题，`@sub` 是结论，不再另写一串解释性小字。
- 一级要点≤6；正文 bullet 每条≤24字；cards 每卡≤4条；table≤6列、≤9行；单页图片≤3张（collage 可作为一张合成图）。
- 图片必须是真实素材、官方渲染、授权 KV 或可追溯截图。无素材时用 cards/table，不放空框。
- 图片裁切：人物保脸、产品保 logo/CMF、IP 保角色完整轮廓；横图用于 cover-crop，竖图适合 split left/right。

## 4. 可复用 slide archetypes

### A01 蓝色 cover/end

- 用法：P1 封面、最后一页。
- Renderer：封面由 `#`、`<!-- deck: -->`、`<!-- deck: meta: -->` 与显式 `P1｜封面` 生成；尾页用 `@layout end`。
- 视觉：整版蓝色渐变（#3458F6 → #1E46E6/深蓝），右侧 2-3 条斜向半透明白光带，vivo 白色字标右上。
- 文案：封面主标题≤2行；副标题说明对象/场景；meta 写部门｜日期｜密级。尾页只放 THANK YOU/谢谢，不加多余 slogan。

### A02 Agenda 目录

- 用法：决策链预览。
- Renderer：标题写 `P2｜目录` 或 `P3｜目录`，脚本会自动识别；也可显式写 `@layout toc`。目录不能与报告目的、评估框架合并。
- 视觉：白底；“目录”20pt；条目 16pt；序号 vivo 蓝 bold；每条下方蓝色细线。
- 规则：只列 4-7 个章节，不混入报告目的、评估框架或方法论细节。

### A03 Report purpose 报告目的

- 用法：解决“为什么开这份报告/本轮回答什么”。尤其用于艺人报告，避免 P2 同时承担目录与框架。
- Renderer：`@layout bullets` + `@sub 一句话结论`。
- 视觉：蓝 title label + 右侧黑色结论；正文 3-5 条。
- 规则：只写业务问题、判断口径、输出物、决策用途。不得写目录条目。

### A04 Section divider 章节分隔

- 用法：进入 IP 价值论证、策略、权益报价、风险等大段落。
- Renderer：`@layout part` 或 `@part`。
- 视觉：白底；Part N + 蓝线 + 章节标题蓝 28pt + 12pt 灰副标题。
- 规则：用于 12 页以上 deck；短 deck 可不用，避免打断节奏。

### A05 RTB / WHAT / WHY / HOW

- 用法：产品力证据、策略解释、方案骨架。
- Renderer：`@layout cards 4`，四张卡标题固定为 RTB、WHAT、WHY、HOW。
- 视觉：四张浅蓝卡片横排；卡标题 vivo 蓝 13pt；短蓝线；每卡 2-4 条。
- 规则：RTB 必须是可验证产品力；WHAT 是动作；WHY 是人群/生意理由；HOW 是执行机制。

### A06 A→A-I→I-P funnel

- 用法：传播漏斗、艺人/IP 资源角色分工、从 awareness 到 interaction 到 purchase 的链路。
- Renderer：优先 `@layout cards 3`；需要细分动作时用 `@layout table`。
- 视觉：三段横向卡片或三列表格，标题固定为 A 认知、A-I 互动转化、I-P 购买转化。
- 规则：每段都要写目标、触点、内容、KPI。不得只写“曝光/种草/转化”三个空词。

### A07 Cards / matrix

- 用法：候选人分层、IP 价值维度、KOL 分层、风险分级、1+N 矩阵。
- Renderer：`@layout cards 2/3/4` 或 `@layout table`。
- 视觉：卡片用于定性并列；表格用于多维对比。浅蓝卡片承载结论，表格承载数据。
- 规则：矩阵必须有行列定义。艺人 1+N 输出固定为主代言/体验官/星光好友，不得把使用思路页写成矩阵页。

### A08 Timeline

- 用法：传播节奏、项目倒排、定制设计/合同/上线节点。
- Renderer：目前无 timeline 原生布局。用 `@layout table` 表达。
- 视觉：列建议为阶段/时间/关键动作/责任方/风险提醒；关键 deadline 用红色语义写入文本，renderer 不会自动标红整格。
- 规则：必须倒排或按预热-爆发-延续顺序；同页≤6个阶段。

### A09 Table / budget

- 用法：权益清单、报价、预算、资源清单、渠道排期。
- Renderer：`@layout table`。
- 视觉：表头蓝底白字，隔行浅蓝，首列深蓝 bold。
- 规则：预算页必须有总计和口径；分项加总=总预算；报价页必须标注含税/服务费/授权范围。

### A10 Image collage / KV

- 用法：核心创意、KV 方向、IP/产品视觉 mood、线下空间。
- Renderer：`@layout split left/right` + `@img`。多图 collage 需先合成为一张图片再引用。
- 视觉：图片侧保比例 cover-crop，高≤5in；文字侧 3-5 条说明；图注 9pt 灰。
- 规则：KV 页只说明视觉概念、产品卖点落点、可执行物料。不把预算/排期混入 KV 页。

### A11 Case study

- 用法：艺人商业案例、IP 联名案例、竞品 campaign 拆解。
- Renderer：`@layout cards 3`（策略/执行/效果）或 `@layout split right`（案例图 + bullets）。
- 视觉：案例图优先；无图则三卡。效果数据必须来自真实来源或标注待补。
- 规则：每个 case study 只拆一个案例。多个案例用多页，不塞成一页长表。

## 5. 模板迁移说明

- 旧模板中“目录与评估框架”必须拆开：P2 固定为目录或报告目的二选一；艺人报告采用 P2 目录、P3 报告目的、P4 方法论，彻底消除 celebrity P2 ambiguity。
- 旧模板中“时间线/预算/权益”如果是 bullets，迁移为 `@layout table`。
- 旧模板中“创意方向/KV/案例”如果没有图片，先保留 cards 结构；有素材后迁移为 `@layout split left/right` + `@img`。
- 旧模板中“三件要事/RTB/策略”迁移为 RTB/WHAT/WHY/HOW 或 A→A-I→I-P funnel，不再写成纯 bullet 堆叠。
- 不迁移到 renderer 不支持的自定义指令。需要让后续 parser 识别 archetype 时，从备注 `> Archetype：...` 读取。

## 6. 验证

- P 编号连续，封面不重复，最后一页为 `@layout end`。
- 每页 layout 显式：除目录页外，建议写 `@layout bullets/cards/table/split/part/end`。
- 内容页含蓝 title label、页脚、右下页码；封面/尾页含蓝渐变、斜向光带、vivo 白字标。
- 色值与字号落在本文件令牌内；campaign 专属风格只出现在 KV/图片页。
- 表格不超过 6 列 9 行；cards 不超过 4 列 8 卡；单页一级要点≤6。
- 纯代码验证不操控 PowerPoint：读回 pptx 检查页数、shape 坐标、文本估算溢出、表格行列数、图片与字标在位。
