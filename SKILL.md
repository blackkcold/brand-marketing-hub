---
name: brand-marketing-hub
description: vivo 通用 Source-to-Deck 工作包（v4.1）。核心目标：输入任意文档、表格、PPT、PDF 或多文件资料，完整摄取与追踪源内容，必要时进行安全的深度联网研究，重组为决策清晰的故事结构，并通过真实 vivo reference/master 或统一 vivo visual template 生成可编辑、视觉一致的 PPT 汇报文件；IP联名、艺人、Campaign、VI 等作为可选领域增强模块。
version: 4.1.1
---

# vivo Universal Source-to-Deck v4.1.1

## 唯一主目标

> **任意业务资料 → 深度理解/搜索 → 结构化整理 → vivo 一致性可编辑 PPT → Render/QA/Repair**

本 Skill 的主流程不是“套某个营销模板”，也不是“Markdown 转 PPT”。所有需要制作或改版汇报文件的任务默认执行 `workflows/source-to-deck.md`。

## 主路由

### A. 用户提供文档 / 表格 / PPT / PDF / 多文件，并要求整理、汇报、做 PPT
**必须先走 Source-to-Deck。**

1. Source Inventory
2. Type-specific Ingestion
3. Content Units + Preserve Policy
4. Confidentiality Gate
5. Research & Evidence
6. Synthesis
7. Story Archetype
8. Deck Spec
9. Coverage Manifest
10. vivo Visual Template / Reference
11. Presentation Runtime
12. Render + QA + Repair

不得跳过 ingestion/coverage 直接开始画 PPT。

### B. 用户只问品牌/IP/艺人/Campaign/VI 问题，不要求 PPT
按需调用领域模块：
- IP/品牌联名 → `references/ip-collab.md`
- 艺人/代言人 → `references/celebrity.md`
- Campaign/上市传播 → `references/campaign.md`
- vivo VI/品牌视觉 → `references/design-spec.md`
- 市场/案例/品牌调研 → `workflows/research.md`

### C. 用户要求 PPT，同时任务属于某个领域
**Source-to-Deck 是主流程；领域模块只是 synthesis 的增强器。**
领域模块不得选择 renderer、不得直接生成 Markdown deck、不得创建另一套视觉模板。

---

# 1. Source Ingestion：任何输入都必须可追踪

读取 `workflows/ingest/file-routing.md` 并按文件类型执行。

优先使用宿主的原生文件能力：
- 文档：完整结构、段落、标题、表格；
- 表格：sheet/range、值、公式、单位、日期、图表；
- PPT：slide order、text、notes、tables、charts、images；作为 style-reference 时额外读取 master/layout/theme；
- PDF：文本 + 必要的页面视觉检查；
- 多文件：统一 FILE ID，保留原始 locator。

本地/CLI fallback：
```bash
python3 assets/ingest_sources.py file.docx file.xlsx old.pptx --out-dir source_bundle
```

标准输出：
- `source_inventory.json` → `schemas/source.schema.json`
- `content_units.json` → `schemas/content-unit.schema.json`

## Source role 强制区分
- content-source：提供业务内容；
- data-source：提供数字/公式；
- style-reference：提供视觉模板；
- evidence-source：提供事实证明；
- mixed：确实同时承担多种角色。

**输入是 PPT ≠ 自动把其旧版式当视觉模板。**

---

# 2. Preservation & Coverage：禁止 AI 总结时漏内容

每个 Content Unit 必须标记：
- `exact`：原意/数字/措辞必须精确保留；
- `semantic`：可改写但含义、数据、条件必须完整；
- `summarize-ok`：允许压缩；
- `optional`：可在有理由时舍弃。

生成 deck_spec 后必须建立 `coverage.json`：
```bash
python3 assets/init_coverage.py --content content_units.json --deck deck_spec.json --out coverage.json
```

规则：
- 每个 UNIT 有且只有一个 disposition；
- exact 只能 body/appendix；
- semantic 被摘要必须验证保真；
- intentionally-excluded 必须写原因；
- missing = P0；
- 用户明确要求保留但未映射到 slide = P0。

Never Silent Loss 不只针对 renderer，也针对**源文件摘要阶段**。

---

# 3. Research：先保护内部信息，再深度搜索

读取 `workflows/research.md`。

## Confidentiality Gate
搜索前将资料分类：
- public
- internal
- confidential
- restricted

严禁把未发布产品名、内部项目代号、预算/报价、合同条款、个人数据、内部业务数字直接放入公网搜索 query。必要时将查询抽象为公开问题。

## Research 要求
- 先做 research-question plan；
- entity resolution；
- 官方/合作方官方优先；
- 必要时本地语言/区域搜索；
- 重要判断做交叉验证；
- 检查 freshness；
- 冲突显式保留；
- 搜索到足够证据即停止，不做无意义堆料。

Evidence 使用 `evidence_refs`：
- 用户文件 → `UNIT-xxxx`
- 网络来源 → `SRC-xxx`

---

# 4. Real Asset Policy：真实案例必须是真图

`schemas/asset.schema.json` 要求事实素材记录：
- subject_entity
- collaboration
- product_model
- visual_role
- source_ref / source_url
- semantic_verified
- verification
- embedded_required

以下角色必须语义验证：
- product-proof
- collaboration-proof
- case-study
- logo

生成图片只能是 `generated-concept`，不能冒充真实产品、真实联名、人物、活动或历史案例。

最终 PPTX 禁止外链事实图片。所有 `embedded_required=true` 的素材必须先 materialize/download，再嵌入 PPTX。

`assets/validate_pptx.py` 会把 external image relationship 作为 **P0**。

---

# 5. Synthesis：不是按原文件顺序搬运

读取 `workflows/synthesis.md`。

必须区分：
1. Source Fact
2. Verified External Fact
3. Analysis / Inference

围绕“这份汇报要推动什么判断/决策”重组内容，而不是简单复制原文目录。

默认论证链：
> Context → Evidence → Interpretation → Choice → Next Action

重要但非正文必需的信息优先放 Appendix，而不是删除。

---

# 6. Story Archetype ≠ Visual Template

`deck_spec.json` 中必须同时指定：

## Story Archetype
- research
- strategy
- campaign
- brand-partnership
- business-review
- project-update
- proposal
- custom

决定**怎么讲故事**。

## Visual Template
默认：
`vivo-house`

由：
`brand/vivo/template-manifest.json`
控制**长什么样**。

不同业务类型不得因此产生互不一致的视觉设计系统。

---

# 7. vivo Visual Template / Reference

优先级：

1. 用户明确提供的最新 vivo master/reference PPTX；
2. 已保存的 `artifact-template-vivo-internal-report`（若宿主存在）；
3. `brand/vivo/template-manifest.json` + `layout-map.json`；
4. `references/deck-style.md` fallback；
5. legacy renderer 最后兜底。

有真实 reference/master 时：
- 复用 master/layout/theme；
- 复用 recurring chrome；
- 保留 native editable table/chart；
- 不把 reference 简化为“蓝色 + 字号”后重画。

合作品牌颜色只能作为经验证的 accent；vivo house system 仍控制整体一致性。

---

# 8. Deck Spec：内容必须 typed

使用 `schemas/deck.schema.json`。

支持的核心 block：
- headline
- body_text
- bullets
- stat
- table
- chart
- image
- comparison
- callout
- timeline
- source_footer

每个 block 应携带适用的：
- unit_ids
- claim_ids
- asset_ids

每张 slide 有：
- intent
- takeaway
- source_unit_ids
- claim_ids
- asset_ids
- layout_candidates
- must_preserve_unit_ids

禁止用无约束 generic object 代替内容结构。

---

# 9. Runtime

读取 `workflows/runtime-adapters.md`。

优先：
1. Host-native Presentations / PowerPoint
2. Artifact presentation runtime
3. PptxGenJS fallback
4. Legacy python-pptx

PptxGenJS fallback 已实际实现：
```bash
npm install
node runtime/pptxgenjs/render.js deck_spec.json output.pptx assets.json brand/vivo/template-manifest.json
```

`assets/md2pptx_vivo.py` 仅为旧 Markdown 兼容，不是默认路径。

---

# 10. Visual QA：XML PASS 不等于 PPT PASS

读取 `workflows/visual-qa.md`。

必须实际 render：
```bash
python3 assets/render_pptx.py output.pptx --out-dir render
```

得到：
- PDF
- per-slide PNG
- montage.png

然后检查：
- hierarchy
- whitespace
- title baseline
- margins
- image crop
- product visibility
- table/chart readability
- logo/master consistency
- partner accents
- cross-slide rhythm
- factual image correctness

---

# 11. QA Delivery Gates

`schemas/qa.schema.json` 七个 gate：
- coverage
- semantic
- embedded_assets
- template_fidelity
- structural
- geometry
- visual

## Severity
- **P0**：事实错误、源内容丢失、错事实图、外链事实图、placeholder、关键品牌错误；
- **P1**：overflow、overlap、不可读、严重 crop、坏图表/表格、证据缺失、模板明显不一致；
- **P2**：明显降低精修度的问题；
- **P3**：可选 cosmetic polish。

P0/P1 未清零禁止交付。

Manifest 校验：
```bash
python3 assets/validate_v4_manifests.py \
  --sources source_inventory.json \
  --content content_units.json \
  --evidence evidence.json \
  --assets assets.json \
  --coverage coverage.json \
  --deck deck_spec.json \
  --template brand/vivo/template-manifest.json
```

---

# 12. Revision：局部修改优先

用户针对现有 PPT 提出修改时读取 `workflows/revision.md`。

- 先 inspect 现有 PPT；
- 定位 slide/object；
- 尽量原位修改；
- 未修改页面保持不变；
- 内容变化同步 coverage/deck_spec；
- factual change 同步 evidence/assets；
- 修改页及相邻页重新 render；
- 最终整套 P0/P1 gate 再跑一次。

禁止“小改一页 → 整套完全重生成”导致版式漂移。

---

# 13. vivo Domain Intelligence

领域模块只负责业务判断，不拥有 Presentation Pipeline。

共享方法论可继续使用：
- 内容容器四维模型；
- 三阶段 IP 演进；
- RTB；
- 广度曝光 + 深度种草；
- 一套内容 + 一个用户事件 + 一个转化载体。

品牌视觉权威：
> 官方原件/VONE > 真实 reference/master（版式） > design-spec rules > fallback deck-style > renderer defaults

---

# 14. Final Definition of Done

只有全部满足才算完成：

- 所有输入文件已进入 Source Inventory；
- 关键源内容 locator 可追溯；
- Coverage 无 missing；
- 关键数字/公式/日期无语义损失；
- 深度研究未泄露内部信息；
- 事实 claims 可追溯；
- 事实图片身份正确；
- 所有事实图片嵌入 PPTX；
- story archetype 与 visual template 已分离；
- PPT 使用一致 vivo visual system；
- 输出可编辑；
- 已实际 render；
- Visual QA 已执行；
- P0/P1 = 0。
