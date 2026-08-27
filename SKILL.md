---
name: brand-marketing-hub
description: 品牌营销一站式工作包（v4）。覆盖品牌调研、IP联名、艺人评估、传播方案、品牌视觉合规与高质量演示文稿生产。v4 默认采用 Research → Evidence → Story → Deck Spec → Presentation Runtime → QA/Repair 闭环；外部事实官方信源优先，真实案例与产品图必须可追溯；优先调用宿主环境原生 Presentations/PowerPoint/Artifact 能力，PptxGenJS 为 CLI fallback，python-pptx 仅保留兼容与验证用途。
version: 4.0.0
---

# 品牌营销调研中心 v4

## 核心原则

1. **Evidence first, slides second.** 外部事实、案例、产品、日期、排名、数据与事实图片先验证，再进入 deck。
2. **deck_spec.json 是演示文稿 Source of Truth。** Markdown 仅作为人类可读投影和 legacy 输入，不再同时承担事实库、页面模型与渲染指令。
3. **Reference/master first.** 有真实 PPTX/POTX/reference deck 时，优先复用其 master、layout、theme、geometry 和 recurring chrome，不从文字规则重画一个“相似模板”。
4. **Native presentation runtime first.** 宿主存在 ChatGPT Presentations / PowerPoint / Artifact presentation capability 时优先使用；CLI/无原生能力时用 PptxGenJS；`assets/md2pptx_vivo.py` 仅作为 legacy fallback。
5. **Never silently lose content.** 禁止静默截断、丢行、丢列、裁掉用户要求保留的事实；放不下必须换版式、压缩非关键文字、拆页或明确失败。
6. **P0/P1 block delivery.** 事实错误、错图、占位符、内容丢失、品牌关键违规、overflow/overlap/不可读等问题未修复前不得交付。

## 路由

| 任务 | 领域工作流 |
|---|---|
| IP 联名、合作提案、权益/报价 | `references/ip-collab.md` |
| 艺人/代言人/TGI/风险/1+N | `references/celebrity.md` |
| campaign、上市传播、RTB、KOL、预算/排期 | `references/campaign.md` |
| VI、logo、颜色、字体、设计审查 | `references/design-spec.md` |
| 市场/品牌/案例/趋势/候选调研 | `workflows/research.md` |
| 任何 PPTX/Slides 产出或改版 | `workflows/deck-production.md` + `workflows/runtime-adapters.md` |

组合任务按业务依赖执行：研究/证据 → 人/IP 判断 → campaign/策略 → deck production；design-spec 贯穿视觉交付。

## v4 五层架构

- **Domain**：IP / celebrity / campaign / design-spec / research。
- **Evidence**：`schemas/evidence.schema.json` + `schemas/asset.schema.json`。
- **Story**：每页 intent、takeaway、claim_ids、asset_ids、layout_candidates，落到 `deck_spec.json`。
- **Runtime**：ChatGPT Presentations / PowerPoint → Artifact presentation runtime → PptxGenJS → legacy Python。
- **QA**：Semantic → Structural → Geometry → Visual → Repair loop，输出遵循 `schemas/qa.schema.json`。

详细架构见 `references/v4-architecture.md`。

## Research / Evidence 强制规则

当任务包含任何外部事实或事实图片时必须读取并执行 `workflows/research.md`。

- 首选品牌/机构官方来源；重要决策事实在可行时做第二权威信源交叉验证。
- 记录 claim → source → slide；事实图片记录 asset → source → slide。
- 对矛盾信源主动报告，不自行圆掉。
- 生成图片仅用于创意概念，不得冒充真实产品、真实联名、真实人物、真实历史案例或事实证据。
- 时效性主题记录来源发布日期/更新时间和 retrieved_at。
- 未验证字段不得为了填模板而编造。

## Deck Production 强制流程

凡需要生成或修改 PPTX/Slides，读取 `workflows/deck-production.md`。

1. Brief parser：目的、受众、决策问题、用户约束、must-preserve。
2. Evidence gate：完成 claims/sources/assets manifests。
3. Story planner：先建立完整判断链，再切页；每页一个判断。
4. 生成 `deck_spec.json` 并按 `schemas/deck.schema.json` 校验。
5. Resolve template/reference：优先 retained/reference PPTX 或宿主模板。
6. Runtime adapter：按 `workflows/runtime-adapters.md` 选择最强可用 presentation runtime。
7. Render。
8. 必须把成品渲染成页面图/PDF，并生成 montage/contact sheet 进行视觉复核。
9. Semantic / structural / geometry / visual 四层 QA。
10. 仅修复失败页并重新 render，直到 P0/P1 清零。

## vivo 品牌基线

品牌规范争议的权威层级：VONE / 官方原件 > design-spec rules > deck style > renderer defaults。

高频：
- 品牌名营销语境统一写 `vivo`。
- vivo 蓝 `#1E46E6`；深蓝 `#06175E`；浅蓝 `#D1EBFE`。
- 官方字体规范以当前 vivo VI 原件为准；旧 vivo type 不得继续作为现行规范。
- logo、联合标识、安全距离、最小尺寸与例外审批均回 `references/design-spec.md` / 原件核验。

## Legacy Markdown → PPTX

旧模板与 `assets/md2pptx_vivo.py` 保留兼容，但不再是新 deck 默认架构。

仅在宿主没有原生 presentation runtime、没有 PptxGenJS 路径、或用户明确要求 legacy Markdown pipeline 时使用：

```bash
python3 assets/md2pptx_vivo.py input.md output.pptx
python3 assets/validate_pptx.py output.pptx --md input.md --json
```

Legacy renderer 也必须遵守 Never-silent-loss：超出支持的表格/布局容量时显式失败，不得切片截断。

## QA Severity

| 等级 | 典型问题 | 交付 |
|---|---|---|
| P0 Critical | 事实错误、错事实图、静默丢内容、placeholder、关键 logo/品牌违规 | 禁止 |
| P1 Major | overflow、overlap、不可读、严重裁切、坏表/坏图、证据页缺来源 | 禁止 |
| P2 Minor | 对齐/间距/字体或色彩轻微偏差 | 应修 |
| P3 Cosmetic | 可选精修 | 可带说明交付 |

不得再按“问题页占比”放行 P0/P1。

## 模板确认

用户已经明确模板/风格/参考文件时直接执行，不重复询问。只有模板选择会实质改变结果且当前上下文无法判断时才确认。参考文件存在时优先 reference-native，不将其降级成纯色值/字号摘要。

## 验证

最终交付至少满足：
- 关键事实可追溯到 evidence manifest；
- factual image 可追溯且语义匹配；
- 用户 must-preserve 内容完整；
- deck 判断链前后一致；
- 无静默数据丢失；
- P0/P1 = 0；
- 已完成实际渲染后的视觉复核，而非仅解析 PPTX XML；
- 视觉交付通过 design-spec 合规检查。
