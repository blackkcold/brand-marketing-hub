# brand-marketing-hub

[![CI](https://github.com/blackkcold/brand-marketing-hub/actions/workflows/ci.yml/badge.svg)](https://github.com/blackkcold/brand-marketing-hub/actions/workflows/ci.yml)
[![Release](https://img.shields.io/github/v/release/blackkcold/brand-marketing-hub)](https://github.com/blackkcold/brand-marketing-hub/releases)
[![Skill Version](https://img.shields.io/badge/skill-v4.0.0-1E46E6)](./SKILL.md)

面向品牌营销、合作策略与高质量演示文稿生产的 AI Skill。覆盖 **品牌调研、IP 联名、艺人评估、Campaign、品牌视觉合规和 PPT/Slides 生产**。

> **v4 核心链路：Research → Evidence → Story → Deck Spec → Presentation Runtime → QA / Repair**

v4 不再把“做 PPT”定义成 Markdown 到固定坐标 PPTX 的转换，而是先验证事实和素材、建立决策叙事，再选择最合适的原生演示文稿运行时，并通过真实渲染做视觉 QA。

## Why v4

| 问题 | v3.x | v4 |
|---|---|---|
| 外部事实 | `@source` 文本标记 | Claim / Source Manifest |
| 真实图片 | 本地路径为主 | Asset Manifest + provenance |
| 页面模型 | Markdown + 单一 `@layout` | `deck_spec.json` + `layout_candidates` |
| 模板 | Style guide 重画 | Reference / master native first |
| Renderer | python-pptx 主导 | Native presentation runtime first |
| QA | XML/结构检查为主 | Semantic + Structural + Geometry + Visual |
| 错误放行 | 问题页比例 | P0/P1 Severity Gate |
| 超容量内容 | 可能截断 | Never Silent Loss |

## Architecture

```text
User Brief
   │
   ▼
Brief Parser
   │
   ▼
Research & Evidence
   ├── claims.json
   ├── sources.json
   └── assets.json
   │
   ▼
Story Planner
   │
   ▼
deck_spec.json
   │
   ▼
Template / Reference Resolver
   │
   ├── ChatGPT Presentations / PowerPoint
   ├── OpenAI Artifact Presentation Runtime
   ├── PptxGenJS
   └── Legacy python-pptx fallback
   │
   ▼
Editable PPTX
   │
   ▼
Render → QA → Repair
```

详细设计见 [`references/v4-architecture.md`](./references/v4-architecture.md)。

## 1. Research & Evidence

外部事实、案例、产品、日期、排名、数据与事实图片必须先进入 Evidence Layer，再进入 deck。

默认规则：

- 官方信源优先；
- 对影响最终建议的关键事实，在可行时做第二权威来源交叉验证；
- 时间敏感信息记录发布日期/更新时间和检索时间；
- 不同信源冲突时主动暴露，不静默选择；
- 未验证的信息不得为了填模板而编造；
- 真实联名、产品、人物、历史案例优先使用可追溯真实素材。

证据关系：

```text
Claim → Source → Slide
Asset → Source → Slide
```

相关文件：

- [`workflows/research.md`](./workflows/research.md)
- [`schemas/evidence.schema.json`](./schemas/evidence.schema.json)
- [`schemas/asset.schema.json`](./schemas/asset.schema.json)

### Real Asset Policy

生成图片可以用于概念创意、Mood/KV Direction、未发生方案可视化；**不得冒充真实品牌产品、真实联名、真实人物、真实历史活动或事实证明材料。**

## 2. Story & Deck Spec

v4 新 deck 的 source of truth 是 `deck_spec.json`，Markdown 仅保留为 human-readable projection / legacy input。

每页至少定义：

- `intent`：为什么存在；
- `takeaway`：看完必须记住什么；
- `claim_ids`：哪些事实支撑；
- `asset_ids`：使用哪些素材；
- `layout_candidates`：多个候选布局；
- `must_preserve`：不可丢失的用户输入。

Schema：[`schemas/deck.schema.json`](./schemas/deck.schema.json)

## 3. Presentation Runtime

brand-marketing-hub 负责 **领域智能、证据、故事结构与品牌 QA**，不重复制造宿主已经具备的完整 presentation engine。

| Runtime | 用途 | 优先级 |
|---|---|---:|
| ChatGPT Presentations / PowerPoint | 基于真实模板创建、编辑、精修 | 最高 |
| OpenAI Artifact Presentation Runtime | ChatGPT 内原生可编辑演示文稿 | 高 |
| PptxGenJS | CLI / server deterministic generation | 高 |
| python-pptx | readback、validation、patch、legacy | 兼容 |

详见 [`workflows/runtime-adapters.md`](./workflows/runtime-adapters.md)。

### Reference-native First

如果用户提供 `.pptx` / `.potx` / master deck，优先复用其原生 **master、layout、theme、typography、geometry、charts、tables 与 recurring chrome**，而不是把参考 PPT 总结成几个颜色和字号后重新绘制。

`references/deck-style.md` 在 v4 中属于 fallback style guide，不是 reference deck 的替代品。

## 4. QA & Delivery Gate

v4 QA 分四层：

1. **Semantic QA**：claim/source、日期数字、真实图片匹配、must-preserve 完整性；
2. **Structural QA**：文件结构、placeholder、对象、字体色值、silent data loss；
3. **Geometry QA**：overflow、overlap、out-of-bounds、alignment、crop、可读性；
4. **Visual QA**：基于真实 render 的层级、节奏、留白、均衡、图片质量与品牌一致性。

```text
PPTX → PDF / Slide PNG → Montage → Visual Review → Repair → Re-render
```

| Severity | 示例 | 交付 |
|---|---|---:|
| **P0 Critical** | 事实错误、错事实图、静默丢内容、placeholder、关键品牌违规 | ❌ |
| **P1 Major** | overflow、overlap、不可读、严重 crop、坏表/坏图、证据页缺来源 | ❌ |
| **P2 Minor** | 明显间距/对齐/一致性问题 | 原则上修复 |
| **P3 Cosmetic** | 可选精修 | ✅ |

**任何 P0 / P1 都阻断交付。**

## 5. Never Silent Loss

> 用户输入、已验证证据和 `must_preserve` 信息不得静默截断或丢失。

如果内容无法放入单页：

1. 更换 layout；
2. 压缩非关键表达但不改变含义；
3. 拆成 continuation slide；
4. 无法处理时明确失败。

Legacy renderer 已移除类似 `rows = rows[:10]` / `cols = min(cols, 6)` 的静默截断逻辑。

## 6. Domain Workflows

- **IP Collaboration**：[`references/ip-collab.md`](./references/ip-collab.md)
- **Celebrity**：[`references/celebrity.md`](./references/celebrity.md)
- **Campaign**：[`references/campaign.md`](./references/campaign.md)
- **Brand / Design Compliance**：[`references/design-spec.md`](./references/design-spec.md)
- **Research**：[`workflows/research.md`](./workflows/research.md)
- **Deck Production**：[`workflows/deck-production.md`](./workflows/deck-production.md)

## 7. Repository Structure

```text
brand-marketing-hub/
├── SKILL.md
├── README.md
├── CHANGELOG.md
├── workflows/
│   ├── research.md
│   ├── deck-production.md
│   └── runtime-adapters.md
├── schemas/
│   ├── evidence.schema.json
│   ├── asset.schema.json
│   ├── deck.schema.json
│   └── qa.schema.json
├── references/
│   ├── v4-architecture.md
│   ├── pptx-workflow.md
│   ├── deck-style.md
│   ├── ip-collab.md
│   ├── celebrity.md
│   ├── campaign.md
│   └── design-spec.md
├── assets/
│   ├── validate_v4_manifests.py
│   ├── validate_pptx.py
│   ├── md2pptx_vivo.py
│   └── tests/
└── .github/workflows/
    └── ci.yml
```

## 8. Validation

安装 QA / legacy 依赖：

```bash
python3 -m pip install -r assets/requirements.txt
```

验证 v4 manifests：

```bash
python3 assets/validate_v4_manifests.py \
  --evidence evidence/claims.json \
  --assets evidence/assets.json \
  --deck deck_spec.json
```

运行 regression tests：

```bash
python3 assets/tests/test_v4_contracts.py
python3 assets/tests/test_archetypes.py
python3 assets/tests/test_render_smoke.py
```

Legacy Markdown renderer：

```bash
python3 assets/md2pptx_vivo.py input.md output.pptx
python3 assets/validate_pptx.py output.pptx --md input.md --json
```

Legacy renderer **不是 v4 新项目默认路径**。

## 9. Brand Asset Boundary

内部品牌规范、版权资料与企业 master deck 不应在未经授权的情况下上传到公开仓库。

公开仓库保存方法论、schema、workflow、validator 与 fallback style rules；运行时通过用户提供或授权位置加载最新官方 VI、内部 master/reference deck、受限版权素材和内部业务数据。

品牌规范冲突时，优先级为：

> **官方原件 / VONE > retained reference deck（版式） > design-spec rules > fallback deck-style > renderer default**

## Version

Current Skill version: **v4.0.0**

- Release: [`v4.0.0`](https://github.com/blackkcold/brand-marketing-hub/releases/tag/v4.0.0)
- Changelog: [`CHANGELOG.md`](./CHANGELOG.md)
- CI: [GitHub Actions](https://github.com/blackkcold/brand-marketing-hub/actions)
