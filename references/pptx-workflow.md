# PPTX / Slides 能力与 QA 工作流（v4）

## 定位

v4 不再把 `python-pptx` Markdown renderer 当作默认演示文稿引擎。brand-marketing-hub 负责 Research、Evidence、Story、Brand QA；具体 PPTX 创建/编辑优先交给宿主环境最强的 presentation runtime。

运行时选择见 `../workflows/runtime-adapters.md`，完整流程见 `../workflows/deck-production.md`。

## 默认顺序

1. ChatGPT Presentations / PowerPoint：有真实 reference/master deck 或需要交互编辑时优先。
2. OpenAI artifact/presentation runtime：宿主支持时优先。
3. PptxGenJS：CLI/server deterministic generation 的推荐 fallback。
4. `assets/md2pptx_vivo.py`：legacy compatibility only。

## v4 Source of Truth

- `evidence/claims.json` / `sources.json`：事实层。
- `evidence/assets.json`：事实图片/素材层。
- `deck_spec.json`：故事与页面意图层。
- PPTX/Markdown：runtime projection，不是事实数据库。

对应 schema：
- `schemas/evidence.schema.json`
- `schemas/asset.schema.json`
- `schemas/deck.schema.json`
- `schemas/qa.schema.json`

可用：
```bash
python3 assets/validate_v4_manifests.py --evidence evidence/claims.json --assets evidence/assets.json --deck deck_spec.json
```

## QA 四层

### 1. Semantic QA
检查：
- claim 是否有来源；
- source 是否支持 claim；
- 时间敏感信息是否检查 freshness；
- 用户 must-preserve 信息是否完整；
- factual image 是否与对应产品/合作/人物/事件匹配；
- deck 内数字、名称、日期是否一致。

### 2. Structural QA
检查：
- 文件可打开；
- slide/object 数量合理；
- placeholder 清零；
- chart/table/image 存在；
- 字体/色值/品牌参数；
- 没有 silent data loss。

### 3. Geometry QA
检查：
- overflow；
- overlap；
- out of bounds；
- margin / alignment / baseline；
- 图片 crop；
- 表格/图表可读性。

### 4. Visual QA
必须基于实际 render，而不是只读 PPTX XML：
- 导出 PDF 或逐页 PNG；
- 生成 montage/contact sheet；
- 检查视觉层级、节奏、留白、均衡、图片质量、品牌一致性；
- 失败页进入 repair loop 后重新 render。

## QA Severity

- P0：事实错误、错事实图、静默丢内容、placeholder、关键品牌/logo错误。
- P1：overflow、overlap、不可读、严重 crop、坏图表/表格、证据页缺来源。
- P2：明显影响精修度但不影响事实与阅读的间距/对齐/风格问题。
- P3：可选 cosmetic polish。

**任何 P0/P1 都阻断交付。禁止按问题页占比放行。**

## Legacy 命令

```bash
python3 assets/md2pptx_vivo.py input.md output.pptx
python3 assets/validate_pptx.py output.pptx --md input.md --json
python3 assets/tests/test_render_smoke.py
python3 assets/tests/test_archetypes.py
python3 assets/tests/test_v4_contracts.py
```

Legacy renderer 超过支持容量时必须失败或由上游拆页，不得通过 `rows[:N]` / `cols[:N]` 静默截断。

## 安全与模板

- 用户提供的 PPTX/POTX 默认只读分析，不原地覆盖；
- 真实 reference/master 优先 clone/import；
- 内部/版权模板不自动上传到公开仓库；
- 本地绝对路径不得写入公共 skill；
- reference template 控制未被用户明确覆盖的 layout / master / typography / geometry。
