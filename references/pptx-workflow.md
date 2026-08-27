# PPTX / Slides 能力与 QA 工作流（v4.1）

## 定位
演示文稿生产是 `workflows/source-to-deck.md` 的下游。brand-marketing-hub 负责 Source Ingestion、Evidence、Coverage、Story、vivo Brand QA；具体 PPTX 创建/编辑优先交给宿主环境最强的 presentation runtime。

## Runtime 顺序
1. Host-native Presentations / PowerPoint：真实 reference/master、交互编辑、局部修订优先。
2. Artifact presentation runtime：结构化生成、master/theme-aware 编辑。
3. PptxGenJS：已实现的 deterministic fallback。
4. `assets/md2pptx_vivo.py`：legacy compatibility only。

详见 `workflows/runtime-adapters.md` 与 `runtime/`。

## v4.1 Source of Truth
```text
source_inventory.json + content_units.json
        ↓
evidence.json + assets.json
        ↓
deck_spec.json + coverage.json
        ↓
vivo template/reference
        ↓
PPTX
        ↓
rendered QA artifacts
```

PPTX/Markdown 只是 runtime projection，不是事实数据库。

## Manifest validation
```bash
python assets/validate_v4_manifests.py \
  --sources source_inventory.json \
  --content content_units.json \
  --evidence evidence.json \
  --assets assets.json \
  --coverage coverage.json \
  --deck deck_spec.json \
  --template brand/vivo/template-manifest.json
```

重点检查 provenance、exhaustive coverage、exact/semantic preservation、factual asset safety、story/visual template consistency。

## PPTX deterministic checks
```bash
python assets/validate_pptx.py output.pptx \
  --template-manifest brand/vivo/template-manifest.json \
  --json
```

检查 unresolved placeholder、out-of-bounds、external linked image、vivo branding signal、template-aware fonts/palette 和 legacy content-loss hazards。**外链事实图片属于 P0。**

## Actual render
Visual QA 必须检查真实 render：
```bash
python assets/render_pptx.py output.pptx --out-dir render
```

本地 fallback 生成 PDF、per-slide PNG 和 montage；宿主存在原生 presentation render 时优先使用宿主能力。

## QA 七层 Gate
1. Coverage
2. Semantic
3. Embedded Assets
4. Template Fidelity
5. Structural
6. Geometry
7. Visual

P0/P1 未清零禁止交付。

## Template fidelity
1. 用户明确提供的最新 vivo master/reference；
2. retained `artifact-template-vivo-internal-report`（若存在）；
3. `brand/vivo/template-manifest.json` + `layout-map.json`；
4. `references/deck-style.md` fallback；
5. legacy renderer。

Story archetype 与 visual template 必须分离。Campaign/IP/Research 不得因此各自形成不同视觉系统。

## Revision
已有 PPTX 的小范围修改应先 inspect，再 patch slide/object；未修改页面保持不变。内容语义变化时同步更新 deck_spec/coverage/evidence/assets，最终重新跑 P0/P1 gate。

## Legacy regression
以下仅用于兼容测试：
```bash
python assets/md2pptx_vivo.py input.md output.pptx
python assets/tests/test_render_smoke.py
python assets/tests/test_archetypes.py
```
