# brand-marketing-hub v4

vivo / 品牌营销工作包。v4 将项目从 **Markdown → 固定坐标 PPTX renderer** 升级为：

**Research → Evidence → Story → Deck Spec → Presentation Runtime → QA / Repair**

## v4 重点

- 官方优先的研究与事实校验；
- claim → source → slide 的 Evidence Manifest；
- factual image → source → slide 的 Asset Manifest；
- `deck_spec.json` 作为演示文稿 source of truth；
- 优先复用真实 reference/master PPTX，而不是从文字规则重画模板；
- 优先调用 ChatGPT Presentations / PowerPoint / Artifact presentation runtime；
- PptxGenJS 作为 CLI/server 新 renderer 的推荐方向；
- `python-pptx` renderer 降级为 legacy fallback / readback / validation；
- Semantic / Structural / Geometry / Visual 四层 QA；
- P0/P1 阻断交付；
- 明确禁止 silent truncation。

## 目录

```text
SKILL.md
workflows/
  research.md
  deck-production.md
  runtime-adapters.md
schemas/
  evidence.schema.json
  asset.schema.json
  deck.schema.json
  qa.schema.json
references/
  v4-architecture.md
  deck-style.md
  ip-collab.md
  celebrity.md
  campaign.md
  design-spec.md
assets/
  md2pptx_vivo.py        # legacy renderer
  validate_pptx.py       # legacy/structural QA
  tests/
```

## Runtime 策略

1. ChatGPT Presentations / PowerPoint：优先用于真实模板驱动的创建与修改。
2. OpenAI artifact/presentation runtime：宿主支持时优先。
3. PptxGenJS：CLI/server deterministic fallback 的目标实现。
4. python-pptx：兼容、读取、验证、轻量 patch 与 legacy generation。

详见 `workflows/runtime-adapters.md`。

## Legacy 使用

```bash
python3 -m pip install --user -r assets/requirements.txt
python3 assets/md2pptx_vivo.py input.md output.pptx
python3 assets/validate_pptx.py output.pptx --md input.md --json
python3 assets/tests/test_render_smoke.py
python3 assets/tests/test_archetypes.py
python3 assets/tests/test_v4_contracts.py
```

Legacy renderer 不得静默截断内容；超出当前表格/布局容量时应明确失败或由上游拆页。

## 资料边界

内部/版权原件不应无授权上传到公开仓库。运行时需要使用品牌规范或真实 master/reference deck 时，通过用户提供、授权的内部位置或宿主 artifact-template/reference 能力加载。
