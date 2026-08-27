# brand-marketing-hub

[![CI](https://github.com/blackkcold/brand-marketing-hub/actions/workflows/ci.yml/badge.svg)](https://github.com/blackkcold/brand-marketing-hub/actions/workflows/ci.yml)
[![Release](https://img.shields.io/github/v/release/blackkcold/brand-marketing-hub)](https://github.com/blackkcold/brand-marketing-hub/releases)
[![Skill Version](https://img.shields.io/badge/skill-v4.1.1-1E46E6)](./SKILL.md)

**vivo Universal Source-to-Deck Skill**

> 任意文档 / 表格 / PPT / PDF → 深度理解与安全研究 → 内容重组 → vivo 一致性可编辑 PPT → Render / QA / Repair

v4.1 将项目从“品牌营销模块 + Markdown/PPTX renderer”重构为通用的 **Source-to-Deck production system**。IP 联名、艺人、Campaign、VI 仍保留，但只是可选领域增强模块，不再拥有 PPT 生成路径。

## Architecture

```text
DOCX / XLSX / PPTX / PDF / CSV / connected sources
                         │
                         ▼
                Source Ingestion
          source_inventory.json
            content_units.json
                         │
                         ▼
        Preservation + Confidentiality
                         │
                         ▼
             Research & Evidence
               evidence.json
                assets.json
                         │
                         ▼
                   Synthesis
                         │
                         ▼
               Story Archetype
                         +
               vivo Visual Template
                         │
                         ▼
                deck_spec.json
                coverage.json
                         │
                         ▼
     Presentations / PowerPoint / Artifact
                or PptxGenJS
                         │
                         ▼
                 Editable PPTX
                         │
                         ▼
          PDF / PNG / Montage Render
                         │
                         ▼
 Coverage / Semantic / Embedded Assets
 Template / Structural / Geometry / Visual QA
                         │
                         ▼
                      Repair
```

## What v4.1 fixes

| Risk | v4.1 |
|---|---|
| Word/Excel/PPT/PDF only loosely treated as “brief” | Universal ingestion + locators |
| AI summarization can silently drop source content | Exhaustive Coverage Manifest |
| User-file evidence has no URL | UNIT-based provenance |
| Story template and visual template mixed | Separate `story_archetype` / `visual_template` |
| Domain module routes back to Markdown renderer | Domain modules are analysis-only |
| PptxGenJS only documented | Executable fallback runtime included |
| Visual QA only described | Executable PDF/PNG/montage renderer included |
| Linked online images can leak into PPT | External image relationship = P0 |
| Any picture could count as vivo logo | Known embedded wordmark verification |
| Fixed 微软雅黑 / fixed palette regardless of reference | Template-aware font/palette contract |
| Partner-brand accent flagged incorrectly | Verified partner accent policy |
| Release tag can be created halfway through refactor | Changelog-gated release workflow |
| CI only checks legacy renderer | Multi-format ingestion + coverage + PPTX + real render CI |

## 1. Universal Source Ingestion

Preferred execution uses host-native file capabilities. Local fallback:

```bash
python assets/ingest_sources.py   brief.docx data.xlsx old-deck.pptx appendix.pdf   --out-dir source_bundle
```

Outputs:

- `source_inventory.json`
- `content_units.json`

Supported fallback inputs:

- DOCX
- XLSX
- PPTX
- PDF
- CSV

The schemas preserve file identity and locators such as paragraph, table, sheet/range, slide/shape and PDF page.

## 2. Preservation & Coverage

Content units are marked:

- `exact`
- `semantic`
- `summarize-ok`
- `optional`

Build an initial coverage manifest:

```bash
python assets/init_coverage.py   --content content_units.json   --deck deck_spec.json   --out coverage.json
```

Every source unit must have exactly one disposition. Missing source content is not silently ignored.

## 3. Research & Security

Research runs after ingestion so the system knows what is already present.

Before public web search, inputs are classified as:

- public
- internal
- confidential
- restricted

Confidential/internal identifiers are sanitized before search. Unreleased product names, internal budgets, contracts, personal data and internal project codes must not be copied into public search queries.

Research includes:

- question planning
- entity resolution
- official-first sourcing
- local-language query expansion
- freshness checking
- conflict handling
- factual image verification
- explicit stop criteria

## 4. Evidence & Real Assets

Claims can point to:

```text
UNIT-xxxx  # user source file
SRC-xxx    # external research source
```

Factual assets record:

- subject entity
- collaboration
- product model
- visual role
- source
- semantic verification
- embedding requirement

Generated concepts cannot be treated as factual proof.

## 5. Story vs Visual Template

These are independent:

### Story Archetype
`research / strategy / campaign / brand-partnership / business-review / project-update / proposal / custom`

### Visual Template
Default: `vivo-house`

Defined by:

- `brand/vivo/template-manifest.json`
- `brand/vivo/layout-map.json`

A real authorized vivo reference/master always takes priority over fallback style rules.

## 6. Presentation Runtime

Priority:

1. Host-native Presentations / PowerPoint
2. Artifact presentation runtime
3. PptxGenJS deterministic fallback
4. Legacy python-pptx

PptxGenJS fallback:

```bash
npm install
node runtime/pptxgenjs/render.js   deck_spec.json output.pptx assets.json brand/vivo/template-manifest.json
```

Legacy Markdown renderer remains only for regression/compatibility.

## 7. Real Render QA

```bash
python assets/render_pptx.py output.pptx --out-dir render
```

Produces:

- PDF
- per-slide PNGs
- `montage.png`

Render integrity gate:
```bash
python assets/validate_render.py render --json
```

Visual QA then inspects the actual rendered slides, not only PPTX XML.

## 8. PPTX Delivery Gate

```bash
python assets/validate_pptx.py output.pptx   --template-manifest brand/vivo/template-manifest.json   --json
```

Delivery-blocking checks include:

- external linked images
- placeholders
- out-of-bounds shapes
- missing/incorrect brand embedding signals
- template-aware fonts/palette
- legacy data-loss hazards

## 9. Full Manifest Validation

```bash
python assets/validate_v4_manifests.py   --sources source_inventory.json   --content content_units.json   --evidence evidence.json   --assets assets.json   --coverage coverage.json   --deck deck_spec.json   --template brand/vivo/template-manifest.json
```

Cross-manifest checks include provenance, orphan IDs, coverage completeness, exact-content preservation, factual-asset safety and template consistency.

## 10. Revision Workflow

Existing PPT changes should be slide/object-level when possible:

- inspect existing deck;
- patch only affected objects/slides;
- preserve untouched slides;
- update evidence/coverage only when meaning changes;
- rerender affected slides and neighboring slides;
- rerun full P0/P1 gate before delivery.

## 11. Repository Structure

```text
SKILL.md
workflows/
  source-to-deck.md
  ingest/
  research.md
  synthesis.md
  deck-production.md
  visual-qa.md
  revision.md
schemas/
  source.schema.json
  content-unit.schema.json
  evidence.schema.json
  asset.schema.json
  coverage.schema.json
  deck.schema.json
  template.schema.json
  qa.schema.json
brand/vivo/
  template-manifest.json
  layout-map.json
  README.md
runtime/
  openai-presentations.md
  powerpoint.md
  artifact-tool.md
  pptxgenjs/
assets/
  ingest_sources.py
  init_coverage.py
  validate_v4_manifests.py
  validate_pptx.py
  render_pptx.py
  md2pptx_vivo.py   # legacy only
tests/
.github/workflows/
```

## 12. CI

CI now performs:

1. v4.1 architecture contract checks;
2. generated DOCX/XLSX/PPTX/PDF/CSV ingestion;
3. exhaustive coverage gate test;
4. legacy regression tests;
5. PptxGenJS fallback generation;
6. PPTX validation;
7. LibreOffice real rendering;
8. PDF + PNG + montage existence checks.

## vivo Template Boundary

The public repository does not commit confidential internal master decks. Runtime order:

1. latest authorized user-supplied vivo master/reference;
2. retained `artifact-template-vivo-internal-report` if available;
3. public `brand/vivo` fallback system.

See `brand/vivo/README.md`.

## Version

Current Skill version: **v4.1.1**
