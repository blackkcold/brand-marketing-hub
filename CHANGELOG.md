# Changelog

本项目遵循 [Keep a Changelog](https://keepachangelog.com/zh-CN/1.1.0/) 与 [语义化版本](https://semver.org/lang/zh-CN/)。

## [Unreleased]

### Added
- （待添加）

## [v4.0.0] - 2026-08-27

### Added
- v4 五层架构：Domain / Evidence / Story / Presentation Runtime / QA。
- `workflows/research.md`：官方优先的 Research & Evidence 工作流。
- `workflows/deck-production.md`：deck_spec 驱动的演示文稿生产与 repair loop。
- `workflows/runtime-adapters.md`：ChatGPT Presentations / PowerPoint / Artifact / PptxGenJS / legacy Python 的 runtime 选择协议。
- `schemas/evidence.schema.json`、`asset.schema.json`、`deck.schema.json`、`qa.schema.json`。
- `references/v4-architecture.md`。
- QA severity：P0/P1 阻断交付，不再按问题页比例放行。

### Changed
- `deck_spec.json` 成为新演示文稿 source of truth；Markdown 降为 human-readable projection / legacy input。
- reference/master/template-native workflow 成为默认，固定坐标重画模板不再是首选。
- `python-pptx` renderer 降级为 legacy fallback / readback / validation。
- 外部事实、真实案例、产品与事实图片进入 deck 前必须建立可追溯 evidence/asset 记录。

### Fixed
- 禁止 legacy renderer 静默截断超出容量的表格行列。
- validator 改为基于 severity 的 delivery gate。

## [v3.2.0] - 2026-08-27

### Added
- 内容类型驱动的 Markdown → PPTX 输出流水线（v3.2）：按显式 archetype 与内容密度生成。
- 原生可编辑柱状图、折线图、环图与多种页面 archetype。

### Changed
- 从统一蓝标题条 + bullet 改为 archetype 驱动渲染。

---

[Unreleased]: https://github.com/blackkcold/brand-marketing-hub/compare/v4.0.0...HEAD
[v4.0.0]: https://github.com/blackkcold/brand-marketing-hub/compare/v3.2.0...v4.0.0
[v3.2.0]: https://github.com/blackkcold/brand-marketing-hub/releases/tag/v3.2.0
