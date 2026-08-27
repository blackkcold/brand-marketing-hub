# Changelog

本项目遵循 [Keep a Changelog](https://keepachangelog.com/zh-CN/1.1.0/) 与 [语义化版本](https://semver.org/lang/zh-CN/)。

## [Unreleased]

### Added
- （待添加）

## [v3.2.0] - 2026-08-27

### Added
- 内容类型驱动的 Markdown → PPTX 输出流水线（v3.2）：按显式 archetype 与内容密度生成，支持自适应框体、必要时拆页。
- 原生可编辑图表：柱状图、折线图、环图。
- 新增页面布局 archetype：`stats`、`framework`、`comparison`、`matrix`、`timeline`、`budget`、`collage`、`case-study`、`chart`。
- 四套独立模板：传播方案、艺人评估、IP 联名、调研洞察，不再共用同一套页面节奏。
- 新增调研洞察模板 `assets/template-research-insight.md` 与创作者品牌合作模板 `assets/creator-brand-partnership-vivo.md`。
- 新增 archetype 测试 `assets/tests/test_archetypes.py`。
- 新增 `references/pptx-workflow.md`：PPTX 能力、依赖、结构验证与视觉 QA 流程。
- 可选 XML 安全依赖：`defusedxml`、`lxml`。

### Changed
- 重构 `assets/md2pptx_vivo.py`：从统一蓝标题条 + bullet 改为 archetype 驱动渲染。
- 重构 `assets/validate_pptx.py`：品牌输出验证器升级。
- 升级传播方案、艺人评估、IP 联名三套模板。
- 更新 `references/deck-style.md` 设计系统说明。
- 生成产物（`output/`）与本地备份（`.SKILL.md.bak`）移出仓库跟踪。

### Fixed
- 生成前必须确认模板，避免默认套用错误模板。

---

[Unreleased]: https://github.com/blackkcold/brand-marketing-hub/compare/v3.2.0...HEAD
[v3.2.0]: https://github.com/blackkcold/brand-marketing-hub/releases/tag/v3.2.0
