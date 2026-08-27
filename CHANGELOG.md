# Changelog

本项目遵循 [Keep a Changelog](https://keepachangelog.com/zh-CN/1.1.0/) 与 [语义化版本](https://semver.org/lang/zh-CN/)。

## [Unreleased]

### Added
- v4.1.2 hardening：PPTX grouped-shape 递归摄取、真实 canvas / text overlap / tiny-font deterministic QA。
- Render 页序自然排序，确保 10+ 页 montage/QA 顺序正确。
- CI 增加 runtime 高危依赖审计。

### Changed
- fallback runtime 切换到经 CI 验证的 `@lofcz/pptxgenjs 4.1.17`，并保持 PptxGenJS API 兼容路径。
- Release 改为仅在 main 的 CI 成功后执行；未存在对应版本 Changelog section 时自动跳过。
- deck block / chart / comparison / timeline schema 改为严格 typed contract。

### Fixed
- 修复 factual proof asset 仅写 `properties`、实际上未强制 source_ref/source_url 存在的 JSON Schema 漏洞。
- 修复 embedded asset 只检查 local_path 字符串、不检查文件实际存在的问题。
- 修复 duplicate source/unit/claim/asset/slide/block ID 被 set 聚合静默吞掉的问题。
- 修复 block 内 claim/asset/source 引用未进入 cross-manifest validation 的问题。
- 修复 table 行宽和 chart labels/values 数量不一致仍可通过验证的问题。
- 修复相对 asset path 依赖当前工作目录，以及图片按固定框拉伸变形的问题。

## [v4.1.1] - 2026-08-27

### Added
- Universal Source-to-Deck 主流程：DOCX / XLSX / PPTX / PDF / CSV → source inventory / content units → evidence / synthesis → deck spec / coverage → vivo PPTX → render / QA / repair。
- 新增 `source.schema.json`、`content-unit.schema.json`、`coverage.schema.json`、`template.schema.json`，补齐输入摄取、内容保真与模板解析契约。
- 新增 Word/Excel/PPT/PDF/CSV fallback ingestion，并保留表格、公式、PPT notes/chart、嵌入图片与来源 locator。
- 新增 exhaustive Coverage Gate，逐个 `UNIT` 追踪 body / appendix / summarized / intentionally-excluded / missing。
- 新增 vivo `template-manifest.json`、`layout-map.json` 与 story archetype registry，将故事结构与视觉模板彻底分离。
- 新增可执行 PptxGenJS fallback runtime、PDF/PNG/montage 真实渲染与 render integrity QA。
- 新增 Source-to-Deck、Coverage、PPTX no-silent-loss、真实渲染等 CI 回归测试。

### Changed
- `SKILL.md` 主目标升级为通用 vivo Source-to-Deck，而不是 IP/艺人/Campaign 模板路由或 Markdown → PPTX。
- IP、艺人、Campaign、VI 模块降为可选领域增强层，不再拥有 renderer 或独立视觉模板。
- PPTX 输入默认作为 content-source；只有显式声明时才作为 style-reference / mixed。
- Presentation runtime 优先级调整为 native Presentations / PowerPoint → Artifact → PptxGenJS → legacy python-pptx。
- 真实 vivo master/reference 优先于 fallback style guide；partner 色只允许作为已验证 accent。
- Research 增加 confidentiality gate、entity resolution、官方优先、freshness 与 factual-image verification。

### Fixed
- 修复表格/比较/时间线/长正文在 renderer 中可能被静默截断的问题，改为 continuation slides / chunking。
- 修复 manifest schema 根目录解析错误。
- 修复任意图片都可能被当作 vivo logo 的弱校验。
- 修复最终 PPTX 可能保留 external linked image 的风险；外链事实图片现在为 P0。
- 修复源文件图片只读文字、不保留图片本体的问题。
- 修复 source manifest 记录本地绝对路径的隐私/可移植性问题。
- 修复 release workflow 在版本修改中途提前创建 tag 的问题。

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

[Unreleased]: https://github.com/blackkcold/brand-marketing-hub/compare/v4.1.1...HEAD
[v4.1.1]: https://github.com/blackkcold/brand-marketing-hub/compare/v4.0.0...v4.1.1
[v4.0.0]: https://github.com/blackkcold/brand-marketing-hub/compare/v3.2.0...v4.0.0
[v3.2.0]: https://github.com/blackkcold/brand-marketing-hub/releases/tag/v3.2.0
