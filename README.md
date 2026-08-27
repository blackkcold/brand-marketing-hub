# brand-marketing-hub

vivo 品牌营销工作包：包含 IP 联名、艺人评估、传播方案、品牌视觉合规，以及内容类型驱动的 Markdown → PPTX 输出流水线（v3.0）。

## 目录

- `SKILL.md`：skill 总入口与路由规则
- `references/`：模块工作流与 vivo 2026 deck 设计系统
- `assets/md2pptx_vivo.py`：vivo 企业风格 Markdown → PPTX renderer
- `assets/validate_pptx.py`：PPTX 结构与品牌输出验证器
- `assets/tests/`：视觉回归 fixture 与冒烟测试
- `references/pptx-workflow.md`：PPTX 能力、依赖、结构验证与视觉 QA 流程
- `assets/samples/v0.2_*.pptx`：历史样张；新版样张应通过模板重新生成
- `assets/vivo-design-spec/规范清单.md`：原始规范资料索引

## 使用

```bash
python3 -m pip install --user -r assets/requirements.txt
python3 assets/md2pptx_vivo.py input.md output.pptx
python3 assets/validate_pptx.py output.pptx --md input.md --json
python3 assets/tests/test_render_smoke.py
python3 assets/tests/test_archetypes.py
```

新版页面布局包括 `stats`、`framework`、`comparison`、`matrix`、`timeline`、`budget`、`collage`、`case-study`、`chart`。四套模板分别对应传播方案、艺人评估、IP 联名、调研洞察，不再共用同一套页面节奏。

PPTX 专项能力包括可编辑的柱状图、折线图和环图；完整依赖与 QA 流程见 `references/pptx-workflow.md`。

## 生成前必须确认模板

每次生成 PPTX 前先询问用户选择：`传播方案`、`艺人评估`、`IP 联名` 或 `自定义 .pptx/.potx`。不得默认套用模板。确认后再根据模板的叙事顺序和对应 profile 生成页面。

## 资料边界

原始 `assets/vivo-design-spec/` 资料不纳入 GitHub 仓库，避免上传大体积内部/版权文档。需要查阅时，按 `规范清单.md` 指向的本地资料进行核对。
