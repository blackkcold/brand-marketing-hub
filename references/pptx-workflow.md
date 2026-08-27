# PPTX 能力与 QA 工作流

## 能力边界

本 skill 默认输出 PowerPoint 原生可编辑对象：文本框、表格、形状、图片、speaker notes，以及柱状图、折线图和环图。内容页使用显式 `@layout`，按内容密度调整框体和间距；无法在合理字号下容纳时自动拆页。

可由 PowerPoint 原生表达的图表不转成图片；只有原生不支持的复杂图形才使用图片。

## Markdown 图表语法

```markdown
@layout chart
@chart bar
| 候选 | 影响力 | 适配度 |
|---|---:|---:|
| A | 80 | 72 |
| B | 65 | 88 |
```

支持 `bar`、`line`、`doughnut`，图表指令必须紧随 Markdown 表格。

## 依赖

```bash
python3 -m pip install --user -r assets/requirements.txt
```

- `python-pptx`：创建和读取可编辑 PPTX；
- `Pillow`：图片尺寸、比例和裁切计算；
- `lxml`：PPTX/XML 辅助处理；
- `defusedxml`：XML 变换时的安全解析。

## QA 命令

```bash
python3 assets/md2pptx_vivo.py input.md output.pptx
python3 assets/validate_pptx.py output.pptx --md input.md --json
python3 assets/tests/test_render_smoke.py
python3 assets/tests/test_archetypes.py
soffice --headless --convert-to pdf --outdir rendered output.pptx
```

检查顺序：文字溢出 → 图片比例/裁切 → 元素重叠 → 色彩对比度 → 页脚/来源 → 封面和末页。

## 安全与版本

- 用户提供的 `.pptx/.potx` 只读分析，不原地覆盖；
- 输出使用 `output/vX.Y_description/` 版本目录；
- 模板文件只迁移视觉参数，不引入复杂母版；
- 结构 validator 通过后再进行 PDF 视觉复核。
