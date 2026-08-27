# brand-marketing-hub v4 Architecture

## Five layers
1. Domain workflows: IP collaboration, celebrity, campaign, research, design-spec.
2. Evidence layer: claims, sources, factual assets.
3. Story layer: decision chain and `deck_spec.json`.
4. Presentation runtime layer: native presentation runtime / artifact runtime / PptxGenJS / legacy Python fallback.
5. QA layer: semantic, structural, geometry and visual review with repair loop.

## Source-of-truth hierarchy
User constraints and supplied source material -> verified evidence manifests -> deck_spec.json -> runtime projection (PPTX/Markdown) -> rendered QA artifacts.

## Why v4
v3.x coupled content, evidence and layout in Markdown and treated structural validation as presentation quality. v4 separates the concerns so factual reliability, template fidelity and visual quality can be tested independently.

## Compatibility
Existing Markdown templates and `md2pptx_vivo.py` remain available as legacy input/output paths. New work should use v4 manifests and adapters whenever the host environment supports them.
