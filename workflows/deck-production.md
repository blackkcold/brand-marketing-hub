# v4.1 Presentation Production Workflow

This workflow is downstream of `workflows/source-to-deck.md`. It does not ingest source files directly.

## Core contract
`deck_spec.json` is the story/layout-intent source of truth. It is generated only after source ingestion, research and synthesis.

## Pipeline
1. Receive validated source/content/evidence/assets plus decision question.
2. Select **story_archetype** independently from **visual_template**.
3. Build slide intents/takeaways first; each slide answers one decision question.
4. Populate typed content blocks with unit_ids / claim_ids / asset_ids.
5. Generate exhaustive coverage manifest. Resolve every missing unit before render.
6. Resolve `brand/vivo/template-manifest.json`; prefer user-supplied/retained vivo reference/master when available.
7. Runtime selection: host-native Presentations/PowerPoint → artifact presentation runtime → PptxGenJS → legacy python-pptx.
8. Render editable PPTX.
9. Deterministic structural checks: placeholders, linked images, template font/palette rules, bounds.
10. Actual render to PDF/PNG + montage.
11. Semantic / coverage / embedded-assets / template-fidelity / structural / geometry / visual QA.
12. Repair the smallest failing scope and repeat until P0/P1 = 0.

## Layout selection
Choose among layout candidates using content type, density, image aspect ratio, hierarchy and the real reference/master. Never let a content archetype create a separate visual design family.

## Delivery
The final deck must be editable where technically possible, contain embedded factual images, retain required source meaning and visually read as one vivo deck rather than a collection of unrelated templates.
