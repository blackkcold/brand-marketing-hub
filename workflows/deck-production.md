# v4 Presentation Production Workflow

## Core contract
The source of truth is `deck_spec.json`, not Markdown. Markdown is an optional human-readable projection only.

## Pipeline
1. Brief Parser: resolve purpose, audience, decision, constraints, template/reference, brand profile, and must-preserve content.
2. Research & Evidence: run `workflows/research.md`; create evidence manifests.
3. Story Planner: produce slide intents and takeaways. Each slide must answer one question and contribute to the deck decision chain.
4. Deck Spec: validate against `schemas/deck.schema.json`.
5. Template Resolver: prefer a retained real reference/master deck or an OpenAI artifact template over reconstructing style from tokens.
6. Runtime selection:
   - ChatGPT/PowerPoint presentation capability when available and appropriate;
   - artifact/presentation native renderer;
   - PptxGenJS for deterministic CLI/server generation;
   - legacy `python-pptx` renderer only as fallback.
7. Render to PPTX.
8. Render slides to images/PDF and create a montage/contact sheet.
9. QA gates in order: semantic -> structural -> geometry -> visual.
10. Repair only failed slides, rerender, and repeat QA until no P0/P1 issues remain.

## Never-silent-loss rule
User content or verified evidence must never be silently truncated, dropped, clipped, or replaced. If content cannot fit:
1. choose a more suitable layout;
2. reduce nonessential prose without changing meaning;
3. split into continuation slides;
4. fail explicitly if none is possible.

## Layout selection
A slide stores multiple `layout_candidates`. The runtime should choose based on content type, density, image aspect ratio, hierarchy, and template fidelity rather than a single hard-coded `@layout`.

## QA severity
- P0 Critical: factual error, wrong factual image, missing verified content, unresolved placeholder, brand/logo violation, silent data loss.
- P1 Major: overflow, overlap, illegibility, materially bad crop, broken chart/table, missing citation on evidence slide.
- P2 Minor: spacing/alignment inconsistency that degrades polish.
- P3 Cosmetic: optional polish only.
P0/P1 block delivery. P2 should be repaired where feasible. P3 may ship with notes.
