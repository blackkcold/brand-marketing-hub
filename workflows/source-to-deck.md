# v4.1 Universal Source-to-Deck Workflow

## Product goal
Turn one or more user-provided documents, spreadsheets, presentations or PDFs into a researched, reorganized, editable, visually consistent vivo presentation without losing source meaning.

## Mandatory pipeline
1. **Source inventory** — classify every input as content-source, data-source, style-reference, evidence-source or mixed. Never assume a PPT is both content and style unless explicitly intended.
2. **Ingestion** — normalize source material into `source_inventory.json` and `content_units.json`. Prefer host-native Files/Documents/Spreadsheets/Presentations/PDF capabilities; local fallback is `assets/ingest_sources.py`.
3. **Preservation policy** — mark units exact / semantic / summarize-ok / optional. User-specified “must keep” content is exact or semantic, never optional.
4. **Confidentiality gate** — classify sources public/internal/confidential/restricted before web research. Never put confidential project names, unreleased products, prices, budgets, personal data or internal identifiers into public search queries.
5. **Research planner** — identify information gaps, entity ambiguities, dates that need freshness checks and factual images that need verification. Run `workflows/research.md`.
6. **Synthesis** — run `workflows/synthesis.md`; separate source facts, verified external facts and analyst inference.
7. **Story architecture** — select a story archetype based on the decision task. Story template and visual template are separate dimensions.
8. **Deck spec** — create typed `deck_spec.json`; every factual block carries unit_ids and/or claim_ids. Every real image block carries asset_id.
9. **Coverage manifest** — create `coverage.json`; every content unit must have exactly one disposition. Run `assets/init_coverage.py` and then review/resolve every missing or excluded item.
10. **Template resolution** — use `brand/vivo/template-manifest.json`. If the user supplies a vivo reference/master, it becomes the preferred visual reference. Do not treat content-source PPT styling as authoritative unless it is explicitly a style reference.
11. **Render** — native ChatGPT Presentation/PowerPoint/artifact runtime first; PptxGenJS deterministic fallback; legacy python-pptx only for compatibility.
12. **QA** — validate manifests, verify all images are embedded, render actual slide PNGs, create montage, run semantic/coverage/template/geometry/visual review.
13. **Repair** — fix only failing slides/objects where possible. Re-run coverage and render QA. Delivery is blocked until P0/P1 = 0.

## Required deliverables
- editable `.pptx`
- `source_inventory.json`
- `content_units.json`
- `evidence.json`
- `assets.json`
- `deck_spec.json`
- `coverage.json`
- `qa.json`
- rendered slide images or PDF + montage used for QA

## Non-negotiable rules
- No silent source loss.
- No generated image presented as a real product/collaboration/person/event.
- No linked web image in final PPTX; factual images must be materialized and embedded.
- No invented number to complete a template.
- No unverified color/logo/logo lockup when a real brand asset is available.
- No full-deck regeneration for a small revision when slide/object-level editing is available.
