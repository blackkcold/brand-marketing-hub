# brand-marketing-hub v4.1 Architecture

## Product definition
A universal vivo Source-to-Deck engine: ingest arbitrary business files, research gaps, reorganize the material, and produce an editable visually consistent vivo presentation.

## Seven layers
1. **Source ingestion** — DOCX/XLSX/PPTX/PDF/CSV and connected sources → source inventory + normalized content units.
2. **Evidence** — user-source facts + verified external claims + factual assets.
3. **Coverage** — exhaustive source-unit disposition so no important input disappears silently.
4. **Story** — decision chain, story archetype and typed deck_spec blocks.
5. **Visual template** — vivo template manifest/reference/master, independent from story archetype.
6. **Presentation runtime** — native presentation / PowerPoint / artifact / PptxGenJS / legacy fallback.
7. **QA & revision** — coverage, semantic, embedded-assets, template fidelity, structural, geometry, visual, repair.

Domain modules (IP, celebrity, campaign, design-spec) are optional intelligence modules attached to layers 2–4. They never own rendering.

## Source-of-truth hierarchy
User instruction → source files/content units → verified evidence → coverage + deck_spec → visual template/reference → PPTX runtime projection → rendered QA artifacts.

## Key invariants
- every source unit has one disposition;
- story archetype != visual template;
- factual image != generated concept;
- linked web image != final embedded asset;
- domain method != renderer;
- structural pass != visual pass.
