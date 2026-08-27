# v4 Presentation Runtime Adapters

## Selection principle
brand-marketing-hub owns domain intelligence, evidence, story structure and brand QA. It should not reimplement a full presentation engine when a native presentation runtime is available.

## 1. ChatGPT Presentations / PowerPoint
Preferred for interactive creation or editing of real decks, especially when the user supplies a reference deck/template. Preserve masters, layouts, theme, recurring chrome, typography and editable objects. Use reference-backed templates where possible.

## 2. OpenAI artifact/presentation runtime
Preferred for native presentation generation in ChatGPT when supported. Feed the validated deck spec, evidence pack and retained template/reference.

## 3. PptxGenJS
Preferred deterministic CLI/server fallback for new generation where native OpenAI presentation runtime is unavailable. Build adaptive layout helpers and render/measure loops; do not port the current fixed-coordinate renderer unchanged.

## 4. Legacy python-pptx
Compatibility fallback only. It may be used for readback, validation, lightweight patches, extraction and legacy generation. New visual capabilities should not default here.

## Runtime output contract
Every runtime must produce:
- editable PPTX where technically possible;
- a render for visual QA;
- QA report;
- provenance link between slide -> claim -> source and slide -> asset;
- no P0/P1 outstanding issues.

## Template fidelity
When a retained reference/master deck exists, clone/import and reuse its native layouts rather than reconstructing them from a prose style guide.
