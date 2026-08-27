# v4.1 Presentation Runtime Adapters

## Selection principle
brand-marketing-hub owns source understanding, research, synthesis, deck semantics and vivo QA. Rendering should use the strongest available presentation engine.

## 1. Host-native Presentations / PowerPoint
Preferred for reference/master-native generation and revision. See `runtime/openai-presentations.md` and `runtime/powerpoint.md`.

## 2. Artifact presentation runtime
Preferred for structured master/theme-aware generation or precise edits when available. See `runtime/artifact-tool.md`.

## 3. PptxGenJS
Executable deterministic fallback:
```bash
npm install
node runtime/pptxgenjs/render.js deck_spec.json output.pptx assets.json brand/vivo/template-manifest.json
```
It is a fallback, not a substitute for a real vivo master/reference.

## 4. Legacy python-pptx
Compatibility only: readback, validation, lightweight patching and old Markdown generation.

## Runtime output contract
Every runtime must yield:
- editable PPTX where possible;
- no silent content loss;
- all factual images embedded;
- source/claim/asset provenance preserved outside or inside the deck metadata;
- a real render for visual QA;
- no P0/P1 outstanding.

## Revision
Prefer inspecting and patching existing slides/objects. Do not regenerate the whole deck for a local revision when stable slide/object editing is available.
