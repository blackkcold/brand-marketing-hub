# PptxGenJS fallback runtime

This is the deterministic CLI fallback for v4.1. It is not the preferred path when a real vivo reference/master can be preserved by a host-native presentation runtime.

Usage:

```bash
npm install
node runtime/pptxgenjs/render.js deck_spec.json output.pptx assets.json brand/vivo/template-manifest.json
```

The fallback supports typed deck blocks and keeps text/tables/charts editable. It deliberately fails visually soft rather than inventing factual imagery; missing verified images render as an explicit placeholder and therefore fail final QA.
