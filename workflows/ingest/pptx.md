# PPTX Ingestion

Treat content and style as separate roles.

## Default
A supplied PPTX is **content-source by default**. Extract and preserve:
- slide order;
- text;
- speaker notes;
- tables;
- native chart data where accessible;
- embedded images;
- presentation metadata.

Do **not** assume the source deck's visual style should carry into the new report.

## Explicit style reference
Only treat a PPTX as `style-reference` when the user explicitly identifies it as the desired visual/master reference.

Local fallback:
```bash
python assets/ingest_sources.py old-deck.pptx \
  --out-dir source_bundle \
  --role old-deck.pptx=style-reference
```

When style-reference is explicit, the host-native presentation runtime should additionally inspect and preserve masters, layouts, theme, recurring chrome, typography and geometry. The local parser records reference metadata but is not a substitute for native master-preserving editing.

## Mixed role
Use `mixed` only when the same PPTX is explicitly intended to provide both report content and visual reference.
