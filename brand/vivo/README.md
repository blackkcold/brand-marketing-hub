# vivo Presentation Template Resolution

The public repository intentionally does **not** embed a confidential/internal vivo master deck. Do not promote the legacy generated sample PPTX files under `assets/samples/` to canonical template status.

## Runtime order
1. User-supplied latest vivo master/reference PPTX.
2. Retained personal/workspace artifact template named `artifact-template-vivo-internal-report`, if available.
3. `template-manifest.json` + `layout-map.json` fallback system.
4. `references/deck-style.md` fallback guidance.
5. Legacy Markdown renderer only for compatibility.

## Registering a canonical retained template
When an authorized, representative vivo PPTX is available:
- create/update the reusable presentation template using the host Template Creator capability;
- retain the original PPTX, not a screenshot or reconstructed copy;
- verify its preview and native slide structures;
- keep the skill/display name stable;
- update `template-manifest.json` only if the retained template slug changes.

The canonical reference should include representative cover, section, content, table/chart, image-led, comparison, recommendation and appendix layouts.

## Security
Do not commit confidential master decks, internal-only imagery or proprietary fonts into this public repository unless explicitly authorized.
