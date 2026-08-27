# v4.1 Visual QA

Visual QA is mandatory and must inspect actual renders.

## Required checks
- hierarchy: takeaway is visible in 3 seconds;
- readable minimum text size and line length;
- no overlap, clipping or out-of-bounds;
- consistent margins, title baselines and section rhythm;
- tables/charts remain editable and legible;
- image crop shows the intended product/object;
- factual image identity matches asset manifest;
- vivo master/chrome/template is consistent across all slides;
- partner accent colors are verified rather than guessed;
- no placeholder or obvious synthetic/incorrect product image;
- no slide looks like an unrelated template family.

## Render
Host-native presentation render is preferred. Local fallback: `assets/render_pptx.py deck.pptx --out-dir render`.

## Repair
Log issue by slide_id and object where possible. Repair the smallest possible scope, rerender the affected slide/deck, and recheck. P0/P1 block delivery.
