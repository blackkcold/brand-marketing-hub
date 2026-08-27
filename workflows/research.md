# v4 Research & Evidence Workflow

## Objective
Research is a mandatory upstream stage for any deck containing external facts, cases, rankings, dates, collaboration claims, product facts, prices, market data, or factual imagery. Slides are never the research database.

## Workflow
1. Parse user-provided facts into provisional claims. Mark exact user-provided wording as `user-provided` and preserve it where requested.
2. Identify claims that require external verification or freshness checks.
3. Search official sources first: brand/company official site, official newsroom, official product/campaign page, official social/channel, regulator or public institution.
4. When the claim is material to a recommendation, use a second independent authoritative source where practical.
5. Record every usable claim in `evidence/claims.json` and every source in `evidence/sources.json` using `schemas/evidence.schema.json`.
6. Report conflicts rather than silently choosing one source. Mark unresolved claims `uncertain`; do not present them as fact.
7. Research factual imagery separately. Product/collaboration/case-study slides require traceable real assets where available.
8. Download/embed only assets permitted by the execution environment and record them in `evidence/assets.json`.
9. Generated images may be used only for concept/creative exploration. They may not be used as factual evidence of a real product, collaboration, person, event, or historical design.
10. Only verified/user-provided claims may enter the final deck. Uncertain claims require explicit caveat or exclusion.

## Source priority
official > partner-official > authority/academic > major-media > other sources.

## Freshness
For changing subjects, check publication/update date and record `retrieved_at`. A source being official does not exempt it from freshness review.

## Evidence gate
Before story planning:
- every material factual claim has at least one source_id;
- recommendation-critical claims have adequate corroboration;
- factual case-study images are verified or clearly user-provided;
- contradictions are surfaced;
- no placeholder or invented metric is treated as evidence.
