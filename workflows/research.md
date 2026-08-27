# v4.1 Research & Evidence Workflow

## Objective
Fill only decision-relevant gaps after source ingestion. Research must never overwrite user-provided source content silently.

## 0. Confidentiality / query sanitization
Before any public web search, classify the source context: public / internal / confidential / restricted.
- Public: search directly.
- Internal: remove internal-only identifiers unless necessary.
- Confidential / restricted: never send unreleased product names, internal project codes, budgets, pricing, personal data, contract terms or other sensitive text into public search queries. Rewrite the query to a public abstraction.
If the task cannot be researched safely without revealing sensitive context, stop external lookup for that fact and use internal/user evidence only.

## 1. Research question planner
For each gap, define:
- exact entity/entities;
- fact type: identity / product / collaboration / market / date / price / performance / audience / risk / design;
- freshness requirement;
- locale/language likely to contain authoritative evidence;
- whether a factual image is required.

## 2. Entity resolution and query expansion
Resolve brand/product/person/collaboration names before collecting evidence. Expand searches across official domains, partner official domains, local-language names and relevant market spellings. Do not treat similarly named products or collaborations as identical.

## 3. Source priority
official > partner-official > regulator/public institution > academic/authority > major media > secondary source.

Material recommendation claims should have corroboration where practical. Conflicts remain visible until resolved.

## 4. Evidence capture
Create claims in `evidence.json`.
- User-file evidence points to `UNIT-xxxx`.
- External evidence points to `SRC-xxx`.
- Mark status verified / user-provided / uncertain / rejected.
- Record publication/update date and retrieved_at for time-sensitive facts.

## 5. Factual image research
Research images independently from text claims.
For product/collaboration/person/event proof:
- prefer official or partner-official imagery;
- verify subject entity, collaboration/product model and visual role;
- materialize the file locally;
- set `semantic_verified=true`;
- set `embedded_required=true`.
Generated concepts may never be marked factual evidence.

## 6. Search stop criteria
Stop when:
- every decision-critical gap is verified or explicitly uncertain;
- additional search is only duplicative;
- the source hierarchy cannot improve;
- privacy/confidentiality prevents safe lookup.

## Evidence gate
Before synthesis:
- every critical factual claim has valid evidence_refs;
- time-sensitive claims have freshness checked;
- conflicts are surfaced;
- factual assets are semantically verified;
- no confidential information was exposed through search;
- no placeholder/invented metric is treated as evidence.
