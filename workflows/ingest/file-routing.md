# File Routing

Use the strongest native reader for the file type and normalize to the v4.1 schemas.

| Input | Preferred reader | Preserve |
|---|---|---|
| DOCX | Documents / Files | headings, paragraphs, tables, images/captions where relevant |
| XLSX/CSV | Spreadsheets | sheet/range, displayed values, formulas, units, dates, charts |
| PPTX | Presentations / Files | slide order, text, notes, tables, charts, images, groups, masters/layouts when style-reference |
| PDF | Files/PDF capability | full text plus visual page inspection for figures/layout |
| Multiple files | Files + type-specific readers | file identity and cross-source provenance |

Always distinguish source role: content-source, data-source, style-reference, evidence-source, mixed.
