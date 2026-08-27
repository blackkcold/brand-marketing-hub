# v4.1 Test Suite

Run from the repository root:

```bash
python assets/tests/test_v4_contracts.py
python assets/tests/test_source_to_deck.py
python assets/tests/test_archetypes.py
python assets/tests/test_render_smoke.py
```

The first two tests cover the current v4.1 architecture: universal source ingestion, coverage, template/runtime contracts and delivery gates. The latter two are retained only as legacy Markdown renderer regression tests.

GitHub CI additionally generates a PPTX through the PptxGenJS fallback, validates it, renders it through LibreOffice, and requires a real PDF/PNG montage.
