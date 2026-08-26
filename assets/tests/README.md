# Visual regression smoke test

Run from the skill root or any directory:

```bash
python3 /Users/11169285/.vbuddy/skills/brand-marketing-hub/assets/tests/test_render_smoke.py
```

The script uses a temporary directory, generates a synthetic image, renders `assets/fixtures/visual-regression.md` through `assets/md2pptx_vivo.py`, and validates structural expectations. It does not overwrite `assets/samples/` or any existing PPTX.
