#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Static v4.1 architecture contract tests."""
from __future__ import annotations
import json
from pathlib import Path

ROOT=Path(__file__).resolve().parents[2]

def read(path:str)->str:
    return (ROOT/path).read_text(encoding="utf-8")

def main()->int:
    skill=read("SKILL.md")
    renderer=read("assets/md2pptx_vivo.py")
    validator=read("assets/validate_pptx.py")
    assert "version: 4.1.1" in skill
    assert "Universal Source-to-Deck" in skill
    assert "coverage.json" in skill
    assert "Story Archetype ≠ Visual Template" in skill
    assert "external image relationship" in skill.lower()

    for name in (
        "source.schema.json","content-unit.schema.json","evidence.schema.json",
        "asset.schema.json","coverage.schema.json","deck.schema.json",
        "template.schema.json","qa.schema.json",
    ):
        data=json.loads(read(f"schemas/{name}"))
        assert data.get("$schema") and data.get("type")=="object",name

    required=(
        "workflows/source-to-deck.md","workflows/research.md","workflows/synthesis.md",
        "workflows/deck-production.md","workflows/visual-qa.md","workflows/revision.md",
        "workflows/ingest/docx.md","workflows/ingest/xlsx.md","workflows/ingest/pptx.md",
        "brand/vivo/template-manifest.json","brand/vivo/layout-map.json",
        "assets/ingest_sources.py","assets/init_coverage.py","assets/render_pptx.py","assets/validate_render.py","assets/tests/test_pptx_content_preservation.py",
        "runtime/pptxgenjs/render.js","story-archetypes/archetypes.json",
    )
    for p in required: assert (ROOT/p).exists(),p

    # Legacy renderer stays quarantined and cannot silently discard content.
    assert "/Users/11169285/" not in renderer
    assert "rows = rows[:10]" not in renderer
    assert "cols = min(cols, 6)" not in renderer
    assert "禁止静默截断" in renderer

    # Delivery gate must detect linked images and cannot use issue-page ratio.
    assert "external_image_relationship" in validator
    assert "issue_page_ratio" not in validator
    assert "small<=10%" not in validator
    assert "P0/P1 blocks delivery" in validator

    # Domain modules may not route back into Markdown/python-pptx.
    domain="\n".join(read(p) for p in ("references/ip-collab.md","references/celebrity.md","references/campaign.md"))
    assert "套用模板填充为md" not in domain
    assert "~/.vbuddy/skills/brand-marketing-hub" not in domain
    design=read("references/design-spec.md")
    assert "默认用 `assets/md2pptx_vivo.py`" not in design

    tm=json.loads(read("brand/vivo/template-manifest.json"))
    assert tm["visual_template"]=="vivo-house"
    assert tm["partner_accent_policy"]=="allow-verified-brand-colors"

    print("PASS v4.1 contracts: universal ingestion, coverage, template/runtime and QA are enforced")
    return 0

if __name__=="__main__": raise SystemExit(main())
