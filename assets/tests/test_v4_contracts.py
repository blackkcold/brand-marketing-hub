#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Static v4 architecture contract tests."""
from __future__ import annotations

import json
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]


def read(path: str) -> str:
    return (ROOT / path).read_text(encoding="utf-8")


def main() -> int:
    skill = read("SKILL.md")
    renderer = read("assets/md2pptx_vivo.py")
    validator = read("assets/validate_pptx.py")

    assert "version: 4.0.0" in skill
    assert "Evidence first, slides second" in skill
    assert "deck_spec.json" in skill
    assert "P0/P1" in skill

    for name in (
        "evidence.schema.json",
        "asset.schema.json",
        "deck.schema.json",
        "qa.schema.json",
    ):
        data = json.loads(read(f"schemas/{name}"))
        assert data.get("$schema")
        assert data.get("type") == "object"

    for path in (
        "workflows/research.md",
        "workflows/deck-production.md",
        "workflows/runtime-adapters.md",
        "references/v4-architecture.md",
    ):
        assert (ROOT / path).exists(), path

    # v4 must not depend on a developer-specific absolute PowerPoint template.
    assert "/Users/11169285/" not in renderer
    # Never silently discard table content.
    assert "rows = rows[:10]" not in renderer
    assert "cols = min(cols, 6)" not in renderer
    assert "禁止静默截断" in renderer

    # Severity gate replaces the former issue-page-ratio exemption.
    assert "issue_page_ratio" not in validator
    assert "small<=10%" not in validator
    assert "any P0/P1 blocks delivery" in validator
    assert "output_allowed" in validator

    print("PASS v4 contracts: evidence/deck/runtime/QA architecture is enforced")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
