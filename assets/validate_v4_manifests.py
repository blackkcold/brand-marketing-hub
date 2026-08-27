#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Validate v4 evidence/assets/deck/QA manifests against repository schemas."""
from __future__ import annotations

import argparse
import json
from pathlib import Path

from jsonschema import Draft202012Validator

ROOT = Path(__file__).resolve().parents[1]
SCHEMAS = {
    "evidence": ROOT.parent / "schemas" / "evidence.schema.json",
    "assets": ROOT.parent / "schemas" / "asset.schema.json",
    "deck": ROOT.parent / "schemas" / "deck.schema.json",
    "qa": ROOT.parent / "schemas" / "qa.schema.json",
}


def load(path: Path):
    return json.loads(path.read_text(encoding="utf-8"))


def validate_one(kind: str, path: Path):
    schema = load(SCHEMAS[kind])
    data = load(path)
    validator = Draft202012Validator(schema)
    errors = sorted(validator.iter_errors(data), key=lambda e: list(e.path))
    return errors


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--evidence", type=Path)
    ap.add_argument("--assets", type=Path)
    ap.add_argument("--deck", type=Path)
    ap.add_argument("--qa", type=Path)
    args = ap.parse_args()

    selected = {
        "evidence": args.evidence,
        "assets": args.assets,
        "deck": args.deck,
        "qa": args.qa,
    }
    selected = {k: v for k, v in selected.items() if v is not None}
    if not selected:
        ap.error("至少提供一个 manifest 参数")

    failed = False
    for kind, path in selected.items():
        if not path.exists():
            print(f"FAIL {kind}: file not found: {path}")
            failed = True
            continue
        errors = validate_one(kind, path)
        if errors:
            failed = True
            print(f"FAIL {kind}: {path}")
            for err in errors:
                where = ".".join(str(x) for x in err.path) or "<root>"
                print(f"  - {where}: {err.message}")
        else:
            print(f"PASS {kind}: {path}")

    return 1 if failed else 0


if __name__ == "__main__":
    raise SystemExit(main())
