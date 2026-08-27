#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Validate v4.1 manifests plus cross-manifest provenance/coverage contracts."""
from __future__ import annotations

import argparse
import json
from pathlib import Path
from collections import Counter

from jsonschema import Draft202012Validator

ROOT=Path(__file__).resolve().parents[1]
SCHEMAS={
 "sources":ROOT/"schemas/source.schema.json",
 "content":ROOT/"schemas/content-unit.schema.json",
 "evidence":ROOT/"schemas/evidence.schema.json",
 "assets":ROOT/"schemas/asset.schema.json",
 "coverage":ROOT/"schemas/coverage.schema.json",
 "deck":ROOT/"schemas/deck.schema.json",
 "template":ROOT/"schemas/template.schema.json",
 "qa":ROOT/"schemas/qa.schema.json",
}

def load(path:Path):
    return json.loads(path.read_text(encoding="utf-8"))

def schema_errors(kind,path):
    schema=load(SCHEMAS[kind]); data=load(path)
    return data,sorted(Draft202012Validator(schema).iter_errors(data),key=lambda e:list(e.path))

def add(errors,code,msg):
    errors.append(f"{code}: {msg}")

def main()->int:
    ap=argparse.ArgumentParser()
    for name in SCHEMAS:
        ap.add_argument(f"--{name}",type=Path)
    args=ap.parse_args()
    selected={name:getattr(args,name) for name in SCHEMAS if getattr(args,name) is not None}
    if not selected: ap.error("至少提供一个 manifest 参数")

    data={}
    failed=False
    for kind,path in selected.items():
        if not path.exists():
            print(f"FAIL {kind}: file not found: {path}"); failed=True; continue
        obj,errs=schema_errors(kind,path); data[kind]=obj
        if errs:
            failed=True; print(f"FAIL {kind}: {path}")
            for e in errs:
                where=".".join(str(x) for x in e.path) or "<root>"
                print(f"  - schema {where}: {e.message}")
        else:
            print(f"PASS schema {kind}: {path}")

    cross=[]
    source_ids={x["source_id"] for x in data.get("sources",{}).get("sources",[])}
    units=data.get("content",{}).get("units",[])
    unit_ids={x["unit_id"] for x in units}
    unit_map={x["unit_id"]:x for x in units}
    ext_sources=data.get("evidence",{}).get("sources",[])
    evidence_source_ids={x["source_id"] for x in ext_sources}
    claims=data.get("evidence",{}).get("claims",[])
    claim_ids={x["claim_id"] for x in claims}
    assets=data.get("assets",{}).get("assets",[])
    asset_ids={x["asset_id"] for x in assets}
    slides=data.get("deck",{}).get("slides",[])
    slide_ids={x["slide_id"] for x in slides}

    # Content units must point to an ingested file.
    for u in units:
        if source_ids and u["source_id"] not in source_ids:
            add(cross,"orphan_unit",f'{u["unit_id"]} references missing {u["source_id"]}')

    # Evidence references may resolve to a web/evidence source or a source content unit.
    valid_evidence_refs=evidence_source_ids|unit_ids
    for cl in claims:
        for ref in cl.get("evidence_refs",[]):
            if valid_evidence_refs and ref not in valid_evidence_refs:
                add(cross,"orphan_evidence_ref",f'{cl["claim_id"]} -> {ref}')
        for sid in cl.get("slide_ids",[]):
            if slide_ids and sid not in slide_ids:
                add(cross,"orphan_claim_slide",f'{cl["claim_id"]} -> {sid}')
        for aid in cl.get("asset_ids",[]):
            if asset_ids and aid not in asset_ids:
                add(cross,"orphan_claim_asset",f'{cl["claim_id"]} -> {aid}')

    for a in assets:
        ref=a.get("source_ref")
        if ref and valid_evidence_refs and ref not in valid_evidence_refs:
            add(cross,"orphan_asset_source",f'{a["asset_id"]} -> {ref}')
        if a.get("factual_evidence_allowed") and (not a.get("semantic_verified") or a.get("verification") not in ("verified","user-provided")):
            add(cross,"unsafe_factual_asset",a["asset_id"])
        if a.get("type")=="generated-concept" and a.get("factual_evidence_allowed"):
            add(cross,"generated_asset_as_fact",a["asset_id"])
        if a.get("embedded_required") and not a.get("local_path"):
            add(cross,"asset_not_materialized",f'{a["asset_id"]} requires embedded local asset')

    # Deck references.
    for s in slides:
        for uid in s.get("source_unit_ids",[])+s.get("must_preserve_unit_ids",[]):
            if unit_ids and uid not in unit_ids:
                add(cross,"orphan_slide_unit",f'{s["slide_id"]} -> {uid}')
        for cid in s.get("claim_ids",[]):
            if claim_ids and cid not in claim_ids:
                add(cross,"orphan_slide_claim",f'{s["slide_id"]} -> {cid}')
        for aid in s.get("asset_ids",[]):
            if asset_ids and aid not in asset_ids:
                add(cross,"orphan_slide_asset",f'{s["slide_id"]} -> {aid}')
        for b in s.get("content",[]):
            for uid in b.get("unit_ids",[]):
                if unit_ids and uid not in unit_ids:
                    add(cross,"orphan_block_unit",f'{s["slide_id"]}/{b.get("block_id")} -> {uid}')

    # Coverage is exhaustive: every normalized source unit gets one disposition.
    if "coverage" in data and units:
        mappings=data["coverage"].get("mappings",[])
        counts=Counter(m["unit_id"] for m in mappings)
        for uid in unit_ids:
            if counts[uid]==0: add(cross,"coverage_missing_mapping",uid)
            elif counts[uid]>1: add(cross,"coverage_duplicate_mapping",uid)
        for m in mappings:
            uid=m["unit_id"]
            if uid not in unit_ids:
                add(cross,"coverage_orphan_mapping",uid); continue
            u=unit_map[uid]; disp=m["disposition"]
            if disp=="missing":
                add(cross,"coverage_content_missing",uid)
            if u.get("preserve_mode")=="exact" and disp not in ("body","appendix"):
                add(cross,"exact_content_not_preserved",f"{uid} -> {disp}")
            if u.get("preserve_mode")=="semantic" and disp=="summarized" and not m.get("preservation_verified"):
                add(cross,"semantic_preservation_unverified",uid)
            if disp=="intentionally-excluded" and not m.get("reason"):
                add(cross,"exclusion_without_reason",uid)
            for sid in m.get("slide_ids",[]):
                if slide_ids and sid not in slide_ids:
                    add(cross,"coverage_orphan_slide",f"{uid} -> {sid}")

    # Story template and visual template are distinct and must align to template manifest.
    if "template" in data and "deck" in data:
        tm=data["template"]; d=data["deck"]["deck"]
        if d.get("visual_template")!=tm.get("visual_template"):
            add(cross,"visual_template_mismatch",f'{d.get("visual_template")} != {tm.get("visual_template")}')
        if d.get("template_manifest") and not str(d["template_manifest"]).endswith("template-manifest.json"):
            add(cross,"template_manifest_path_unexpected",d["template_manifest"])

    if cross:
        failed=True
        print("FAIL cross-manifest contracts:")
        for e in cross: print("  - "+e)
    else:
        print("PASS cross-manifest contracts")

    return 1 if failed else 0

if __name__=="__main__":
    raise SystemExit(main())
