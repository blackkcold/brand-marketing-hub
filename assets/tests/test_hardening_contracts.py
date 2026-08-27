#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Focused regression tests for v4.1.2 hardening contracts."""
from __future__ import annotations
import importlib.util
import json
import subprocess
import sys
import tempfile
from pathlib import Path

from jsonschema import Draft202012Validator

ROOT=Path(__file__).resolve().parents[2]

def load(path):
    return json.loads(Path(path).read_text(encoding="utf-8"))

def module(path,name):
    spec=importlib.util.spec_from_file_location(name,path)
    mod=importlib.util.module_from_spec(spec)
    assert spec.loader
    spec.loader.exec_module(mod)
    return mod

def assert_schema_valid(schema,obj):
    errs=list(Draft202012Validator(schema).iter_errors(obj))
    assert not errs,[e.message for e in errs]

def assert_schema_invalid(schema,obj):
    assert list(Draft202012Validator(schema).iter_errors(obj)),obj

def test_natural_render_sort():
    rp=module(ROOT/"assets/render_pptx.py","render_pptx")
    vr=module(ROOT/"assets/validate_render.py","validate_render")
    files=[Path("slide-10.png"),Path("slide-2.png"),Path("slide-1.png")]
    expected=["slide-1.png","slide-2.png","slide-10.png"]
    assert [p.name for p in sorted(files,key=rp.slide_sort_key)]==expected
    assert [p.name for p in sorted(files,key=vr.slide_sort_key)]==expected

def test_asset_schema_conditions():
    schema=load(ROOT/"schemas/asset.schema.json")
    generated={"version":"4.1","assets":[{
        "asset_id":"IMG-001","type":"generated-concept","verification":"generated",
        "semantic_verified":False,"visual_role":"concept"
    }]}
    assert_schema_valid(schema,generated)

    missing_source={"version":"4.1","assets":[{
        "asset_id":"IMG-001","type":"official-product-image","verification":"verified",
        "semantic_verified":True,"visual_role":"product-proof","embedded_required":True
    }]}
    assert_schema_invalid(schema,missing_source)

    missing_embed={"version":"4.1","assets":[{
        "asset_id":"IMG-001","type":"official-product-image","verification":"verified",
        "semantic_verified":True,"visual_role":"product-proof","source_url":"https://example.com/a.png"
    }]}
    assert_schema_invalid(schema,missing_embed)

    valid_proof={"version":"4.1","assets":[{
        "asset_id":"IMG-001","type":"official-product-image","verification":"verified",
        "semantic_verified":True,"visual_role":"product-proof","embedded_required":True,
        "source_url":"https://example.com/a.png"
    }]}
    assert_schema_valid(schema,valid_proof)

def test_cross_manifest_duplicate_and_table_gate():
    validator=ROOT/"assets/validate_v4_manifests.py"
    template=ROOT/"brand/vivo/template-manifest.json"
    with tempfile.TemporaryDirectory(prefix="bmh-hardening-") as td:
        d=Path(td)
        sources={"version":"4.1","sources":[{
            "source_id":"FILE-001","kind":"docx","display_name":"a.docx",
            "origin":"user-upload","confidentiality":"internal","role":"content-source"
        }]}
        content={"version":"4.1","units":[
            {"unit_id":"UNIT-0001","source_id":"FILE-001","unit_type":"paragraph","locator":{"paragraph":1},
             "content":"A","preserve_mode":"semantic","importance":"high"},
            {"unit_id":"UNIT-0001","source_id":"FILE-001","unit_type":"paragraph","locator":{"paragraph":2},
             "content":"B","preserve_mode":"semantic","importance":"high"},
        ]}
        deck={"version":"4.1","deck":{
            "title":"QA","purpose":"QA","story_archetype":"custom","visual_template":"vivo-house",
            "template_manifest":"brand/vivo/template-manifest.json","runtime_preference":["pptxgenjs"],
            "source_of_truth":"deck_spec.json","confidentiality":"internal"
        },"slides":[{
            "slide_id":"S01","intent":"qa","takeaway":"qa","layout_candidates":["table"],
            "content":[{"block_id":"BLK-001","type":"table","headers":["A","B"],
                        "rows":[["1","2"],["3"]]}]
        }]}
        for name,obj in (("sources",sources),("content",content),("deck",deck)):
            (d/f"{name}.json").write_text(json.dumps(obj),encoding="utf-8")
        p=subprocess.run([
            sys.executable,str(validator),"--sources",str(d/"sources.json"),
            "--content",str(d/"content.json"),"--deck",str(d/"deck.json"),
            "--template",str(template)
        ],capture_output=True,text=True)
        assert p.returncode==1,p.stdout+p.stderr
        assert "duplicate_unit_id" in p.stdout,p.stdout
        assert "table_width_mismatch" in p.stdout,p.stdout

def main():
    test_natural_render_sort()
    test_asset_schema_conditions()
    test_cross_manifest_duplicate_and_table_gate()
    print("PASS v4.1.2 hardening contracts")
    return 0

if __name__=="__main__":
    raise SystemExit(main())
