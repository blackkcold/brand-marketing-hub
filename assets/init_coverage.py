#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Build an initial exhaustive source-to-slide coverage manifest from deck_spec."""
from __future__ import annotations
import argparse, json
from pathlib import Path

def load(p): return json.loads(Path(p).read_text(encoding="utf-8"))

def main():
    ap=argparse.ArgumentParser()
    ap.add_argument("--content",required=True,type=Path)
    ap.add_argument("--deck",required=True,type=Path)
    ap.add_argument("--out",required=True,type=Path)
    args=ap.parse_args()
    content=load(args.content); deck=load(args.deck)
    refs={}
    appendix=set()
    for slide in deck.get("slides",[]):
        sid=slide["slide_id"]
        is_appendix="appendix" in (slide.get("section") or "").lower() or "附录" in (slide.get("section") or "")
        uids=set(slide.get("source_unit_ids",[]))|set(slide.get("must_preserve_unit_ids",[]))
        for b in slide.get("content",[]): uids.update(b.get("unit_ids",[]))
        for uid in uids:
            refs.setdefault(uid,[]).append(sid)
            if is_appendix: appendix.add(uid)
    mappings=[]
    covered=0; critical_missing=0
    for u in content.get("units",[]):
        uid=u["unit_id"]
        if uid in refs:
            mappings.append({"unit_id":uid,"disposition":"appendix" if uid in appendix else "body",
                             "slide_ids":sorted(set(refs[uid])),"reason":"Referenced by deck_spec",
                             "preservation_verified":False})
            covered+=1
        else:
            mappings.append({"unit_id":uid,"disposition":"missing","slide_ids":[],
                             "reason":"Not yet mapped to deck_spec","preservation_verified":False})
            if u.get("importance") in ("critical","high") or u.get("preserve_mode") in ("exact","semantic"):
                critical_missing+=1
    out={"version":"4.1","mappings":mappings,
         "summary":{"total_units":len(mappings),"covered_units":covered,"critical_missing":critical_missing}}
    args.out.parent.mkdir(parents=True,exist_ok=True)
    args.out.write_text(json.dumps(out,ensure_ascii=False,indent=2),encoding="utf-8")
    print(f"coverage: {covered}/{len(mappings)} mapped; critical_missing={critical_missing}")
    return 1 if critical_missing else 0

if __name__=="__main__": raise SystemExit(main())
