#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Deterministic sanity checks for rendered slide images.

This does not replace model/vision review. It blocks corrupt, blank, inconsistent
render output before higher-level visual QA.
"""
from __future__ import annotations
import argparse, json
from pathlib import Path
from PIL import Image, ImageStat

def main():
    ap=argparse.ArgumentParser()
    ap.add_argument("render_dir",type=Path)
    ap.add_argument("--json",action="store_true")
    args=ap.parse_args()
    slides=sorted(args.render_dir.glob("slide-*.png"))
    issues=[]
    if not slides:
        issues.append({"severity":"P1","code":"render_missing","message":"No slide PNGs found"})
    base_ratio=None
    for i,p in enumerate(slides,1):
        try:
            im=Image.open(p).convert("RGB")
            if im.width<640 or im.height<360:
                issues.append({"severity":"P1","code":"render_too_small","slide":i,"size":[im.width,im.height]})
            ratio=im.width/im.height
            if base_ratio is None:base_ratio=ratio
            elif abs(ratio-base_ratio)>.01:
                issues.append({"severity":"P1","code":"render_ratio_mismatch","slide":i,"ratio":ratio})
            stat=ImageStat.Stat(im.resize((160,90)))
            spread=max((hi-lo) for lo,hi in im.getextrema())
            variance=max(stat.var)
            if spread<8 or variance<1.0:
                issues.append({"severity":"P1","code":"render_blank_or_corrupt","slide":i,"spread":spread,"variance":variance})
        except Exception as e:
            issues.append({"severity":"P1","code":"render_unreadable","slide":i,"message":str(e)})
    montage=args.render_dir/"montage.png"
    if not montage.exists() or montage.stat().st_size==0:
        issues.append({"severity":"P1","code":"montage_missing","message":"montage.png missing"})
    result={"version":"4.1","slide_count":len(slides),"issues":issues,"pass":not issues}
    if args.json:print(json.dumps(result,ensure_ascii=False,indent=2))
    else:
        print(("PASS" if not issues else "FAIL")+f" render sanity: {len(slides)} slides")
        for x in issues:print("  - "+json.dumps(x,ensure_ascii=False))
    return 0 if not issues else 1

if __name__=="__main__":raise SystemExit(main())
