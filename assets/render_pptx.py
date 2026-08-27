#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Render PPTX to PDF/PNGs and a contact-sheet montage using LibreOffice + poppler."""
from __future__ import annotations
import argparse, math, shutil, subprocess, tempfile
from pathlib import Path
from PIL import Image, ImageOps, ImageDraw

def cmd(name):
    p=shutil.which(name)
    if not p: raise SystemExit(f"required executable not found: {name}")
    return p

def main():
    ap=argparse.ArgumentParser()
    ap.add_argument("pptx",type=Path)
    ap.add_argument("--out-dir",type=Path,default=Path("render"))
    ap.add_argument("--cols",type=int,default=4)
    args=ap.parse_args()
    if not args.pptx.exists(): raise SystemExit(f"missing pptx: {args.pptx}")
    args.out_dir.mkdir(parents=True,exist_ok=True)
    soffice=cmd("soffice"); pdftoppm=cmd("pdftoppm")
    subprocess.run([soffice,"--headless","--convert-to","pdf","--outdir",str(args.out_dir),str(args.pptx)],check=True)
    pdf=args.out_dir/(args.pptx.stem+".pdf")
    if not pdf.exists(): raise SystemExit("LibreOffice did not create PDF")
    prefix=args.out_dir/"slide"
    subprocess.run([pdftoppm,"-png","-r","120",str(pdf),str(prefix)],check=True)
    slides=sorted(args.out_dir.glob("slide-*.png"))
    if not slides: raise SystemExit("No slide PNGs were rendered")
    thumbs=[]
    for i,p in enumerate(slides,1):
        im=Image.open(p).convert("RGB")
        im.thumbnail((480,270))
        canvas=Image.new("RGB",(500,310),"white")
        canvas.paste(im,((500-im.width)//2,20))
        ImageDraw.Draw(canvas).text((12,286),f"{i:02d}",fill="black")
        thumbs.append(canvas)
    cols=max(1,args.cols); rows=math.ceil(len(thumbs)/cols)
    montage=Image.new("RGB",(cols*500,rows*310),(235,235,235))
    for i,im in enumerate(thumbs):
        montage.paste(im,((i%cols)*500,(i//cols)*310))
    montage_path=args.out_dir/"montage.png"; montage.save(montage_path)
    print(f"PASS render: {len(slides)} slides; pdf={pdf}; montage={montage_path}")
    return 0

if __name__=="__main__": raise SystemExit(main())
