#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Semantic and structural validator for brand-marketing-hub PPTX output."""
from __future__ import annotations
import argparse, json, os, re, sys
from pathlib import Path
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

SW, SH = 13.333, 7.5
TOKENS = {'1E46E6','3458F6','06175E','D1EBFE','EAF2FF','F2F7FF','D6E0F2','AAB9D5','111111','565656','374151','687586','788395','9AA3B2','F6C84C','E6001E','FFFFFF','1020D0','4661F4','80A0D0','E0C0E0','F0B419','FFF4D6','242424'}
PLACEHOLDER_RE = re.compile(r'【[^】]*】')
LAYOUTS = {'bullets','cards','split','table','part','end','toc','stats','framework','funnel','comparison','matrix','timeline','budget','collage','case-study','purpose','evidence-grid','chart'}

def shape_text(shape):
    if getattr(shape, 'has_text_frame', False): return shape.text_frame.text
    if getattr(shape, 'has_table', False): return '\n'.join(cell.text for row in shape.table.rows for cell in row.cells)
    return ''

def md_pages(path):
    pages=[]; current=None
    for line in Path(path).read_text(encoding='utf-8').splitlines():
        m=re.match(r'^##\s+P(\d+)[｜|]\s*(.+)$', line)
        if m:
            current={'no':int(m.group(1)),'title':m.group(2).strip(),'layout':None,'sub':None,'source':None,'bullets':0,'cards':0,'table':0,'stats':0}; pages.append(current); continue
        if current is None: continue
        if line.startswith('@layout'): current['layout']=line.split()[1]
        elif line.startswith('@sub'): current['sub']=line[4:].strip()
        elif line.startswith('@source'): current['source']=line[7:].strip()
        elif line.startswith('@stat'): current['stats']+=1
        elif line.startswith('### '): current['cards']+=1
        elif re.match(r'^-\s+', line): current['bullets']+=1
        elif line.strip().startswith('|') and not set(line.strip().strip('|').replace('|','').strip()) <= set('-:'): current['table']+=1
    return pages

def rgb_hex(color):
    try: return str(color).upper()
    except Exception: return None

def validate(pptx_path, md_path=None):
    prs=Presentation(pptx_path); issues=[]; info={'file':pptx_path,'slide_count':len(prs.slides),'canvas_in':[round(prs.slide_width/914400,3),round(prs.slide_height/914400,3)],'issues':[],'checks':{}}
    source_pages=md_pages(md_path) if md_path and os.path.exists(md_path) else []
    if source_pages:
        info['md_pages']=len(source_pages)
        if len(prs.slides)!=len(source_pages): issues.append({'type':'page_count','code':'page_count_mismatch','expected':len(source_pages),'actual':len(prs.slides)})
        nums=[p['no'] for p in source_pages]
        if nums!=list(range(1,len(nums)+1)): issues.append({'type':'markdown','code':'page_numbers_not_contiguous','numbers':nums})
        for p in source_pages:
            if p['sub'] and len(p['sub'])>22: issues.append({'type':'semantic','code':'sub_too_long','page':p['no'],'length':len(p['sub'])})
            if p['bullets']>6 and p['layout'] in (None,'bullets','purpose'): issues.append({'type':'semantic','code':'bullet_count_exceeded','page':p['no'],'count':p['bullets']})
            if p['cards']>8: issues.append({'type':'semantic','code':'card_count_exceeded','page':p['no'],'count':p['cards']})
            if p['layout'] in ('table','budget','matrix','timeline') and p['table']>10: issues.append({'type':'semantic','code':'table_row_count_exceeded','page':p['no'],'count':p['table']})
            if p['layout'] in ('stats','evidence-grid') and not p['source']: issues.append({'type':'semantic','code':'source_missing','page':p['no']})
    for idx,slide in enumerate(prs.slides):
        cover_or_end=idx in (0,len(prs.slides)-1)
        has_logo=False; has_footer=False; has_page=False
        for shape in slide.shapes:
            try:
                l,t,w,h=shape.left/914400,shape.top/914400,shape.width/914400,shape.height/914400
                if not cover_or_end and (l<0 or t<0 or l+w>SW+.01 or t+h>SH+.01): issues.append({'type':'bounds','code':'shape_out_of_bounds','slide':idx+1,'shape':shape.shape_id})
            except Exception: pass
            if shape.shape_type==MSO_SHAPE_TYPE.PICTURE: has_logo=True
            text=shape_text(shape)
            if getattr(shape, 'has_text_frame', False):
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        if run.font.name and run.font.name != '微软雅黑':
                            issues.append({'type':'font','code':'font_off_contract','slide':idx+1,'shape':shape.shape_id,'font':run.font.name})
            try:
                if 'outerShdw' in shape._element.xml or 'innerShdw' in shape._element.xml:
                    issues.append({'type':'style','code':'shadow_forbidden','slide':idx+1,'shape':shape.shape_id})
            except Exception:
                pass
            if '内部汇报资料' in text or 'THANK YOU' in text: has_footer=has_footer or '内部汇报资料' in text
            if re.fullmatch(r'\d+',text.strip()): has_page=True
            for match in PLACEHOLDER_RE.finditer(text): issues.append({'type':'placeholder','code':'placeholder_remaining','slide':idx+1,'token':match.group(0)})
            if getattr(shape,'fill',None) and shape.fill.type:
                try:
                    value=rgb_hex(shape.fill.fore_color.rgb)
                    if value and value not in TOKENS and value not in {'00000000'}: issues.append({'type':'color','code':'color_off_palette','slide':idx+1,'shape':shape.shape_id,'color':value})
                except Exception: pass
        if cover_or_end and not has_logo and len(prs.slide_masters) < 2: issues.append({'type':'brand','code':'logo_missing','slide':idx+1})
        if not cover_or_end and not has_page: issues.append({'type':'chrome','code':'page_num_missing','slide':idx+1})
    issue_pages = {issue.get('slide') for issue in issues if issue.get('slide')}
    checked_pages = max(len(prs.slides) - 2, 1)
    ratio = len(issue_pages) / checked_pages
    info['issue_page_count'] = len(issue_pages)
    info['issue_page_ratio'] = round(ratio, 4)
    info['issue_policy'] = 'small<=10%: output-with-mark; large>10%: report'
    info['issues']=issues; info['pass']=not issues; info['output_allowed']=ratio <= 0.10 or not issues
    return info

def main():
    ap=argparse.ArgumentParser(); ap.add_argument('pptx'); ap.add_argument('--md'); ap.add_argument('--json',action='store_true'); args=ap.parse_args()
    if not os.path.exists(args.pptx): print(json.dumps({'error':'pptx 不存在','path':args.pptx},ensure_ascii=False)); return 2
    info=validate(args.pptx,args.md)
    if args.json: print(json.dumps(info,ensure_ascii=False,indent=2))
    else:
        print(f'文件: {info["file"]}\n页数: {info["slide_count"]}')
        print('全部通过' if info['pass'] else f'发现 {len(info["issues"])} 项问题:')
        for issue in info['issues']: print('  - '+json.dumps(issue,ensure_ascii=False))
    return 0 if info['pass'] else 1

if __name__=='__main__': raise SystemExit(main())
