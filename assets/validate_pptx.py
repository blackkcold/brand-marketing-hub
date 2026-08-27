#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""v4 semantic/structural validator for brand-marketing-hub PPTX output.

This validator is a compatibility QA layer for PPTX/legacy Markdown. It applies
severity-based delivery gates. P0/P1 always block output; page-ratio exemptions
were removed in v4.
"""
from __future__ import annotations

import argparse
import json
import os
import re
from pathlib import Path

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

SW, SH = 13.333, 7.5
TOKENS = {
    '1E46E6','3458F6','06175E','D1EBFE','EAF2FF','F2F7FF','D6E0F2','AAB9D5',
    '111111','565656','374151','687586','788395','9AA3B2','F6C84C','E6001E',
    'FFFFFF','1020D0','4661F4','80A0D0','E0C0E0','F0B419','FFF4D6','242424'
}
PLACEHOLDER_RE = re.compile(r'【[^】]*】')
LAYOUTS = {
    'bullets','cards','split','table','part','end','toc','stats','framework',
    'funnel','comparison','matrix','timeline','budget','collage','case-study',
    'purpose','evidence-grid','chart'
}

SEVERITY = {
    'page_count_mismatch': 'P0',
    'page_numbers_not_contiguous': 'P0',
    'placeholder_remaining': 'P0',
    'table_row_count_exceeded': 'P0',
    'table_col_count_exceeded': 'P0',
    'source_missing': 'P1',
    'shape_out_of_bounds': 'P1',
    'logo_missing': 'P1',
    'bullet_count_exceeded': 'P1',
    'card_count_exceeded': 'P1',
    'sub_too_long': 'P2',
    'font_off_contract': 'P2',
    'shadow_forbidden': 'P2',
    'color_off_palette': 'P2',
    'page_num_missing': 'P2',
}


def issue(code, **kwargs):
    return {'severity': SEVERITY.get(code, 'P2'), 'code': code, **kwargs}


def shape_text(shape):
    if getattr(shape, 'has_text_frame', False):
        return shape.text_frame.text
    if getattr(shape, 'has_table', False):
        return '\n'.join(cell.text for row in shape.table.rows for cell in row.cells)
    return ''


def md_pages(path):
    pages = []
    current = None
    for line in Path(path).read_text(encoding='utf-8').splitlines():
        m = re.match(r'^##\s+P(\d+)[｜|]\s*(.+)$', line)
        if m:
            current = {
                'no': int(m.group(1)),
                'title': m.group(2).strip(),
                'layout': None,
                'sub': None,
                'source': None,
                'bullets': 0,
                'cards': 0,
                'table': 0,
                'table_cols': 0,
                'stats': 0,
            }
            pages.append(current)
            continue
        if current is None:
            continue
        if line.startswith('@layout'):
            parts = line.split()
            current['layout'] = parts[1] if len(parts) > 1 else None
        elif line.startswith('@sub'):
            current['sub'] = line[4:].strip()
        elif line.startswith('@source'):
            current['source'] = line[7:].strip()
        elif line.startswith('@stat'):
            current['stats'] += 1
        elif line.startswith('### '):
            current['cards'] += 1
        elif re.match(r'^-\s+', line):
            current['bullets'] += 1
        elif line.strip().startswith('|') and not set(line.strip().strip('|').replace('|','').strip()) <= set('-:'):
            current['table'] += 1
            current['table_cols'] = max(current['table_cols'], len(line.strip().strip('|').split('|')))
    return pages


def rgb_hex(color):
    try:
        return str(color).upper()
    except Exception:
        return None


def validate(pptx_path, md_path=None):
    prs = Presentation(pptx_path)
    issues = []
    info = {
        'version': '4.0',
        'file': pptx_path,
        'slide_count': len(prs.slides),
        'canvas_in': [round(prs.slide_width / 914400, 3), round(prs.slide_height / 914400, 3)],
        'issues': [],
        'checks': {},
    }

    source_pages = md_pages(md_path) if md_path and os.path.exists(md_path) else []
    if source_pages:
        info['md_pages'] = len(source_pages)
        if len(prs.slides) != len(source_pages):
            issues.append(issue('page_count_mismatch', expected=len(source_pages), actual=len(prs.slides)))
        nums = [p['no'] for p in source_pages]
        if nums != list(range(1, len(nums) + 1)):
            issues.append(issue('page_numbers_not_contiguous', numbers=nums))
        for p in source_pages:
            page = p['no']
            if p['sub'] and len(p['sub']) > 22:
                issues.append(issue('sub_too_long', page=page, length=len(p['sub'])))
            if p['bullets'] > 6 and p['layout'] in (None, 'bullets', 'purpose'):
                issues.append(issue('bullet_count_exceeded', page=page, count=p['bullets']))
            if p['cards'] > 8:
                issues.append(issue('card_count_exceeded', page=page, count=p['cards']))
            if p['layout'] in ('table','budget','matrix','timeline') and p['table'] > 10:
                issues.append(issue('table_row_count_exceeded', page=page, count=p['table']))
            if p['layout'] in ('table','budget','matrix','timeline') and p['table_cols'] > 6:
                issues.append(issue('table_col_count_exceeded', page=page, count=p['table_cols']))
            if p['layout'] in ('stats','evidence-grid') and not p['source']:
                issues.append(issue('source_missing', page=page))

    for idx, slide in enumerate(prs.slides):
        slide_no = idx + 1
        cover_or_end = idx in (0, len(prs.slides) - 1)
        has_logo = False
        has_page = False

        for shape in slide.shapes:
            try:
                l = shape.left / 914400
                t = shape.top / 914400
                w = shape.width / 914400
                h = shape.height / 914400
                if not cover_or_end and (l < 0 or t < 0 or l + w > SW + .01 or t + h > SH + .01):
                    issues.append(issue('shape_out_of_bounds', slide=slide_no, shape=shape.shape_id))
            except Exception:
                pass

            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                has_logo = True

            text = shape_text(shape)

            if getattr(shape, 'has_text_frame', False):
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        if run.font.name and run.font.name != '微软雅黑':
                            issues.append(issue('font_off_contract', slide=slide_no, shape=shape.shape_id, font=run.font.name))

            try:
                if 'outerShdw' in shape._element.xml or 'innerShdw' in shape._element.xml:
                    issues.append(issue('shadow_forbidden', slide=slide_no, shape=shape.shape_id))
            except Exception:
                pass

            if re.fullmatch(r'\d+', text.strip()):
                has_page = True

            for match in PLACEHOLDER_RE.finditer(text):
                issues.append(issue('placeholder_remaining', slide=slide_no, token=match.group(0)))

            if getattr(shape, 'fill', None) and shape.fill.type:
                try:
                    value = rgb_hex(shape.fill.fore_color.rgb)
                    if value and value not in TOKENS and value not in {'00000000'}:
                        issues.append(issue('color_off_palette', slide=slide_no, shape=shape.shape_id, color=value))
                except Exception:
                    pass

        if cover_or_end and not has_logo and len(prs.slide_masters) < 2:
            issues.append(issue('logo_missing', slide=slide_no))
        if not cover_or_end and not has_page:
            issues.append(issue('page_num_missing', slide=slide_no))

    blocking = [x for x in issues if x['severity'] in ('P0', 'P1')]
    info['issues'] = issues
    info['counts_by_severity'] = {
        level: sum(1 for x in issues if x['severity'] == level)
        for level in ('P0','P1','P2','P3')
    }
    info['issue_policy'] = 'v4 severity gate: any P0/P1 blocks delivery'
    info['pass'] = not issues
    info['output_allowed'] = not blocking
    info['status'] = 'pass' if not issues else ('fail' if blocking else 'pass-with-cosmetic-notes')
    return info


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('pptx')
    ap.add_argument('--md')
    ap.add_argument('--json', action='store_true')
    args = ap.parse_args()

    if not os.path.exists(args.pptx):
        print(json.dumps({'error':'pptx 不存在','path':args.pptx}, ensure_ascii=False))
        return 2

    info = validate(args.pptx, args.md)
    if args.json:
        print(json.dumps(info, ensure_ascii=False, indent=2))
    else:
        print(f'文件: {info["file"]}\n页数: {info["slide_count"]}\n状态: {info["status"]}')
        for item in info['issues']:
            print('  - ' + json.dumps(item, ensure_ascii=False))

    return 0 if info['output_allowed'] else 1


if __name__ == '__main__':
    raise SystemExit(main())
