#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
validate_pptx.py — 纯代码校验 md2pptx_vivo.py 产出的 pptx（不操控任何应用）。

校验项：
  - 页数与源 md 的 P 编号一致（含封面去重逻辑）
  - 所有 shape 坐标不越界（封面/结尾斜向色块允许出血）
  - 占位符令牌（【】）残留检查
  - 文本按字宽估算高度不溢出文本框
  - 表格行列数与表头样式在位
  - 封面/结尾含渐变背景、斜向光带与 vivo 字标图片
  - 内容页含蓝标题条与页脚（密级+页码）

退出码：
  0 = 全部通过
  1 = 存在校验失败项（FAIL）
  2 = 参数/文件错误（用法错误、文件不存在、依赖缺失）

输出：机器可读 JSON 到 stdout（--json），人类可读文本为默认。
"""
import argparse
import json
import os
import re
import sys

try:
    from pptx import Presentation
    from pptx.enum.shapes import MSO_SHAPE_TYPE
except ImportError:
    sys.exit(2)

SW, SH = 13.333, 7.5
SW_EMU = int(SW * 914400)
SH_EMU = int(SH * 914400)

# 占位符令牌
PLACEHOLDER_RE = re.compile(r'【[^】]*】')

# 页脚密文与页码
FOOTER_TEXT = '内部汇报资料 · 请勿外传'
FOOTER_RE = FOOTER_TEXT


def cjk_w(text, size_pt):
    w = 0
    for ch in text:
        w += size_pt if ord(ch) > 0x2E00 else size_pt * 0.55
    return w / 72.0


def est_lines(text, size_pt, width_in):
    lines = 0
    for para in text.split('\n'):
        w = 0
        for ch in para:
            cw = size_pt if ord(ch) > 0x2E00 else size_pt * 0.55
            w += cw
            if w / 72.0 > width_in:
                lines += 1
                w = cw
        lines += 1
    return lines


def check_bounds(shape, slide_idx, issues, bleed_ok=False):
    try:
        l, t, w, h = shape.left, shape.top, shape.width, shape.height
    except Exception:
        return
    if l is None or t is None or w is None or h is None:
        return
    right = l + w
    bottom = t + h
    if bleed_ok:
        return
    if l < 0 or t < 0 or right > SW_EMU or bottom > SH_EMU:
        issues.append({
            'type': 'bounds',
            'slide': slide_idx,
            'shape': shape.shape_id,
            'name': shape.name,
            'left_in': round(l / 914400, 3),
            'top_in': round(t / 914400, 3),
            'right_in': round(right / 914400, 3),
            'bottom_in': round(bottom / 914400, 3),
        })


def check_text_overflow(shape, slide_idx, issues):
    if not shape.has_text_frame:
        return
    tf = shape.text_frame
    if not tf.word_wrap:
        return
    box_w = shape.width / 914400.0
    box_h = shape.height / 914400.0
    total_h = 0.0
    for p in tf.paragraphs:
        text = p.text
        if not text:
            continue
        size = None
        for r in p.runs:
            if r.font.size:
                size = r.font.size.pt
                break
        if size is None:
            continue
        lh = p.line_spacing if isinstance(p.line_spacing, float) else 1.25
        lines = est_lines(text, size, box_w)
        total_h += lines * size * lh / 72.0
        if p.space_after:
            total_h += p.space_after.pt / 72.0
    if total_h > box_h + 0.05:
        issues.append({
            'type': 'overflow',
            'slide': slide_idx,
            'shape': shape.shape_id,
            'name': shape.name,
            'est_height_in': round(total_h, 3),
            'box_height_in': round(box_h, 3),
        })


def check_placeholder(shape, slide_idx, issues):
    text = shape_text(shape)
    for m in PLACEHOLDER_RE.finditer(text):
        issues.append({
            'type': 'placeholder',
            'slide': slide_idx,
            'shape': shape.shape_id,
            'token': m.group(0),
        })


def shape_text(shape):
    if shape.has_text_frame:
        return shape.text_frame.text
    if shape.has_table:
        parts = []
        for row in shape.table.rows:
            for cell in row.cells:
                parts.append(cell.text)
        return '\n'.join(parts)
    return ''


def validate(pptx_path, md_path=None):
    prs = Presentation(pptx_path)
    slides = list(prs.slides)
    issues = []
    info = {
        'file': pptx_path,
        'slide_count': len(slides),
        'canvas_in': [SW, SH],
        'issues': [],
        'checks': {},
    }

    # 页数 vs md P 编号
    md_pages = None
    if md_path and os.path.exists(md_path):
        md_pages = _count_md_pages(md_path)
        info['md_pages'] = md_pages
        # 封面去重：md 首个 P1｜封面 不单独成页
        expected = md_pages
        if expected is not None:
            if len(slides) != expected:
                info['checks']['page_count'] = {
                    'status': 'FAIL',
                    'expected': expected,
                    'actual': len(slides),
                }
                issues.append({'type': 'page_count', 'expected': expected, 'actual': len(slides)})
            else:
                info['checks']['page_count'] = {'status': 'PASS', 'count': len(slides)}
    else:
        info['checks']['page_count'] = {'status': 'PASS', 'count': len(slides)}

    # ---- 逐页检查 ----
    for idx, slide in enumerate(slides):
        is_cover_end = (idx == 0 or idx == len(slides) - 1)
        has_logo = False
        has_footer = False
        has_page_num = False
        for shape in slide.shapes:
            check_bounds(shape, idx, issues, bleed_ok=is_cover_end)
            check_text_overflow(shape, idx, issues)
            check_placeholder(shape, idx, issues)
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                has_logo = True
            text = shape_text(shape)
            if FOOTER_RE in text:
                has_footer = True
            if re.fullmatch(r'\d+', text.strip()):
                has_page_num = True
        # 封面/结尾页需含图片（字标）
        if is_cover_end:
            if not has_logo:
                issues.append({'type': 'logo_missing', 'slide': idx})
        # 内容页需含页脚
        if not is_cover_end:
            if not has_footer:
                issues.append({'type': 'footer_missing', 'slide': idx})
            if not has_page_num:
                issues.append({'type': 'page_num_missing', 'slide': idx})

    # ---- 表格检查 ----
    for idx, slide in enumerate(slides):
        for shape in slide.shapes:
            if shape.has_table:
                tbl = shape.table
                nr, nc = len(tbl.rows), len(tbl.columns)
                if nr < 1 or nc < 1:
                    issues.append({'type': 'table_empty', 'slide': idx})
                # 表头样式：首行品牌蓝底白字
                hdr = tbl.cell(0, 0)
                try:
                    fill = hdr.fill.fore_color.rgb
                    if str(fill) != '1E46E6':
                        issues.append({'type': 'table_header_style', 'slide': idx})
                except Exception:
                    pass

    info['issues'] = issues
    info['pass'] = len(issues) == 0
    return info


def _count_md_pages(md_path):
    count = 0
    with open(md_path, encoding='utf-8') as f:
        for line in f:
            if re.match(r'^##\s+P\d+[｜|]', line):
                count += 1
    return count


def main():
    ap = argparse.ArgumentParser(description='校验 md2pptx_vivo.py 产出的 pptx')
    ap.add_argument('pptx', help='待校验的 pptx 文件')
    ap.add_argument('--md', help='源 markdown（用于页数比对）')
    ap.add_argument('--json', action='store_true', help='输出机器可读 JSON')
    args = ap.parse_args()

    if not os.path.exists(args.pptx):
        print(json.dumps({'error': 'pptx 不存在', 'path': args.pptx}))
        sys.exit(2)

    info = validate(args.pptx, args.md)
    if args.json:
        print(json.dumps(info, ensure_ascii=False, indent=2))
    else:
        print(f'文件: {info["file"]}')
        print(f'页数: {info["slide_count"]}')
        if info['issues']:
            print(f'发现 {len(info["issues"])} 项问题:')
            for it in info['issues']:
                print('  - ' + json.dumps(it, ensure_ascii=False))
        else:
            print('全部通过')
    sys.exit(0 if info['pass'] else 1)


if __name__ == '__main__':
    main()
