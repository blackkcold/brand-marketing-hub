#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
md2pptx_vivo.py — vivo 企业风格 deck 转换器（brand-marketing-hub 输出流水线 v2.1）

设计系统见 references/deck-style.md（提炼自 8 份 vivo 内部真实 deck）。
版式：封面(渐变+字标) / 目录 / Part章节页 / 内容页(蓝标题条+bullet/卡片/表格/图文) / 结尾页。

md 方言（与 v2 完全兼容）：
  # 文档标题                  → 封面主标题（40pt 白）
  <!-- deck: 副标题 -->        → 封面副标题（20pt）
  <!-- deck: meta: 部门｜日期｜密级 --> → 封面左下小字（9pt）
  ## P{n}｜页面标题            → 一页 slide
  @layout bullets|cards N|split left/right|table|part|end   （默认 bullets）
  @sub 右侧黑色副标题          → 标题条右侧的补充说明
  @img 路径 | 图注             → split 版式的图片（cover-crop 填充；缺失时保留图槽占位）
  ### 卡片标题                → cards 版式的卡片（后接 - 要点）
  - 要点 / 缩进两空格 - 子要点   → 两级 bullet
  | a | b |                   → table 版式（markdown 表格）
  > 备注                      → 演讲者备注

v2.1 修复（相对 v2）：
  - 封面去重：显式 `## P1｜封面` 视为封面页，不再额外生成重复封面；
  - @layout 健壮解析：`cards N` 的 N 与 `split left|right` 的侧向均被保留；
  - split 侧向不再丢失；cards 尊重 ncol（支持多行排布）；
  - 渲染期溢出闸门：按字宽估算文本高度，超框即告警；
  - 字体安全回退：微软雅黑 → PingFang SC → Arial/Helvetica（不抑制任何脚本类型）；
  - 图片 cover-crop：split 图片铺满目标框并裁切溢出，而非留白；
  - split 缺失图片：显式 @img 但文件缺失时保留清晰图槽占位，无 @img 才退化为 bullets；
  - 表格健壮回退：行/列不齐自动补齐，空表回退 bullets；
  - 页码：直接使用 md 的 P 编号，与源文档一致。

用法：python3 md2pptx_vivo.py 输入.md 输出.pptx
依赖：python-pptx、Pillow（见 requirements.txt）
"""
import re
import os
import sys
import glob

try:
    from pptx import Presentation
    from pptx.util import Inches, Pt
    from pptx.dml.color import RGBColor
    from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.oxml.ns import qn
except ImportError:
    sys.exit("缺少依赖：python3 -m pip install --user -r requirements.txt")

# ---------- 设计令牌（VI 3.1 速查 + 模板提炼） ----------
BLUE   = '1E46E6'   # 品牌蓝
NAVY   = '06175E'   # 深蓝
LIGHT  = 'D1EBFE'   # 浅蓝
CARD   = 'EAF2FF'   # 卡片浅蓝底
INK    = '111111'
GRAY   = '565656'
LGRAY  = '9AA3B2'
RED    = 'E6001E'
WHITE  = 'FFFFFF'
GRAD_A = '3458F6'   # 封面渐变亮端（2026 参考页 house blue）
GRAD_B = '1E46E6'   # 封面渐变深端（VI 3.1 品牌蓝）
META_C = 'C9D8FF'   # 封面左下小字
TABLE_ALT = 'F2F7FF'  # 表格隔行

# 字体安全回退链（不抑制任何脚本类型：latin/ea/cs 均写入同一 typeface）
FONT_CHAIN = ['微软雅黑', 'PingFang SC', 'Arial', 'Helvetica']

LOGO = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                    'vivo-deck', 'vivo_wordmark_white.png')

SW, SH = 13.333, 7.5

# 渲染期溢出告警收集（供 build 汇总输出）
OVERFLOW_WARNINGS = []


def C(hexstr):
    return RGBColor.from_string(hexstr)


# ---------- 字体回退 ----------
def _font_available(name):
    """按平台探测系统字体是否存在，覆盖 macOS 系统字体资产目录。"""
    if sys.platform == 'darwin':
        if name == 'PingFang SC':
            candidates = [
                '/System/Library/Fonts/PingFang.ttc',
                os.path.expanduser('~/Library/Fonts/PingFang.ttc'),
                '/Library/Fonts/PingFang.ttc',
            ]
            candidates.extend(glob.glob('/System/Library/AssetsV2/**/PingFang*.ttc', recursive=True))
            candidates.extend(glob.glob(os.path.expanduser('~/Library/Application Support/*/PingFang*.ttc')))
            return any(os.path.exists(path) for path in candidates)
        if name == '微软雅黑':
            return False  # macOS 默认无微软雅黑
        if name == 'Arial':
            return (os.path.exists('/Library/Fonts/Arial.ttf')
                    or os.path.exists('/System/Library/Fonts/Supplemental/Arial.ttf'))
        if name == 'Helvetica':
            return os.path.exists('/System/Library/Fonts/Helvetica.ttc')
    elif sys.platform.startswith('win'):
        if name == '微软雅黑':
            return os.path.exists(r'C:\Windows\Fonts\msyh.ttc')
        if name == 'PingFang SC':
            return False
        if name == 'Arial':
            return os.path.exists(r'C:\Windows\Fonts\arial.ttf')
        if name == 'Helvetica':
            return False
    return True  # 未知平台：默认可用


def resolve_font():
    """按回退链返回第一个可用的字体名。"""
    for name in FONT_CHAIN:
        if _font_available(name):
            return name
    # macOS 的 PingFang 可能由 CoreText 提供但不暴露为固定路径；
    # 中文内容宁可声明 PingFang，也不要回退到没有 CJK 字形的 Arial。
    if sys.platform == 'darwin':
        return 'PingFang SC'
    return FONT_CHAIN[0]


FONT = resolve_font()


def set_font(run, name=FONT):
    """为 run 写入 latin/ea/cs 三种脚本的 typeface（不抑制任何类型）。"""
    run.font.name = name
    rPr = run._r.get_or_add_rPr()
    for tag in ('a:latin', 'a:ea', 'a:cs'):
        el = rPr.find(qn(tag))
        if el is None:
            el = rPr.makeelement(qn(tag), {})
            rPr.append(el)
        el.set('typeface', name)


# ---------- 基础形状 ----------
def alpha(shape, pct):
    """给纯色填充加透明度，pct=透明百分比(0-100)，如 90 表示 10% 不透明。"""
    clr = shape.fill._xPr.find(qn('a:solidFill')).find(qn('a:srgbClr'))
    clr.append(clr.makeelement(qn('a:alpha'), {'val': str(int((100 - pct) * 1000))}))


def textbox(slide, l, t, w, h, runs, size, color, bold=False, align=PP_ALIGN.LEFT,
            lh=None, anchor=MSO_ANCHOR.TOP, space_after=None):
    """runs: str 或 [(text,color,bold),...]"""
    tb = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    p = tf.paragraphs[0]
    p.alignment = align
    if lh:
        p.line_spacing = lh
    if space_after:
        p.space_after = Pt(space_after)
    if isinstance(runs, str):
        runs = [(runs, color, bold)]
    for txt, col, bd in runs:
        r = p.add_run()
        r.text = txt
        r.font.size = Pt(size)
        r.font.bold = bd
        r.font.color.rgb = C(col)
        set_font(r)
    return tb


def bar(slide, l, t, w, h, color):
    sh = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                Inches(l), Inches(t), Inches(w), Inches(h))
    sh.fill.solid()
    sh.fill.fore_color.rgb = C(color)
    sh.line.fill.background()
    sh.shadow.inherit = False
    try:
        sh.adjustments[0] = 0.12
    except Exception:
        pass
    return sh


# ---------- 文本估算（溢出闸门） ----------
def cjk_w(text, size_pt):
    """估算文本宽度(in)：CJK 按字号宽，拉丁按 0.55 倍。"""
    w = 0
    for ch in text:
        w += size_pt if ord(ch) > 0x2E00 else size_pt * 0.55
    return w / 72.0


def est_lines(text, size_pt, width_in):
    """按字宽估算文本在给定宽度下占用的行数。"""
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


def est_bullets_height(bullets, l1=12.0, width_in=12.1):
    """估算两级 bullet 列表的渲染高度(in)。"""
    total = 0.0
    for lv, txt in bullets:
        size = l1 if lv == 0 else l1 - 1
        lines = est_lines(txt, size, width_in)
        total += lines * size * 1.25 / 72.0
        total += (8 if lv == 0 else 4) / 72.0
    return total


def _warn_overflow(page, label, est_h, box_h):
    if est_h > box_h:
        OVERFLOW_WARNINGS.append(
            f'P{page["no"]}「{page["title"]}」{label} 估算高度 {est_h:.2f}in > 框高 {box_h:.2f}in')


# ---------- 解析 ----------
SLIDE_RE = re.compile(r'^##\s+P(\d+)[｜|]\s*(.+)$')


def parse_md(path):
    deck = {'title': None, 'sub': None, 'meta': None, 'pages': []}
    cur = None
    for raw in open(path, encoding='utf-8'):
        line = raw.rstrip('\n')
        m = re.search(r'<!--\s*deck:\s*(.+?)\s*-->', line)
        if m and cur is None:
            body = m.group(1)
            if body.startswith('meta:'):
                deck['meta'] = body[5:].strip()
            else:
                deck['sub'] = body
            continue
        if line.startswith('# ') and cur is None:
            deck['title'] = line[2:].strip()
            continue
        m = SLIDE_RE.match(line)
        if m:
            cur = {'no': int(m.group(1)), 'title': m.group(2).strip(),
                   'layout': None, 'layout_arg': None, 'sub': None,
                   'img': None, 'imgcap': None,
                   'bullets': [], 'cards': [], 'table': [], 'notes': []}
            deck['pages'].append(cur)
            continue
        if cur is None:
            continue
        if line.startswith('@layout'):
            parts = line.split()
            cur['layout'] = parts[1] if len(parts) > 1 else 'bullets'
            cur['layout_arg'] = parts[2] if len(parts) > 2 else None
            continue
        if line.startswith('@part'):
            cur['layout'] = 'part'
            continue
        if line.startswith('@sub'):
            cur['sub'] = line[4:].strip()
            continue
        if line.startswith('@img'):
            rest = line[4:].strip()
            if '|' in rest:
                p, cap = rest.split('|', 1)
                cur['img'], cur['imgcap'] = p.strip(), cap.strip()
            else:
                cur['img'] = rest
            continue
        if line.startswith('### '):
            cur['cards'].append({'title': line[4:].strip(), 'items': []})
            continue
        m = re.match(r'^>\s*(.*)$', line)
        if m and m.group(1):
            cur['notes'].append(m.group(1))
            continue
        if line.strip().startswith('|'):
            cells = [c.strip() for c in line.strip().strip('|').split('|')]
            if not all(set(c) <= set('-: ') for c in cells):
                cur['table'].append(cells)
            continue
        m2 = re.match(r'^\s{2,}-\s+(.+)$', line)
        if m2:
            if cur['layout'] == 'cards' and cur['cards']:
                cur['cards'][-1]['items'].append(m2.group(1))
            else:
                cur['bullets'].append((1, m2.group(1)))
            continue
        m1 = re.match(r'^-\s+(.+)$', line)
        if m1:
            if cur['layout'] == 'cards' and cur['cards']:
                cur['cards'][-1]['items'].append(m1.group(1))
            else:
                cur['bullets'].append((0, m1.group(1)))
            continue
    return deck


# ---------- 页面渲染 ----------
def gradient_bg(slide):
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(SW), Inches(SH))
    bg.fill.gradient()
    bg.fill.gradient_angle = 45
    stops = bg.fill.gradient_stops
    # 确保至少两个渐变停靠点（默认即 2 个，防御性补齐）
    while len(stops) < 2:
        stops.add_position(1.0)
    stops[0].color.rgb = C(GRAD_A)
    stops[0].position = 0.0
    stops[1].color.rgb = C(GRAD_B)
    stops[1].position = 1.0
    bg.line.fill.background()
    bg.shadow.inherit = False
    # 斜向半透明色块（模板封面的对角光带，允许出血）
    for (x, y, w, h, rot, a) in [(8.2, -2.5, 3.2, 14, 24, 88),
                                 (10.6, -2.5, 1.6, 14, 24, 92),
                                 (6.6, -2.5, 0.9, 14, 24, 94)]:
        s = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y),
                                   Inches(w), Inches(h))
        s.rotation = rot
        s.fill.solid()
        s.fill.fore_color.rgb = C(WHITE)
        s.line.fill.background()
        s.shadow.inherit = False
        alpha(s, a)
    return bg


def logo(slide, top=0.42, height=0.42):
    if os.path.exists(LOGO):
        slide.shapes.add_picture(LOGO, Inches(SW - 1.9), Inches(top), height=Inches(height))


def page_chrome(slide, no, total):
    textbox(slide, 0.55, 7.05, 6, 0.3, '内部汇报资料 · 请勿外传', 9, LGRAY)
    textbox(slide, SW - 1.4, 7.05, 0.85, 0.3, str(no), 9, LGRAY, align=PP_ALIGN.RIGHT)


def render_cover(prs, deck, cover_pg=None):
    """封面页。cover_pg 为显式 `## P1｜封面` 页时，其 bullets 作为封面附加行渲染。"""
    s = prs.slides.add_slide(prs.slide_layouts[6])
    gradient_bg(s)
    logo(s)
    textbox(s, 0.9, 2.35, 10.5, 2.2, deck['title'] or '未命名', 40, WHITE, bold=True, lh=1.15)
    if deck['sub']:
        textbox(s, 0.92, 4.55, 10.5, 0.8, deck['sub'], 20, LIGHT)
    if deck['meta']:
        textbox(s, 0.92, 6.55, 8, 0.4, deck['meta'], 9, META_C)
    # 显式封面页的要点（提报方/日期/密级等）渲染为封面下方小字
    if cover_pg:
        y = 5.6
        for lv, txt in cover_pg['bullets']:
            textbox(s, 0.92, y, 10.5, 0.35, txt, 9, META_C)
            y += 0.38
    return s


def render_end(s, title='Thank you'):
    gradient_bg(s)
    logo(s)
    textbox(s, 0.9, 3.0, 10, 1.2, title, 40, WHITE, bold=True)
    textbox(s, 0.92, 4.1, 10, 0.6, '谢谢', 20, LIGHT)


def render_toc(s, page):
    textbox(s, 0.7, 0.55, 4, 0.6, '目录', 20, INK, bold=True)
    y = 1.7
    for (lv, txt) in page['bullets']:
        parts = txt.split(' ', 1)
        num, label = (parts[0], parts[1]) if len(parts) == 2 and parts[0].isdigit() else ('', txt)
        runs = []
        if num:
            runs.append((num, BLUE, True))
        runs.append(('  ' + label if num else label, INK, False))
        textbox(s, 1.0, y, 10.5, 0.5, runs, 16 if num else 15, INK)
        bar(s, 1.0, y + 0.52, 4.6, 0.035, BLUE)
        y += 0.85


def render_part(s, page):
    m = re.search(r'(\d+)', page['title'])
    num = m.group(1) if m else ''
    title_txt = re.sub(r'^Part\s*\d*\s*', '', page['title'], flags=re.I).strip() or page['title']
    textbox(s, 0.9, 2.5, 3, 0.7, [('Part ', INK, True), (num, BLUE, True)], 28, INK)
    bar(s, 0.92, 3.35, 2.2, 0.04, BLUE)
    textbox(s, 0.9, 3.6, 11, 0.9, title_txt, 28, BLUE, bold=True)
    if page['sub']:
        textbox(s, 0.92, 4.5, 11, 0.5, page['sub'], 12, GRAY)


def header(s, page):
    t = page['title']
    w = min(cjk_w(t, 14) + 0.55, 8.5)
    bar(s, 0.55, 0.42, w, 0.44, BLUE)
    textbox(s, 0.78, 0.455, w - 0.4, 0.4, t, 14, WHITE, bold=True, anchor=MSO_ANCHOR.MIDDLE)
    if page['sub']:
        textbox(s, 0.6 + w + 0.25, 0.47, 12 - w, 0.4, page['sub'], 12, INK, bold=True)


def bullet_paras(tf, bullets, l1=12.0, start_first=True):
    first = start_first
    for lv, txt in bullets:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.line_spacing = 1.25
        if lv == 0:
            if '：' in txt:
                lead, rest = txt.split('：', 1)
                for ttxt, col, bd in [('• ', BLUE, True), (lead + '：', NAVY, True), (rest, INK, False)]:
                    r = p.add_run()
                    r.text = ttxt
                    r.font.size = Pt(l1)
                    r.font.bold = bd
                    r.font.color.rgb = C(col)
                    set_font(r)
            else:
                r = p.add_run()
                r.text = '• ' + txt
                r.font.size = Pt(l1)
                r.font.bold = True
                r.font.color.rgb = C(INK)
                set_font(r)
            p.space_after = Pt(8)
        else:
            r = p.add_run()
            r.text = '   – ' + txt
            r.font.size = Pt(l1 - 1)
            r.font.color.rgb = C(GRAY)
            set_font(r)
            p.space_after = Pt(4)


def render_bullets(s, page):
    header(s, page)
    box_l, box_t, box_w, box_h = 0.62, 1.35, 12.1, 5.4
    _warn_overflow(page, 'bullets', est_bullets_height(page['bullets'], 12, box_w), box_h)
    tb = s.shapes.add_textbox(Inches(box_l), Inches(box_t), Inches(box_w), Inches(box_h))
    tf = tb.text_frame
    tf.word_wrap = True
    bullet_paras(tf, page['bullets'])


def render_cards(s, page, ncol=None):
    header(s, page)
    cards = page['cards'] or [{'title': '', 'items': [t for _, t in page['bullets']]}]
    ncol = ncol or len(cards)
    ncol = max(1, min(ncol, len(cards)))
    rows = (len(cards) + ncol - 1) // ncol
    gap, margin = 0.25, 0.55
    top = 1.45
    avail_h = 5.1
    row_h = (avail_h - gap * (rows - 1)) / rows
    cw = (SW - margin * 2 - gap * (ncol - 1)) / ncol
    for i, cd in enumerate(cards):
        r, c = i // ncol, i % ncol
        x = margin + c * (cw + gap)
        y = top + r * (row_h + gap)
        bar(s, x, y, cw, row_h, CARD)
        textbox(s, x + 0.22, y + 0.2, cw - 0.44, 0.5, cd['title'], 13, BLUE, bold=True)
        bar(s, x + 0.22, y + 0.67, 0.5, 0.03, BLUE)
        tb = s.shapes.add_textbox(Inches(x + 0.22), Inches(y + 0.85),
                                  Inches(cw - 0.44), Inches(row_h - 1.0))
        tf = tb.text_frame
        tf.word_wrap = True
        first = True
        for it in cd['items']:
            p = tf.paragraphs[0] if first else tf.add_paragraph()
            first = False
            p.line_spacing = 1.2
            p.space_after = Pt(5)
            r = p.add_run()
            r.text = '• ' + it
            r.font.size = Pt(10.5)
            r.font.color.rgb = C(INK)
            set_font(r)
        _warn_overflow(page, f'卡片{i + 1}',
                       est_bullets_height([(0, it) for it in cd['items']], 10.5, cw - 0.44),
                       row_h - 1.0)


def add_cover_crop(slide, path, x, y, w, h):
    """图片 cover-crop：铺满目标框并裁切溢出（而非留白）。"""
    from PIL import Image as PILImage
    iw, ih = PILImage.open(path).size
    src_ar = iw / ih
    tgt_ar = w / h
    pic = slide.shapes.add_picture(path, Inches(x), Inches(y),
                                   width=Inches(w), height=Inches(h))
    if src_ar > tgt_ar:
        keep = tgt_ar / src_ar
        c = (1 - keep) / 2
        pic.crop_left = c
        pic.crop_right = c
    else:
        keep = src_ar / tgt_ar
        c = (1 - keep) / 2
        pic.crop_top = c
        pic.crop_bottom = c
    return pic


def add_missing_image_slot(slide, path, cap, x, y, w, h):
    """显式 @img 缺失时保留 split 图槽，避免误退化为全宽文字。"""
    slot = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y),
                                  Inches(w), Inches(h))
    slot.fill.solid()
    slot.fill.fore_color.rgb = C(CARD)
    slot.line.color.rgb = C(BLUE)
    slot.line.width = Pt(1.25)
    slot.shadow.inherit = False
    textbox(slide, x + 0.35, y + h / 2 - 0.42, w - 0.7, 0.32,
            'MISSING IMAGE', 13, BLUE, bold=True, align=PP_ALIGN.CENTER,
            anchor=MSO_ANCHOR.MIDDLE)
    missing_name = os.path.basename(path) if path else 'empty @img path'
    textbox(slide, x + 0.35, y + h / 2 + 0.02, w - 0.7, 0.36,
            missing_name, 8.5, GRAY, align=PP_ALIGN.CENTER,
            anchor=MSO_ANCHOR.MIDDLE)
    if cap:
        textbox(slide, x, y + h + 0.08, w, 0.3, cap, 9, LGRAY,
                align=PP_ALIGN.CENTER)
    return slot


def render_split(s, page, side='left'):
    header(s, page)
    right = (side == 'right')
    img_w = 5.4
    img_h = 5.0
    if page['img']:
        x = SW - 0.6 - img_w if right else 0.6
        if os.path.exists(page['img']):
            add_cover_crop(s, page['img'], x, 1.5, img_w, img_h)
            if page['imgcap']:
                textbox(s, x, 1.5 + img_h + 0.08, img_w, 0.3, page['imgcap'], 9, LGRAY,
                        align=PP_ALIGN.CENTER)
        else:
            add_missing_image_slot(s, page['img'], page['imgcap'], x, 1.5, img_w, img_h)
        bx, bw = (0.6, x - 1.0) if right else (0.6 + img_w + 0.4, SW - 1.2 - img_w)
    else:
        bx, bw = 0.6, 12.1
    _warn_overflow(page, 'split bullets', est_bullets_height(page['bullets'], 11.5, bw), 5.4)
    tb = s.shapes.add_textbox(Inches(bx), Inches(1.35), Inches(bw), Inches(5.4))
    tf = tb.text_frame
    tf.word_wrap = True
    bullet_paras(tf, page['bullets'], l1=11.5)


def render_table(s, page):
    header(s, page)
    rows = page['table']
    if not rows:
        # 空表回退 bullets
        render_bullets(s, page)
        return
    nc = max(len(r) for r in rows)
    # 补齐不齐的行
    rows = [r + [''] * (nc - len(r)) for r in rows]
    nr = len(rows)
    gt = s.shapes.add_table(nr, nc, Inches(0.55), Inches(1.4), Inches(12.2), Inches(0.4 * nr))
    tbl = gt.table
    for i in range(nc):
        tbl.columns[i].width = Inches(12.2 / nc)
    for ri in range(nr):
        tbl.rows[ri].height = Inches(0.45 if ri == 0 else 0.55)
    for ri, row in enumerate(rows):
        for ci in range(nc):
            cell = tbl.cell(ri, ci)
            cell.text = row[ci]
            p = cell.text_frame.paragraphs[0]
            r = p.runs[0] if p.runs else p.add_run()
            r.font.size = Pt(10.5)
            set_font(r)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            if ri == 0:
                cell.fill.solid()
                cell.fill.fore_color.rgb = C(BLUE)
                r.font.color.rgb = C(WHITE)
                r.font.bold = True
            else:
                cell.fill.solid()
                cell.fill.fore_color.rgb = C(TABLE_ALT) if ri % 2 == 0 else C(WHITE)
                r.font.color.rgb = C(INK)
                if ci == 0:
                    r.font.bold = True
                    r.font.color.rgb = C(NAVY)
    # 表格高度估算（表头0.45 + 数据行0.55）
    est_h = 0.45 + (nr - 1) * 0.55
    _warn_overflow(page, 'table', est_h, 5.4)


def build(deck, out):
    prs = Presentation()
    prs.slide_width = Inches(SW)
    prs.slide_height = Inches(SH)
    blank = prs.slide_layouts[6]

    # 封面去重：显式 P1｜封面 视为封面页，不再额外生成重复封面
    pages = list(deck['pages'])
    cover_pg = None
    if pages and pages[0]['title'].strip() == '封面':
        cover_pg = pages.pop(0)
    render_cover(prs, deck, cover_pg)

    total = len(pages) + 1  # 含封面
    for pg in pages:
        lay = pg['layout'] or ('toc' if pg['title'] == '目录' else
                               ('end' if 'Thank' in pg['title'] or '谢谢' in pg['title'] else 'bullets'))
        s = prs.slides.add_slide(blank)
        if lay == 'end':
            render_end(s, pg['title'] if pg['title'] not in ('Thank you', '谢谢') else 'Thank you')
        elif lay == 'toc':
            render_toc(s, pg)
        elif lay == 'part':
            render_part(s, pg)
        elif lay == 'cards':
            ncol = int(pg['layout_arg']) if pg['layout_arg'] and pg['layout_arg'].isdigit() else None
            render_cards(s, pg, ncol)
        elif lay == 'split':
            side = pg['layout_arg'] if pg['layout_arg'] in ('left', 'right') else 'left'
            render_split(s, pg, side)
        elif lay == 'table':
            render_table(s, pg)
        else:
            render_bullets(s, pg)
        if lay not in ('end',):
            page_chrome(s, pg['no'], total)
        if pg['notes']:
            s.notes_slide.notes_text_frame.text = ' '.join(pg['notes'])
    prs.save(out)
    return len(prs.slides._sldIdLst)


def main():
    if len(sys.argv) != 3:
        sys.exit(__doc__)
    deck = parse_md(sys.argv[1])
    nos = [p['no'] for p in deck['pages']]
    if nos != list(range(1, len(nos) + 1)):
        print(f'[警告] 页码不连续: {nos}')
    for p in deck['pages']:
        l1 = sum(1 for lv, _ in p['bullets'] if lv == 0)
        if l1 > 6 and not p['layout']:
            print(f'[警告] P{p["no"]} 一级要点 {l1} 条，超过6条')
    n = build(deck, sys.argv[2])
    print(f'[完成] {sys.argv[2]} 共 {n} 页（含封面）')
    if OVERFLOW_WARNINGS:
        print('[溢出] 渲染期文本高度估算超框：')
        for w in OVERFLOW_WARNINGS:
            print('  - ' + w)
    print(f'[字体] 回退链解析为: {FONT}')


if __name__ == '__main__':
    main()
