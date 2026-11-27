# -*- coding: utf-8 -*-
"""排版引擎 v2：断行 + 字号自适应（下限 20pt）+ 允许行数 + 加高容器

对外函数：
    break_lines(text, width_px, pt)                     → 行列表
    fit_block(text, w, h, max_pt, min_pt, max_lines)     → dict(pt, lines, need_h, overflow, action)
    process_pptx(src, dst, entries, grow=True)           → 写回 OOXML 并返回报告

规则：
    CJK 按字断行；Latin 按词断行；数字+单位不拆；行首/行尾禁则
    装不下时依次：缩字号（≥min_pt）→ 加高容器（grow）→ 报告需减字
"""
import re
import zipfile

from lxml import etree
from PIL import ImageFont

P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
EMU = 9525
FONT = r'<ABS_PATH>'

NO_LINE_START = set('，。、；：！？）》」』】〕．,.;:!?)]}%”’')
NO_LINE_END = set('（《「『【〔([{“‘')
UNBREAKABLE = re.compile(r'(\d+(?:[.,]\d+)?\s*(?:%|pt|px|元|人|页|场|天|次|年|月|日|个|张|份|km|kg|cm|mm|GB|MB|s))',
                         re.IGNORECASE)


def load_font(pt):
    return ImageFont.truetype(FONT, int(round(pt * 4 / 3)))


def is_cjk(ch):
    return '\u4e00' <= ch <= '\u9fff' or '\u3000' <= ch <= '\u303f' or '\uff00' <= ch <= '\uffef'


def tokenize(text):
    tokens, i = [], 0
    while i < len(text):
        m = UNBREAKABLE.match(text, i)
        if m:
            tokens.append(m.group(0))
            i = m.end()
            continue
        ch = text[i]
        if is_cjk(ch) or ch in NO_LINE_START or ch in NO_LINE_END or ch == ' ':
            tokens.append(ch)
            i += 1
        else:
            j = i
            while j < len(text) and not is_cjk(text[j]) and text[j] not in ' ，。、；：！？()（）「」『』':
                j += 1
            tokens.append(text[i:j])
            i = j
    return tokens


def break_lines(text, width_px, pt):
    font = load_font(pt)
    lines, cur = [], ''
    for tk in tokenize(text.replace('\r', '')):
        if tk == '\n':
            lines.append(cur)
            cur = ''
            continue
        trial = cur + tk
        if font.getlength(trial.rstrip()) <= width_px or not cur:
            cur = trial
        else:
            if tk and tk[0] in NO_LINE_START and len(cur) > 1:
                moved = cur[-1]
                cur = cur[:-1]
                lines.append(cur.rstrip())
                cur = moved + tk
            else:
                lines.append(cur.rstrip())
                cur = '' if tk == ' ' else tk
    if cur.strip():
        lines.append(cur.rstrip())
    return lines


def line_height(pt, gap=1.35):
    return pt * 4 / 3 * gap


def fit_block(text, w, h, max_pt=32, min_pt=20, max_lines=1, gap=1.35):
    """在 max_pt→min_pt 间找最大可用字号；返回 dict

    下限约束：min_pt 不会高于 max_pt（不允许把小字号放大）
    """
    min_pt = min(min_pt, max_pt)
    pt = max_pt
    while pt >= min_pt:
        lines = break_lines(text, w, pt)
        need = len(lines) * line_height(pt, gap)
        if len(lines) <= max_lines and need <= h:
            return {'pt': pt, 'lines': lines, 'need_h': round(need), 'overflow': False, 'action': 'fit'}
        pt -= 1
    lines = break_lines(text, w, min_pt)
    need = round(len(lines) * line_height(min_pt, gap))
    return {'pt': min_pt, 'lines': lines, 'need_h': need,
            'overflow': len(lines) > max_lines or need > h, 'action': 'grow'}


# ---------- OOXML ----------
def shape_box(sp):
    xf = sp.find('.//{%s}xfrm' % A)
    if xf is None:
        return None
    off, ext = xf.find('{%s}off' % A), xf.find('{%s}ext' % A)
    return (int(off.get('x')) // EMU, int(off.get('y')) // EMU,
            int(ext.get('cx')) // EMU, int(ext.get('cy')) // EMU)


def set_box_height(sp, h):
    xf = sp.find('.//{%s}xfrm' % A)
    ext = xf.find('{%s}ext' % A)
    ext.set('cy', str(int(h * EMU)))


def set_lines(sp, lines, pt):
    tx = sp.find('.//{%s}txBody' % A)
    if tx is None:
        return False
    bodyPr = tx.find('{%s}bodyPr' % A)
    for tag in ('normAutofit', 'spAutoFit', 'noAutofit'):
        for el in bodyPr.findall('{%s}%s' % (A, tag)):
            bodyPr.remove(el)
    etree.SubElement(bodyPr, '{%s}noAutofit' % A)
    for para in tx.findall('{%s}p' % A):
        tx.remove(para)
    p = etree.SubElement(tx, '{%s}p' % A)
    pPr = etree.SubElement(p, '{%s}pPr' % A)
    pPr.set('algn', 'l')
    for i, line in enumerate(lines):
        r = etree.SubElement(p, '{%s}r' % A)
        rPr = etree.SubElement(r, '{%s}rPr' % A)
        rPr.set('lang', 'zh-CN')
        rPr.set('sz', str(int(pt * 100)))
        rPr.set('dirty', '0')
        t = etree.SubElement(r, '{%s}t' % A)
        t.text = line
        if i != len(lines) - 1:
            etree.SubElement(p, '{%s}br' % A)
    return True


def process_pptx(src, dst, entries, grow=True, pad=16):
    """entries: [(slide_part, shape_name, text, max_pt, min_pt, max_lines), ...]"""
    with zipfile.ZipFile(src) as z:
        items = {n: z.read(n) for n in z.namelist()}
    report = []
    for part, name, text, max_pt, min_pt, max_lines in entries:
        root = etree.fromstring(items[part])
        target = None
        for sp in root.iter('{%s}sp' % P):
            nv = sp.find('.//{%s}cNvPr' % P)
            if nv is not None and nv.get('name') == name:
                target = sp
                break
        if target is None:
            report.append({'shape': name, 'status': '未找到'})
            continue
        x, y, w, h = shape_box(target) or (0, 0, 300, 100)
        r = fit_block(text, w - pad * 2, h - pad * 2, max_pt, min_pt, max_lines)
        if r['overflow'] and grow:
            new_h = r['need_h'] + pad * 2
            set_box_height(target, new_h)
            r['action'] = 'grew'
            r['new_h'] = new_h
        set_lines(target, r['lines'], r['pt'])
        r.update({'shape': name, 'box_h': h})
        report.append(r)
        items[part] = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
    with zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as zo:
        for k, b in items.items():
            zo.writestr(k, b)
    return report
