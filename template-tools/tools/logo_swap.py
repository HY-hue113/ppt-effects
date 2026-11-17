# -*- coding: utf-8 -*-
"""校徽 / 校名替换（母版·版式级，一次替换全篇）

设计原则（吸取教训）：
    · 只"替换"，不"删除"任何其他对象 —— 绝不按坐标批量删东西
    · 定位按语义，优先级：
        1) 文字为「校徽」的占位（模板自带的明确语义）
        2) 每个版式里左上角**面积最大**的候选（真正的 logo 位；小图标不会被选中）
    · 深色底自动用反白版，浅色底用彩色版
    · 等比缩放居中，不变形

用法：
    python logo_swap.py <模板.pptx> <彩色校徽> [学校名称] [反白校徽]
"""
import os
import re
import sys
import zipfile

from lxml import etree
from PIL import Image

P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PKG_R = 'http://schemas.openxmlformats.org/package/2006/relationships'
IMG_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'
EMU = 9525


def rels_name(part):
    d, f = os.path.split(part)
    return '%s/_rels/%s.rels' % (d, f)


def box_of(el):
    xf = el.find('.//{%s}xfrm' % A)
    if xf is None:
        return None
    off, ext = xf.find('{%s}off' % A), xf.find('{%s}ext' % A)
    if off is None or ext is None:
        return None
    return (int(off.get('x')) // EMU, int(off.get('y')) // EMU,
            int(ext.get('cx')) // EMU, int(ext.get('cy')) // EMU)


def set_image(sp, box, rid, iw, ih):
    spPr = sp.find('{%s}spPr' % P)
    if spPr is None:
        return False
    for old in spPr.findall('{%s}solidFill' % A) + spPr.findall('{%s}blipFill' % A):
        spPr.remove(old)
    fill = etree.fromstring(
        '<a:blipFill xmlns:a="%s" xmlns:r="%s"><a:blip r:embed="%s"/>'
        '<a:stretch><a:fillRect/></a:stretch></a:blipFill>' % (A, R, rid))
    ln = spPr.find('{%s}ln' % A)
    if ln is not None:
        spPr.remove(ln)
    spPr.append(fill)

    # 等比缩放居中
    xf = spPr.find('{%s}xfrm' % A)
    if xf is not None:
        off, ext = xf.find('{%s}off' % A), xf.find('{%s}ext' % A)
        if off is not None and ext is not None:
            bx, by = int(off.get('x')), int(off.get('y'))
            bw, bh = int(ext.get('cx')), int(ext.get('cy'))
            scale = min(bw / iw, bh / ih)
            nw, nh = int(iw * scale), int(ih * scale)
            off.set('x', str(bx + (bw - nw) // 2))
            off.set('y', str(by + (bh - nh) // 2))
            ext.set('cx', str(nw))
            ext.set('cy', str(nh))
    return True


def add_rel(items, part, img_name):
    rn = rels_name(part)
    root = etree.fromstring(items[rn]) if rn in items else etree.fromstring(
        ('<Relationships xmlns="%s"/>' % PKG_R).encode())
    rid = 'rIdLogo%d' % (len(root) + 100)
    rel = etree.SubElement(root, '{%s}Relationship' % PKG_R)
    rel.set('Id', rid)
    rel.set('Type', IMG_REL)
    rel.set('Target', '../media/%s' % img_name)
    items[rn] = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
    return rid


def swap(pptx, logo, school=None, white_logo=None):
    with zipfile.ZipFile(pptx) as z:
        items = {n: z.read(n) for n in z.namelist()}

    def ensure_media(path, name):
        if name in items:
            return name
        items['ppt/media/' + name] = open(path, 'rb').read()
        ext = os.path.splitext(path)[1].lower().lstrip('.')
        ct = items['[Content_Types].xml'].decode('utf-8')
        if ('Extension="%s"' % ext) not in ct:
            ct = ct.replace('</Types>', '<Default Extension="%s" ContentType="image/%s"/></Types>'
                            % (ext, 'jpeg' if ext in ('jpg', 'jpeg') else ext))
            items['[Content_Types].xml'] = ct.encode('utf-8')
        return name

    col_name = ensure_media(logo, 'brand' + os.path.splitext(logo)[1].lower())
    wht_name = ensure_media(white_logo, 'brand_white' + os.path.splitext(white_logo)[1].lower()) if white_logo else None
    iw, ih = Image.open(logo).size

    parts = [n for n in items
             if re.match(r'ppt/(slideLayouts/slideLayout\d+|slideMasters/slideMaster\d+|slides/slide\d+)\.xml$', n)]
    n_logo = n_name = 0

    for part in parts:
        root = etree.fromstring(items[part])
        hit = False

        # ① 明确的「校徽」占位
        for sp in root.iter('{%s}sp' % P):
            txt = ''.join(t.text or '' for t in sp.iter('{%s}t' % A)).strip()
            if txt == '学校名称' and school:
                for t in sp.iter('{%s}t' % A):
                    if (t.text or '').strip() == '学校名称':
                        t.text = school
                hit = True
                n_name += 1
                continue
            if txt != '校徽':
                continue
            b = box_of(sp)
            if not b:
                continue
            dark = b[0] < 400
            name = wht_name if (dark and wht_name) else col_name
            rid = add_rel(items, part, name)
            if set_image(sp, b, rid, iw, ih):
                for t in sp.iter('{%s}t' % A):
                    t.text = ''
                hit = True
                n_logo += 1

        # ② 没有占位时：只挑"左上角面积最大的那个"，其余对象一律不动
        if 'slideMasters' in part and not hit and not any((''.join(t.text or '' for t in sp.iter('{%s}t' % A)).strip() in ('校徽',))
                                  for sp in root.iter('{%s}sp' % P)):
            if not hit:
                cands = []
                for el in root.iter():
                    if etree.QName(el).localname not in ('grpSp', 'pic', 'sp'):
                        continue
                    if el.find('.//{%s}t' % A) is not None:
                        continue                      # 有文字的不是 logo 位
                    b = box_of(el)
                    if not b:
                        continue
                    x, y, w, h = b
                    if x > 420 or y > 230 or w > 300 or h > 300 or w < 60 or h < 40:
                        continue
                    cands.append((w * h, el, b))
                if cands:
                    cands.sort(key=lambda t: -t[0])
                    area, el, b = cands[0]            # 面积最大 = 真 logo 位
                    dark = b[0] < 400
                    name = wht_name if (dark and wht_name) else col_name
                    rid = add_rel(items, part, name)
                    if set_image(el, b, rid, iw, ih):
                        hit = True
                        n_logo += 1

        if hit:
            items[part] = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)

    with zipfile.ZipFile(pptx, 'w', zipfile.ZIP_DEFLATED) as zo:
        for k, b in items.items():
            zo.writestr(k, b)
    return n_logo, n_name


if __name__ == '__main__':
    src, logo = sys.argv[1], sys.argv[2]
    school = sys.argv[3] if len(sys.argv) > 3 and sys.argv[3] else None
    white = sys.argv[4] if len(sys.argv) > 4 and sys.argv[4] else None
    a, b = swap(src, logo, school, white)
    print('校徽替换 %d 处；学校名称替换 %d 处（未删除任何其他对象）' % (a, b))
