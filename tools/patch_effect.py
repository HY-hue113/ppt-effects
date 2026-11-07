#!/usr/bin/env python3
"""把 ppt-master 导出的两页效果稿校正成 spec.json 里的精确参数。

用法:
    python patch_effect.py --in 导出.pptx --spec spec.json --out 定稿.pptx [--text 新标题]

做四件事:
  1. 第 1 页的 reveal 图与标题按 spec.layers[].start 移出画布
  2. 第 2 页按 spec.layers[].end 精确对齐（EMU 原值）
  3. 给 reveal 图补 outerShdw，给标题补字体/字号/颜色
  4. 第 N 页写入 morph 切换（其他页不加）
"""
import argparse
import json
import zipfile
from lxml import etree

A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006'
P159 = 'http://schemas.microsoft.com/office/powerpoint/2015/09/main'
NS = {'a': A, 'p': P}

TRANSITION = (
    '<mc:AlternateContent xmlns:mc="%s" xmlns:p="%s" xmlns:p159="%s">'
    '<mc:Choice Requires="p159"><p:transition spd="%s"><p159:morph option="%s"/></p:transition></mc:Choice>'
    '<mc:Fallback><p:transition spd="%s"><p:fade/></p:transition></mc:Fallback>'
    '</mc:AlternateContent>'
)


def shadow_xml(sh):
    return (
        '<a:effectLst xmlns:a="%s"><a:outerShdw blurRad="%d" dist="%d" dir="%d" algn="%s" rotWithShape="%d">'
        '<a:prstClr val="%s"><a:alpha val="%d"/></a:prstClr></a:outerShdw></a:effectLst>'
        % (A, sh['blurRad'], sh['dist'], sh['dir'], sh['algn'], sh.get('rotWithShape', 0),
           sh['prstClr'], sh['alpha'])
    )


def set_xfrm(shape, emu):
    xfrm = shape.find('.//a:xfrm', NS)
    off = xfrm.find('a:off', NS)
    ext = xfrm.find('a:ext', NS)
    off.set('x', str(emu['x'])); off.set('y', str(emu['y']))
    ext.set('cx', str(emu['cx'])); ext.set('cy', str(emu['cy']))


def pics(slide):
    return [e for e in slide.iter('{%s}pic' % P)]


def text_boxes(slide):
    out = []
    for sp in slide.iter('{%s}sp' % P):
        c = sp.find('.//{%s}cNvSpPr' % P)
        if c is not None and c.get('txBox') == '1':
            out.append(sp)
    return out


def style_text(sp, layer, text, size_pt=None, align=None, anchor=None):
    rpr = sp.find('.//a:rPr', NS)
    rpr.set('sz', str(int(round((size_pt or layer['size_pt']) * 100))))
    for tag in ('latin', 'ea', 'cs'):
        el = rpr.find('a:' + tag, NS)
        if el is None:
            el = etree.SubElement(rpr, '{%s}%s' % (A, tag))
        el.set('typeface', layer['font'])
    fill = rpr.find('a:solidFill', NS)
    if fill is None:
        fill = etree.SubElement(rpr, '{%s}solidFill' % A)
    for child in list(fill):
        fill.remove(child)
    etree.SubElement(fill, '{%s}srgbClr' % A).set('val', layer.get('color', '000000'))
    body = sp.find('.//a:bodyPr', NS)
    body.set('wrap', layer.get('bodyPr', {}).get('wrap', 'square'))
    for attr in ('lIns', 'tIns', 'rIns', 'bIns'):
        if body.get(attr) is not None:
            del body.attrib[attr]
    if anchor:
        body.set('anchor', anchor)
    if align:
        for p in sp.iter('{%s}p' % A):
            ppr = p.find('a:pPr', NS)
            if ppr is None:
                ppr = etree.Element('{%s}pPr' % A)
                p.insert(0, ppr)
            ppr.set('algn', align)
    if text:
        for t in sp.iter('{%s}t' % A):
            t.text = text


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--in', dest='src', required=True)
    ap.add_argument('--spec', required=True)
    ap.add_argument('--out', dest='dst', required=True)
    ap.add_argument('--text', default=None)
    ap.add_argument('--variant', default=None,
                    help="title variant key from spec.title_variants (e.g. short)")
    args = ap.parse_args()

    spec = json.load(open(args.spec, encoding='utf-8'))
    layers = {l['role']: l for l in spec['layers']}
    reveal, title = layers['reveal'], layers['title']
    variant = (spec.get('title_variants') or {}).get(args.variant) if args.variant else None
    if args.variant and not variant:
        raise SystemExit('unknown title variant: %s' % args.variant)

    with zipfile.ZipFile(args.src) as z:
        items = {n: z.read(n) for n in z.namelist()}

    s1 = etree.fromstring(items['ppt/slides/slide1.xml'])
    s2 = etree.fromstring(items['ppt/slides/slide2.xml'])

    for slide, phase in ((s1, 'start'), (s2, 'end')):
        imgs = pics(slide)
        # 约定: 第 1 张图是背景，第 2 张是 reveal 斜切图
        if len(imgs) < 2:
            raise SystemExit('expected at least 2 pictures per slide, got %d' % len(imgs))
        rev = imgs[1]
        set_xfrm(rev, reveal[phase]['emu'])
        spPr = rev.find('.//p:spPr', NS)
        old = spPr.find('a:effectLst', NS)
        if old is not None:
            spPr.remove(old)
        if spec.get('shadow'):
            spPr.append(etree.fromstring(shadow_xml(spec['shadow'])))
        tbs = text_boxes(slide)
        if not tbs:
            raise SystemExit('no text box found')
        if variant:
            # 变体给的是"第 2 页盒子"，第 1 页沿用相同的相对位移
            end = title['end']['emu']
            start = title['start']['emu']
            box = dict(variant['box'])
            if phase == 'start':
                box['x'] += start['x'] - end['x']
                box['y'] += start['y'] - end['y']
            set_xfrm(tbs[0], box)
        else:
            set_xfrm(tbs[0], title[phase]['emu'])
        style_text(
            tbs[0], title,
            args.text if phase == 'end' else (args.text or title.get('default_text')),
            size_pt=variant.get('size_pt') if variant else None,
            align=variant.get('align') if variant else None,
            anchor=variant.get('anchor') if variant else None,
        )

    tr = spec.get('transition')
    if tr:
        page = int(tr['slide'])
        target = {1: s1, 2: s2}[page]
        for old in [e for e in target if etree.QName(e).localname == 'AlternateContent']:
            target.remove(old)
        target.append(etree.fromstring(TRANSITION % (MC, P, P159, tr.get('spd', 'slow'),
                                                     tr.get('morph_by', 'object'), tr.get('spd', 'slow'))))

    items['ppt/slides/slide1.xml'] = etree.tostring(s1, xml_declaration=True, encoding='UTF-8', standalone=True)
    items['ppt/slides/slide2.xml'] = etree.tostring(s2, xml_declaration=True, encoding='UTF-8', standalone=True)

    with zipfile.ZipFile(args.dst, 'w', zipfile.ZIP_DEFLATED) as z:
        for n, d in items.items():
            z.writestr(n, d)
    print('patched ->', args.dst)


if __name__ == '__main__':
    main()
