# -*- coding: utf-8 -*-
"""样式集可编辑性 / 换色可行性检查

输出每份 PPT：
    页数、形状数、组合数、图片数、图片填充形状数
    颜色统计（srgbClr 频次 Top）
    可疑项：纯图片页、图片填充形状、无文字但有文本样式的图片
"""
import os
import re
import sys
import zipfile
from collections import Counter

from lxml import etree

P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'


def order_parts(z):
    pres = etree.fromstring(z.read('ppt/presentation.xml'))
    rels = etree.fromstring(z.read('ppt/_rels/presentation.xml.rels'))
    m = {r.get('Id'): r.get('Target') for r in rels}
    return ['ppt/' + m[s.get('{%s}id' % R)].replace('../', '')
            for s in pres.find('{%s}sldIdLst' % P)]


def check(path):
    z = zipfile.ZipFile(path)
    parts = order_parts(z)
    stat = dict(pages=len(parts), sp=0, grp=0, pic=0, blip_sp=0, texts=0)
    colors = Counter()
    pics_pages = []
    blip_pages = []

    for i, part in enumerate(parts, 1):
        root = etree.fromstring(z.read(part))
        sp = len(root.findall('.//{%s}sp' % P))
        grp = len(root.findall('.//{%s}grpSp' % P))
        pic = len(root.findall('.//{%s}pic' % P))
        blip_sp = 0
        for s in root.iter('{%s}sp' % P):
            if s.find('.//{%s}blipFill' % A) is not None:
                blip_sp += 1
        texts = sum(1 for t in root.iter('{%s}t' % A) if (t.text or '').strip())
        stat['sp'] += sp
        stat['grp'] += grp
        stat['pic'] += pic
        stat['blip_sp'] += blip_sp
        stat['texts'] += texts
        if pic:
            pics_pages.append((i, pic))
        if blip_sp:
            blip_pages.append((i, blip_sp))
        for c in root.iter('{%s}srgbClr' % A):
            v = (c.get('val') or '').upper()
            if v:
                colors[v] += 1
    return stat, colors, pics_pages, blip_pages


if __name__ == '__main__':
    for f in sys.argv[1:]:
        st, colors, pp, bp = check(f)
        print('=' * 60)
        print(os.path.basename(f))
        print('  页 %d | 形状 %d | 组合 %d | 图片 %d | 图片填充形状 %d | 文本条 %d'
              % (st['pages'], st['sp'], st['grp'], st['pic'], st['blip_sp'], st['texts']))
        print('  颜色 Top6:', ', '.join('%s×%d' % (c, n) for c, n in colors.most_common(6)))
        if pp:
            print('  含图片的页:', pp)
        if bp:
            print('  含图片填充的页:', bp)
