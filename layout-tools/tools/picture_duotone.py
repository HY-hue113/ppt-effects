# -*- coding: utf-8 -*-
"""给保留的两张图片加"双色调"，使其可随配色换色

做法：在 a:blip 里插入 a:duotone（暗调 + 亮调两个颜色），
      图片灰度映射到这两个颜色之间 → 改这两个值即整体换色。
"""
import os
import sys
import zipfile

from lxml import etree

P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
NS = {'a': A}


def apply(src, dst, dark, light):
    with zipfile.ZipFile(src) as z:
        items = {n: z.read(n) for n in z.namelist()}
    pres = etree.fromstring(items['ppt/presentation.xml'])
    rels = etree.fromstring(items['ppt/_rels/presentation.xml.rels'])
    R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    m = {r.get('Id'): r.get('Target') for r in rels}
    parts = ['ppt/' + m[s.get('{%s}id' % R)].replace('../', '')
             for s in pres.find('{%s}sldIdLst' % P)]
    n = 0
    for part in parts:
        root = etree.fromstring(items[part])
        changed = False
        for blip in root.iter('{%s}blip' % A):
            for old in blip.findall('{%s}duotone' % A):
                blip.remove(old)
            duo = etree.SubElement(blip, '{%s}duotone' % A)
            for v in (dark, light):
                c = etree.SubElement(duo, '{%s}srgbClr' % A)
                c.set('val', v)
            changed = True
            n += 1
        if changed:
            items[part] = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
    with zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as zo:
        for k, b in items.items():
            zo.writestr(k, b)
    return n


if __name__ == '__main__':
    src = sys.argv[1]
    dst = sys.argv[2] if len(sys.argv) > 2 else src.replace('.pptx', '-双色调.pptx')
    dark = sys.argv[3] if len(sys.argv) > 3 else '404040'
    light = sys.argv[4] if len(sys.argv) > 4 else 'F2F2F2'
    print('双色调处理 %d 张图片（%s → %s）→ %s' % (apply(src, dst, dark, light), dark, light, os.path.basename(dst)))
