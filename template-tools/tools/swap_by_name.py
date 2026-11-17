# -*- coding: utf-8 -*-
"""按「对象名称」替换指定 logo 位（不按坐标、不删除任何对象）

用法：python swap_by_name.py <pptx> <对象名> <图片>
      python swap_by_name.py x.pptx "矩形 11" logo.png
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


def swap(path, obj_name, img):
    with zipfile.ZipFile(path) as z:
        items = {n: z.read(n) for n in z.namelist()}

    name = 'objlogo' + os.path.splitext(img)[1].lower()
    items['ppt/media/' + name] = open(img, 'rb').read()
    ct = items['[Content_Types].xml'].decode('utf-8')
    ext = os.path.splitext(img)[1].lower().lstrip('.')
    if ('Extension="%s"' % ext) not in ct:
        ct = ct.replace('</Types>', '<Default Extension="%s" ContentType="image/%s"/></Types>'
                        % (ext, 'jpeg' if ext in ('jpg', 'jpeg') else ext))
        items['[Content_Types].xml'] = ct.encode('utf-8')

    iw, ih = Image.open(img).size
    hit = 0
    for part in [n for n in items
                 if re.match(r'ppt/(slideLayouts/slideLayout\d+|slideMasters/slideMaster\d+|slides/slide\d+)\.xml$', n)]:
        root = etree.fromstring(items[part])
        changed = False
        for sp in root.iter('{%s}sp' % P):
            nv = sp.find('.//{%s}cNvPr' % P)
            if nv is None or nv.get('name') != obj_name:
                continue
            spPr = sp.find('{%s}spPr' % P)
            if spPr is None:
                continue
            rn = '%s/_rels/%s.rels' % (os.path.dirname(part), os.path.basename(part))
            rels = etree.fromstring(items[rn]) if rn in items else etree.fromstring(
                ('<Relationships xmlns="%s"/>' % PKG_R).encode())
            rid = 'rIdByName%d' % (len(rels) + 300)
            rel = etree.SubElement(rels, '{%s}Relationship' % PKG_R)
            rel.set('Id', rid); rel.set('Type', IMG_REL); rel.set('Target', '../media/%s' % name)
            items[rn] = etree.tostring(rels, xml_declaration=True, encoding='UTF-8', standalone=True)

            for old in spPr.findall('{%s}solidFill' % A) + spPr.findall('{%s}blipFill' % A):
                spPr.remove(old)
            fill = etree.fromstring(
                '<a:blipFill xmlns:a="%s" xmlns:r="%s"><a:blip r:embed="%s"/>'
                '<a:stretch><a:fillRect/></a:stretch></a:blipFill>' % (A, R, rid))
            spPr.append(fill)
            # 等比居中
            xf = spPr.find('{%s}xfrm' % A)
            if xf is not None:
                off, ext2 = xf.find('{%s}off' % A), xf.find('{%s}ext' % A)
                if off is not None and ext2 is not None:
                    bx, by = int(off.get('x')), int(off.get('y'))
                    bw, bh = int(ext2.get('cx')), int(ext2.get('cy'))
                    scale = min(bw / iw, bh / ih)
                    nw, nh = int(iw * scale), int(ih * scale)
                    off.set('x', str(bx + (bw - nw) // 2))
                    off.set('y', str(by + (bh - nh) // 2))
                    ext2.set('cx', str(nw)); ext2.set('cy', str(nh))
            hit += 1
            changed = True
        if changed:
            items[part] = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)

    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as zo:
        for k, b in items.items():
            zo.writestr(k, b)
    return hit


if __name__ == '__main__':
    print('按名称替换 %d 处' % swap(sys.argv[1], sys.argv[2], sys.argv[3]))
