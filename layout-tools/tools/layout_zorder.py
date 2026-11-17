# -*- coding: utf-8 -*-
"""把"大图占位"沉到底层，保证内容压在上面（z 序修正）

规则：图片占位面积 > 版面 30% → 移动到 spTree 最底层（仅次 nvGrpSpPr/grpSpPr）
"""
import os
import sys
import zipfile

from lxml import etree

P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
EMU = 9525
AREA = 1280 * 720


def fix(path, out=None):
    out = out or path
    with zipfile.ZipFile(path) as z:
        items = {n: z.read(n) for n in z.namelist()}
    moved = 0
    for name in [x for x in items if x.startswith('ppt/slides/slide') and x.endswith('.xml')]:
        root = etree.fromstring(items[name])
        spTree = root.find('.//{%s}spTree' % P)
        if spTree is None:
            continue
        kids = list(spTree)
        for sp in kids:
            nv = sp.find('.//{%s}cNvPr' % P)
            if nv is None or '图片占位' not in (nv.get('name') or ''):
                continue
            xf = sp.find('.//{%s}xfrm' % A)
            if xf is None:
                continue
            ext = xf.find('{%s}ext' % A)
            w = int(ext.get('cx')) // EMU
            h = int(ext.get('cy')) // EMU
            if w * h < AREA * 0.30:
                continue
            spTree.remove(sp)
            idx = 2 if len(spTree) >= 2 else 0
            spTree.insert(idx, sp)
            moved += 1
        items[name] = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as zo:
        for k, b in items.items():
            zo.writestr(k, b)
    return moved


if __name__ == '__main__':
    for p in sys.argv[1:]:
        print('%s → 下沉大图 %d 个' % (os.path.basename(p), fix(p)))
