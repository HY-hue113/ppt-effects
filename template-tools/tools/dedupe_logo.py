# -*- coding: utf-8 -*-
"""删除左侧栏中部多余校徽（保留靠上那处）

规则：x < 400 且 160 ≤ y ≤ 460 的图片填充形状视为"中部多余校徽"，删除。
封面右上角（x > 900）与栏顶那处不受影响。
"""
import re
import sys
import zipfile

from lxml import etree

P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
EMU = 9525


def clean(path):
    with zipfile.ZipFile(path) as z:
        items = {n: z.read(n) for n in z.namelist()}
    total = 0
    for n in [x for x in items
              if re.match(r'ppt/(slideLayouts/slideLayout\d+|slideMasters/slideMaster\d+|slides/slide\d+)\.xml$', x)]:
        root = etree.fromstring(items[n])
        removed = 0
        for sp in list(root.iter('{%s}sp' % P)):
            if sp.find('.//{%s}blipFill' % A) is None:
                continue
            xf = sp.find('.//{%s}xfrm' % A)
            if xf is None:
                continue
            off = xf.find('{%s}off' % A)
            if off is None:
                continue
            x = int(off.get('x')) // EMU
            y = int(off.get('y')) // EMU
            if x < 400 and 160 <= y <= 460:
                sp.getparent().remove(sp)
                removed += 1
        if removed:
            total += removed
            items[n] = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
    if total:
        with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as zo:
            for k, b in items.items():
                zo.writestr(k, b)
    return total


if __name__ == '__main__':
    print('删除左侧中部多余校徽 %d 处' % clean(sys.argv[1]))
