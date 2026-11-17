# -*- coding: utf-8 -*-
"""只做目录页：左侧「目录」+ 右侧「01 标题」…（与版1一致），不动其他页"""
import os
import re
import sys
import zipfile

from lxml import etree

P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
EMU = 9525
TOC_PAGE = 6


def order(z):
    pres = etree.fromstring(z.read('ppt/presentation.xml'))
    rels = etree.fromstring(z.read('ppt/_rels/presentation.xml.rels'))
    m = {r.get('Id'): r.get('Target') for r in rels}
    return ['ppt/' + m[s.get('{%s}id' % R)].replace('../', '') for s in pres.find('{%s}sldIdLst' % P)]


def box(sp):
    xf = sp.find('.//{%s}xfrm' % A)
    if xf is None:
        return None
    off, ext = xf.find('{%s}off' % A), xf.find('{%s}ext' % A)
    if off is None or ext is None:
        return None
    return (int(off.get('x')) // EMU, int(off.get('y')) // EMU,
            int(ext.get('cx')) // EMU, int(ext.get('cy')) // EMU)


def fix(path):
    with zipfile.ZipFile(path) as z:
        items = {n: z.read(n) for n in z.namelist()}
        part = order(z)[TOC_PAGE - 1]
    root = etree.fromstring(items[part])
    before = len([t for t in root.iter('{%s}t' % A) if (t.text or '').strip()])
    items_list = []
    for sp in root.iter('{%s}sp' % P):
        b = box(sp)
        if not b:
            continue
        runs = [t for t in sp.iter('{%s}t' % A)]
        if not runs:
            continue
        if b[0] < 500 and b[2] > 150 and b[3] > 60:
            for t in runs:
                if (t.text or '').strip():
                    t.text = '目录'
                    break
        elif b[0] > 600 and b[2] > 200:
            items_list.append((b[1], runs))
    items_list.sort(key=lambda v: v[0])
    for i, (_, runs) in enumerate(items_list, 1):
        for t in runs:
            if (t.text or '').strip():
                t.text = '%02d 标题' % i
                break
    after = len([t for t in root.iter('{%s}t' % A) if (t.text or '').strip()])
    items[part] = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as zo:
        for k, b in items.items():
            zo.writestr(k, b)
    return len(items_list), before, after


if __name__ == '__main__':
    for p in sys.argv[1:]:
        n, b, a = fix(p)
        print('%s → 目录条目 %d 个；run %d→%d %s' % (os.path.basename(p), n, b, a, 'OK' if a == b else '⚠'))
