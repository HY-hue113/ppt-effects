# -*- coding: utf-8 -*-
"""只改版6左侧导航标签：文本 → 标题（按左栏窄框特征，不动其他）"""
import sys
import zipfile

from lxml import etree

P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
EMU = 9525


def fix(path):
    with zipfile.ZipFile(path) as z:
        items = {n: z.read(n) for n in z.namelist()}
    n = 0
    for name in [x for x in items if x.endswith('.xml')
                 and (x.startswith('ppt/slides/') or x.startswith('ppt/slideLayouts/')
                      or x.startswith('ppt/slideMasters/'))]:
        r = etree.fromstring(items[name])
        changed = False
        for sp in r.iter('{%s}sp' % P):
            t = ''.join(x.text or '' for x in sp.iter('{%s}t' % A)).strip()
            if t != '文本':
                continue
            xf = sp.find('.//{%s}xfrm' % A)
            if xf is None:
                continue
            off, ext = xf.find('{%s}off' % A), xf.find('{%s}ext' % A)
            if off is None or ext is None:
                continue
            x = int(off.get('x')) // EMU
            w = int(ext.get('cx')) // EMU
            h = int(ext.get('cy')) // EMU
            if x < 230 and 140 <= w <= 210 and h <= 70:
                for e in sp.iter('{%s}t' % A):
                    if (e.text or '').strip() == '文本':
                        e.text = '标题'
                n += 1
                changed = True
        if changed:
            items[name] = etree.tostring(r, xml_declaration=True, encoding='UTF-8', standalone=True)
    if n:
        with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as zo:
            for k, b in items.items():
                zo.writestr(k, b)
    return n


if __name__ == '__main__':
    for p in sys.argv[1:]:
        print('%s → 左侧导航标签改标题 %d 处' % (p.split('\\')[-1], fix(p)))
