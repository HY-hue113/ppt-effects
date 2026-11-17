# -*- coding: utf-8 -*-
"""把导航带（y<135，含组合内无自身坐标的形状）里的「文本」改成「标题」

要点：组合内的形状常没有自己的 xfrm，位置由祖先组合决定；
     这里向上找最近的带 xfrm 的祖先来确定 y。
"""
import os
import re
import sys
import zipfile

from lxml import etree

P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
EMU = 9525


def eff_geom(el):
    """返回 (y, h)：自身无坐标时取最近的祖先，且只接受扁条（h<=150）"""
    cur = el
    while cur is not None:
        xf = cur.find('{%s}xfrm' % A) if cur is not None else None
        if xf is not None:
            off, ext = xf.find('{%s}off' % A), xf.find('{%s}ext' % A)
            if off is not None and ext is not None:
                return int(off.get('y')) // EMU, int(ext.get('cy')) // EMU
        cur = cur.getparent()
    return -1, 9999


def fix(path):
    with zipfile.ZipFile(path) as z:
        items = {n: z.read(n) for n in z.namelist()}
    n_fix = 0
    for name in [x for x in items if x.endswith('.xml')
                 and (x.startswith('ppt/slides/') or x.startswith('ppt/slideLayouts/') or x.startswith('ppt/slideMasters/'))]:
        r = etree.fromstring(items[name])
        changed = False
        for sp in r.iter('{%s}sp' % P):
            t = ''.join(x.text or '' for x in sp.iter('{%s}t' % A)).strip()
            if t != '文本':
                continue
            y, h = eff_geom(sp)
            if y < 135 and h <= 150:
                for x in sp.iter('{%s}t' % A):
                    if (x.text or '').strip() == '文本':
                        x.text = '标题'
                n_fix += 1
                changed = True
        if changed:
            items[name] = etree.tostring(r, xml_declaration=True, encoding='UTF-8', standalone=True)
    if n_fix:
        with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as zo:
            for k, b in items.items():
                zo.writestr(k, b)
    return n_fix


if __name__ == '__main__':
    for p in sys.argv[1:]:
        print('%s → 导航带修复 %d 处' % (os.path.basename(p), fix(p)))
