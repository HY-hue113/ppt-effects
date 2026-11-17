# -*- coding: utf-8 -*-
"""精确修：
1) 第 6 页（目录页）：左侧大标题 → 「目录」，右侧条目 → 「标题」
2) 所有部件（含母版/版式）里位于左侧/顶部的「文本」→「标题」
"""
import os
import re
import shutil
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


def write_txt(sp, text):
    ts = list(sp.iter('{%s}t' % A))
    if not ts:
        return False
    ts[0].text = text
    for t in ts[1:]:
        t.text = ''
    return True


def fix(path):
    with zipfile.ZipFile(path) as z:
        items = {n: z.read(n) for n in z.namelist()}
        toc_part = order(z)[TOC_PAGE - 1]

    toc = 0
    root = etree.fromstring(items[toc_part])
    for sp in root.iter('{%s}sp' % P):
        b = box(sp)
        if not b:
            continue
        txt = ''.join(t.text or '' for t in sp.iter('{%s}t' % A)).strip()
        if not txt:
            continue
        if b[0] < 500 and b[2] > 150 and b[3] > 60:
            write_txt(sp, '目录')
            toc += 1
        elif b[0] > 600 and b[2] > 200:
            m = re.match(r'^\s*(\d{1,2})\s*', txt)      # 保留数字序号
            write_txt(sp, ('%s 标题' % m.group(1)) if m else '标题')
            toc += 1
    items[toc_part] = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)

    nav = 0
    for name in [x for x in items if x.endswith('.xml')
                 and (x.startswith('ppt/slides/') or x.startswith('ppt/slideLayouts/') or x.startswith('ppt/slideMasters/'))]:
        r = etree.fromstring(items[name])
        changed = False
        for sp in r.iter('{%s}sp' % P):
            txt = ''.join(t.text or '' for t in sp.iter('{%s}t' % A)).strip()
            if txt != '文本':
                continue
            b = box(sp)
            if not b:
                continue
            is_layout = name.startswith(('ppt/slideLayouts/', 'ppt/slideMasters/'))
            if b[1] < 120 or (is_layout and b[0] < 160 and b[2] < 160):
                write_txt(sp, '标题')
                nav += 1
                changed = True
        if changed:
            items[name] = etree.tostring(r, xml_declaration=True, encoding='UTF-8', standalone=True)

    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as zo:
        for k, b in items.items():
            zo.writestr(k, b)
    return toc, nav


if __name__ == '__main__':
    for p in sys.argv[1:]:
        t, n = fix(p)
        print('%s → 目录 %d 处；导航 %d 处' % (os.path.basename(p), t, n))
