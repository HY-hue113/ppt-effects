# -*- coding: utf-8 -*-
"""run 级修正：目录页标题/条目 + 版式级导航标签

原则：只改 run 的文本，绝不增删 run/段落/形状；改完逐项校验 run 数量不变。

用法：python fix_toc_nav2.py <pptx> [...]
"""
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


def run_texts(sp):
    return [t for t in sp.iter('{%s}t' % A)]


def count_runs(root):
    return len([t for t in root.iter('{%s}t' % A) if (t.text or '').strip()])


def fix(path):
    with zipfile.ZipFile(path) as z:
        items = {n: z.read(n) for n in z.namelist()}
        toc_part = order(z)[TOC_PAGE - 1]

    report = {}

    # ---- 1) 目录页（run 级） ----
    root = etree.fromstring(items[toc_part])
    before = count_runs(root)
    items_list = []
    for sp in root.iter('{%s}sp' % P):
        b = box(sp)
        if not b:
            continue
        runs = run_texts(sp)
        if not runs:
            continue
        joined = ''.join(t.text or '' for t in runs).strip()
        if not joined:
            continue
        if b[0] < 500 and b[2] > 150 and b[3] > 60:
            for t in runs:                       # 左侧大标题：只改有文字的那一个 run
                if (t.text or '').strip():
                    t.text = '目录'
                    break
        elif b[0] > 600 and b[2] > 200:
            items_list.append((b[1], sp, runs))   # 右侧条目稍后按 y 排序加序号
    # 按 y 排序后写成「01 标题」「02 标题」…
    items_list.sort(key=lambda v: v[0])
    for i, (_, sp, runs) in enumerate(items_list, 1):
        done = False
        for t in runs:
            if (t.text or '').strip() in ('文本', '正文', '标题'):
                t.text = '%02d 标题' % i
                done = True
                break
        if not done and runs:
            runs[0].text = '%02d 标题' % i

    after = count_runs(root)
    report['toc_runs'] = (before, after)
    items[toc_part] = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)

    # ---- 2) 版式 / 母版里的导航标签 ----
    nav_before = nav_after = 0
    for name in [x for x in items if x.endswith('.xml') and (x.startswith('ppt/slides/slide')
                 or re.match(r'ppt/(slideLayouts/slideLayout\d+|slideMasters/slideMaster\d+)\.xml$', x))]:
        r = etree.fromstring(items[name])
        nav_before += sum(1 for t in r.iter('{%s}t' % A) if (t.text or '').strip() == '文本')
        changed = False
        for sp in r.iter('{%s}sp' % P):
            b = box(sp)
            if not b:
                continue
            is_nav_label = (b[1] < 130) or (b[0] < 220 and 140 <= b[2] <= 200 and b[3] <= 60)
            if not is_nav_label:
                continue
            for t in sp.iter('{%s}t' % A):
                if (t.text or '').strip() == '文本':
                    t.text = '标题'
                    changed = True
        if changed:
            nav_after += sum(1 for t in r.iter('{%s}t' % A) if (t.text or '').strip() == '文本')
            items[name] = etree.tostring(r, xml_declaration=True, encoding='UTF-8', standalone=True)
    report['nav'] = (nav_before, nav_after)

    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as zo:
        for k, b in items.items():
            zo.writestr(k, b)
    return report


if __name__ == '__main__':
    for p in sys.argv[1:]:
        r = fix(p)
        tb, ta = r['toc_runs']
        nb, na = r['nav']
        flag = 'OK' if ta == tb else '⚠ run 数变化'
        print('%s | 目录页 run %d→%d %s | 版式「文本」%d→%d' %
              (os.path.basename(p), tb, ta, flag, nb, na))
