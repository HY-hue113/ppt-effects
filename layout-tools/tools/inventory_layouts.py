# -*- coding: utf-8 -*-
"""盘点版式库 15 页栅格页的构成：原生形状 vs 图片"""
import os
import zipfile

from lxml import etree

LIB = r'<LIB_ROOT>\layout-tools'
SRC = os.path.join(LIB, 'reference', '版式库-01-06.pptx')
OUTMD = os.path.join(LIB, '盘点清单.md')
P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
EMU = 9525
PAGES = [5, 6, 7, 8, 9, 10, 11, 15, 19, 21, 23, 25, 26, 27, 30]


def box(el):
    xf = el.find('.//{%s}xfrm' % A)
    if xf is None:
        xf = el.find('.//{%s}xfrm' % P)
    if xf is None:
        return None
    off = xf.find('{%s}off' % A)
    if off is None:
        off = xf.find('{%s}off' % P)
    ext = xf.find('{%s}ext' % A)
    if ext is None:
        ext = xf.find('{%s}ext' % P)
    if off is None or ext is None:
        return None
    if off is None or ext is None:
        return None
    return (int(off.get('x')) // EMU, int(off.get('y')) // EMU,
            int(ext.get('cx')) // EMU, int(ext.get('cy')) // EMU)


def main():
    z = zipfile.ZipFile(SRC)
    pres = etree.fromstring(z.read('ppt/presentation.xml'))
    rels = etree.fromstring(z.read('ppt/_rels/presentation.xml.rels'))
    m = {r.get('Id'): r.get('Target') for r in rels}
    parts = ['ppt/' + m[s.get('{%s}id' % R)].replace('../', '')
             for s in pres.find('{%s}sldIdLst' % P)]

    lines = ['# 版式库 15 页栅格页 · 构成盘点', '',
             '> 目的：分清每页「哪些能直接编辑 / 哪些是图片」，重画时**只替换图片部分**，不碰原生形状。', '']
    for idx in PAGES:
        root = etree.fromstring(z.read(parts[idx - 1]))
        sps, pics = [], []
        for sp in root.iter('{%s}sp' % P):
            b = box(sp)
            if not b:
                continue
            txt = ''.join(t.text or '' for t in sp.iter('{%s}t' % A)).strip()
            name = sp.find('.//{%s}cNvPr' % P)
            sps.append(((name.get('name') if name is not None else '')[:16], b, txt[:16]))
        for pic in root.iter('{%s}pic' % P):
            b = box(pic)
            if b:
                pics.append(b)
        area = 1280 * 720
        pic_area = sum(w * h for _, _, w, h in pics)
        kind = '整页图片' if pic_area > area * 0.6 else ('局部图片' if pics else '全原生')
        lines.append('## p%d（%s）' % (idx, kind))
        lines.append('')
        lines.append('| 类型 | 数量 |')
        lines.append('|---|---|')
        lines.append('| 原生形状 | %d |' % len(sps))
        lines.append('| 图片 | %d（合计 %d%% 版面） |' % (len(pics), round(pic_area / area * 100)))
        lines.append('')
        if sps:
            lines.append('**原生形状**')
            lines.append('')
            lines.append('| 名称 | 位置 | 尺寸 | 文本 |')
            lines.append('|---|---|---|---|')
            for nm, (x, y, w, h), tx in sps[:12]:
                lines.append('| %s | %d,%d | %d×%d | %s |' % (nm, x, y, w, h, tx))
            lines.append('')
        if pics:
            lines.append('**图片**')
            lines.append('')
            for x, y, w, h in pics:
                lines.append('- %d,%d · %d×%d' % (x, y, w, h))
            lines.append('')
    open(OUTMD, 'w', encoding='utf-8').write('\n'.join(lines))
    print('盘点清单已生成：%s' % OUTMD)
    # 控制台摘要
    for idx in PAGES:
        root = etree.fromstring(z.read(parts[idx - 1]))
        n_sp = len([1 for sp in root.iter('{%s}sp' % P) if box(sp)])
        n_pic = len([1 for pic in root.iter('{%s}pic' % P) if box(pic)])
        print('p%-2d 原生形状 %-3d 图片 %-2d' % (idx, n_sp, n_pic))


if __name__ == '__main__':
    main()
