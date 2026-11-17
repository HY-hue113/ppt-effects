# -*- coding: utf-8 -*-
"""只把「导航栏」里的文字恢复成「标题」（正文小标题不动）

判定（精确）：
    顶部导航：y < 120
    左侧导航：x < 160 且 宽 < 160
其余位置一律不动。

用法：python fix_nav_title.py <pptx> [...]
"""
import sys
import zipfile

from lxml import etree

P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
EMU = 9525


def fix(path):
    with zipfile.ZipFile(path) as z:
        items = {n: z.read(n) for n in z.namelist()}
    total = 0
    for name in [x for x in items if x.startswith('ppt/slides/slide') and x.endswith('.xml')]:
        root = etree.fromstring(items[name])
        changed = False
        for sp in root.iter('{%s}sp' % P):
            txt = ''.join(t.text or '' for t in sp.iter('{%s}t' % A)).strip()
            if txt != '文本':
                continue
            xf = sp.find('.//{%s}xfrm' % A)
            if xf is None:
                continue
            off, ext = xf.find('{%s}off' % A), xf.find('{%s}ext' % A)
            if off is None or ext is None:
                continue
            x = int(off.get('x')) // EMU
            y = int(off.get('y')) // EMU
            w = int(ext.get('cx')) // EMU
            if y < 120 or (x < 160 and w < 160):
                for t in sp.iter('{%s}t' % A):
                    if (t.text or '').strip() == '文本':
                        t.text = '标题'
                total += 1
                changed = True
        if changed:
            items[name] = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
    if total:
        with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as zo:
            for k, b in items.items():
                zo.writestr(k, b)
    return total


if __name__ == '__main__':
    for p in sys.argv[1:]:
        print('%s → 导航文字恢复 %d 处' % (p.split('\\')[-1], fix(p)))
