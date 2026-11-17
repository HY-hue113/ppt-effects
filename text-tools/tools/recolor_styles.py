# -*- coding: utf-8 -*-
"""样式集换色：把红色/蓝色主色系映射到指定配色

用法：python recolor_styles.py <pptx> <配色名> [输出]
    配色名见 palette-tools/palettes.json（学术蓝/红/绿/紫）
"""
import os
import re
import sys
import zipfile

sys.path.insert(0, r'<LIB_ROOT>\palette-tools\tools')
from palette import get, tint, shade  # noqa: E402

from lxml import etree  # noqa: E402

P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'

RED = {  # 标题 / 文框（红系）
    'C00000': 'primary', '920000': 'dark', 'AC0000': 'dark',
    'FEF5F6': 'light', 'FCE9EB': 'light',
}
BLUE = {  # 导航栏（蓝系）
    '1843CE': 'primary', '133CA3': 'dark', '325BE7': 'primary',
    '5D7FED': 'secondary', '829CF1': 'light',
}


def build_map(palette):
    p = get(palette)
    out = {}
    for table in (RED, BLUE):
        for old, role in table.items():
            out[old] = p[role] if role in ('primary', 'dark', 'secondary', 'light') else p['primary']
    # 浅色再淡化一点，避免过重
    out['FEF5F6'] = tint(p['primary'], 0.92)
    out['FCE9EB'] = tint(p['primary'], 0.88)
    out['829CF1'] = tint(p['primary'], 0.72)
    return out


def recolor(src, palette, dst):
    m = build_map(palette)
    with zipfile.ZipFile(src) as z:
        items = {n: z.read(n) for n in z.namelist()}
    n = 0
    for name in [x for x in items if x.startswith('ppt/slides/slide') and x.endswith('.xml')]:
        root = etree.fromstring(items[name])
        for c in root.iter('{%s}srgbClr' % A):
            v = (c.get('val') or '').upper()
            if v in m:
                c.set('val', m[v])
                n += 1
        items[name] = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
    with zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as zo:
        for k, b in items.items():
            zo.writestr(k, b)
    return n


if __name__ == '__main__':
    src, pal = sys.argv[1], sys.argv[2]
    dst = sys.argv[3] if len(sys.argv) > 3 else src.replace('.pptx', '-%s.pptx' % pal)
    print('换色 %d 处 → %s' % (recolor(src, pal, dst), os.path.basename(dst)))
