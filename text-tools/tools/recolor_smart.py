# -*- coding: utf-8 -*-
"""按色相智能换色：把红系/蓝系的所有深浅（含浅粉、浅蓝）映射到目标配色

原理：读每个 srgbClr → 转 HSL →
    若色相落在红系(340–20°) 或蓝系(200–250°) 且有一定饱和度 → 用目标主色的色相替换，
    保留原来的明度/饱和度关系（浅色仍浅、深色仍深）。
非主色系（黑白灰、金属色、青色霓虹）不动。

用法：python recolor_smart.py <pptx> <配色名> [输出]
"""
import colorsys
import os
import sys
import zipfile

sys.path.insert(0, r'<LIB_ROOT>\palette-tools\tools')
from palette import get  # noqa: E402

from lxml import etree  # noqa: E402

P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'


def hex2rgb(h):
    return tuple(int(h[i:i + 2], 16) / 255 for i in (0, 2, 4))


def rgb2hex(r, g, b):
    return '%02X%02X%02X' % (int(round(r * 255)), int(round(g * 255)), int(round(b * 255)))


def in_red(h):
    return h >= 0.88 or h <= 0.10


def in_blue(h):
    return 0.55 <= h <= 0.70


def recolor(src, palette, dst):
    p = get(palette)
    pr, pg, pb = hex2rgb(p['primary'])
    base_h, base_l, base_s = colorsys.rgb_to_hls(pr, pg, pb)

    with zipfile.ZipFile(src) as z:
        items = {n: z.read(n) for n in z.namelist()}
    n = 0
    for name in [x for x in items if x.endswith('.xml') and (x.startswith('ppt/slides/slide') or x.startswith('ppt/slideLayouts/') or x.startswith('ppt/slideMasters/'))]:
        root = etree.fromstring(items[name])
        for c in root.iter('{%s}srgbClr' % A):
            v = (c.get('val') or '').upper()
            if len(v) != 6:
                continue
            r, g, b = hex2rgb(v)
            h, l, s = colorsys.rgb_to_hls(r, g, b)
            if s < 0.04:          # 仅纯灰阶不动
                continue
            if not (in_red(h) or in_blue(h)):
                continue
            # 保留明度关系：接近白色的浅色 → 目标浅色；深色 → 目标深色
            new_l = base_l if l < 0.5 else min(0.97, base_l + (l - 0.5) * 1.4)
            new_s = min(1.0, max(base_s * 0.6, s * base_s / max(base_s, 0.01)))
            nr, ng, nb = colorsys.hls_to_rgb(base_h, new_l, new_s)
            c.set('val', rgb2hex(nr, ng, nb))
            n += 1
        items[name] = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
    with zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as zo:
        for k, b in items.items():
            zo.writestr(k, b)
    return n


if __name__ == '__main__':
    src, pal = sys.argv[1], sys.argv[2]
    dst = sys.argv[3] if len(sys.argv) > 3 else src.replace('.pptx', '-%s.pptx' % pal)
    print('智能换色 %d 处 → %s' % (recolor(src, pal, dst), os.path.basename(dst)))
