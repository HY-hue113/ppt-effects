# -*- coding: utf-8 -*-
"""按配色重生成偏学术类参考件（换色不改结构）

用法：
    python recolor.py 学术蓝 [输出.pptx]
    python recolor.py 红
    不给输出名时，默认写到 reference/学术导航0-纵横导航栏-偏学术-<配色>.pptx
"""
import os
import re
import shutil
import sys
import zipfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from palette import get, tint, shade

DEFAULT_SRC = r'<LIB_ROOT>\effects\academic\学术导航0\reference\学术导航0-纵横导航栏-偏学术.pptx'
MOD_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
REF_DIR = r'<LIB_ROOT>\effects\academic\学术导航0\reference'


def mapping(palette):
    p = get(palette)
    primary, dark = p["primary"], p["dark"]
    return {
        "17335F": dark,
        "0D3063": dark,
        "2F76B8": primary,
        "3E7CC8": tint(primary, 0.35),
        "2A5E9E": shade(primary, 0.12),
        "1E447A": dark,
        "1F4E79": primary,
        "EAF1FA": tint(primary, 0.90),
        "DCE8F6": tint(primary, 0.74),
        "C3AE53": p["accent"],
        "D8DFC9": tint(primary, 0.72),
    }


def recolor(src, dst, palette):
    m = mapping(palette)
    shutil.copyfile(src, dst)
    z = zipfile.ZipFile(dst)
    items = {n: z.read(n) for n in z.namelist()}
    z.close()
    changed = 0
    for n in list(items):
        if not re.match(r'ppt/(slides/slide\d+\.xml|slideMasters/.+|slideLayouts/.+|theme/.+)', n):
            continue
        s = items[n].decode('utf-8')
        o = s
        for old, new in m.items():
            s = re.sub(r'(srgbClr val=")' + old + r'(")', r'\g<1>' + new + r'\g<2>', s, flags=re.I)
        if s != o:
            items[n] = s.encode('utf-8')
            changed += 1
    with zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as zo:
        for k, b in items.items():
            zo.writestr(k, b)
    return changed, m


if __name__ == '__main__':
    palette = sys.argv[1] if len(sys.argv) > 1 else "学术蓝"
    p = get(palette)
    out = sys.argv[2] if len(sys.argv) > 2 else os.path.join(
        REF_DIR, "学术导航0-纵横导航栏-偏学术-%s.pptx" % p["name"])
    n, m = recolor(DEFAULT_SRC, out, palette)
    print("配色 %s | 主 %s | 替换 %d 个文件 | %s" % (p["name"], p["primary"], n, out))
