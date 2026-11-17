# -*- coding: utf-8 -*-
"""去除版式库水印（两种形态）

1. 独立水印小图（141×42）→ 直接删对象（之前已处理，这里兜底）
2. 烘焙在大图里的水印 → 在图片右下角区域做"亮像素回填"：把比背景更亮的像素
   替换成该区域的中位色，视觉上去掉水印文字
"""
import io
import os
import zipfile

import numpy as np
from PIL import Image
from lxml import etree

LIB = r'<LIB_ROOT>\layout-tools'
SRC = r'<USER_HOME>\Desktop\ppt版式库(01-06).pptx'
OUT = os.path.join(LIB, 'reference', '版式库-无水印.pptx')

P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
EMU = 9525


def clean_watermark(data):
    """在大图右下角去掉白色水印文字"""
    try:
        im = Image.open(io.BytesIO(data))
    except Exception:
        return data, False
    fmt = im.format or 'PNG'
    if im.mode not in ('RGB', 'RGBA'):
        im = im.convert('RGB')
    a = np.asarray(im).astype(np.uint8).copy()
    h, w = a.shape[:2]
    y0, y1 = int(h * 0.84), h
    x0, x1 = int(w * 0.74), w
    region = a[y0:y1, x0:x1]
    rgb = region[..., :3]
    gray = rgb.mean(axis=2)
    bg = float(np.median(gray))
    mask = gray > bg + 16          # 比背景更亮的像素 = 水印文字
    if mask.sum() == 0:
        return data, False
    # 用周围中位色回填
    keep = rgb[~mask].reshape(-1, 3)
    fill = np.median(keep, axis=0) if len(keep) else np.array([bg, bg, bg], dtype=float)
    rgb[mask] = fill.astype(np.uint8)
    region[..., :3] = rgb
    a[y0:y1, x0:x1] = region
    out = Image.fromarray(a)
    buf = io.BytesIO()
    out.save(buf, format='PNG' if fmt == 'PNG' else fmt)
    return buf.getvalue(), True


def main():
    with zipfile.ZipFile(SRC) as z:
        items = {n: z.read(n) for n in z.namelist()}

    # 1. 独立水印对象
    pres = etree.fromstring(items['ppt/presentation.xml'])
    rels = etree.fromstring(items['ppt/_rels/presentation.xml.rels'])
    m = {r.get('Id'): r.get('Target') for r in rels}
    parts = ['ppt/' + m[s.get('{%s}id' % R)].replace('../', '')
             for s in pres.find('{%s}sldIdLst' % P)]
    n_obj = 0
    for part in parts:
        root = etree.fromstring(items[part])
        for pic in list(root.iter('{%s}pic' % P)):
            xf = pic.find('.//{%s}xfrm' % A)
            ext = xf.find('{%s}ext' % A) if xf is not None else None
            if ext is None:
                continue
            w = int(ext.get('cx')) // EMU
            h = int(ext.get('cy')) // EMU
            if 120 <= w <= 170 and 30 <= h <= 55:
                pic.getparent().remove(pic)
                n_obj += 1
        items[part] = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)

    # 2. 图片内烘焙水印
    n_img = 0
    for name in list(items):
        if not name.startswith('ppt/media/'):
            continue
        ext = os.path.splitext(name)[1].lower()
        if ext not in ('.png', '.jpg', '.jpeg'):
            continue
        data, changed = clean_watermark(items[name])
        if changed:
            items[name] = data
            n_img += 1

    with zipfile.ZipFile(OUT, 'w', zipfile.ZIP_DEFLATED) as zo:
        for k, b in items.items():
            zo.writestr(k, b)
    print('删除独立水印对象 %d 个，处理图片水印 %d 张 → %s' % (n_obj, n_img, OUT))


if __name__ == '__main__':
    main()
