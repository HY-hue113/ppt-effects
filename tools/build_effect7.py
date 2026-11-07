#!/usr/bin/env python3
"""效果7 · 环形开场：只替换背景素材，不碰任何形状与动画。

用法:
    python build_effect7.py --in reference/效果7.pptx --image 新图.jpg --out 输出.pptx
"""
import argparse
import io
import os
import zipfile

from PIL import Image

MEDIA = 'ppt/media/image1.jpeg'


def cover(im, tw, th):
    sw, sh = im.size
    s = max(tw / sw, th / sh)
    nw, nh = int(sw * s + .5), int(sh * s + .5)
    im = im.resize((nw, nh), Image.LANCZOS)
    l, t = (nw - tw) // 2, (nh - th) // 2
    return im.crop((l, t, l + tw, t + th))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--in', dest='src', required=True)
    ap.add_argument('--image', required=True)
    ap.add_argument('--out', dest='dst', required=True)
    args = ap.parse_args()

    with zipfile.ZipFile(args.src) as z:
        items = {n: z.read(n) for n in z.namelist()}

    if MEDIA not in items:
        raise SystemExit('参考件里没有 %s' % MEDIA)

    tgt = Image.open(io.BytesIO(items[MEDIA]))
    buf = io.BytesIO()
    cover(Image.open(args.image).convert('RGB'), *tgt.size).save(buf, 'JPEG', quality=93)
    items[MEDIA] = buf.getvalue()

    with zipfile.ZipFile(args.dst, 'w', zipfile.ZIP_DEFLATED) as z:
        for n, d in items.items():
            z.writestr(n, d)
    print('built', args.dst, os.path.getsize(args.dst) // 1024, 'KB | image', tgt.size)


if __name__ == '__main__':
    main()
