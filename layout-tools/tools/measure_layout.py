# -*- coding: utf-8 -*-
"""按原图精确测量 p8 / p11 的色块与圆形位置，再重画"""
import os
import sys

import numpy as np
from PIL import Image

OUT = r'<USER_HOME>\Documents\<DSH_HOME>\default-workspace\out\layouts'
SLIDE_W, SLIDE_H = 1280, 720


def comps(mask):
    h, w = mask.shape
    seen = np.zeros_like(mask, dtype=bool)
    res = []
    for y in range(h):
        for x in range(w):
            if not mask[y, x] or seen[y, x]:
                continue
            stack = [(y, x)]
            seen[y, x] = True
            ys, xs = [], []
            while stack:
                cy, cx = stack.pop()
                ys.append(cy)
                xs.append(cx)
                for dy, dx in ((1, 0), (-1, 0), (0, 1), (0, -1)):
                    ny, nx = cy + dy, cx + dx
                    if 0 <= ny < h and 0 <= nx < w and mask[ny, nx] and not seen[ny, nx]:
                        seen[ny, nx] = True
                        stack.append((ny, nx))
            res.append((min(ys), min(xs), max(ys), max(xs), len(ys)))
    return res


def measure(png, bands):
    im = Image.open(png).convert('L')
    W, H = im.size
    scale = 260 / max(W, H)
    small = im.resize((max(1, int(W * scale)), max(1, int(H * scale))))
    a = np.asarray(small).astype(int)
    out = []
    for lo, hi, label in bands:
        mask = (a >= lo) & (a <= hi)
        for (y0, x0, y1, x1, area) in sorted(comps(mask), key=lambda c: -c[4])[:8]:
            w, hh = x1 - x0 + 1, y1 - y0 + 1
            if area < 40 or w < 5 or hh < 5:
                continue
            fr = area / (w * hh)
            out.append(dict(label=label,
                            x=round(x0 / small.size[0] * SLIDE_W),
                            y=round(y0 / small.size[1] * SLIDE_H),
                            w=round(w / small.size[0] * SLIDE_W),
                            h=round(hh / small.size[1] * SLIDE_H),
                            round=0.62 <= fr <= 0.86))
    return out


if __name__ == '__main__':
    for png in sys.argv[1:]:
        print('===', os.path.basename(png))
        for b in measure(png, [(235, 250, 'F2'), (160, 180, 'A6'), (205, 225, 'D9')]):
            print('  %-3s x=%-4d y=%-4d w=%-4d h=%-4d %s' %
                  (b['label'], b['x'], b['y'], b['w'], b['h'], '圆形' if b['round'] else ''))
