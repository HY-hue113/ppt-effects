# -*- coding: utf-8 -*-
"""图片处理区 · 形状遮罩库（v2，10 种形状）

库函数（无副作用，可在任意效果里调用）：
    mask(w, h, kind)                     -> 灰度遮罩图
    shaped_image(src, dst, w, h, kind)   -> 形状遮罩的透明 PNG
    outline_svg(x, y, w, h, kind, ...)   -> 与遮罩同源的 SVG 描边

kind 取值：
    cut 切角 / circle 圆 / hex 六边 / arch 拱形 / diag 斜切
    rhombus 菱形 / capsule 胶囊 / wave 波浪边 / ticket 票券齿孔 / broken 破格矩形

参数（与 README 一致）：
    切角 18%　六边 22%　拱形 45%　斜切 20%　
    菱形取短边 50%　胶囊半径=短边/2　波浪振幅 6% 波长 25%　
    票券齿半径 8%（左右各一个）　破格缺口 = 短边 30%（右下角）
"""
import math
from PIL import Image, ImageDraw


def mask(w, h, kind):
    m = Image.new("L", (w, h), 0)
    d = ImageDraw.Draw(m)

    if kind == "circle":
        d.ellipse([0, 0, w - 1, h - 1], fill=255)

    elif kind == "capsule":
        d.rounded_rectangle([0, 0, w - 1, h - 1], radius=min(w, h) // 2, fill=255)

    elif kind == "cut":
        c = int(min(w, h) * 0.18)
        d.polygon([(c, 0), (w - 1, 0), (w - 1, h - 1), (0, h - 1), (0, c)], fill=255)

    elif kind == "hex":
        c = int(min(w, h) * 0.22)
        d.polygon([(c, 0), (w - 1 - c, 0), (w - 1, h // 2),
                   (w - 1 - c, h - 1), (c, h - 1), (0, h // 2)], fill=255)

    elif kind == "arch":
        r = int(min(w, h) * 0.45)
        d.rounded_rectangle([0, 0, w - 1, h - 1], radius=r, fill=255)
        d.rectangle([0, r, w - 1, h - 1], fill=255)

    elif kind == "diag":
        k = int(w * 0.2)
        d.polygon([(k, 0), (w - 1, 0), (w - 1 - k, h - 1), (0, h - 1)], fill=255)

    elif kind == "rhombus":
        d.polygon([(w // 2, 0), (w - 1, h // 2), (w // 2, h - 1), (0, h // 2)], fill=255)

    elif kind == "wave":
        amp = max(4, int(min(w, h) * 0.06))
        wave_len = max(18, int(w * 0.25))
        pts = [(0, 0), (w - 1, 0)]
        x = w - 1
        while x >= 0:
            y = h - 1 - int(amp * math.sin(2 * math.pi * (w - 1 - x) / wave_len))
            pts.append((x, y))
            x -= 2
        d.polygon(pts, fill=255)

    elif kind == "ticket":
        r = max(6, int(min(w, h) * 0.08))
        d.rectangle([0, 0, w - 1, h - 1], fill=255)
        d.ellipse([-r, h // 2 - r, r, h // 2 + r], fill=0)
        d.ellipse([w - 1 - r, h // 2 - r, w - 1 + r, h // 2 + r], fill=0)

    elif kind == "broken":
        # 右下角缺一块，形成"破格"
        c = int(min(w, h) * 0.30)
        d.polygon([(0, 0), (w - 1, 0), (w - 1, h - 1 - c), (w - 1 - c, h - 1 - c),
                   (w - 1 - c, h - 1), (0, h - 1)], fill=255)

    else:
        raise ValueError("未知形状: %s" % kind)

    return m


def shaped_image(src, dst, w, h, kind="hex"):
    im = Image.open(src).convert("RGB")
    s = max(w / im.width, h / im.height)
    im = im.resize((int(im.width * s + 1), int(im.height * s + 1)), Image.LANCZOS)
    l = (im.width - w) // 2
    t = (im.height - h) // 2
    im = im.crop((l, t, l + w, t + h))
    out = Image.new("RGBA", (w, h), (0, 0, 0, 0))
    out.paste(im, (0, 0), mask(w, h, kind))
    out.save(dst)
    return dst


def _poly_pts(w, h, kind):
    if kind == "cut":
        c = int(min(w, h) * 0.18)
        return [(c, 0), (w, 0), (w, h), (0, h), (0, c)]
    if kind == "hex":
        c = int(min(w, h) * 0.22)
        return [(c, 0), (w - c, 0), (w, h / 2), (w - c, h), (c, h), (0, h / 2)]
    if kind == "diag":
        k = int(w * 0.2)
        return [(k, 0), (w, 0), (w - k, h), (0, h)]
    if kind == "rhombus":
        return [(w / 2, 0), (w, h / 2), (w / 2, h), (0, h / 2)]
    if kind == "broken":
        c = int(min(w, h) * 0.30)
        return [(0, 0), (w, 0), (w, h - c), (w - c, h - c), (w - c, h), (0, h)]
    return None


def outline_svg(x, y, w, h, kind, color="#C3AE53", sw=3, bg="#F6F4EE"):
    pts = _poly_pts(w, h, kind)
    common = 'fill="none" stroke="%s" stroke-width="%g"' % (color, sw)
    if pts:
        s = " ".join("%g,%g" % (x + px, y + py) for px, py in pts)
        return '<polygon points="%s" %s/>' % (s, common)
    if kind == "circle":
        return '<circle cx="%g" cy="%g" r="%g" %s/>' % (x + w / 2, y + h / 2, min(w, h) / 2, common)
    if kind == "capsule":
        return '<rect x="%g" y="%g" width="%g" height="%g" rx="%g" %s/>' % (
            x, y, w, h, min(w, h) / 2, common)
    if kind == "arch":
        r = min(w, h) * 0.45
        return ('<path d="M %g %g L %g %g A %g %g 0 0 1 %g %g L %g %g Z" %s/>' % (
            x, y + h, x, y + r, r, r, x + w, y + r, x + w, y + h, common))
    if kind == "ticket":
        r = min(w, h) * 0.08
        return ('<rect x="%g" y="%g" width="%g" height="%g" %s/>'
                '<circle cx="%g" cy="%g" r="%g" fill="%s" stroke="%s" stroke-width="%g"/>'
                '<circle cx="%g" cy="%g" r="%g" fill="%s" stroke="%s" stroke-width="%g"/>' % (
                    x, y, w, h, common,
                    x, y + h / 2, r, bg, color, sw,
                    x + w, y + h / 2, r, bg, color, sw))
    if kind == "wave":
        amp = max(4, int(min(w, h) * 0.06))
        wl = max(18, int(w * 0.25))
        d = ["M %g %g" % (x, y)]
        xx = 0
        while xx <= w:
            yy = h - amp * math.sin(2 * math.pi * xx / wl)
            d.append("L %g %g" % (x + xx, y + yy))
            xx += w / 16.0
        d.append("L %g %g Z" % (x + w, y))
        return '<path d="%s" %s/>' % (" ".join(d), common)
    return ""
