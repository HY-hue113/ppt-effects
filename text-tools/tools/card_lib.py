# -*- coding: utf-8 -*-
"""文本底卡处理区 · 卡片库（v1）

用法：
    from card_lib import card, chip, tab, badge, grid_cards
    card(dst, w, h, style="plain")      # 通用底卡
    chip(dst, w, h)                      # 胶囊条（单行文字用）
    tab(dst, w, h)                       # 顶部带缺口的小标签
    badge(dst, w, h)                     # 序号方块

style 取值：
    plain   白底 + 细边（最稳）
    tint    浅蓝渐变 + 细边
    bar     白底 + 左侧金色竖条
    line    仅细边、无填充（叠在图上用）
"""
import os
from PIL import Image, ImageDraw, ImageFilter, ImageChops

SS = 2
GOLD = (0xC3, 0xAE, 0x53)
EDGE = (0xDC, 0xE8, 0xF6)
INK = (0x17, 0x33, 0x5F)


def _vgrad(w, h, c1, c2):
    im = Image.new("RGB", (1, h))
    for y in range(h):
        t = y / max(1, h - 1)
        im.putpixel((0, y), tuple(int(c1[k] + (c2[k] - c1[k]) * t) for k in range(3)))
    return im.resize((w, h), Image.BILINEAR)


def _canvas(w, h):
    return Image.new("RGBA", (w * SS, h * SS), (0, 0, 0, 0))


def card(dst, w, h, style="plain", r=12, shadow=True):
    W, H = w * SS, h * SS
    m = Image.new("L", (W, H), 0)
    ImageDraw.Draw(m).rounded_rectangle([0, 0, W - 1, H - 1], radius=r * SS, fill=255)
    if style == "tint":
        fill = _vgrad(W, H, (0xFA, 0xFC, 0xFF), (0xE8, 0xF1, 0xFB)).convert("RGBA")
    elif style == "line":
        fill = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    else:
        fill = Image.new("RGBA", (W, H), (0xFF, 0xFF, 0xFF, 245))
    base = _canvas(w, h)
    base.paste(fill, (0, 0), m)

    if shadow and style != "line":
        sh = Image.new("RGBA", (W, H), (0, 0, 0, 0))
        ImageDraw.Draw(sh).rounded_rectangle([3 * SS, 5 * SS, W - 3 * SS, H - 3 * SS], radius=r * SS,
                                             fill=(0x1B, 0x3A, 0x6B, 30))
        sh = sh.filter(ImageFilter.GaussianBlur(5 * SS))
        base.alpha_composite(sh)
        base.paste(fill, (0, 0), m)

    d = ImageDraw.Draw(base)
    d.rounded_rectangle([0, 0, W - 1, H - 1], radius=r * SS, outline=EDGE + (255,), width=2 * SS)
    if style == "corner":
        # 折角卡：右上角切一块并描金
        cut = int(min(w, h) * 0.22)
        d = ImageDraw.Draw(base)
        d.polygon([(W - cut * SS, 0), (W - 1, 0), (W - 1, cut * SS)], fill=(0xFA, 0xFC, 0xFF, 255))
        d.line([(W - cut * SS, 0), (W - 1, cut * SS)], fill=GOLD + (255,), width=2 * SS)
    if style == "duo":
        # 双层线卡：外框 + 内缩 10px 细线
        d = ImageDraw.Draw(base)
        d.rounded_rectangle([10 * SS, 10 * SS, W - 10 * SS, H - 10 * SS],
                            radius=max(2, r - 4) * SS, outline=(0xE3, 0xD3, 0xA0, 230), width=1 * SS)
    if style == "topbar":
        # 顶条卡：顶部 6px 金条
        d = ImageDraw.Draw(base)
        d.rectangle([0, 0, W, 6 * SS], fill=GOLD + (255,))
    if style == "bar":
        d.rounded_rectangle([SS, 8 * SS, 7 * SS, H - 8 * SS], radius=3 * SS, fill=GOLD + (255,))
    base.save(dst)
    return dst


def chip(dst, w, h):
    """胶囊条：两端全圆，浅蓝底 + 左侧圆点"""
    W, H = w * SS, h * SS
    m = Image.new("L", (W, H), 0)
    ImageDraw.Draw(m).rounded_rectangle([0, 0, W - 1, H - 1], radius=H // 2, fill=255)
    base = _canvas(w, h)
    base.paste(_vgrad(W, H, (0xF7, 0xFB, 0xFF), (0xE4, 0xEF, 0xFA)).convert("RGBA"), (0, 0), m)
    d = ImageDraw.Draw(base)
    d.rounded_rectangle([0, 0, W - 1, H - 1], radius=H // 2, outline=EDGE + (255,), width=2 * SS)
    cy, rr = H // 2, int(4.5 * SS)
    d.ellipse([10 * SS, cy - rr, 10 * SS + 2 * rr, cy + rr], fill=GOLD + (255,))
    base.save(dst)
    return dst


def tab(dst, w, h):
    """小标题标签：顶部深蓝，底部圆角，像书签"""
    W, H = w * SS, h * SS
    m = Image.new("L", (W, H), 0)
    ImageDraw.Draw(m).rounded_rectangle([0, 0, W - 1, H - 1], radius=8 * SS, fill=255)
    base = _canvas(w, h)
    base.paste(_vgrad(W, H, (0x24, 0x4F, 0x8A), (0x3E, 0x7C, 0xC8)).convert("RGBA"), (0, 0), m)
    d = ImageDraw.Draw(base)
    d.rounded_rectangle([0, 0, W - 1, H - 1], radius=8 * SS, outline=GOLD + (255,), width=2 * SS)
    d.rectangle([0, H - 5 * SS, W, H], fill=GOLD + (255,))
    base.save(dst)
    return dst


def badge(dst, size, text=None):
    """序号方块（圆角，深蓝底金边）"""
    W = H = size * SS
    m = Image.new("L", (W, H), 0)
    ImageDraw.Draw(m).rounded_rectangle([0, 0, W - 1, H - 1], radius=int(size * SS * 0.22), fill=255)
    base = _canvas(size, size)
    base.paste(_vgrad(W, H, (0x1E, 0x44, 0x7A), (0x3E, 0x7C, 0xC8)).convert("RGBA"), (0, 0), m)
    d = ImageDraw.Draw(base)
    d.rounded_rectangle([0, 0, W - 1, H - 1], radius=int(size * SS * 0.22), outline=GOLD + (255,), width=2 * SS)
    base.save(dst)
    return dst
