# -*- coding: utf-8 -*-
"""选色工具：四套学术配色 + 通用辅色

用法：
    from palette import get, rgb, shade, tint, apply_theme

    p = get("academic-blue")        # 返回配色 dict（hex 字符串）
    p["primary"]                    # "183A6A"
    rgb("183A6A")                   # (24, 58, 106)
    tint("183A6A", 0.2)             # 变浅 20% → "3C5A87"
    shade("183A6A", 0.2)            # 变深 20% → "0F2547"
    css(p)                          # {"primary": "#183A6A", ...}
"""
import json
import os

_HERE = os.path.dirname(os.path.abspath(__file__))
_DATA = json.load(open(os.path.join(_HERE, '..', 'palettes.json'), encoding='utf-8'))

PALETTES = _DATA['palettes']
COMMON = _DATA['common']
ALIASES = {
    "橄榄黄绿": "academic-olive", "青蓝湖色": "academic-lake",
    "灰紫丁香": "academic-lilac", "暖橙金棕": "academic-amber",
    "砖红暖棕": "academic-brick", "森系墨绿": "academic-forest",
    "深海蓝系": "academic-ocean",
    "学术蓝": "academic-blue", "学术红": "academic-red",
    "学术绿": "academic-green", "学术紫": "academic-purple",
    "蓝": "academic-blue", "红": "academic-red",
    "绿": "academic-green", "紫": "academic-purple",
    "blue": "academic-blue", "red": "academic-red",
    "green": "academic-green", "purple": "academic-purple",
}


def list_palettes():
    return [(k, v['name'], v['note']) for k, v in PALETTES.items()]


def resolve(name):
    if name in PALETTES:
        return name
    key = ALIASES.get(str(name).strip())
    if key:
        return key
    raise KeyError("未知配色：%s（可选：%s）" % (name, "、".join(PALETTES)))


def get(name):
    """返回该配色的完整 dict（hex，不带 #）"""
    key = resolve(name)
    p = dict(PALETTES[key])
    p.update({k: v for k, v in COMMON.items() if k not in p})
    p['key'] = key
    return p


def rgb(h):
    h = h.lstrip('#')
    return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))


def _to_hex(t):
    return "".join("%02X" % max(0, min(255, int(round(v)))) for v in t)


def tint(h, ratio=0.2, base=255):
    """向白色混合（变浅）"""
    r, g, b = rgb(h)
    return _to_hex((r + (base - r) * ratio, g + (base - g) * ratio, b + (base - b) * ratio))


def shade(h, ratio=0.2, base=0):
    """向黑色混合（变深）"""
    r, g, b = rgb(h)
    return _to_hex((r + (base - r) * ratio, g + (base - g) * ratio, b + (base - b) * ratio))


def css(p):
    return {k: "#" + v for k, v in p.items() if isinstance(v, str) and len(v) == 6
            and all(c in "0123456789ABCDEFabcdef" for c in v)}


def apply_theme(palette="academic-blue"):
    """给出常用语义色，供生成器直接取用"""
    p = get(palette)
    return {
        "primary": p["primary"],
        "secondary": p["secondary"],
        "light": p["light"],
        "dark": p["dark"],
        "accent": p["accent"],
        "card_bg": "FFFFFF",
        "card_edge": p["light"],
        "title": p["primary"],
        "body": p["ink"],
        "muted": "8A8578",
        "nav_bar_dark": p["dark"],
        "nav_bar_light": p["secondary"],
        "chip_bg": p["light"],
        "mask_light": "FFFFFF",
        "mask_dark": "12212E",
    }


if __name__ == '__main__':
    for k, n, note in list_palettes():
        p = get(k)
        print("%-14s %s | 主 %s 辅 %s 浅 %s 金 %s | %s" %
              (k, n, p['primary'], p['secondary'], p['light'], p['accent'], note))
