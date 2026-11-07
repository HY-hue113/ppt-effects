import os, math
from PIL import Image

BASE = r'<USER_HOME>\Documents\<DSH_HOME>\default-workspace\projects'
PROJ = os.path.join(BASE, '效果10_ppt169_20261106')
SVG = os.path.join(PROJ, 'svg_output')
IMG = os.path.join(PROJ, 'images')
os.makedirs(SVG, exist_ok=True)
os.makedirs(IMG, exist_ok=True)

SRC = r'<DSH_ATTACHMENTS>\v1\objects\f2\f2e3a3164d01a19c95802efae711da0762ffa4302229fda31249833dcf9482ce'
im = Image.open(SRC).convert('RGBA')
im = im.crop((0, 0, im.width, int(im.height*0.74)))
# 白底转透明：alpha 由亮度反推，颜色压成深墨
px = im.load()
w, h = im.size
out = Image.new('RGBA', (w, h), (0, 0, 0, 0))
op = out.load()
for y in range(h):
    for x in range(w):
        r, g, b, _ = px[x, y]
        lum = int(0.299 * r + 0.587 * g + 0.114 * b)
        a = max(0, min(255, int((235 - lum) * 3.2)))
        if a > 6:
            op[x, y] = (236, 206, 150, min(255, int(a*0.85)))
out = out.resize((int(w * 0.62), int(h * 0.62)), Image.LANCZOS)
out.save(os.path.join(IMG, 'tower.png'))
print('tower', out.size)

W, H = 1280, 720
RED = '#2B0908'
DEEP = '#120303'
GOLD = '#E8C98A'


def svg(body):
    return ('<svg xmlns="http://www.w3.org/2000/svg" width="1280" height="720" viewBox="0 0 1280 720" lang="en-US" '
            'font-family="Arial">\n' + body + '\n</svg>')


# ---------- P1：黑底，散字 ----------
letters = list('PRESENTATION')
b = [f'  <rect x="0" y="0" width="{W}" height="{H}" fill="{DEEP}"/>']
n = len(letters)
x0, span = 90, 1100
for i, ch in enumerate(letters):
    x = x0 + i * (span / (n - 1))
    b.append(f'  <text x="{x:.0f}" y="380" font-family="Times New Roman" font-size="132" fill="#FFFFFF" '
             f'text-anchor="middle">{ch}</text>')
b += [f'  <circle cx="1130" cy="150" r="7" fill="{GOLD}"/>',
      f'  <circle cx="1180" cy="120" r="5" fill="{GOLD}" opacity="0.9"/>']
open(os.path.join(SVG, 'slide_01.svg'), 'w', encoding='utf-8').write(svg('\n'.join(b)))

# ---------- P2：红底 + 塔 + 大字 ----------
b = [f'  <rect x="0" y="0" width="{W}" height="{H}" fill="{RED}"/>',
     f'  <circle cx="640" cy="400" r="420" fill="#4A0E0C" opacity="0.55"/>',
     '  <image href="../images/tower.png" x="470" y="120" width="340" height="560" preserveAspectRatio="xMidYMid meet"/>',
     '  <polygon points="880,720 1160,120 1280,120 1280,720" fill="#F3D9A4" opacity="0.10"/>',
     f'  <text x="640" y="150" font-family="Times New Roman" font-size="104" fill="#FFFFFF" '
     f'text-anchor="middle" letter-spacing="18">PRESENTATION</text>',
     f'  <circle cx="1120" cy="640" r="9" fill="{GOLD}"/>']
open(os.path.join(SVG, 'slide_02.svg'), 'w', encoding='utf-8').write(svg('\n'.join(b)))

# ---------- P3：塔 + 三条标签 ----------
b = [f'  <rect x="0" y="0" width="{W}" height="{H}" fill="{RED}"/>',
     f'  <circle cx="420" cy="380" r="330" fill="#4A0E0C" opacity="0.5"/>',
     '  <image href="../images/tower.png" x="270" y="150" width="300" height="500" preserveAspectRatio="xMidYMid meet"/>',
     '  <polygon points="0,720 220,120 380,120 300,720" fill="#F3D9A4" opacity="0.08"/>',
     f'  <text x="80" y="110" font-family="Times New Roman" font-size="68" fill="#FFFFFF" letter-spacing="8">PRESENTATION</text>']
labels = [('Supported by wooden frames', 300), ('Curve like spreading bird wings', 380), ('Glen like cover the sloping roof', 460)]
for txt, y in labels:
    b += [f'  <rect x="640" y="{y-40}" width="520" height="58" rx="29" fill="#8E1B18"/>',
          f'  <text x="672" y="{y-2}" font-family="Arial" font-size="22" fill="#FFFFFF">{txt}</text>']
open(os.path.join(SVG, 'slide_03.svg'), 'w', encoding='utf-8').write(svg('\n'.join(b)))

import json
json.dump({"version": 1, "slides": {
    "slide_02": {"transition": {"effect": "morph", "effect_options": {"morph_by": "object"}, "duration": 1.8}},
    "slide_03": {"transition": {"effect": "morph", "effect_options": {"morph_by": "object"}, "duration": 1.8}},
}}, open(os.path.join(PROJ, 'animations.json'), 'w', encoding='utf-8'), ensure_ascii=False, indent=2)
print('built', PROJ)
