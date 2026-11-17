# -*- coding: utf-8 -*-
"""效果17 偏学术风版 · 参数化生成器

用法：
  python build_academic.py --chapters "标题1;标题2;标题3;标题4;标题5" \
                           --subs "标题1;标题2;标题3;标题4" \
                           [--title 大标题] [--out 输出.pptx]

规则：
  - 左栏项数 = 章节数（自上而下），顶栏项数 = 每章子页数（自左向右）
  - 页数 = 章节数 × 子页数
  - 槽位高度/间隙按项数自动计算（项多则压缩，项少则放松）
"""
import argparse, io, os, sys, subprocess
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
from PIL import Image, ImageDraw, ImageFilter

SK = r"<USER_HOME>\<DSH_DIR>\skills\ppt-master"
PY = sys.executable
W, H, SS = 1280, 720, 2
TOP_H = 82

ap = argparse.ArgumentParser()
ap.add_argument("--chapters", required=True, help="左栏章节名，分号分隔")
ap.add_argument("--subs", required=True, help="顶栏子页名，分号分隔")
ap.add_argument("--title", default="大标题")
ap.add_argument("--out", default=os.path.join(os.path.expanduser("~"), "Desktop", "效果17-偏学术风版.pptx"))
ap.add_argument("--work", default=os.path.join(os.path.dirname(os.path.abspath(__file__)), "_build"))
a = ap.parse_args()

CH = [x.strip() for x in a.chapters.split(";") if x.strip()]
_groups = [g.strip() for g in a.subs.split("|") if g.strip()]
SUB_GROUPS = [[x.strip() for x in g.split(";") if x.strip()] for g in _groups]
SUB = SUB_GROUPS[0]          # 顶栏宽度按最长一组计算
NC = len(CH)
NS = max(len(g) for g in SUB_GROUPS) if SUB_GROUPS else 0
WS = a.work
SVG = os.path.join(WS, "svg_output"); IMG = os.path.join(WS, "images")
for d in (SVG, IMG):
    os.makedirs(d, exist_ok=True)
    for f in os.listdir(d):
        os.remove(os.path.join(d, f))

# ---------- 顶栏几何：右对齐、等距、激活项拉长 ----------
TAB_ZONE_X0, TAB_ZONE_X1 = 560, W - 32      # 可用区（右边界贴画布右）
TAB_H, TAB_Y = 46, 18
TAB_GAP = 12 if NS <= 5 else 10
ACTIVE_SCALE = 1.45                          # 激活标签相对普通标签的宽度倍数
_zone = TAB_ZONE_X1 - TAB_ZONE_X0
_base_w = int((_zone - (NS - 1) * TAB_GAP) / (NS - 1 + ACTIVE_SCALE))
BASE_TAB_W = max(96, min(150, _base_w))

# 右对齐排布：从右往左累加
tab_w_list = [BASE_TAB_W] * NS
tab_x = [0] * NS


def _layout_tabs(active):
    w_active = int(BASE_TAB_W * ACTIVE_SCALE)
    widths = [w_active if i == active else BASE_TAB_W for i in range(NS)]
    total = sum(widths) + (NS - 1) * TAB_GAP
    x = TAB_ZONE_X1 - total
    xs = []
    for i in range(NS):
        xs.append(int(x))
        x += widths[i] + TAB_GAP
    return widths, xs


# ---------- 左栏几何：等间隙、整体居中、激活拉长 ----------
RAIL_X, RAIL_W = 32, 104
RAIL_Y = TOP_H + 12
RAIL_TOTAL = H - 12 - RAIL_Y
ACTIVE_H = 140
avail = RAIL_TOTAL - ACTIVE_H

if NC == 1:
    NORMAL_H, GAP = min(120, avail), 0
else:
    # 先给普通项一个合理高度，再用「等间隙」把剩余空间均分
    NORMAL_H = int(max(56, min(96, avail * 0.60 / (NC - 1))))
    GAP = int((avail - (NC - 1) * NORMAL_H) / (NC - 1))
    if GAP < 12:                      # 间隙太小 → 压矮项
        NORMAL_H = int(max(44, (avail - 12 * (NC - 1)) / (NC - 1)))
        GAP = int((avail - (NC - 1) * NORMAL_H) / (NC - 1))
    elif GAP > 64:                    # 间隙太大 → 加高项
        NORMAL_H = int(min(120, (avail - 64 * (NC - 1)) / (NC - 1)))
        GAP = int((avail - (NC - 1) * NORMAL_H) / (NC - 1))
GAP_ACTIVE = GAP                      # 等间隙：激活两侧与普通一致

CONTENT = (162, TOP_H + 16, W - 32, H - 28)


def vgrad(w, h, c1, c2):
    im = Image.new("RGB", (1, h))
    for y in range(h):
        t = y / max(1, h - 1)
        im.putpixel((0, y), tuple(int(c1[k] + (c2[k] - c1[k]) * t) for k in range(3)))
    return im.resize((w, h), Image.BILINEAR)


def dgrad(w, h, c1, c2):
    im = Image.new("RGB", (w, 1))
    for x in range(w):
        t = x / max(1, w - 1)
        im.putpixel((x, 0), tuple(int(c1[k] + (c2[k] - c1[k]) * t) for k in range(3)))
    return im.resize((w, h), Image.BILINEAR)


def rounded_png(name, w, h, r, top, bottom, border=None, vertical=True, bw=2):
    ww, hh = w * SS, h * SS
    im = (vgrad(ww, hh, top, bottom) if vertical else dgrad(ww, hh, top, bottom)).convert("RGBA")
    m = Image.new("L", (ww, hh), 0)
    ImageDraw.Draw(m).rounded_rectangle([0, 0, ww - 1, hh - 1], radius=r * SS, fill=255)
    out = Image.new("RGBA", (ww, hh), (0, 0, 0, 0)); out.paste(im, (0, 0), m)
    if border:
        ImageDraw.Draw(out).rounded_rectangle([0, 0, ww - 1, hh - 1], radius=r * SS, outline=border, width=bw * SS)
    out.save(os.path.join(IMG, name))
    return name


def bg_panel():
    w, h = W * SS, H * SS
    im = vgrad(w, h, (0xFD, 0xFE, 0xFF), (0xE4, 0xED, 0xF8)).convert("RGBA")
    x0, y0, x1, y1 = [v * SS for v in CONTENT]
    sh = Image.new("RGBA", (w, h), (0, 0, 0, 0))
    ImageDraw.Draw(sh).rounded_rectangle([x0 + 4, y0 + 6, x1 + 4, y1 + 6], radius=12 * SS, fill=(0x1B, 0x3A, 0x6B, 40))
    im = Image.alpha_composite(im, sh.filter(ImageFilter.GaussianBlur(8 * SS)))
    board = Image.new("RGBA", (w, h), (0, 0, 0, 0))
    d = ImageDraw.Draw(board)
    d.rounded_rectangle([x0, y0, x1, y1], radius=12 * SS, fill=(0xFF, 0xFF, 0xFF, 255))
    d.rounded_rectangle([x0, y0, x1, y1], radius=12 * SS, outline=(0xC9, 0xD8, 0xEA, 255), width=2 * SS)
    Image.alpha_composite(im, board).save(os.path.join(IMG, "bg.png"))


def topbar():
    w, h = W * SS, TOP_H * SS
    im = vgrad(w, h, (0x17, 0x33, 0x60), (0x3E, 0x7C, 0xC8)).convert("RGBA")
    sheen = Image.new("RGBA", (w, h), (0, 0, 0, 0))
    ImageDraw.Draw(sheen).polygon([(0, 0), (w, 0), (w, h * 0.5), (0, h * 0.72)], fill=(255, 255, 255, 34))
    im = Image.alpha_composite(im, sheen)
    d = ImageDraw.Draw(im)
    d.rectangle([0, h - 6 * SS, w, h], fill=(0x0E, 0x24, 0x45, 255))
    d.rectangle([0, h - 8 * SS, w, h - 6 * SS], fill=(0xE8, 0xB8, 0x4B, 255))
    im.save(os.path.join(IMG, "bar.png"))


def rail_panel():
    im = vgrad(RAIL_W * SS, RAIL_TOTAL * SS, (0xEE, 0xF4, 0xFB), (0xDC, 0xE8, 0xF6)).convert("RGBA")
    im.save(os.path.join(IMG, "rail.png"))


bg_panel(); topbar(); rail_panel()
button = rounded_png("button.png", RAIL_W, NORMAL_H, 8, (0xFF, 0xFF, 0xFF), (0xFF, 0xFF, 0xFF), (0xBF, 0xD3, 0xE8, 255))
hole_v = rounded_png("hole_v.png", RAIL_W + 10, ACTIVE_H + 10, 10, (0xF7, 0xFB, 0xFF), (0xE2, 0xEE, 0xFA), (0xB9, 0xCF, 0xE6, 255))
slider_v = rounded_png("slider_v.png", RAIL_W, ACTIVE_H, 8, (0x3A, 0x82, 0xC4), (0x18, 0x46, 0x84), (0xE8, 0xB8, 0x4B, 220))
hole_h = rounded_png("hole_h.png", BASE_TAB_W + 10, TAB_H + 10, 11, (0xF7, 0xFB, 0xFF), (0xDF, 0xEB, 0xF9), (0xB9, 0xCF, 0xE6, 255))
slider_h = rounded_png("slider_h.png", BASE_TAB_W, TAB_H, 9, (0x2F, 0x76, 0xB8), (0x6C, 0xB0, 0xEA), (0xFF, 0xFF, 0xFF, 200), vertical=False)
hole_h_big = rounded_png("hole_h_big.png", int(BASE_TAB_W * ACTIVE_SCALE) + 10, TAB_H + 10, 11, (0xF7, 0xFB, 0xFF), (0xDF, 0xEB, 0xF9), (0xB9, 0xCF, 0xE6, 255))
slider_h_big = rounded_png("slider_h_big.png", int(BASE_TAB_W * ACTIVE_SCALE), TAB_H, 9, (0x2F, 0x76, 0xB8), (0x6C, 0xB0, 0xEA), (0xFF, 0xFF, 0xFF, 200), vertical=False)

PAGES = [(t, c) for c in range(NC) for t in range(len(SUB_GROUPS[c]) if c < len(SUB_GROUPS) else NS)]
for idx, (ta, ra) in enumerate(PAGES, 1):
    heights = [ACTIVE_H if i == ra else NORMAL_H for i in range(NC)]
    gaps = [GAP_ACTIVE if (i == ra or i + 1 == ra) else GAP for i in range(NC - 1)]
    total = sum(heights) + sum(gaps)
    cur = RAIL_Y + max(6, (RAIL_TOTAL - total) // 2)
    ys = []
    for i, hh in enumerate(heights):
        ys.append(cur); cur += hh + (gaps[i] if i < len(gaps) else 0)

    p = ['<image href="../images/bg.png" x="0" y="0" width="%d" height="%d"/>' % (W, H)]
    p.append('<image href="../images/bar.png" x="0" y="0" width="%d" height="%d"/>' % (W, TOP_H))
    p.append('<image href="../images/rail.png" x="%d" y="%d" width="%d" height="%d"/>' % (RAIL_X, RAIL_Y, RAIL_W, RAIL_TOTAL))
    tw, tx = _layout_tabs(ta)
    p.append('<image href="../images/hole_h_big.png" x="%d" y="%d" width="%d" height="%d"/>'
             % (tx[ta] - 5, TAB_Y - 5, tw[ta] + 10, TAB_H + 10))
    p.append('<image href="../images/slider_h_big.png" x="%d" y="%d" width="%d" height="%d"/>'
             % (tx[ta], TAB_Y, tw[ta], TAB_H))
    for i in range(NC):
        if i == ra:
            p.append('<image href="../images/hole_v.png" x="%d" y="%d" width="%d" height="%d"/>'
                     % (RAIL_X - 5, ys[i] - 5, RAIL_W + 10, heights[i] + 10))
            p.append('<image href="../images/slider_v.png" x="%d" y="%d" width="%d" height="%d"/>'
                     % (RAIL_X, ys[i], RAIL_W, heights[i]))
        else:
            p.append('<image href="../images/button.png" x="%d" y="%d" width="%d" height="%d"/>'
                     % (RAIL_X, ys[i], RAIL_W, heights[i]))
    p.append('<text x="64" y="56" font-size="28" font-weight="bold" fill="#FFFFFF" '
             'font-family="Microsoft YaHei">%s</text>' % a.title)
    _subs = SUB_GROUPS[ra] if ra < len(SUB_GROUPS) else SUB
    NS_PAGE = len(_subs)
    for i, t in enumerate(_subs):
        col = "#FFFFFF" if i == ta else "#CFE0F3"
        fs = 18 if tw[i] >= 118 else 15
        p.append('<text x="%d" y="%d" font-size="%d" fill="%s" font-family="Microsoft YaHei" '
                 'text-anchor="middle">%s</text>' % (tx[i] + tw[i] // 2, TAB_Y + 31, fs, col, t))
    for i, t in enumerate(CH):
        col = "#FFFFFF" if i == ra else "#17335F"
        fs = 14 if heights[i] >= 90 else 12
        step = fs + 2
        y0 = ys[i] + heights[i] // 2 - (len(t) - 1) * step // 2 + fs // 3
        for k, ch in enumerate(t):
            p.append('<text x="%d" y="%d" font-size="%d" fill="%s" font-family="Microsoft YaHei" '
                     'text-anchor="middle">%s</text>' % (RAIL_X + RAIL_W // 2, y0 + k * step, fs, col, ch))
    cx0, cy0, cx1, cy1 = CONTENT
    bar_x = cx0 + 46; bar_w = int((cx1 - cx0) * 0.46)
    for (yy, hh, ww) in [(cy0 + 56, 18, bar_w * 0.55), (cy0 + 96, 11, bar_w), (cy0 + 118, 11, bar_w * 0.86),
                         (cy0 + 140, 11, bar_w * 0.72), (cy0 + 190, 11, bar_w * 0.94), (cy0 + 212, 11, bar_w * 0.66)]:
        p.append('<rect x="%d" y="%d" width="%d" height="%d" rx="4" fill="#C9D8EA" opacity="0.55"/>'
                 % (bar_x, int(yy), int(ww), hh))
    img_x = bar_x + bar_w + 40; img_w = cx1 - img_x - 40
    p.append('<rect x="%d" y="%d" width="%d" height="%d" rx="10" fill="#DCE8F6" stroke="#B9CFE6" stroke-width="2"/>'
             % (img_x, cy0 + 56, img_w, (cy1 - cy0) - 112))
    p.append('<polygon points="%d,%d %d,%d %d,%d" fill="#E8B84B"/>'
             % (RAIL_X + RAIL_W + 2, ys[ra] + heights[ra] // 2 - 10,
                RAIL_X + RAIL_W + 18, ys[ra] + heights[ra] // 2,
                RAIL_X + RAIL_W + 2, ys[ra] + heights[ra] // 2 + 10))
    open(os.path.join(SVG, "slide_%02d.svg" % idx), "w", encoding="utf-8").write(
        '<svg xmlns="http://www.w3.org/2000/svg" width="%d" height="%d" viewBox="0 0 %d %d" lang="zh-CN">\n%s\n</svg>'
        % (W, H, W, H, "\n".join(p)))

print("章节 %d × 子页 %d = %d 页 | 左栏 激活%d/普通%d 等间隙%d | 顶栏 普通%d/激活%d 间隙%d 右对齐"
      % (NC, NS, len(PAGES), ACTIVE_H, NORMAL_H, GAP, BASE_TAB_W, int(BASE_TAB_W * ACTIVE_SCALE), TAB_GAP))

subprocess.run([PY, os.path.join(SK, "scripts", "svg_quality_checker.py"), WS,
                "--quick-generate", "--canonical-authoring", "--stage", "final", "--json"], capture_output=True)
r = subprocess.run([PY, os.path.join(SK, "scripts", "svg_to_pptx.py"), WS, "-f", "ppt169",
                    "--quick-generate", "-t", "none", "-o", a.out], capture_output=True, text=True,
                   encoding="utf-8", errors="ignore")
if r.returncode != 0:
    print("export 失败", (r.stderr or "")[-400:]); sys.exit(1)

# 注入 morph(byChar) + 关闭换行
import zipfile, re
BLOCK = ('<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">'
         '<mc:Choice xmlns:p159="http://schemas.microsoft.com/office/powerpoint/2015/09/main" Requires="p159">'
         '<p:transition spd="slow" p14:dur="1500" '
         'xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main">'
         '<p159:morph option="byChar"/></p:transition></mc:Choice>'
         '<mc:Fallback><p:transition spd="slow"><p:fade/></p:transition></mc:Fallback></mc:AlternateContent>')
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
z = zipfile.ZipFile(a.out); items = {n: z.read(n) for n in z.namelist()}; z.close()
for name in list(items):
    if not re.fullmatch(r'ppt/slides/slide\d+\.xml', name):
        continue
    s = items[name].decode('utf-8')
    s = re.sub(r'<mc:AlternateContent\b.*?</mc:AlternateContent>', '', s, flags=re.S)
    s = re.sub(r'<p:transition\b[^>]*/>', '', s)
    s = s.replace('</p:clrMapOvr>', '</p:clrMapOvr>' + BLOCK, 1)
    from lxml import etree
    root = etree.fromstring(s.encode('utf-8'))
    for body in root.iter('{%s}bodyPr' % A):
        body.set('wrap', 'none')
    for sp in root.iter('{%s}sp' % P):
        off = sp.find('{%s}spPr/{%s}xfrm/{%s}off' % (P, A, A))
        ext = sp.find('{%s}spPr/{%s}xfrm/{%s}ext' % (P, A, A))
        if off is None or ext is None or sp.find('.//{%s}t' % A) is None:
            continue
        if int(off.get('y')) < 1200000 and int(off.get('x')) > 600 * 9525:
            off.set('x', str(int(off.get('x')) - 200000))
            ext.set('cx', str(int(ext.get('cx')) + 400000))
    items[name] = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
with zipfile.ZipFile(a.out, 'w', zipfile.ZIP_DEFLATED) as zo:
    for k, b in items.items():
        zo.writestr(k, b)
print("saved", a.out, round(os.path.getsize(a.out) / 1024), "KB")
