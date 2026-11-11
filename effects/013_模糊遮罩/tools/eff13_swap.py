"""效果13 · 模糊遮罩（模板克隆法）

支持两种用法：
  1 张图  → 5 页共用（原设计）
  N 张图  → 每页一张，N ≤ 5；不足则循环复用

每页固定三组媒体，按角色整体替换：
  清晰层 image1/3/5、hdphoto*  ← 原图
  模糊层 image2/4/6、grass_blur ← 同图模糊版
造型、文字、层级、morph 一律不动。

用法:
    python eff13_swap.py --template 效果13.pptx --out 输出.pptx 图1.jpg [图2 ...]
    python eff13_swap.py --template 效果13.pptx --out 输出.pptx --blur 60 --bright -15 图.jpg
"""
import argparse
import io
import os
import re
import zipfile

from PIL import Image, ImageEnhance, ImageFilter

ap = argparse.ArgumentParser()
ap.add_argument('--template', required=True)
ap.add_argument('--out', required=True)
ap.add_argument('--blur', type=float, default=50)
ap.add_argument('--bright', type=float, default=-10)
ap.add_argument('images', nargs='+')
a = ap.parse_args()

TW, TH = 1920, 1080


def variants(path):
    im = Image.open(path).convert('RGB')
    sc = max(TW / im.width, TH / im.height)
    im2 = im.resize((int(im.width * sc + .5), int(im.height * sc + .5)), Image.LANCZOS)
    l, t = (im2.width - TW) // 2, (im2.height - TH) // 2
    sharp = im2.crop((l, t, l + TW, t + TH))
    blur = ImageEnhance.Brightness(
        sharp.filter(ImageFilter.GaussianBlur(a.blur))).enhance(1 + a.bright / 100.0)
    b1, b2 = io.BytesIO(), io.BytesIO()
    sharp.save(b1, 'JPEG', quality=92)
    blur.save(b2, 'JPEG', quality=90)
    return b1.getvalue(), b2.getvalue()


PAIRS = [variants(p) for p in a.images]

SHARP_ROLES = {'image1.jpg', 'image3.jpg', 'image5.jpg',
               'hdphoto1.jpg', 'hdphoto2.jpg', 'hdphoto3.jpg'}
BLUR_ROLES = {'image2.jpg', 'image4.jpg', 'image6.jpg', 'grass_blur.jpg'}

z = zipfile.ZipFile(a.template)
items = {n: z.read(n) for n in z.namelist()}
z.close()

slides = sorted((n for n in items if re.fullmatch(r'ppt/slides/slide\d+\.xml', n)),
                key=lambda s: int(re.search(r'slide(\d+)', s).group(1)))

for idx, sname in enumerate(slides):
    sharp, blur = PAIRS[idx % len(PAIRS)]
    rname = sname.replace('slides/', 'slides/_rels/') + '.rels'
    if rname not in items:
        continue
    rels = items[rname].decode('utf-8')
    targets = re.findall(r'Target="\.\./media/([^"]+)"', rels)

    for base in targets:
        if base in SHARP_ROLES:
            new = 'eff13_p%d_sharp.jpg' % (idx + 1)
            items['ppt/media/' + new] = sharp
        elif base in BLUR_ROLES:
            new = 'eff13_p%d_blur.jpg' % (idx + 1)
            items['ppt/media/' + new] = blur
        else:
            continue
        rels = rels.replace('Target="../media/%s"' % base, 'Target="../media/%s"' % new)

    # 仍指向 .wdp 的旧参考件
    rels = re.sub(r'Target="\.\./media/[^"]+\.wdp"', 'Target="../media/eff13_p%d_sharp.jpg"' % (idx + 1), rels)
    items[rname] = rels.encode('utf-8')

ct = items['[Content_Types].xml'].decode('utf-8')
if 'Extension="jpg"' not in ct:
    ct = ct.replace('</Types>', '<Default Extension="jpg" ContentType="image/jpeg"/></Types>')
items['[Content_Types].xml'] = ct.encode('utf-8')

with zipfile.ZipFile(a.out, 'w', zipfile.ZIP_DEFLATED) as zo:
    for k, v in items.items():
        zo.writestr(k, v)
print('输出', a.out, round(os.path.getsize(a.out) / 1024), 'KB | 图片', len(PAIRS), '张 |', len(slides), '页')
