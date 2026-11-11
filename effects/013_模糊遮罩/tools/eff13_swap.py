"""效果13 · 模糊遮罩（模板克隆法）

复制模板 → 替换 image1.jpg（清晰）与 grass_blur.jpg（模糊）→ 把各页指向 .wdp 的关系改指模糊图。
造型、文字、层级、morph 一律不动。

用法:
    python eff13_swap.py --template 效果13.pptx --out 输出.pptx 图.jpg
"""
import zipfile, re, os, io, argparse
from PIL import Image, ImageFilter, ImageEnhance

ap = argparse.ArgumentParser()
ap.add_argument('--template', required=True)
ap.add_argument('--out', required=True)
ap.add_argument('--blur', type=float, default=50)
ap.add_argument('--bright', type=float, default=-10)
ap.add_argument('image')
a = ap.parse_args()

im = Image.open(a.image).convert('RGB')
tw, th = 1920, 1080
sc = max(tw / im.width, th / im.height)
im2 = im.resize((int(im.width * sc + .5), int(im.height * sc + .5)), Image.LANCZOS)
l, t = (im2.width - tw) // 2, (im2.height - th) // 2
sharp = im2.crop((l, t, l + tw, t + th))
blur = ImageEnhance.Brightness(sharp.filter(ImageFilter.GaussianBlur(a.blur))).enhance(1 + a.bright / 100.0)

b1, b2 = io.BytesIO(), io.BytesIO()
sharp.save(b1, 'JPEG', quality=92)
blur.save(b2, 'JPEG', quality=90)

z = zipfile.ZipFile(a.template)
items = {n: z.read(n) for n in z.namelist()}
items['ppt/media/image1.jpg'] = b1.getvalue()
items['ppt/media/grass_blur.jpg'] = b2.getvalue()
for k in list(items):
    if re.match(r'ppt/slides/_rels/slide\d+\.xml\.rels$', k):
        s = items[k].decode('utf-8')
        items[k] = re.sub(r'Target="\.\./media/[^"]+\.wdp"',
                          'Target="../media/grass_blur.jpg"', s).encode('utf-8')

with zipfile.ZipFile(a.out, 'w', zipfile.ZIP_DEFLATED) as zo:
    for k, v in items.items():
        zo.writestr(k, v)
print('输出', a.out, round(os.path.getsize(a.out) / 1024), 'KB')
