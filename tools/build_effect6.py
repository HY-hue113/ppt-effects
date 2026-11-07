import zipfile, io, os
from PIL import Image

SRC = r'<USER_HOME>\Documents\<DSH_HOME>\default-workspace\_ref6\ref.pptx'
DST = r'<USER_HOME>\Documents\<DSH_HOME>\default-workspace\preview\效果6.pptx'
IMG = r'<DSH_ATTACHMENTS>\v1\objects\46\469e386a512252f3920a9fc9d1fbbaafcf0760a8d520952a700341f8926b2633'


def cover(im, tw, th):
    sw, sh = im.size
    s = max(tw / sw, th / sh)
    nw, nh = int(sw * s + 0.5), int(sh * s + 0.5)
    im = im.resize((nw, nh), Image.LANCZOS)
    l, t = (nw - tw) // 2, (nh - th) // 2
    return im.crop((l, t, l + tw, t + th))


with zipfile.ZipFile(SRC) as z:
    items = {n: z.read(n) for n in z.namelist()}

part = 'ppt/media/image1.jpeg'
tgt = Image.open(io.BytesIO(items[part]))
im = cover(Image.open(IMG).convert('RGB'), *tgt.size)
buf = io.BytesIO()
im.save(buf, 'JPEG', quality=93)
items[part] = buf.getvalue()

with zipfile.ZipFile(DST, 'w', zipfile.ZIP_DEFLATED) as z:
    for n, d in items.items():
        z.writestr(n, d)
print('built', DST, os.path.getsize(DST) // 1024, 'KB', '| image', tgt.size)
