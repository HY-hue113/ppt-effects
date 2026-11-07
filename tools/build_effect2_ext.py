import zipfile, re, io, os
import numpy as np
from lxml import etree
from PIL import Image

SRC = r'<USER_HOME>\Documents\<DSH_HOME>\default-workspace\_ref4\ref.pptx'
DST = r'<USER_HOME>\Documents\<DSH_HOME>\default-workspace\preview\效果2_长版.pptx'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
CT = 'http://schemas.openxmlformats.org/package/2006/content-types'
PR = 'http://schemas.openxmlformats.org/package/2006/relationships'
KEEP = list(range(5, 24))

ATT = r'<DSH_ATTACHMENTS>\v1\objects'
POOL = {
    '草原': ATT + r'\be\beeab8ad8c130c6da6560e4626f40ffd7ffdb2dd4211bc322a68ee03e1238f31',
    '国门': ATT + r'\a8\a8eb323b40a77068b37401a8a34cfff3a6e9aff7542b01ef20f4a536e441ffd9',
    '湖':   ATT + r'\46\469e386a512252f3920a9fc9d1fbbaafcf0760a8d520952a700341f8926b2633',
    '开学': ATT + r'\f5\f544eae007a239534e0e8bcee7c7de03e9d4122587d4977bae3b643d4b6c01e9',
    '盘山': ATT + r'\0d\0d3944e073c337e85bdf488e2650367d661045bc068c10524fbc8bcc485a7bf6',
}


def sig(im):
    a = np.asarray(im.convert('RGB').resize((48, 27))).astype(float)
    return np.concatenate([a.mean(2).ravel(), a.std(2).ravel()])


def cover(im, tw, th):
    sw, sh = im.size
    s = max(tw / sw, th / sh)
    nw, nh = int(sw * s + 0.5), int(sh * s + 0.5)
    im = im.resize((nw, nh), Image.LANCZOS)
    l, t = (nw - tw) // 2, (nh - th) // 2
    return im.crop((l, t, l + tw, t + th))


with zipfile.ZipFile(SRC) as z:
    items = {n: z.read(n) for n in z.namelist()}

pool_imgs = {k: Image.open(v).convert('RGB') for k, v in POOL.items()}
pool_sigs = {k: sig(im) for k, im in pool_imgs.items()}

media = [n for n in items if re.match(r'ppt/media/image\d+\.(jpe?g|png)$', n)]
mapping = {}
for part in media:
    blob = items[part]
    try:
        cur = Image.open(io.BytesIO(blob)).convert('RGB')
    except Exception:
        continue
    cs = sig(cur)
    best = min(pool_sigs, key=lambda k: np.linalg.norm(pool_sigs[k] - cs))
    dist = np.linalg.norm(pool_sigs[best] - cs)
    mapping[part] = (best, round(float(dist), 1))
    # 只替换真正被 5-23 页用到的槽位，避免误改其它页
print('media match:', {k.split("/")[-1]: v for k, v in mapping.items()})

# 找出 5-23 页实际引用的媒体
used = set()
for i in KEEP:
    rels = items['ppt/slides/_rels/slide%d.xml.rels' % i].decode()
    s = items['ppt/slides/slide%d.xml' % i].decode()
    for rid in dict.fromkeys(re.findall(r'r:embed="(rId\d+)"', s)):
        m = re.search(r'Id="%s"[^>]*Target="\.\./media/([^"]+)"' % rid, rels)
        if m:
            used.add('ppt/media/' + m.group(1))
print('used by 5-23:', sorted(x.split('/')[-1] for x in used))

for part in used:
    name, dist = mapping.get(part, (None, None))
    if name is None:
        continue
    tgt = Image.open(io.BytesIO(items[part]))
    im = cover(pool_imgs[name], *tgt.size)
    buf = io.BytesIO()
    if part.lower().endswith('.png'):
        im.save(buf, 'PNG', optimize=True)
    else:
        im.save(buf, 'JPEG', quality=90)
    items[part] = buf.getvalue()
    print('replaced', part.split('/')[-1], '<-', name, 'dist', dist)

pres = etree.fromstring(items['ppt/presentation.xml'])
rels = etree.fromstring(items['ppt/_rels/presentation.xml.rels'])
rid2target = {r.get('Id'): r.get('Target') for r in rels.findall('{%s}Relationship' % PR)}
sld_ids = pres.find('.//{%s}sldIdLst' % P)
drop = set()
for sld in list(sld_ids):
    tgt = rid2target.get(sld.get('{%s}id' % R), '')
    m = re.match(r'slides/slide(\d+)\.xml$', tgt)
    num = int(m.group(1)) if m else -1
    if num not in KEEP:
        sid = sld.get('{%s}id' % R)
        sld_ids.remove(sld)
        for rel in rels.findall('{%s}Relationship' % PR):
            if rel.get('Id') == sid:
                rels.remove(rel)
        if num > 0:
            drop.add(num)
items['ppt/presentation.xml'] = etree.tostring(pres, xml_declaration=True, encoding='UTF-8', standalone=True)
items['ppt/_rels/presentation.xml.rels'] = etree.tostring(rels, xml_declaration=True, encoding='UTF-8', standalone=True)

for n in list(items):
    if n.startswith('ppt/slides/') or n.startswith('ppt/notesSlides/'):
        m = re.search(r'(?:slide|notesSlide)(\d+)\.xml', n)
        if m and int(m.group(1)) in drop:
            items.pop(n)

ct = etree.fromstring(items['[Content_Types].xml'])
for ov in list(ct.findall('{%s}Override' % CT)):
    if ov.get('PartName', '').lstrip('/') not in items:
        ct.remove(ov)
items['[Content_Types].xml'] = etree.tostring(ct, xml_declaration=True, encoding='UTF-8', standalone=True)

with zipfile.ZipFile(DST, 'w', zipfile.ZIP_DEFLATED) as z:
    for n, d in items.items():
        z.writestr(n, d)
print('built', DST, os.path.getsize(DST) // 1024, 'KB')
