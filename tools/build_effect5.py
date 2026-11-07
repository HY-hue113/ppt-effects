import zipfile, re, io, os
from lxml import etree
from PIL import Image

SRC = r'<USER_HOME>\Documents\<DSH_HOME>\default-workspace\_ref3\ref.pptx'
DST = r'<USER_HOME>\Documents\<DSH_HOME>\default-workspace\preview\效果5.pptx'
IMG = r'<DSH_ATTACHMENTS>\v1\objects\9a\9a352b70e3df5cc941a189c231b9f3476d54a3f0ac96432e36ba13bc897db400'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
CT = 'http://schemas.openxmlformats.org/package/2006/content-types'
PR = 'http://schemas.openxmlformats.org/package/2006/relationships'
KEEP = [17, 18, 19]


def cover(im, tw, th):
    sw, sh = im.size
    s = max(tw / sw, th / sh)
    nw, nh = int(sw * s + 0.5), int(sh * s + 0.5)
    im = im.resize((nw, nh), Image.LANCZOS)
    l, t = (nw - tw) // 2, (nh - th) // 2
    return im.crop((l, t, l + tw, t + th))


with zipfile.ZipFile(SRC) as z:
    items = {n: z.read(n) for n in z.namelist()}

part = 'ppt/media/image2.jpeg'
tgt = Image.open(io.BytesIO(items[part]))
im = cover(Image.open(IMG).convert('RGB'), *tgt.size)
buf = io.BytesIO()
im.save(buf, 'JPEG', quality=93)
items[part] = buf.getvalue()

pres = etree.fromstring(items['ppt/presentation.xml'])
rels = etree.fromstring(items['ppt/_rels/presentation.xml.rels'])
rid2target = {r.get('Id'): r.get('Target') for r in rels.findall('{%s}Relationship' % PR)}
sld_ids = pres.find('.//{%s}sldIdLst' % P)
drop = set()
for sld in list(sld_ids):
    tgt_ = rid2target.get(sld.get('{%s}id' % R), '')
    m = re.match(r'slides/slide(\d+)\.xml$', tgt_)
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
