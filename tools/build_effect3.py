import zipfile, re, io, os
from lxml import etree
from PIL import Image

SRC = r'<USER_HOME>\Documents\<DSH_HOME>\default-workspace\_ref3\ref.pptx'
DST = r'<USER_HOME>\Documents\<DSH_HOME>\default-workspace\preview\哲学存在.pptx'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
CT = 'http://schemas.openxmlformats.org/package/2006/content-types'
PR = 'http://schemas.openxmlformats.org/package/2006/relationships'

# 画幅比例匹配：3-巴门尼德 1.76 / 5-海 1.58 / 7-亚 1.33 / 9-波 1.89 / 11-康 1.68
IMG = {
    'ppt/media/image2.jpeg': r'<DSH_ATTACHMENTS>\v1\objects\9a\9a352b70e3df5cc941a189c231b9f3476d54a3f0ac96432e36ba13bc897db400',
    'ppt/media/image3.png':  r'<DSH_ATTACHMENTS>\v1\objects\f6\f6d0a9604b7bc0fee1b0004b8a7180f3865734e5707f12d4f0c0bfc6f1a411c3',
    'ppt/media/image4.jpeg': r'<DSH_ATTACHMENTS>\v1\objects\7a\7ae154e849e04052d9c0276879632a512bb021b5521624fd764543750783e648',
    'ppt/media/image5.jpeg': r'<DSH_ATTACHMENTS>\v1\objects\8b\8bd85f6b50ca6ed6edb352774b562239d3fdd993d829569b033b5df7e9bed03a',
    'ppt/media/image6.jpeg': r'<DSH_ATTACHMENTS>\v1\objects\79\79f61ae7e25e75a6be8c10ac46da78382778946ce7856bc37ee5e9a8c000d08c',
}
KEEP = 12

with zipfile.ZipFile(SRC) as z:
    items = {n: z.read(n) for n in z.namelist()}

# 1) 媒体替换：按目标尺寸等比缩放（两者画幅已匹配）
for part, src in IMG.items():
    im = Image.open(src).convert('RGB')
    tgt = Image.open(io.BytesIO(items[part]))
    im = im.resize(tgt.size, Image.LANCZOS)
    out = io.BytesIO()
    if part.lower().endswith('.png'):
        im.save(out, 'PNG', optimize=True)
    else:
        im.save(out, 'JPEG', quality=92)
    items[part] = out.getvalue()

# 2) 只保留前 12 页
pres = etree.fromstring(items['ppt/presentation.xml'])
rels = etree.fromstring(items['ppt/_rels/presentation.xml.rels'])
rid2target = {r.get('Id'): r.get('Target') for r in rels.findall('{%s}Relationship' % PR)}

sld_ids = pres.find('.//{%s}sldIdLst' % P)
removed = set()
for sld in list(sld_ids):
    rid = sld.get('{%s}id' % R)
    tgt = rid2target.get(rid, '')
    m = re.match(r'slides/slide(\d+)\.xml$', tgt)
    if not m or int(m.group(1)) > KEEP:
        sld_ids.remove(sld)
        for rel in rels.findall('{%s}Relationship' % PR):
            if rel.get('Id') == rid:
                rels.remove(rel)
                removed.add(rel.get('Target'))

items['ppt/presentation.xml'] = etree.tostring(pres, xml_declaration=True, encoding='UTF-8', standalone=True)
items['ppt/_rels/presentation.xml.rels'] = etree.tostring(rels, xml_declaration=True, encoding='UTF-8', standalone=True)

# 3) 删除多余页与其备注、清理内容类型
drop = set()
for n in list(items):
    m = re.match(r'ppt/(slides|notesSlides)/slide(\d+)\.xml$', n) or re.match(r'ppt/(slides|notesSlides)/_rels/slide(\d+)\.xml\.rels$', n)
    if n.startswith('ppt/slides/') or n.startswith('ppt/notesSlides/'):
        m2 = re.search(r'(?:slide|notesSlide)(\d+)\.xml', n)
        if m2 and int(m2.group(1)) > KEEP:
            drop.add(n)
for n in drop:
    items.pop(n, None)

ct = etree.fromstring(items['[Content_Types].xml'])
for ov in list(ct.findall('{%s}Override' % CT)):
    part = ov.get('PartName', '').lstrip('/')
    if part not in items:
        ct.remove(ov)
items['[Content_Types].xml'] = etree.tostring(ct, xml_declaration=True, encoding='UTF-8', standalone=True)

with zipfile.ZipFile(DST, 'w', zipfile.ZIP_DEFLATED) as z:
    for n, d in items.items():
        z.writestr(n, d)

print('built', DST, os.path.getsize(DST) // 1024, 'KB')
print('slides kept:', len(sld_ids))
