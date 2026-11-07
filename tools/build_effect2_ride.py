import zipfile, re, io, os
from lxml import etree
from PIL import Image

SRC = r'<USER_HOME>\Documents\<DSH_HOME>\default-workspace\_ref5\ref.pptx'
DST = r'<USER_HOME>\Documents\<DSH_HOME>\default-workspace\preview\效果2_骑行.pptx'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
CT = 'http://schemas.openxmlformats.org/package/2006/content-types'
PR = 'http://schemas.openxmlformats.org/package/2006/relationships'
NS = {'a': A, 'p': P, 'r': R}
KEEP = list(range(5, 24))          # 原稿 5-23 页 -> 输出 1-19 页

ATT = r'<DSH_ATTACHMENTS>\v1\objects'
PHOTO = {
    '骑行1': ATT + r'\54\545f179f431c8f72c2b8bc8a1e696938dfc2b17549ade2d58c2fa9fb7c2b1489',
    '骑行3': ATT + r'\2d\2d853732b7c19c117b1f8bc3fceb2c3bf987351c51c179b001270ce17f722ac9',
    '骑行4': ATT + r'\e1\e199aa1e3d2100c020a00aa2311d63d2e3289eacf28b65cccdbb4d2d5b8fa87b',
    '骑行6': ATT + r'\de\de404a25a93629914e41a74c0bc9455d2efc9effff28438d302ce3d83b45bcf1',
}
# 输出页 1-6 / 7-12 / 13-18 / 19
PAGE_PHOTO = {}
for n in range(1, 20):
    PAGE_PHOTO[n] = '骑行1' if n <= 6 else '骑行3' if n <= 12 else '骑行4' if n <= 18 else '骑行6'


def cover(im, tw, th):
    sw, sh = im.size
    s = max(tw / sw, th / sh)
    nw, nh = int(sw * s + 0.5), int(sh * s + 0.5)
    im = im.resize((nw, nh), Image.LANCZOS)
    l, t = (nw - tw) // 2, (nh - th) // 2
    return im.crop((l, t, l + tw, t + th))


with zipfile.ZipFile(SRC) as z:
    items = {n: z.read(n) for n in z.namelist()}

photos = {k: Image.open(v).convert('RGB') for k, v in PHOTO.items()}

# 先确定四栏拼贴页的栏位顺序：取一页，按 x 从左到右取 4 个 9144000×6858000 的图片
probe = etree.fromstring(items['ppt/slides/slide16.xml'])
probe_rels = etree.fromstring(items['ppt/slides/_rels/slide16.xml.rels'])
rel_ns_p = probe_rels.nsmap[None]
rid2media = {}
for rel in probe_rels.findall('{%s}Relationship' % rel_ns_p):
    rid2media[rel.get('Id')] = rel.get('Target').replace('../', 'ppt/')

cols = []
for pic in probe.iter('{%s}pic' % P):
    xf = pic.find('.//a:xfrm', NS)
    if xf is None:
        continue
    ext = xf.find('a:ext', NS)
    if ext is None:
        continue
    if abs(int(ext.get('cx')) - 9144000) < 50000 and abs(int(ext.get('cy')) - 6858000) < 50000:
        off = xf.find('a:off', NS)
        blip = pic.find('.//a:blip', NS)
        cols.append((int(off.get('x')), rid2media.get(blip.get('{%s}embed' % R))))
cols.sort()
print('collage columns ->', [c[1] for c in cols])

COL_PHOTO = ['骑行1', '骑行3', '骑行4', '骑行6']
for (_, media_part), name in zip(cols, COL_PHOTO):
    if not media_part:
        continue
    tgt = Image.open(io.BytesIO(items[media_part]))
    im = cover(photos[name], *tgt.size)
    buf = io.BytesIO()
    if media_part.lower().endswith('.png'):
        im.save(buf, 'PNG', optimize=True)
    else:
        im.save(buf, 'JPEG', quality=90)
    items[media_part] = buf.getvalue()
    print('collage', media_part, '<-', name)


for idx, out_page in zip(KEEP, range(1, 20)):
    sname = 'ppt/slides/slide%d.xml' % idx
    rname = 'ppt/slides/_rels/slide%d.xml.rels' % idx
    root = etree.fromstring(items[sname])
    rels = etree.fromstring(items[rname])
    rel_ns = rels.nsmap[None]

    # 找铺满页面的背景形状（有 blipFill 的那个）
    target = None
    for sp in root.iter('{%s}sp' % P):
        xf = sp.find('.//a:xfrm', NS)
        if xf is None:
            continue
        ext = xf.find('a:ext', NS)
        try:
            cx, cy = int(ext.get('cx')), int(ext.get('cy'))
        except Exception:
            continue
        if abs(cx - 12192000) < 200000 and abs(cy - 6858000) < 200000:
            blip = sp.find('.//a:blip', NS)
            if blip is not None:
                target = (sp, blip)
                break
    if target is None:
        bg = root.find('.//p:bg//a:blip', NS)
        if bg is not None:
            target = (None, bg)
    if target is None:
        print('no bg blip on slide', idx)
        continue

    sp, blip = target
    old_rid = blip.get('{%s}embed' % R)
    old_media = None
    for rel in rels.findall('{%s}Relationship' % rel_ns):
        if rel.get('Id') == old_rid:
            old_media = rel.get('Target').replace('../', 'ppt/')
    size = Image.open(io.BytesIO(items[old_media])).size

    name = PAGE_PHOTO[out_page]
    im = cover(photos[name], *size)
    buf = io.BytesIO()
    im.save(buf, 'JPEG', quality=90)
    part = 'ppt/media/bgp%02d.jpg' % out_page
    items[part] = buf.getvalue()

    new_rid = 'rIdBg%d' % out_page
    rel = etree.SubElement(rels, '{%s}Relationship' % rel_ns)
    rel.set('Id', new_rid)
    rel.set('Type', R + '/image')
    rel.set('Target', '../media/bgp%02d.jpg' % out_page)
    blip.set('{%s}embed' % R, new_rid)

    items[sname] = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
    items[rname] = etree.tostring(rels, xml_declaration=True, encoding='UTF-8', standalone=True)
    print('slide', idx, '-> page', out_page, name, size)

# 四栏拼贴页：去掉「盘山/喀纳斯/国门/草原」四个标签，只留数字
LABELS = {'盘山', '喀纳斯', '国门', '草原'}
for idx in range(16, 24):
    sname = 'ppt/slides/slide%d.xml' % idx
    if sname not in items:
        continue
    root = etree.fromstring(items[sname])
    tree = root.find('.//p:spTree', NS)
    removed = 0
    for sp in list(tree):
        if etree.QName(sp).localname != 'sp':
            continue
        txt = ''.join(t.text or '' for t in sp.iter('{%s}t' % A)).strip()
        if txt in LABELS:
            tree.remove(sp)
            removed += 1
    if removed:
        items[sname] = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
    print('slide', idx, 'labels removed:', removed)

# 裁页
pres = etree.fromstring(items['ppt/presentation.xml'])
p_rels = etree.fromstring(items['ppt/_rels/presentation.xml.rels'])
rid2target = {r.get('Id'): r.get('Target') for r in p_rels.findall('{%s}Relationship' % PR)}
sld_ids = pres.find('.//{%s}sldIdLst' % P)
drop = set()
for sld in list(sld_ids):
    tgt = rid2target.get(sld.get('{%s}id' % R), '')
    m = re.match(r'slides/slide(\d+)\.xml$', tgt)
    num = int(m.group(1)) if m else -1
    if num not in KEEP:
        sid = sld.get('{%s}id' % R)
        sld_ids.remove(sld)
        for rel in p_rels.findall('{%s}Relationship' % PR):
            if rel.get('Id') == sid:
                p_rels.remove(rel)
        if num > 0:
            drop.add(num)
items['ppt/presentation.xml'] = etree.tostring(pres, xml_declaration=True, encoding='UTF-8', standalone=True)
items['ppt/_rels/presentation.xml.rels'] = etree.tostring(p_rels, xml_declaration=True, encoding='UTF-8', standalone=True)

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
