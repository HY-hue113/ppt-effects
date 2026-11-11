"""效果11 · 严格对应填充（按形状名）

规则：小四边形副本K图 ← 第 K 张图片；小四边形副本K = 纯色小色块 K；米色板不变。
只在 ppt-master 导出的模板上改图片关系，不动形状/动画。
"""
import os, re, zipfile, shutil
from lxml import etree

SRC = r'<USER_HOME>\<OUTPUT_DIR>\效果11-模板套用.pptx'
OUT = r'<USER_HOME>\<OUTPUT_DIR>\效果11-严格对应.pptx'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
PR = 'http://schemas.openxmlformats.org/package/2006/relationships'

IMGS = {
    1: r'<DSH_ATTACHMENTS>\v1\objects\a8\a8eb323b40a77068b37401a8a34cfff3a6e9aff7542b01ef20f4a536e441ffd9',
    2: r'<DSH_ATTACHMENTS>\v1\objects\46\469e386a512252f3920a9fc9d1fbbaafcf0760a8d520952a700341f8926b2633',
    3: r'<DSH_ATTACHMENTS>\v1\objects\e9\e9466e1a860cc7cf36247873fad1fecaf32eed76615d522ac3ce24d95e2e5a8b',
    4: r'<DSH_ATTACHMENTS>\v1\objects\0d\0d3944e073c337e85bdf488e2650367d661045bc068c10524fbc8bcc485a7bf6',
}

z = zipfile.ZipFile(SRC)
items = {n: z.read(n) for n in z.namelist()}
z.close()

slides = sorted([n for n in items if re.match(r'ppt/slides/slide\d+\.xml$', n)],
                key=lambda s: int(re.search(r'slide(\d+)', s).group(1)))

for sname in slides:
    idx = int(re.search(r'slide(\d+)', sname).group(1))
    rels_name = 'ppt/slides/_rels/slide%d.xml.rels' % idx
    root = etree.fromstring(items[sname])
    rels = etree.fromstring(items[rels_name])
    next_rid = max(int(r.get('Id')[3:]) for r in rels) + 1
    changed = []

    for sp in root.iter('{%s}sp' % P):
        cNv = sp.find('.//{%s}cNvPr' % P)
        if cNv is None:
            continue
        nm = cNv.get('name', '')
        m = re.fullmatch(r'小四边形副本(\d+)图', nm)
        if not m:
            continue
        k = int(m.group(1))
        if k not in IMGS:
            continue
        spPr = sp.find('{%s}spPr' % P)
        blip = spPr.find('{%s}blipFill/{%s}blip' % (A, A)) if spPr is not None else None
        if blip is None:
            continue
        media = 'ppt/media/eff11k_%d.jpg' % k
        if media not in items:
            items[media] = open(IMGS[k], 'rb').read()
        rid = 'rId%d' % next_rid
        next_rid += 1
        e = etree.SubElement(rels, '{%s}Relationship' % PR)
        e.set('Id', rid)
        e.set('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image')
        e.set('Target', '../media/' + os.path.basename(media))
        blip.set('{%s}embed' % R, rid)
        changed.append(nm)

    items[sname] = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
    items[rels_name] = etree.tostring(rels, xml_declaration=True, encoding='UTF-8', standalone=True)
    print('p%d 填充: %s' % (idx, changed))

with zipfile.ZipFile(OUT, 'w', zipfile.ZIP_DEFLATED) as zo:
    for n, b in items.items():
        zo.writestr(n, b)
print('done', OUT, round(os.path.getsize(OUT) / 1024), 'KB')
