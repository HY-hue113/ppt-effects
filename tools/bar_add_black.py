import copy, shutil, os
from pptx import Presentation
from lxml import etree

F = r'<USER_HOME>\<OUTPUT_DIR>\效果10.pptx'
BAK = r'<USER_HOME>\<OUTPUT_DIR>\效果10-备份.pptx'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
NS = {'a': A, 'p': P}
MARGIN = 60000     # 每边外扩 60000 EMU ≈ 0.17cm

if not os.path.exists(BAK):
    shutil.copyfile(F, BAK)

prs = Presentation(F)
added = []

for idx, slide in enumerate(prs.slides, start=1):
    spTree = slide.shapes._spTree
    targets = []
    for sh in slide.shapes:
        if sh.shape_type == 6:
            continue
        w = (sh.width or 0) / 360000.0
        h = (sh.height or 0) / 360000.0
        if not (11.0 <= w <= 16.0 and 1.2 <= h <= 2.0):
            continue
        txt = sh.text_frame.text.strip() if sh.has_text_frame else ''
        if txt:
            continue
        geom = sh._element.spPr.find('{%s}prstGeom' % A)
        if geom is None or geom.get('prst') != 'parallelogram':
            continue
        targets.append(sh)

    for sh in targets:
        el = sh._element
        new = copy.deepcopy(el)

        # 唯一 id / name
        cNvPr = new.find('.//{%s}cNvPr' % P)
        maxid = max(int(e.get('id')) for e in spTree.iter('{%s}cNvPr' % P))
        cNvPr.set('id', str(maxid + 1))
        cNvPr.set('name', (cNvPr.get('name') or 'bar') + '-黑底')

        spPr = new.find('.//{%s}spPr' % A) if new.find('.//{%s}spPr' % A) is not None else new.find('{%s}spPr' % P)
        xfrm = spPr.find('{%s}xfrm' % A)
        off = xfrm.find('{%s}off' % A)
        ext = xfrm.find('{%s}ext' % A)
        x, y = int(off.get('x')), int(off.get('y'))
        cx, cy = int(ext.get('cx')), int(ext.get('cy'))
        off.set('x', str(x - MARGIN)); off.set('y', str(y - MARGIN))
        ext.set('cx', str(cx + 2 * MARGIN)); ext.set('cy', str(cy + 2 * MARGIN))

        # 黑色填充，无描边
        for tag in ('solidFill', 'gradFill', 'blipFill', 'noFill'):
            for e in spPr.findall('{%s}%s' % (A, tag)):
                spPr.remove(e)
        fill = etree.SubElement(spPr, '{%s}solidFill' % A)
        clr = etree.SubElement(fill, '{%s}srgbClr' % A)
        clr.set('val', '000000')
        for e in spPr.findall('{%s}ln' % A):
            spPr.remove(e)
        ln = etree.SubElement(spPr, '{%s}ln' % A)
        etree.SubElement(ln, '{%s}noFill' % A)

        # 去掉文字
        tx = new.find('{%s}txBody' % P)
        if tx is not None:
            new.remove(tx)

        el.addprevious(new)
        added.append((idx, cNvPr.get('name')))

prs.save(F)
print('black backdrops added:', len(added))
for a in added[:8]:
    print('  p%d %s' % a)
