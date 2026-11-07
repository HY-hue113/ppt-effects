import shutil, os
from pptx import Presentation
from lxml import etree

F = r'<USER_HOME>\<OUTPUT_DIR>\效果10.pptx'
BAK = r'<USER_HOME>\<OUTPUT_DIR>\效果10-备份.pptx'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
ADJ = '18000'   # 斜切量，18000 = 18%（默认平行四边形是 25000）

if not os.path.exists(BAK):
    shutil.copyfile(F, BAK)

prs = Presentation(F)
changed = []
for idx, slide in enumerate(prs.slides, start=1):
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
        spPr = sh._element.spPr
        geom = spPr.find('{%s}prstGeom' % A)
        if geom is None or geom.get('prst') not in ('rect', 'parallelogram', 'roundRect'):
            continue
        geom.set('prst', 'parallelogram')
        for c in list(geom):
            geom.remove(c)
        av = etree.SubElement(geom, '{%s}avLst' % A)
        gd = etree.SubElement(av, '{%s}gd' % A)
        gd.set('name', 'adj')
        gd.set('fmla', 'val ' + ADJ)
        changed.append((idx, sh.name, round(w, 1), round(h, 1)))

prs.save(F)
print('changed to parallelogram:', len(changed))
for c in changed[:10]:
    print('  p%d %s  %.1f x %.1f cm' % c)
