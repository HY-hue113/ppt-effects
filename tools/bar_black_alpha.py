import shutil, os
from pptx import Presentation
from lxml import etree

F = r'<USER_HOME>\<OUTPUT_DIR>\效果10.pptx'
BAK = r'<USER_HOME>\<OUTPUT_DIR>\效果10-备份.pptx'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
ALPHA = '80000'   # 80% 不透明 = 20% 透出底色

if not os.path.exists(BAK):
    shutil.copyfile(F, BAK)

prs = Presentation(F)
n = 0
for slide in prs.slides:
    for sh in slide.shapes:
        if not sh.name.endswith('-黑底'):
            continue
        sppr = sh._element.spPr
        fill = sppr.find('{%s}solidFill' % A)
        if fill is None:
            continue
        clr = fill.find('{%s}srgbClr' % A)
        if clr is None:
            continue
        for old in clr.findall('{%s}alpha' % A):
            clr.remove(old)
        al = etree.SubElement(clr, '{%s}alpha' % A)
        al.set('val', ALPHA)
        n += 1

prs.save(F)
print('alpha set on', n, 'backdrops ->', ALPHA)
